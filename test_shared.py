import copy
import io
import json
from pathlib import Path
import shutil
import unittest
import zipfile
from unittest.mock import patch
from main import Api
from store import Store
from shared_sync import SharedSync, fingerprint, PREFIX
from bitrix_disk import BitrixDisk
from project_tests import ProjectTests

class FakeDrive:
    """Shared remote server; each client has independent credentials/network state."""
    def __init__(self,store,server):
        self.store,self.server=store,server;self.offline=False;self.fail_upload='';self.fail_reports=False
        self.c={'portal':'example.org','folder_id':123,'secret':'test','auto':False};self.status={}
    def config(self):return copy.deepcopy(self.c)
    def _write(self,c):self.c=copy.deepcopy(c)
    def check(self):
        if self.offline:raise ValueError('Нет ответа Битрикс24.')
    def children(self,folder=None):
        self.check();return [copy.deepcopy(r) for r in self.server.values() if str(r['PARENT_ID'])==str(folder or 123)]
    def call(self,method,p):
        self.check();assert method=='disk.folder.addsubfolder',method
        id=str(len(self.server)+1000);row={'ID':id,'NAME':p['data']['NAME'],'TYPE':'folder','PARENT_ID':p['id']}
        self.server[id]=row;return row
    def upload_file(self,path,name=None,folder=None):
        self.check();name=name or Path(path).name
        if self.fail_reports and not name.startswith(PREFIX):raise ValueError('Report unavailable')
        fail=self.fail_upload;self.fail_upload=''
        if fail=='before':raise ValueError('Upload disconnected')
        assert not any(r['NAME']==name and str(r['PARENT_ID'])==str(folder or 123) for r in self.server.values())
        id=str(len(self.server)+1000)
        row={'ID':id,'TYPE':'file','PARENT_ID':folder or 123,'NAME':name,'bytes':Path(path).read_bytes()}
        self.server[id]=row
        if fail=='after':raise ValueError('Reply lost after commit')
        return row
    def download(self,id):
        self.check();row=self.server[str(id)];assert row['PARENT_ID']==123
        file=self.store.root/'Резервные копии'/('download-'+str(id)+'.zip');file.write_bytes(row['bytes'])
        return self.store.restore_preview(file)
    def test(self):self.check();return {'message':'Подключено'}
    def public(self):return {'configured':True,'folder_id':123,'auto':False,'message':'Тест'}
    def close(self,**kw):pass
    def schedule(self):pass

class SharedTests(unittest.TestCase):
    call=ProjectTests.call
    def setUp(self):
        ProjectTests.setUp(self)
        self.server={}
        self.a=self.client(self.s)
        self.other=Store(Path(self.temp.name)/'other',self.s.resources)
        self.other.settings({'operator':'Коллега','role':'Администратор','revision':0})
        self.b=self.client(self.other)
    def client(self,store):
        api=Api(store);api._bitrix=FakeDrive(store,self.server);api._sync=SharedSync(store,api._bitrix);return api
    def ok(self,api,action,p=None):
        r=api.request(action,p);self.assertTrue(r['ok'],r);return r['data']
    def connect(self,api):
        p=self.ok(api,'sync_preview');r=self.ok(api,'sync_enable',{'token':p['token'],'confirmed':True})
        self.assertIn(r['state'],['ready','reports'],r);return p
    def mutation(self,api,name='Добавленный'):
        return api.request('employee',{'name':name,'role':'Техник','revision':api._store.snapshot()['revision']})
    def connect_both(self):self.connect(self.a);self.connect(self.b)
    def test_join_roundtrip_preserves_actor_history_and_revision(self):
        self.connect_both();self.assertEqual(self.other.snapshot()['settings']['operator'],'Коллега')
        self.assertEqual(self.other.snapshot()['revision'],self.s.snapshot()['revision'])
        self.assertTrue(self.mutation(self.a)['ok']);self.ok(self.b,'sync_tick')
        self.assertEqual(fingerprint(self.s.root),fingerprint(self.other.root))
        self.assertTrue(self.mutation(self.b,'Сотрудник коллеги')['ok']);self.ok(self.a,'sync_tick')
        self.assertEqual(self.s.snapshot()['history'][0]['actor'],'Коллега')
        self.assertEqual(self.s.snapshot()['settings']['operator'],'Оператор')
        self.assertEqual(self.s.snapshot()['history'],self.other.snapshot()['history'])
    def test_offline_preflight_blocks_mutation(self):
        self.connect(self.a);before=self.s.snapshot();self.a._bitrix.offline=True
        r=self.mutation(self.a);self.assertFalse(r['ok']);self.assertEqual(before,self.s.snapshot())
    def test_upload_failure_is_local_success_then_retries_on_restart(self):
        self.connect(self.a);self.a._bitrix.fail_upload='before'
        r=self.mutation(self.a);self.assertTrue(r['ok']);self.assertEqual(r['data']['sync']['state'],'blocked')
        self.assertTrue(self.a._sync.config()['pending'])
        self.a=self.client(self.s);s=self.ok(self.a,'snapshot')
        self.assertEqual(s['sync']['state'],'ready');self.assertFalse(self.a._sync.config().get('pending'))
    def test_lost_upload_response_does_not_duplicate(self):
        self.connect(self.a);self.a._bitrix.fail_upload='after';r=self.mutation(self.a);self.assertTrue(r['ok'])
        count=sum(r['NAME'].startswith(PREFIX) for r in self.server.values())
        self.ok(self.a,'sync_tick');self.assertEqual(count,sum(r['NAME'].startswith(PREFIX) for r in self.server.values()))
        self.assertEqual(self.a._sync.state,'ready')
    def test_stale_form_pulls_but_rejects_original_action(self):
        self.connect_both();self.assertTrue(self.mutation(self.a)['ok'])
        r=self.mutation(self.b,'Не должен добавиться');self.assertFalse(r['ok']);self.assertIn('заново',r['error'])
        self.assertFalse(any(e['name']=='Не должен добавиться' for e in self.other.snapshot()['employees']))
        self.assertEqual(fingerprint(self.s.root),fingerprint(self.other.root))
    def test_simultaneous_fork_never_overwrites(self):
        self.connect_both()
        old,nodes=self.a._sync.remote()
        for api,name in [(self.a,'Один'),(self.b,'Два')]:
            api._sync.before_change();api._store.employee({'name':name,'role':'x','revision':api._store.snapshot()['revision']})
            api._sync._prepare(api._sync.config())
        # Model two serverside uploads that raced after the same successful preflight.
        for api in [self.a,self.b]:
            p=api._sync.config()['pending'];api._bitrix.upload_file(p['file'],p['name'])
        for api in [self.a,self.b]:
            r=self.ok(api,'sync_tick');self.assertEqual(r['state'],'blocked');self.assertIn('Конфликт',r['message'])
            self.assertFalse(self.mutation(api,'Три')['ok'])
        self.assertEqual(sum(r['NAME'].startswith(PREFIX) for r in self.server.values()),3)
    def test_dirty_local_and_remote_preserved(self):
        self.connect_both();self.other.employee({'name':'Местный','role':'x','revision':self.other.snapshot()['revision']})
        self.mutation(self.a);before=self.other.snapshot()
        self.assertEqual(self.ok(self.b,'sync_tick')['state'],'blocked');self.assertEqual(before,self.other.snapshot())
    def test_deletion_of_parent_blocks(self):
        self.connect(self.a);self.mutation(self.a)
        key=next(k for k,r in self.server.items() if r['NAME'].startswith(PREFIX));del self.server[key]
        self.assertFalse(self.mutation(self.a,'x')['ok'])
    def test_join_preview_invalidated_by_new_remote(self):
        self.connect(self.a);p=self.ok(self.b,'sync_preview');self.mutation(self.a)
        before=self.other.snapshot();r=self.b.request('sync_enable',{'token':p['token'],'confirmed':True})
        self.assertFalse(r['ok']);self.assertEqual(before,self.other.snapshot())
    def test_join_requires_confirmation(self):
        p=self.ok(self.a,'sync_preview');self.assertFalse(self.a.request('sync_enable',{'token':p['token']})['ok'])
        self.assertFalse(self.server)
    def test_backup_contains_no_connection_secrets(self):
        self.connect(self.a)
        record=next(r for r in self.server.values() if r['NAME'].startswith(PREFIX))
        with zipfile.ZipFile(io.BytesIO(record['bytes'])) as z:
            self.assertFalse(any('Подключение' in n or 'Общий_учет' in n for n in z.namelist()))
    def test_reports_failure_does_not_claim_db_failed(self):
        self.a._bitrix.fail_reports=True;self.connect(self.a)
        self.assertEqual(self.a._sync.state,'reports')
        self.a._bitrix.fail_reports=False;self.ok(self.a,'sync_tick');self.assertEqual(self.a._sync.state,'ready')
    def test_receive_act_and_report_reach_second_pc(self):
        self.connect_both()
        r=self.ok(self.a,'receive',{'revision':self.s.snapshot()['revision'],'mode':'Новая поставка','document':'Счет',
            'items':[{'type':'Ноутбук','name':'Новый','serial':'NEW-SN','kit':'БП','condition':'Новое','quantity':1,'purchase_price':''}]})
        self.ok(self.b,'sync_tick');self.assertEqual(fingerprint(self.s.root),fingerprint(self.other.root))
        act=self.other.snapshot()['acts'][0]
        self.assertTrue((self.other.root/act['file']).exists())
        self.assertTrue(any(r['NAME'].endswith('.docx') for r in self.server.values()))
        self.assertTrue(any(r['NAME'].endswith('.xlsx') for r in self.server.values()))
    def test_manual_restore_and_rebind_blocked_while_shared(self):
        self.connect(self.a)
        self.assertFalse(self.a.request('restore',{})['ok']);self.assertFalse(self.a.request('bitrix_save',{})['ok'])
    def test_external_doc_change_is_published(self):
        self.connect_both();path=self.s.root/'Акты'/'external.docx';path.write_bytes(b'example')
        self.ok(self.a,'sync_tick');self.ok(self.b,'sync_tick')
        self.assertEqual((self.other.root/'Акты'/'external.docx').read_bytes(),b'example')
    def test_pending_cannot_be_disconnected(self):
        self.connect(self.a);self.a._bitrix.fail_upload='before';self.mutation(self.a)
        self.assertFalse(self.a.request('sync_disable',{})['ok'])
    def test_corrupt_remote_not_installed(self):
        self.connect(self.a);row=next(r for r in self.server.values() if r['NAME'].startswith(PREFIX));row['bytes']=b'invalid'
        before=self.other.snapshot();self.assertFalse(self.b.request('sync_preview')['ok']);self.assertEqual(before,self.other.snapshot())
    def test_downloaded_copy_restarts_without_restore_journal_loop(self):
        self.connect_both();before=self.other.snapshot();self.b=self.client(self.other)
        self.ok(self.b,'snapshot');self.ok(self.b,'sync_tick');self.assertEqual(before,self.other.snapshot())
    def test_api_pagination_includes_later_pages(self):
        disk=BitrixDisk(self.s);calls=[]
        def page(method,p):
            calls.append(p['start']);return {'result':[{'ID':p['start']+1}],**({'next':50} if p['start']==0 else {})}
        disk.call_page=page
        self.assertEqual(len(disk.children(123)),2);self.assertEqual(calls,[0,50])
    def test_ambiguous_bootstrap_two_databases_blocked(self):
        for api in [self.a,self.b]:
            c={'enabled':True,'portal':'example.org','folder_id':123,'device':'f'*32,'dataset':('a' if api==self.a else 'b')*32,'head':'0','fingerprint':fingerprint(api._store.root)}
            api._sync.save(c);api._sync._prepare(c);p=c['pending'];api._bitrix.upload_file(p['file'],p['name'])
        self.assertEqual(self.ok(self.a,'sync_tick')['state'],'blocked')

    def test_interrupted_install_completes_offline_from_cache(self):
        self.connect_both()
        self.ok(self.a,'receive',{'revision':self.s.snapshot()['revision'],'mode':'Новая поставка','document':'Счет',
            'items':[{'type':'Ноутбук','name':'Новый','serial':'RECOVERY-SN','kit':'БП','condition':'Новое','quantity':1}]})
        import os
        real_replace=os.replace
        def crash(src,dst):
            real_replace(src,dst)
            if Path(src).parent.name=='new' and Path(src).name=='Учет.sqlite':raise KeyboardInterrupt('Simulated power loss after database replacement')
        with patch('store.os.replace',side_effect=crash):
            with self.assertRaises(KeyboardInterrupt):self.b._sync.tick()
        self.assertTrue(self.b._sync.config().get('pull'))
        self.b=self.client(Store(self.other.root,self.other.resources));self.b._bitrix.offline=True
        self.ok(self.b,'snapshot')
        self.assertEqual(fingerprint(self.s.root),fingerprint(self.other.root))
        self.assertEqual(self.other.snapshot()['settings']['operator'],'Коллега')
        self.assertEqual(self.b._sync.state,'blocked')

def run(resources,test_root):
    import project_tests
    from test_prices import PriceTests
    project_tests.RESOURCES=Path(resources);project_tests.TEST_ROOT=Path(test_root)/'regression';project_tests.TEST_ROOT.mkdir(parents=True,exist_ok=True)
    import test_prices
    test_prices.ROOT=Path(resources)
    suite=unittest.TestSuite([unittest.defaultTestLoader.loadTestsFromTestCase(PriceTests),unittest.defaultTestLoader.loadTestsFromTestCase(SharedTests)])
    import test_admin
    suite.addTests(test_admin.suite())
    stream=io.StringIO();result=unittest.TextTestRunner(stream=stream,verbosity=2).run(suite)
    (Path(test_root)/'regression.json').write_text(json.dumps({'run':result.testsRun,'success':result.wasSuccessful(),'details':stream.getvalue()},ensure_ascii=False,indent=2),encoding='utf8')
    print(stream.getvalue());return 0 if result.wasSuccessful() else 1

if __name__=='__main__':raise SystemExit(run(Path(__file__).parent,Path(__file__).parent/'test'))
