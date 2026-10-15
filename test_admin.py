import unittest,copy,shutil,hashlib
from pathlib import Path
from test_prices import PriceTests
from test_shared import SharedTests
from main import Api
from shared_sync import fingerprint

class AdminTests(PriceTests):
    def test_price_preserved_and_editable_without_act(self):
        self.set_price(3,'11799,50');r=self.issue((1,));old=hashlib.sha256((self.s.root/r['file']).read_bytes()).hexdigest();acts=copy.deepcopy(self.s.snapshot()['acts'])
        self.call('admin_asset',id=3,name='Исправлено',reason='Правка');self.assertEqual(self.asset(3)['purchase_price'],'11799.50')
        self.call('admin_asset',id=3,purchase_price='12000,25',reason='Счет');self.assertEqual(self.asset(3)['purchase_price'],'12000.25')
        self.assertEqual(self.s.snapshot()['acts'],acts);self.assertEqual(hashlib.sha256((self.s.root/r['file']).read_bytes()).hexdigest(),old)
    def test_assignment_status_and_all_fields_without_act(self):
        self.call('admin_asset',id=3,type='Ноутбук',name='Правка',serial='NEW-SN',target_kind='Объект',target_id='OBJ-001',responsible_id=2,status='Поломка',location='Кабинет 10',kit='',condition='',notes='Заметка',purpose='Назначение',purchase_date='2026-10-15',supply_doc='Счет',reason='Уточнение')
        a=self.asset(3);self.assertEqual(a['owner'],'Кабинет');self.assertEqual(a['responsible_id'],2);self.assertEqual(a['purchase_date'],'2026-10-15');self.assertIn('Кабинет 10',self.s.snapshot()['locations']);self.assertFalse(self.s.snapshot()['acts'])
    def test_create_and_clear_without_act(self):
        r=self.call('admin_asset',type='Ноутбук',name='Новый',status='На складе',purchase_price='25,99',reason='Ранее не учтено');self.assertEqual(self.asset(r['id'])['purchase_price'],'25.99')
        self.call('admin_asset',id=3,status='На складе',target_kind='',target_id='',responsible_id=None,reason='Уточнение');self.assertFalse(self.asset(3)['owner']);self.assertFalse(self.s.snapshot()['acts'])
    def test_employee_and_dictionary(self):
        self.call('admin_employee',id=1,name='Исправленный',role='Начальник',status='Уволен',reason='Сверка');self.assertEqual(self.asset(3)['owner'],'Исправленный')
        self.call('admin_dictionary',key='types',old='Ноутбук',value='Компьютер',reason='Тип');self.assertEqual(self.asset(3)['type'],'Компьютер');self.assertFalse(self.s.snapshot()['acts'])
    def test_repair_edit_is_consistent(self):
        self.call('admin_asset',id=3,status='В ремонте',repair_provider='Сервис',reason='Сверка');self.assertEqual(self.s.snapshot()['repairs'][0]['state'],'Открыт')
        self.call('admin_asset',id=3,status='В эксплуатации',reason='Готово');self.assertEqual(self.s.snapshot()['repairs'][0]['state'],'Завершен');self.assertFalse(self.s.snapshot()['acts'])
    def test_invalid_data_and_stale_revision(self):
        self.atomic_reject('admin_asset',id=3,serial='SN-1',reason='Тест');self.atomic_reject('admin_asset',id=3,purchase_price='-1',reason='Тест');self.atomic_reject('admin_asset',id=3,status='На складе',reason='Тест')
        with self.assertRaises(ValueError):self.s.admin_asset({'id':3,'name':'Нет','reason':'Тест','revision':0})
    def test_gate_and_session_reset(self):
        api=Api(self.s);p={'id':3,'name':'Правка','reason':'Тест','revision':self.s.snapshot()['revision']}
        self.assertFalse(api.request('admin_asset',p)['ok']);api.request('admin_mode',{'enabled':True});self.assertTrue(api.request('admin_asset',p)['ok']);self.assertFalse(Api(self.s).request('snapshot')['data']['admin_mode'])
    def test_write_failure_rolls_back(self):
        before=self.s.snapshot();old=self.s._put
        def fail(*args):raise RuntimeError('Injected')
        self.s._put=fail
        try:
            with self.assertRaises(RuntimeError):self.call('admin_asset',id=3,name='Нет',reason='Тест')
        finally:self.s._put=old
        self.assertEqual(self.s.snapshot(),before)

class AdminSyncTests(SharedTests):
    def enable_admin(self,api):self.ok(api,'admin_mode',{'enabled':True})
    def edit(self,api,**changes):
        return api.request('admin_asset',{'id':3,'reason':'Исправление','revision':api._store.snapshot()['revision'],**changes})
    def test_admin_sync_propagates_with_price(self):
        self.connect_both();self.enable_admin(self.a);before=len(self.s.snapshot()['acts'])
        r=self.edit(self.a,name='Общая правка',purchase_price='321,45');self.assertTrue(r['ok'],r);self.ok(self.b,'sync_tick')
        a=next(a for a in self.other.snapshot()['assets'] if a['id']==3);self.assertEqual(a['purchase_price'],'321.45');self.assertEqual(a['name'],'Общая правка');self.assertEqual(len(self.s.snapshot()['acts']),before);self.assertEqual(fingerprint(self.s.root),fingerprint(self.other.root))
    def test_admin_does_not_bypass_offline_preflight(self):
        self.connect(self.a);self.enable_admin(self.a);before=self.s.snapshot();self.a._bitrix.offline=True;r=self.edit(self.a,name='Нет');self.assertFalse(r['ok']);self.assertEqual(self.s.snapshot(),before)
    def test_admin_stale_remote_form_rejected(self):
        self.connect_both();self.enable_admin(self.b);self.assertTrue(self.mutation(self.a)['ok']);r=self.edit(self.b,name='Нет');self.assertFalse(r['ok']);self.assertNotEqual(next(a for a in self.other.snapshot()['assets'] if a['id']==3)['name'],'Нет')

def suite():return unittest.TestSuite(cls(name) for cls in [AdminTests,AdminSyncTests] for name in cls.__dict__ if name.startswith('test_'))
