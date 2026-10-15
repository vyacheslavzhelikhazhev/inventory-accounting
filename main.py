import json
import logging
import os
from pathlib import Path
import sys
import tempfile
import threading
import traceback
import uuid
from store import Store
from bitrix_disk import BitrixDisk
from shared_sync import SharedSync

RESOURCES=Path(getattr(sys,'_MEIPASS',Path(__file__).parent))
ROOT=Path(os.environ.get('MATUCHET_DATA', str(Path(sys.executable).parent/'Данные' if getattr(sys,'frozen',False) else Path(__file__).parent/'Данные')))

class Api:
    def __init__(self,store):
        self._store=store;self._window=None;self._import_file=None;self._import_plan=None;self._restore_plan=None
        self._bitrix=BitrixDisk(store);self._sync=SharedSync(store,self._bitrix)
        self._request_lock=threading.RLock();self._started=False;self._admin=False
    def _preferences(self,p=None):
        path=Path(os.environ.get('MATUCHET_PREFERENCES',str(Path(os.environ.get('LOCALAPPDATA',tempfile.gettempdir()))/'Матучет'/'appearance.json')))
        if p is not None:
            if p.get('theme') not in ['light','dark']:raise ValueError('Выберите светлую или темную тему.')
            path.parent.mkdir(parents=True,exist_ok=True)
            pending=path.with_suffix('.'+uuid.uuid4().hex+'.tmp')
            try:
                pending.write_text(json.dumps({'theme':p['theme']}),encoding='utf-8')
                os.replace(pending,path)
            finally:pending.unlink(missing_ok=True)
        try:
            value=json.loads(path.read_text(encoding='utf-8')).get('theme')
        except (OSError,ValueError,AttributeError):value='light'
        return {'theme':value if value in ['light','dark'] else 'light'}
    def request(self,action,p=None):
        with self._request_lock:
            return self._request(action,p)

    def _request(self,action,p=None):
        p=p or {}
        try:
            if action=='snapshot':
                if not self._started:self._started=True;self._sync.tick()
                result={**self._store.snapshot(),'bitrix':self._bitrix.public(),'sync':self._sync.public(),'preferences':self._preferences(),'admin_mode':self._admin}
            elif action=='admin_mode':
                self._admin=p.get('enabled') is True;result={'enabled':self._admin}
            elif action=='preferences':result=self._preferences(p)
            elif action=='sync_preview':result=self._sync.preview()
            elif action=='sync_enable':result=self._sync.enable(p)
            elif action=='sync_tick':result=self._sync.tick()
            elif action=='sync_disable':result=self._sync.disable()
            elif action in ['settings','employee','operation','receive','correct_asset','responsible','asset_status','signed','dictionary','import_commit','movement','repair','admin_asset','admin_employee','admin_dictionary']:
                if action.startswith('admin_') and not self._admin:raise ValueError('Включите режим редактирования в настройках.')
                self._sync.before_change()
                if action=='import_commit':
                    if not self._import_file: raise ValueError('Сначала выберите выгрузку.')
                    p={**p,'file':self._import_file,'preview':self._import_plan}
                result=getattr(self._store,action)(p)
                try:
                    if self._sync.config().get('enabled'):result['sync']=self._sync.after_change()
                    else:self._bitrix.schedule()
                except Exception:self._bitrix.status={'state':'error','message':'Изменения сохранены локально. Проверьте настройки Битрикс24.'}
            elif action=='restore_preview':
                import webview
                files=self._window.create_file_dialog(webview.FileDialog.OPEN,allow_multiple=False,file_types=('Полный архив Матучета (*.zip)',))
                if not files:return {'ok':True,'data':None}
                self._restore_plan=self._store.restore_preview(files[0]);result=self._restore_plan
            elif action=='restore':
                if self._sync.config().get('enabled'):raise ValueError('В общем учете ручное восстановление отключено. Сначала согласуйте восстановление с коллегами.')
                if not self._restore_plan:raise ValueError('Сначала выберите и проверьте архив.')
                self._bitrix.close()
                result=self._store.restore({**p,**self._restore_plan,'revision':p.get('revision')});self._restore_plan=None
            elif action=='bitrix_save':
                if self._sync.config().get('enabled'):raise ValueError('Подключение используется общим учетом. Сначала отключите общий учет в настройках.')
                result=self._bitrix.save(p)
            elif action=='bitrix_test':result=self._bitrix.test()
            elif action=='bitrix_status':result=self._bitrix.public()
            elif action=='bitrix_upload':result=self._bitrix.upload()
            elif action=='bitrix_download':
                self._restore_plan=self._bitrix.download(p.get('file_id'));result=self._restore_plan
            elif action in ['backup','export']:result=getattr(self._store,action)()
            elif action=='open':
                path=(self._store.root/p.get('file','')).resolve()
                if not path.is_relative_to(self._store.root.resolve()) or not path.exists():raise ValueError('Файл не найден.')
                if path.is_file() and path.suffix.lower() not in ['.docx','.xlsx','.sqlite']:raise ValueError('Неподдерживаемый файл.')
                os.startfile(str(path));result={}
            elif action=='import_preview':
                import webview
                files=self._window.create_file_dialog(webview.FileDialog.OPEN,allow_multiple=False,file_types=('Выгрузка Excel (*.xlsx)',))
                if not files: return {'ok':True,'data':None}
                self._import_file=str(files[0]);self._import_plan=self._store.import_preview(self._import_file);result=self._import_plan
            else:raise ValueError('Действие недоступно.')
            return {'ok':True,'data':result}
        except ValueError as e: return {'ok':False,'error':str(e)}
        except Exception:
            logging.exception('Action %s failed',action)
            return {'ok':False,'error':'Не удалось завершить действие. Проверьте состояние учета. Подробности записаны в журнал ошибок.'}

def main():
    ROOT.mkdir(parents=True,exist_ok=True)
    logging.basicConfig(filename=ROOT/'Ошибки.log',level=logging.ERROR,encoding='utf-8',format='%(asctime)s %(message)s')
    import msvcrt
    instance=open(ROOT/'Сеанс.lock','a+b')
    try:
        instance.seek(0);instance.write(b'0');instance.flush();instance.seek(0)
        msvcrt.locking(instance.fileno(),msvcrt.LK_NBLCK,1)
    except OSError:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None,'Приложение уже открыто на этом компьютере.','Матучет',0)
        instance.close();return
    try:
        from runtime_bootstrap import prepare
        prepare()
        import webview
        api=Api(Store(ROOT,RESOURCES))
        window=webview.create_window('Матучет',html=(RESOURCES/'ui/index.html').read_text(encoding='utf-8'),js_api=api,
                    width=1180,height=790,min_size=(900,640),background_color='#FFFFFF',hidden=('--smoke' in sys.argv or '--ui-smoke' in sys.argv))
        api._window=window
        smoke='--smoke' in sys.argv or '--ui-smoke' in sys.argv
        def check():
            window.events.loaded.wait(25)
            try:
                import time
                for attempt in range(100):
                    result=window.evaluate_js('JSON.stringify({title:document.title,ready:!!window.pywebview,nav:document.querySelectorAll(".nav").length,error:document.querySelector(".globalerror")?.textContent})')
                    if json.loads(result).get('nav')==6:break
                    time.sleep(.1)
                if '--ui-smoke' in sys.argv:
                    checks=window.evaluate_js('''JSON.stringify((()=>{const checks={};view='settings';render();checks.settings=!!document.querySelector('#bitrixform');show('quality');checks.quality=!!document.querySelector('#dialogtitle');const a=data.assets.find(x=>x.status==='В эксплуатации'&&x.target_id)||data.assets[0];show('asset',{id:a.id});checks.movementButton=!!document.querySelector('[data-action="move"]');show('movement',{id:a.id,operation:'Передача'});checks.transfer=!!document.querySelector('#movementform');show('movement',{id:a.id,operation:'Перемещение'});checks.move=!!document.querySelector('#movementform');show('repair',{id:a.id,phase:'start'});checks.repair=!!document.querySelector('#repairform');close();return checks})())''')
                    window.evaluate_js("view='settings';render();document.querySelector('[data-action=adminprompt]').click();document.querySelector('[data-action=adminenable]').click();")
                    for attempt in range(50):
                        if window.evaluate_js('!!data.admin_mode'):break
                        time.sleep(.1)
                    adminchecks=window.evaluate_js('''JSON.stringify((()=>{const checks={adminEnabled:data.admin_mode===true,themeToggle:!!document.querySelector('.theme-toggle'),syncControls:!!document.querySelector('[data-action=syncpreview]')};const a=data.assets.find(x=>x.status==='В эксплуатации'&&x.target_id);show('adminasset',{id:a.id});checks.adminFields=['type','name','serial','status','recipient','responsible_id','location','kit','condition','purpose','purchase_date','supply_doc','notes','reason','purchase_price'].every(n=>!!document.querySelector('#adminassetform [name="'+n+'"]'));const dialog=document.querySelector('.dialog'),r=dialog.getBoundingClientRect();checks.buttonsFit=[...dialog.querySelectorAll(':scope>.foot button')].every(b=>{const x=b.getBoundingClientRect();return x.left>=r.left&&x.right<=r.right});show('adminemployee',{id:data.employees[0].id});checks.adminEmployee=!!document.querySelector('#adminemployeeform [name=status]');show('admindict',{key:'types',old:data.types[0]});checks.adminDictionary=!!document.querySelector('#admindictform');close();return checks})())''')
                    result=json.dumps({'startup':json.loads(result),'ui':{**json.loads(checks),**json.loads(adminchecks)}},ensure_ascii=False)
                (ROOT/'smoke.json').write_text(str(result),encoding='utf-8')
            finally:window.destroy()
        webview.start(check if smoke else None,gui='edgechromium',debug=False,private_mode=True,
                      storage_path=str(Path(tempfile.gettempdir())/('MatUchetWebView_'+uuid.uuid4().hex)))
    except Exception as e:
        logging.exception('Startup failed')
        if '--smoke' in sys.argv or '--ui-smoke' in sys.argv:return 1
        import ctypes
        detail=('Не найден компонент программы: '+str(e.name)+'. Распакуйте полный архив целиком.') if isinstance(e,ModuleNotFoundError) else str(e)
        ctypes.windll.user32.MessageBoxW(None,'Не удалось открыть Матучет.\n'+detail+'\nЖурнал: '+str(ROOT/'Ошибки.log'),'Матучет',0)
    finally:
        if 'api' in locals():api._bitrix.close(flush=True)
        instance.seek(0);msvcrt.locking(instance.fileno(),msvcrt.LK_UNLCK,1);instance.close()

if __name__=='__main__':
    if '--runtime-test' in sys.argv:
        from runtime_bootstrap import check
        sys.exit(check(RESOURCES,ROOT))
    if '--self-test' in sys.argv:
        import test_shared
        sys.exit(test_shared.run(RESOURCES,ROOT))
    sys.exit(main())
