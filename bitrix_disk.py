"""Opt-in Bitrix24 Drive snapshots. Does not merge databases or overwrite shared files."""
import base64,json,os,re,threading,ctypes,ctypes.wintypes as wt
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request,urlopen,HTTPRedirectHandler,build_opener

class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self,*args,**kwargs):return None

def protect(value,decode=False):
    class Blob(ctypes.Structure):_fields_=[('size',wt.DWORD),('data',ctypes.POINTER(ctypes.c_ubyte))]
    raw=base64.b64decode(value) if decode else value.encode('utf8')
    buf=ctypes.create_string_buffer(raw);source=Blob(len(raw),ctypes.cast(buf,ctypes.POINTER(ctypes.c_ubyte)));dest=Blob()
    call=ctypes.windll.crypt32.CryptUnprotectData if decode else ctypes.windll.crypt32.CryptProtectData
    if not call(ctypes.byref(source),None,None,None,None,1,ctypes.byref(dest)):
        raise ValueError('Не удалось сохранить/прочитать секрет подключения Windows.')
    try:
        result=ctypes.string_at(dest.data,dest.size)
        return result.decode('utf8') if decode else base64.b64encode(result).decode('ascii')
    finally:ctypes.windll.kernel32.LocalFree(dest.data)

class BitrixDisk:
    def __init__(self,store):
        self.store=store;self.path=store.root/'Подключение_Битрикс.json';self.lock=threading.RLock();self.timer=None
        self.status={'message':'Битрикс24 не настроен','state':'off'}
        self.opener=build_opener(NoRedirect)

    def config(self):
        if not self.path.exists():return {}
        try:return json.loads(self.path.read_text(encoding='utf8'))
        except Exception:raise ValueError('Настройки Битрикс24 повреждены; настройте подключение заново.')

    def public(self):
        with self.lock:
            c=self.config();return {'configured':bool(c.get('secret')),'portal':c.get('portal',''),'folder_id':c.get('folder_id',''),'auto':c.get('auto',False),'last_file_id':c.get('last_file_id',''),**self.status}

    def save(self,p):
        with self.lock:
            c=self.config();webhook=str(p.get('webhook','')).strip()
            if webhook:
                parsed=urlparse(webhook)
                if parsed.scheme!='https' or not parsed.hostname or parsed.username or parsed.password or parsed.query or parsed.fragment or not re.fullmatch(r'/rest/\d+/[A-Za-z0-9_-]+/?',parsed.path):
                    raise ValueError('Введите полный HTTPS-адрес входящего вебхука вида https://ваш-портал/rest/1/секрет/.')
                c.update(secret=protect(webhook.rstrip('/')+'/'),portal=parsed.hostname)
            if not c.get('secret'):raise ValueError('Введите адрес входящего вебхука. Не отправляйте его в чат.')
            folder=str(p.get('folder_id','')).strip()
            if not folder.isdigit() or int(folder)<1:raise ValueError('Укажите числовой ID папки Битрикс24 Диска.')
            c.update(folder_id=int(folder),auto=p.get('auto') in [True,'true','1','on'])
            self._write(c);self.status={'message':'Настройки сохранены. Проверьте подключение.','state':'ready'}
            return self.public()

    def _write(self,c):
        temp=self.path.with_suffix('.pending');temp.write_text(json.dumps(c,ensure_ascii=False),encoding='utf8');os.replace(temp,self.path)

    def call(self,method,p):
        return self.call_page(method,p)['result']

    def call_page(self,method,p):
        c=self.config()
        if not c.get('secret'):raise ValueError('Сначала настройте Битрикс24.')
        request=Request(protect(c['secret'],True)+method+'.json',data=json.dumps(p).encode('utf8'),headers={'Content-Type':'application/json'},method='POST')
        try:
            with self.opener.open(request,timeout=45) as response:
                raw=response.read(2*1024*1024+1)
            if len(raw)>2*1024*1024:raise ValueError('Слишком большой ответ Битрикс24.')
            answer=json.loads(raw)
        except Exception:
            raise ValueError('Нет ответа Битрикс24. Проверьте сеть, адрес вебхука и права доступа.') from None
        if answer.get('error'):
            code=re.sub('[^A-Za-z0-9_.-]','',str(answer['error']))[:80]
            raise ValueError('Битрикс24 отклонил запрос: '+code)
        return answer

    def children(self,folder=None):
        folder=folder or self.config()['folder_id'];start=0;seen=set();rows={}
        while True:
            if start in seen:raise ValueError('Битрикс24 повторил страницу списка файлов.')
            seen.add(start)
            answer=self.call_page('disk.folder.getChildren',{'id':folder,'order':{'ID':'ASC'},'start':start})
            for row in answer['result']:rows[str(row['ID'])]=row
            if 'next' not in answer:break
            start=int(answer['next'])
            if len(seen)>2000:raise ValueError('Слишком много файлов в папке. Обратитесь к администратору.')
        return list(rows.values())

    def upload_file(self,path,name=None,folder=None):
        path=Path(path)
        if path.stat().st_size>20*1024*1024:raise ValueError('Файл больше 20 МБ — отправка этой версии не поддерживается.')
        name=name or path.name
        return self.call('disk.folder.uploadFile',{'id':folder or self.config()['folder_id'],'data':{'NAME':name},'fileContent':[name,base64.b64encode(path.read_bytes()).decode('ascii')],'generateUniqueName':False})

    def test(self):
        result=self.call('disk.folder.get',{'id':self.config()['folder_id']})
        with self.lock:self.status={'state':'ready','message':'Подключено. Папка: '+str(result.get('NAME',''))}
        return self.public()

    def upload(self):
        with self.lock:
            self.status={'state':'uploading','message':'Создание и загрузка полной копии…'}
            try:
                if not self.config().get('folder_id') or not self.config().get('secret'):raise ValueError('Сначала настройте Битрикс24.')
                file=Path(self.store.backup()['file'])
                if file.stat().st_size>20*1024*1024:raise ValueError('Архив больше 20 МБ. Используйте ручную загрузку или настройте передачу больших файлов.')
                result=self.call('disk.folder.uploadFile',{'id':self.config()['folder_id'],'data':{'NAME':file.name},'fileContent':[file.name,base64.b64encode(file.read_bytes()).decode('ascii')],'generateUniqueName':True})
                c=self.config();c['last_file_id']=str(result['ID']);self._write(c)
                self.status={'state':'ready','message':'Копия загружена в Битрикс24. ID '+c['last_file_id']}
                return {'file_id':c['last_file_id'],'message':self.status['message']}
            except Exception as e:
                self.status={'state':'error','message':str(e)};raise

    def download(self,file_id):
        if not str(file_id).isdigit() or int(file_id)<1:raise ValueError('Введите ID архивного файла.')
        meta=self.call('disk.file.get',{'id':int(file_id)})
        if str(meta.get('PARENT_ID'))!=str(self.config()['folder_id']):raise ValueError('Архив находится вне настроенной папки.')
        if not str(meta.get('NAME','')).startswith('Матучет_') or not str(meta.get('NAME','')).endswith('.zip'):raise ValueError('Выберите полный ZIP-архив Матучета.')
        url=meta.get('DOWNLOAD_URL','');parsed=urlparse(url)
        if parsed.scheme!='https' or parsed.hostname!=self.config()['portal'] or parsed.username or parsed.password:raise ValueError('Битрикс24 вернул неподдерживаемый адрес скачивания.')
        try:
            with self.opener.open(url,timeout=45) as response:data=response.read(20*1024*1024+1)
        except Exception:raise ValueError('Не удалось скачать архив. Проверьте доступ и сеть.') from None
        if len(data)>20*1024*1024:raise ValueError('Архив больше 20 МБ.')
        import uuid
        file=self.store.root/'Резервные копии'/('Загружено_'+uuid.uuid4().hex+'.zip');file.write_bytes(data)
        try:return self.store.restore_preview(file)
        except Exception:file.unlink(missing_ok=True);raise

    def schedule(self):
        with self.lock:
            if not self.config().get('auto'):return
            if self.timer:self.timer.cancel()
            def run():
                with self.lock:
                    self.timer=None
                    try:self.upload()
                    except Exception:pass # Status remains visible; a saved local operation must not be reported as failed.
            self.timer=threading.Timer(15,run);self.timer.daemon=True;self.timer.start()

    def close(self,flush=False):
        with self.lock:
            pending=self.timer is not None
            if self.timer:self.timer.cancel();self.timer=None
            if flush and pending and self.config().get('auto'):
                try:self.upload()
                except Exception:pass
