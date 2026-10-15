"""Sequential Bitrix Drive collaboration using immutable, causally linked snapshots.

Drive is not a transactional database. Concurrent writers produce branches, never
overwrite a snapshot. Branches stop further edits; no automatic merge is attempted.
Local outbox is persisted before upload, so ambiguous HTTP replies can be retried.
"""
import contextlib
import hashlib
import json
import os
from pathlib import Path
import re
import sqlite3
import tempfile
import uuid
import zipfile
from store import require, now

PREFIX='Матучет_общий_'
NAME=re.compile(r'^Матучет_общий_([a-f0-9]{32})_([a-f0-9]{32})_(0|[a-f0-9]{32})\.zip$')

def write_json(path,value):
    pending=path.with_suffix('.tmp')
    pending.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf8')
    os.replace(pending,path)

def fingerprint(root):
    """Logical DB content + document bytes; operator identity stays per computer."""
    root=Path(root);result={}
    with contextlib.closing(sqlite3.connect(root/'Учет.sqlite')) as c:
        for table,key in [('employees','id'),('assets','id'),('meta','key'),('history','id'),('acts','id')]:
            rows=[]
            for row in c.execute(f'SELECT * FROM {table} ORDER BY {key}'):
                row=list(row)
                if table=='meta' and row[0]=='settings':
                    settings=json.loads(row[1])
                    for name in ['operator','role']:settings.pop(name,None)
                    row[1]=json.dumps(settings,sort_keys=True,ensure_ascii=False)
                rows.append(row)
            result[table]=rows
    paths=list((root/'Акты').rglob('*.docx'))
    paths += [root/n for n in ['Шаблон_акта.docx','Шаблон_приема_передачи.docx'] if (root/n).exists()]
    result['files']={p.relative_to(root).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(paths)}
    return hashlib.sha256(json.dumps(result,ensure_ascii=False,sort_keys=True).encode('utf8')).hexdigest()

class SharedSync:
    def __init__(self,store,disk):
        self.store,self.disk=store,disk
        self.path=store.root/'Общий_учет.json'
        self.cache=store.root/'Обмен';self.cache.mkdir(exist_ok=True)
        self.state='checking' if self.config().get('enabled') else 'off'
        self.message='Проверяем общую базу…' if self.state=='checking' else 'Общий учет выключен'
        self.plan=None

    def config(self):
        if not self.path.exists():return {}
        try:return json.loads(self.path.read_text(encoding='utf8'))
        except Exception:raise ValueError('Повреждены настройки общего учета. Сохраните папку Данные и обратитесь за помощью.') from None

    def save(self,c):write_json(self.path,c)

    def public(self):
        c=self.config()
        return {'enabled':bool(c.get('enabled')),'state':self.state,'message':self.message,
                'last_sync':c.get('last_sync'),'pending':bool(c.get('pending')),
                'head':c.get('head',''),'folder_id':c.get('folder_id')}

    def _bound(self,c):
        d=self.disk.config()
        require(c.get('portal')==d.get('portal') and str(c.get('folder_id'))==str(d.get('folder_id')),
                'Подключение указывает на другую папку. Верните прежний портал и ID папки.')

    def remote(self):
        nodes={};datasets=set()
        for row in self.disk.children():
            name=str(row.get('NAME',''))
            if not name.startswith(PREFIX):continue
            m=NAME.fullmatch(name)
            require(m and row.get('TYPE')=='file','Неизвестный файл общего учета. Не меняйте имена служебных архивов.')
            dataset,head,parent=m.groups()
            require(head not in nodes,'На Диске повторяется версия общей базы. Требуется проверка копий.')
            nodes[head]={'dataset':dataset,'head':head,'parent':parent,'file_id':str(row['ID']),'name':name}
            datasets.add(dataset)
        require(len(datasets)<=1,'В папке обнаружены разные общие базы. Изменения остановлены; используйте отдельную папку для каждой базы.')
        for head,node in nodes.items():
            require(node['parent']=='0' or node['parent'] in nodes,'На Диске отсутствует предыдущая версия. Восстановите удаленные архивы.')
            seen=set();current=head
            while current!='0':
                require(current not in seen,'Повреждена последовательность версий общей базы.')
                seen.add(current);current=nodes[current]['parent']
        parents={n['parent'] for n in nodes.values()}
        leaves=[n for h,n in nodes.items() if h not in parents]
        require(len(leaves)<=1,'Конфликт: два компьютера изменили одну версию. Обе копии сохранены на Диске; новые операции остановлены. Нужна сверка данных.')
        return (leaves[0] if leaves else None),nodes

    def _read(self,node):
        path=self.cache/(node['head']+'.zip')
        if not path.exists():
            plan=self.disk.download(node['file_id'])
            os.replace(plan['file'],path)
        plan=self.store.restore_preview(path)
        with zipfile.ZipFile(path) as z:
            manifest=json.loads(z.read('manifest.json'));meta=manifest.get('shared',{})
            require(all(meta.get(k)==node[k] for k in ['dataset','head','parent']), 'Архив не соответствует версии на Диске.')
            with tempfile.TemporaryDirectory(dir=self.cache) as tmp:
                root=Path(tmp)
                for name in manifest['files']:
                    dest=root/name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(z.read(name))
                require(fingerprint(root)==meta.get('fingerprint'),'Состав общей базы не прошел проверку.')
        return plan,meta

    def preview(self):
        require(not self.config().get('enabled'),'Общий учет уже включен.')
        self.disk.test();node,_=self.remote()
        local=fingerprint(self.store.root)
        plan={'mode':'join' if node else 'publish','token':uuid.uuid4().hex,'local':local,'node':node}
        if node:
            archive,meta=self._read(node)
            plan.update(archive=archive,meta=meta)
        self.plan=plan
        return {'token':plan['token'],'mode':plan['mode'],
                'revision':plan.get('archive',{}).get('revision',self.store.snapshot()['revision']),
                'documents':plan.get('archive',{}).get('documents',len(self.store.snapshot()['acts']))}

    def enable(self,p):
        plan=self.plan
        require(plan and p.get('token')==plan['token'],'Проверьте общую базу заново.')
        require(p.get('confirmed') is True,'Подтвердите подключение общей базы.')
        require(fingerprint(self.store.root)==plan['local'],'Местная база изменилась. Проверьте подключение заново.')
        node,_=self.remote()
        require(node==plan['node'],'Общая база изменилась. Проверьте подключение заново.')
        d=self.disk.config()
        c={'enabled':True,'portal':d['portal'],'folder_id':d['folder_id'],'device':uuid.uuid4().hex,
           'dataset':node['dataset'] if node else uuid.uuid4().hex,'head':'0','fingerprint':plan['local']}
        # Save the durable join intent before replacing local files.
        if node:c['pull']={'node':node,'plan':plan['archive']}
        self.disk.close();d['auto']=False;self.disk._write(d)
        self.save(c);self.plan=None
        if node:self._pull(c,node)
        else:self._prepare(c)
        return self.tick()

    def _prepare(self,c):
        """Persist a complete immutable outbox before making any remote writes."""
        archive=Path(self.store.backup()['file'])
        head=uuid.uuid4().hex
        meta={'dataset':c['dataset'],'head':head,'parent':c['head'],'fingerprint':fingerprint(self.store.root),
              'device':c['device'],'created':now()}
        name=PREFIX+c['dataset']+'_'+head+'_'+c['head']+'.zip'
        dest=self.cache/name;tmp=dest.with_suffix('.tmp')
        with zipfile.ZipFile(archive) as src,zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as out:
            manifest=json.loads(src.read('manifest.json'));manifest['shared']=meta
            for name_in in manifest['files']:out.writestr(name_in,src.read(name_in))
            out.writestr('manifest.json',json.dumps(manifest,ensure_ascii=False))
        self.store.restore_preview(tmp);os.replace(tmp,dest)
        c['pending']={'file':str(dest),'name':name,'meta':meta,'sha256':hashlib.sha256(dest.read_bytes()).hexdigest()}
        self.save(c)

    def _pull(self,c,node):
        plan,meta=self._read(node)
        intent=c.get('pull',{})
        if intent.get('node')!=node or not intent.get('safety'):
            intent={'node':node,'plan':plan,'safety':self.store.backup()['file'],
                    'settings':self.store.snapshot()['settings']}
        c['pull']=intent;self.save(c)
        self.store.restore({**plan,'revision':self.store.snapshot()['revision']},_shared=True,
                           _safety=intent['safety'],_local_settings=intent['settings'])
        require(fingerprint(self.store.root)==meta['fingerprint'],'Полученная база не прошла проверку после установки.')
        c.update(head=node['head'],fingerprint=meta['fingerprint'],last_sync=now())
        c.pop('pull',None);c.pop('pending',None);c.pop('intent',None);self.save(c)

    def _push(self,c,node,nodes):
        pending=c['pending'];meta=pending['meta']
        require(fingerprint(self.store.root)==meta['fingerprint'],'Локальные данные изменены после подготовки отправки. Обе копии сохранены; требуется сверка.')
        require(hashlib.sha256(Path(pending['file']).read_bytes()).hexdigest()==pending['sha256'],'Повреждена подготовленная копия. Операции остановлены.')
        # If the server committed an upload but its response was lost, reuse it.
        if meta['head'] not in nodes:
            require((node['head'] if node else '0')==meta['parent'],'В Битриксе уже есть изменения коллег, а на этом ПК — неотправленные изменения. Обе копии сохранены; требуется сверка.')
            self.disk.upload_file(pending['file'],pending['name'])
        latest,nodes=self.remote()
        require(meta['head'] in nodes,'Загрузка пока не подтверждена Битриксом. Повторите синхронизацию.')
        _,remote_meta=self._read(nodes[meta['head']])
        require(remote_meta==meta,'Загруженная версия не прошла обратную проверку.')
        c.update(head=meta['head'],fingerprint=meta['fingerprint'],last_sync=now(),projection=meta['head'])
        c.pop('pending',None);c.pop('intent',None);self.save(c)
        if latest['head']!=c['head']:self._pull(c,latest)

    def _projection(self,c):
        """Readable DOCX and Excel are derived copies, never used to restore the DB."""
        if not c.get('projection'):return
        node,_=self.remote()
        if not node or node['head']!=c['head']:
            c.pop('projection',None);self.save(c);return
        children=self.disk.children()
        def folder(name):
            matches=[r for r in children if r.get('NAME')==name and r.get('TYPE')=='folder']
            require(len(matches)<=1,'Повтор папки '+name)
            if matches:return matches[0]['ID']
            return self.disk.call('disk.folder.addsubfolder',{'id':c['folder_id'],'data':{'NAME':name}})['ID']
        acts_folder=folder('Акты');existing={r['NAME'] for r in self.disk.children(acts_folder)}
        for act in self.store.snapshot()['acts']:
            path=(self.store.root/act['file']).resolve()
            require(path.is_relative_to(self.store.root.resolve()) and path.is_file(),'Не найден файл акта для отправки.')
            digest=hashlib.sha256(path.read_bytes()).hexdigest()[:16]
            name=path.stem[:100]+'_'+digest+'.docx'
            if name not in existing:self.disk.upload_file(path,name,acts_folder);existing.add(name)
        reports=folder('Отчеты');names={r['NAME'] for r in self.disk.children(reports)}
        # Immutable per-version reports cannot overwrite a newer colleague's report.
        name=f"Матучет_версия_{self.store.snapshot()['revision']}_{c['head'][:12]}.xlsx"
        if name not in names:self.disk.upload_file(self.store.export()['file'],name,reports)
        c.pop('projection',None);self.save(c)

    def tick(self):
        c=self.config()
        if not c.get('enabled'):return {'changed':False,**self.public()}
        before=c.get('head')
        try:
            self._bound(c)
            if c.get('pull'):
                # Finish an interrupted local installation from its verified cache,
                # even when the network is currently unavailable.
                self._pull(c,c['pull']['node'])
            node,nodes=self.remote()
            require(not node or node['dataset']==c['dataset'],'Выбрана другая общая база. Изменения остановлены.')
            if c.get('pending'):
                self._push(c,node,nodes)
                node,nodes=self.remote()
            local=fingerprint(self.store.root)
            if local!=c['fingerprint'] or c['head']=='0':
                require((node['head'] if node else '0')==c['head'],'Есть местные изменения и новая версия коллег. Ничего не заменено; требуется сверка копий.')
                self._prepare(c);self._push(c,node,nodes);node,nodes=self.remote()
            require(node is not None and c['head'] in nodes,'Общая версия исчезла с Диска. Восстановите удаленные архивы.')
            if node['head']!=c['head']:self._pull(c,node)
            if c.pop('intent',None) is not None:self.save(c)
            self.state='ready';self.message='Сохранено в Битрикс24'
            try:self._projection(c)
            except Exception:
                self.state='reports';self.message='База и акты в архиве сохранены. Отдельные Word / Excel ожидают отправки.'
        except Exception as e:
            self.state='blocked'
            self.message=str(e) if isinstance(e,ValueError) else 'Обмен не завершен. Местные данные сохранены. Проверьте сеть и повторите.'
        return {'changed':before!=self.config().get('head'),**self.public()}

    def before_change(self):
        if not self.config().get('enabled'):return
        result=self.tick()
        require(self.state in ['ready','reports'],self.message)
        require(not result['changed'],'Получена новая версия из Битрикса. Закройте окно операции и откройте его заново.')
        c=self.config();c['intent']=now();self.save(c)

    def after_change(self):
        # Local transaction is already committed; errors must never imply rollback.
        return self.tick()

    def disable(self):
        c=self.config()
        require(not c.get('pending') and not c.get('pull') and fingerprint(self.store.root)==c.get('fingerprint'),
                'Сначала отправьте местные изменения или сохраните копии для сверки.')
        c['enabled']=False;self.save(c);self.state='off';self.message='Общий учет выключен'
        return self.public()
