"""Loopback-only graphical workbench. All business results come from oa.core."""
from __future__ import annotations
import argparse, base64, copy, json, mimetypes, os, secrets, threading, time, webbrowser, zipfile
from dataclasses import asdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from .core import DEFAULT, Item, Ledger, analyze, analyze_local, discover, load_config, save_config, validate_config, safe_name

VERSION='1.0.0'
def data_root():
    return Path(os.environ.get('OA_DATA',str(Path(os.environ.get('LOCALAPPDATA',Path.home()))/'OAArchive'/'workspace'))).resolve()

class Workbench:
    def __init__(self):
        self.base=data_root(); self.base.mkdir(parents=True,exist_ok=True)
        self.output=self.base/'归档库'
        settings=self.base/'workspace-settings.json'
        if settings.exists():
            try:self.output=Path(json.loads(settings.read_text(encoding='utf-8'))['output']).resolve()
            except (ValueError,KeyError,TypeError):pass
        self.items=[];self.busy=False;self.phase='等待导入';self.done=0
        self.events=[];self.lock=threading.RLock();self.token=secrets.token_urlsafe(32)
        self.batch_dir=self.base/'待处理'/time.strftime('%Y%m%d-%H%M%S')
    def event(self,message):
        self.events.append({'time':time.strftime('%H:%M:%S'),'message':message})
    def state(self):
        with self.lock:
            return dict(version=VERSION,busy=self.busy,phase=self.phase,done=self.done,total=len(self.items),
                items=[{**asdict(x),'source':Path(x.source).name,'text':x.text[:18000]} for x in self.items],
                output=str(self.output),events=self.events[-100:])
    def start(self,mode='local'):
        with self.lock:
            if self.busy:raise ValueError('任务正在运行')
            if not self.items:raise ValueError('请先导入文件')
            c=load_config();validate_config(c)
            if mode not in ('local','model'):raise ValueError('识别模式无效')
            if mode=='model' and not c.get('api_key'):raise ValueError('请先设置 API Key')
            self.busy=True;self.done=0;self.phase='识别中';self.event('开始本地提取与'+('模型分类' if mode=='model' else '规则分类'))
        def worker():
            try:
                for i,x in enumerate(list(self.items)):
                    result=(analyze if mode=='model' else analyze_local)(x.source,c,self.output)
                    with self.lock:
                        self.items[i]=result;self.done=i+1;self.event(Path(x.source).name+' · '+result.status)
            finally:
                with self.lock:self.busy=False;self.phase='识别完成，请核对后归档'
        threading.Thread(target=worker,daemon=True).start()
    def archive(self):
        with self.lock:
            if self.busy:raise ValueError('请等待识别完成')
            c=load_config();ledger=Ledger(self.output);n=0
            for x in self.items:
                if x.status=='可归档':
                    try:
                        if ledger.archive(x,c):n+=1
                    except Exception as e:x.status='失败';x.notes=str(e)[:250]
            self.phase=f'本次归档 {n} 份';self.event(self.phase+'，原件保持不变');return n

def make_handler(app):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self,*args):pass
        def send(self,body,status=200,kind='application/json; charset=utf-8',filename=None):
            if not isinstance(body,bytes):body=json.dumps(body,ensure_ascii=False).encode('utf-8')
            self.send_response(status);self.send_header('Content-Type',kind);self.send_header('Content-Length',str(len(body)))
            self.send_header('Cache-Control','no-store');self.send_header('X-Content-Type-Options','nosniff')
            self.send_header('Content-Security-Policy',"default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; frame-ancestors 'none'")
            if filename:self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+__import__('urllib.parse',fromlist=['quote']).quote(filename))
            self.end_headers();self.wfile.write(body)
        def do_GET(self):
            try:
                if urlparse('//'+self.headers.get('Host','')).hostname not in ('127.0.0.1','localhost'):return self.send({'error':'非本机主机名已拒绝'},403)
                p=urlparse(self.path);q=parse_qs(p.query)
                if p.path=='/':
                    body=(Path(__file__).parent/'web'/'index.html').read_text(encoding='utf-8').replace('__TOKEN__',app.token)
                    return self.send(body.encode(),kind='text/html; charset=utf-8')
                if p.path=='/api/state':return self.send(app.state())
                if p.path=='/api/config':
                    c=load_config();configured=bool(c.pop('api_key',''));return self.send({**c,'configured':configured})
                if p.path=='/api/ledger':return self.send(Ledger(app.output).rows(q.get('q',[''])[0]))
                if p.path=='/api/export':
                    query=q.get('q',[''])[0];p=app.base/'导出'/f'文件目录台账_{time.time_ns()}.xlsx';n=Ledger(app.output).export(p,query)
                    app.event(f'导出 Excel 台账 {n} 条');return self.send(p.read_bytes(),kind='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename='文件目录台账.xlsx')
                if p.path=='/api/package':
                    rows=Ledger(app.output).rows();dest=app.base/'导出'/f'归档成果_{time.time_ns()}.zip';dest.parent.mkdir(exist_ok=True)
                    ledger=dest.with_suffix('.xlsx');Ledger(app.output).export(ledger)
                    with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as z:
                        z.write(ledger,'文件目录台账.xlsx')
                        z.writestr('归档记录.json',json.dumps(rows,ensure_ascii=False,indent=2))
                        for r in rows:z.write(app.output/r['destination'],'归档文件/'+r['destination'].replace('\\','/'))
                    app.event(f'导出归档包 {len(rows)} 份');return self.send(dest.read_bytes(),kind='application/zip',filename='归档成果.zip')
                if p.path=='/api/original':
                    rows=Ledger(app.output).rows();r=next(x for x in rows if x['sha256']==q.get('sha',[''])[0]);f=app.output/r['destination']
                    return self.send(f.read_bytes(),kind='application/octet-stream',filename=f.name)
                return self.send({'error':'资源不存在'},404)
            except Exception as e:return self.send({'error':str(e)[:250]},400)
        def do_POST(self):
            try:
                if urlparse('//'+self.headers.get('Host','')).hostname not in ('127.0.0.1','localhost'):return self.send({'error':'非本机主机名已拒绝'},403)
                # A random per-process capability and Origin check prevent cross-site local writes.
                if self.headers.get('X-OA-Token')!=app.token: return self.send({'error':'无效会话，请刷新'},403)
                origin=self.headers.get('Origin')
                if origin and origin!='http://'+self.headers.get('Host',''):return self.send({'error':'跨站请求已拒绝'},403)
                n=int(self.headers.get('Content-Length',0))
                if n>115*1024*1024:raise ValueError('上传过大，单文件上限 80 MB')
                body=json.loads(self.rfile.read(n) or b'{}');p=urlparse(self.path).path
                if app.busy and p not in ('/api/noop',):raise ValueError('任务处理中，请稍候')
                if p=='/api/import':
                    paths=discover(body['folder'],app.output)
                    app.items=[Item(str(x),title=x.stem) for x in paths];app.phase=f'已导入 {len(paths)} 份';app.event(app.phase)
                elif p=='/api/upload':
                    raw=base64.b64decode(body['data'],validate=True)
                    if len(raw)>80*1024*1024:raise ValueError('单文件超过 80 MB')
                    name=safe_name(Path(body['name']).stem)+Path(body['name']).suffix.lower()
                    from .core import SUPPORTED
                    if Path(name).suffix.lower() not in SUPPORTED:raise ValueError('不支持此格式')
                    app.batch_dir.mkdir(parents=True,exist_ok=True);f=app.batch_dir/(secrets.token_hex(4)+'_'+name);f.write_bytes(raw)
                    app.items.append(Item(str(f),title=Path(name).stem));app.phase=f'已导入 {len(app.items)} 份'
                elif p=='/api/clear':app.items=[];app.done=0;app.phase='等待导入';app.batch_dir=app.base/'待处理'/str(time.time_ns())
                elif p=='/api/output':
                    output=Path(body['output']).expanduser().resolve()
                    if not output.is_absolute():raise ValueError('请选择绝对路径')
                    output.mkdir(parents=True,exist_ok=True);app.output=output
                    (app.base/'workspace-settings.json').write_text(json.dumps({'output':str(output)},ensure_ascii=False),encoding='utf-8');app.event('归档目录已更新')
                elif p=='/api/start':app.start(body.get('mode','local'))
                elif p=='/api/review':
                    x=app.items[int(body['index'])];c=load_config()
                    if x.status not in ('可归档','待复核'):raise ValueError('该文件不能复核，请先完成识别')
                    fields={k:str(body[k]).strip() for k in ('department','type','date','title')}
                    from datetime import datetime
                    datetime.strptime(fields['date'],'%Y-%m-%d')
                    if fields['department'] not in c['departments'] or fields['type'] not in c['document_types'] or not fields['title']:raise ValueError('请填写有效部门、类型和标题')
                    for k,v in fields.items():setattr(x,k,v)
                    x.status='可归档';x.notes+='；人工核对确认';app.event('人工确认：'+Path(x.source).name)
                elif p=='/api/archive':return self.send({'count':app.archive()})
                elif p=='/api/config':
                    c=load_config()
                    for k in DEFAULT:
                        if k in body and (k!='api_key' or body[k]):c[k]=body[k]
                    save_config(c);app.event('配置已保存到本机')
                else:return self.send({'error':'接口不存在'},404)
                return self.send({'ok':True})
            except Exception as e:return self.send({'error':str(e)[:250]},400)
    return Handler

def main():
    p=argparse.ArgumentParser();p.add_argument('--port',type=int,default=0);p.add_argument('--no-browser',action='store_true');args=p.parse_args()
    app=Workbench();server=ThreadingHTTPServer(('127.0.0.1',args.port),make_handler(app));url=f'http://127.0.0.1:{server.server_port}'
    if not args.no_browser:threading.Timer(.6,lambda:webbrowser.open(url)).start()
    print('OA Archive '+VERSION+' '+url,flush=True)
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()

if __name__=='__main__':main()
