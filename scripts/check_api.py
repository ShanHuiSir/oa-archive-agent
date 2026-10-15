"""Bounded, real remote-service probe. Records events, never credentials."""
import json,sys,time,urllib.request,urllib.error
from pathlib import Path
from datetime import datetime
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from oa.core import load_config
c=load_config();report=dict(time=datetime.now().astimezone().isoformat(),endpoint=c['base_url'],model_requested=c['model'],credential_present=bool(c.get('api_key')),events=[],classification_response_received=False)
t=time.monotonic();body=dict(model=c['model'],messages=[dict(role='user',content='只返回一个JSON对象，内容是ok为true。')],max_tokens=64,thinking={'type':'disabled'},stream=True)
req=urllib.request.Request(c['base_url'].rstrip('/')+'/chat/completions',data=json.dumps(body).encode(),headers={'Authorization':'Bearer '+c['api_key'],'Content-Type':'application/json'})
try:
    with urllib.request.urlopen(req,timeout=35) as r:
        report['http_status']=r.status
        for line in r:
            s=line.decode('utf-8','replace').strip()
            if s:
                report['events'].append(dict(elapsed_s=round(time.monotonic()-t,3),line=s[:1000]))
                if s.startswith('data:') and '"content"' in s:report['classification_response_received']=True
            if time.monotonic()-t>32:report['outcome']='stopped_after_keepalives_without_classification';break
except urllib.error.HTTPError as e:report.update(http_status=e.code,outcome='http_error')
except Exception as e:report.update(outcome='network_or_read_timeout',error_type=type(e).__name__)
report['elapsed_s']=round(time.monotonic()-t,3)
(ROOT/'validation/api_service.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report,ensure_ascii=False,indent=2))
