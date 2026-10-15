"""Private connectivity probe. Never prints credentials."""
import json,sys,time,urllib.request
from pathlib import Path
c=json.loads(Path('.private/local.json').read_text(encoding='utf-8'))
model=sys.argv[1] if len(sys.argv)>1 else 'deepseek-flash'
body=dict(model=model,messages=[dict(role='user',content='只返回一个包含ok字段的JSON对象，值为true。')],max_tokens=64,thinking={'type':'disabled'})
req=urllib.request.Request(c['base_url'].rstrip('/')+'/chat/completions',data=json.dumps(body).encode(),headers={'Authorization':'Bearer '+c['api_key'],'Content-Type':'application/json'})
t=time.time()
try:
    with urllib.request.urlopen(req,timeout=40) as r: result=json.load(r)
    print(json.dumps(result,ensure_ascii=False));print('seconds',round(time.time()-t,2))
except Exception as e: print(type(e).__name__,str(e))
