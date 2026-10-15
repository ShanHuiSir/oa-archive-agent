import os,tempfile,threading,unittest,urllib.request,urllib.error,json,sys
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from oa.server import Workbench,make_handler,ThreadingHTTPServer
from oa.core import Item
class LocalServerTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.env=patch.dict(os.environ,{'OA_DATA':self.tmp.name});self.env.start();self.app=Workbench()
        self.server=ThreadingHTTPServer(('127.0.0.1',0),make_handler(self.app));threading.Thread(target=self.server.serve_forever,daemon=True).start();self.url='http://127.0.0.1:'+str(self.server.server_port)
    def tearDown(self):self.server.shutdown();self.server.server_close();self.env.stop();self.tmp.cleanup()
    def test_missing_capability_is_rejected(self):
        req=urllib.request.Request(self.url+'/api/clear',data=b'{}',headers={'Content-Type':'application/json'})
        with self.assertRaises(urllib.error.HTTPError) as e:urllib.request.urlopen(req)
        self.assertEqual(e.exception.code,403)
    def test_untrusted_hostname_rejected(self):
        req=urllib.request.Request(self.url+'/',headers={'Host':'attacker.invalid'})
        with self.assertRaises(urllib.error.HTTPError) as e:urllib.request.urlopen(req)
        self.assertEqual(e.exception.code,403)
    def test_no_key_is_returned_in_config(self):
        with urllib.request.urlopen(self.url+'/api/config') as r:c=json.load(r)
        self.assertNotIn('api_key',c);self.assertIn('configured',c)
    def test_invalid_review_does_not_mutate_result(self):
        self.app.items=[Item('x.txt',department='行政部',type='通知',date='2026-04-20',title='原标题',status='可归档')]
        body=dict(index=0,department='无效部门',type='通知',date='bad-date',title='changed')
        req=urllib.request.Request(self.url+'/api/review',data=json.dumps(body).encode(),headers={'Content-Type':'application/json','X-OA-Token':self.app.token})
        with self.assertRaises(urllib.error.HTTPError):urllib.request.urlopen(req)
        self.assertEqual(self.app.items[0].title,'原标题');self.assertEqual(self.app.items[0].date,'2026-04-20')
if __name__=='__main__':unittest.main()
