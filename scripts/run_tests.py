import unittest,json,time,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
suite=unittest.defaultTestLoader.discover(str(ROOT/'tests'));start=time.time()
with (ROOT/'validation/unit_tests.txt').open('w',encoding='utf-8') as f:r=unittest.TextTestRunner(stream=f,verbosity=2).run(suite)
summary=dict(tests_run=r.testsRun,failures=len(r.failures),errors=len(r.errors),successful=r.wasSuccessful(),seconds=round(time.time()-start,3))
(ROOT/'validation/test_summary.json').write_text(json.dumps(summary,indent=2),encoding='utf-8');print(summary)
sys.exit(0 if r.wasSuccessful() else 1)
