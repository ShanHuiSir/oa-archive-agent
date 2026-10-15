"""Reproducible live API verification; keeps original attachment details private."""
import argparse,concurrent.futures,json,sys,time
from dataclasses import asdict
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from oa.core import *
def main():
    p=argparse.ArgumentParser();p.add_argument('--attachments',action='store_true');p.add_argument('--local',action='store_true');args=p.parse_args()
    c=load_config();c['template_hint']='归口单位代表所属部门；制发日期代表生成日期；材料类别代表文件类型。'
    folder=next((ROOT/'.private/attachment').iterdir()) if args.attachments else ROOT/'samples'
    paths=discover(folder)
    if args.attachments:paths=[x for x in paths if not x.name.startswith('赛题')]
    dest=ROOT/'.private'/('attachment_run' if args.attachments else 'sample_run')/str(time.time_ns());ledger=Ledger(dest/'归档库');start=time.time()
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as ex:items=list(ex.map(lambda f:(analyze_local if args.local else analyze)(f,c),paths))
    for x in items:
        if x.status=='可归档':ledger.archive(x,c)
    ledger.export(dest/'文件目录台账.xlsx')
    report=dict(version='1.0.0',time=__import__('datetime').datetime.now().astimezone().isoformat(),input_count=len(items),
        total_seconds=round(time.time()-start,3),mode='local' if args.local else 'model',model_requested='local-rules-v1' if args.local else c['model'],records=[])
    expected=json.loads((ROOT/'tests/expected_samples.json').read_text(encoding='utf-8'))
    for i,x in enumerate(items):
        r=asdict(x);r.pop('text');r['source']=Path(x.source).name;r['destination']=str(Path(x.destination).relative_to(dest)) if x.destination else ''
        if not args.attachments:
            e=next(z for z in expected if z['file']==r['source']);r['field_checks']={k:r[k]==e[k] for k in ('department','type','date')}
        report['records'].append(r)
        print(r['source'],r['status'],r['department'],r['type'],r['date'],r['notes'],flush=True)
    target=ROOT/('.private/attachment_actual.json' if args.attachments else ('validation/local_samples.json' if args.local else 'validation/live_samples.json'));target.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    if args.attachments:
        public={k:v for k,v in report.items() if k!='records'};public['cases']=[dict(case=i+1,format=paths[i].suffix,extraction=x.extraction,status=x.status,model=x.model,request_id=x.request_id,elapsed_s=x.elapsed_s,notes=x.notes) for i,x in enumerate(items)]
        (ROOT/'validation/attachment_summary.json').write_text(json.dumps(public,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Live report saved; successful requests:',sum(bool(x.request_id) for x in items))
if __name__=='__main__':main()
