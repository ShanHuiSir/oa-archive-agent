"""Exact secret/participant scans, recursively inspecting ZIP-based artifacts."""
from pathlib import Path
import json,zipfile,io,hashlib,sys,re
ROOT=Path(__file__).resolve().parents[1]
private=json.loads((ROOT/'.private/local.json').read_text(encoding='utf-8'));person=json.loads((ROOT/'.private/participant.json').read_text(encoding='utf-8'))
needles=[private['api_key'],person['contact'],person['name']];issues=[];checked=0
def scan(data,name,depth=0):
    global checked
    checked+=1
    for needle in needles:
        if any(needle.encode(enc) in data for enc in ['utf-8','utf-16le','utf-16be']):issues.append({'file':name,'kind':'exact confidential value'})
    if data[:2]==b'PK' and depth<4:
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for item in z.infolist():
                    if not item.is_dir():scan(z.read(item),name+'!'+item.filename,depth+1)
        except zipfile.BadZipFile:pass
    if name.endswith(('.py','.json','.md','.html','.txt','.yml','.cmd')):
        for hit in re.finditer(rb'(?i)(?:sk-[a-z0-9]{24,}|gh[pousr]_[a-z0-9]{24,})',data):issues.append({'file':name,'kind':'credential-like text'})
target=ROOT/'delivery/OAArchive-source-v1.0.0.zip';scan(target.read_bytes(),target.name)
scan((ROOT/'delivery/OAArchive-proposal-public.docx').read_bytes(),'proposal-public.docx')
scan((ROOT/'delivery/OAArchive-demo.mp4').read_bytes(),'demo.mp4')
with zipfile.ZipFile(target) as z:
    prefix='OAArchive-v1.0.0/';manifest=json.loads(z.read(prefix+'PACKAGE_MANIFEST.json'))
    for name,expected in manifest.items():assert hashlib.sha256(z.read(prefix+name)).hexdigest()==expected,name
    assert not any('/.private/' in n or '/.git/' in n or '/build/' in n or n.endswith('项目申报书.docx') for n in z.namelist())
    assert any(n.endswith('/OAArchive.exe') for n in z.namelist())
report=dict(version='1.0.0',files_and_nested_members_scanned=checked,confidential_findings=issues,manifest_entries_verified=len(manifest),archive_integrity_test=zipfile.ZipFile(target).testzip())
(ROOT/'.private/public_scan.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report,ensure_ascii=False,indent=2));sys.exit(1 if issues else 0)
