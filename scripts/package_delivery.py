"""Allow-list public source and executable package. Private material is excluded."""
from pathlib import Path
import zipfile,hashlib,json,shutil
ROOT=Path(__file__).resolve().parents[1];DEST=ROOT/'delivery';DEST.mkdir(exist_ok=True)
allow_dirs=['oa','scripts','tests','samples','docs','demo','validation','licenses','.github']
allow_files=['main.py','requirements.txt','requirements-dev.txt','package.json','config.example.json','README.md','LICENSE','THIRD_PARTY_NOTICES.md','install.cmd','start.cmd','build_windows.cmd','.gitignore']
files=[]
for name in allow_files:files.append((ROOT/name,name))
for name in allow_dirs:
    for p in (ROOT/name).rglob('*'):
        if p.is_file() and '__pycache__' not in p.parts and p.suffix!='.pyc':files.append((p,p.relative_to(ROOT).as_posix()))
for p in (ROOT/'dist/OAArchive').rglob('*'):
    if p.is_file():files.append((p,'Windows/OAArchive/'+p.relative_to(ROOT/'dist/OAArchive').as_posix()))
archive=DEST/'OAArchive-source-v1.0.0.zip'
manifest={name:hashlib.sha256(p.read_bytes()).hexdigest() for p,name in files}
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
    for p,name in files:z.write(p,'OAArchive-v1.0.0/'+name)
    z.writestr('OAArchive-v1.0.0/PACKAGE_MANIFEST.json',json.dumps(manifest,ensure_ascii=False,indent=2))
shutil.copy2(ROOT/'docs/项目申报书公开版.docx',DEST/'OAArchive-proposal-public.docx')
shutil.copy2(DEST/'归序演示.mp4',DEST/'OAArchive-demo.mp4')
assets=['OAArchive-source-v1.0.0.zip','OAArchive-demo.mp4','OAArchive-proposal-public.docx']
(DEST/'SHA256SUMS.txt').write_text('\n'.join(hashlib.sha256((DEST/name).read_bytes()).hexdigest()+'  '+name for name in assets)+'\n',encoding='utf-8')
print(json.dumps({'source_zip':str(archive),'files':len(files),'bytes':archive.stat().st_size},ensure_ascii=False))
