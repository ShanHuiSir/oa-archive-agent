from pathlib import Path
import importlib.metadata as m
import ast
ROOT=Path(__file__).resolve().parents[1];dest=ROOT/'licenses';dest.mkdir(exist_ok=True)
names=['PyMuPDF','python-docx','openpyxl','Pillow','rapidocr_onnxruntime','xlrd','numpy','onnxruntime','opencv-python','PyYAML','coloredlogs','flatbuffers','protobuf','sympy','humanfriendly','mpmath','typing_extensions','packaging','lxml','et-xmlfile','shapely','six']
toc=ROOT/'build/OAArchive/Analysis-00.toc'
if toc.exists():
    mapping=m.packages_distributions();modules=set()
    def visit(v):
        if isinstance(v,(tuple,list)):
            if len(v)==3 and isinstance(v[0],str) and v[2] in ('PYMODULE','EXTENSION'):modules.add(v[0].split('.')[0])
            else:
                for x in v:visit(x)
    visit(ast.literal_eval(toc.read_text(encoding='utf-8')))
    for mod in modules:names.extend(mapping.get(mod,[]))
names=sorted(set(names),key=str.casefold)
rows=[]
for name in names:
    try:d=m.distribution(name)
    except m.PackageNotFoundError:continue
    rows.append((name,d.version,d.metadata.get('License-Expression') or d.metadata.get('License') or 'See included license',d.metadata.get('Home-page') or 'https://pypi.org/project/'+name))
    for f in d.files or []:
        if any(k in Path(str(f)).name.upper() for k in ['LICENSE','COPYING','NOTICE']) and ('.dist-info/' in str(f).replace('\\','/') or len(Path(f).parts)<3):
            p=Path(d.locate_file(f))
            if p.is_file():
                out=dest/(name+'_'+Path(f).name);out.write_bytes(p.read_bytes())
text='# Third-party notices\n\nThis application is distributed under GNU AGPL v3, matching the included PyMuPDF component. Third-party components retain their own licenses. Corresponding upstream source is available at the project links and package source distributions.\n\n'
for name,version,license,url in rows:text+=f'## {name} {version}\n\n{license}\n\nSource: {url}\n\n'
text+='Python runtime source and license: https://www.python.org/downloads/source/ . PyInstaller bootloader is distributed with its license exception; see https://pyinstaller.org/en/stable/license.html . RapidOCR model files retain upstream terms: https://github.com/RapidAI/RapidOCR .\n'
(ROOT/'THIRD_PARTY_NOTICES.md').write_text(text,encoding='utf-8');print('Collected licenses for',len(rows),'components')
