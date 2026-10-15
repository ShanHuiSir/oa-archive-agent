"""Narrate genuine browser recordings, shorten waits, burn synchronized subtitles."""
import asyncio,json,re,subprocess,shutil
from pathlib import Path
import edge_tts,imageio_ffmpeg
ROOT=Path(__file__).resolve().parents[1];RAW=ROOT/'.private/video';DEMO=ROOT/'demo';FF=imageio_ffmpeg.get_ffmpeg_exe()
SECTIONS=[
('业务问题','办公文件一多，通知、合同、表格，还有扫描件，很容易混在一起。这里是归序，一个企业文件归档工作台。接下来用明确标注的模拟资料，实际走一遍，从导入到导出。'),
('输入资料','先看企业规则。这里把归口单位解释为部门，制发日期解释为生成日期，企业自己的模板也能使用。回到工作台，填入模拟资料文件夹，再点击导入。六份文件都在这里，Word、Excel、PDF和图片都有。'),
('操作处理','点击本地智能识别。程序先提取正文，扫描件通过本地文字识别处理，再按部门、类型和日期规则生成候选。这里清楚标注了本地规则来源。旁边的模型增强入口会直接请求配置的模型服务，两种处理来源分别记录。这次演示使用本地规则模式。'),
('查看结果','识别后，先别急着归档。点开扫描会议纪要，能看到提取的文字、字段依据，还有明确标注的处理来源。再看这份没有日期的报告，程序没有把今天填进去。我按模拟资料负责人提供的信息，补录四月九日，并保存人工核对。确认后归档，原件保留，副本按部门、类型和年月进入目录。'),
('搜索与导出','归档后打开台账，每份文件的部门、类型、日期和路径都能查到。输入法务部，只看这一类资料，导出的表格也只保留当前结果。清空搜索，再导出全量台账和完整归档包。这些下载文件，都是刚才的实际操作生成的，后面查文件、核对去向，就有依据了。')]
def duration(p):
    r=subprocess.run([FF,'-i',str(p)],capture_output=True);m=re.search(rb'Duration: (\d+):(\d+):(\d+\.\d+)',r.stderr)
    if not m:raise ValueError('Could not read media duration')
    return int(m[1])*3600+int(m[2])*60+float(m[3])
def run(args,cwd=RAW):
    r=subprocess.run([FF,'-hide_banner','-y',*args],cwd=cwd,capture_output=True)
    if r.returncode:raise RuntimeError(r.stderr.decode('utf-8','replace')[-3000:])
def stamp(t):
    ms=round(t*1000);h,ms=divmod(ms,3600000);m,ms=divmod(ms,60000);s,ms=divmod(ms,1000);return f'{h:02}:{m:02}:{s:02},{ms:03}'
def astamp(t):
    cs=round(t*100);h,cs=divmod(cs,360000);m,cs=divmod(cs,6000);s,cs=divmod(cs,100);return f'{h}:{m:02}:{s:02}.{cs:02}'
async def main():
    events=json.loads((RAW/'events.json').read_text(encoding='utf-8'));ev={x['name']:x['t'] for x in events};timeline=[];offset=0;captions=[]
    for i,(title,text) in enumerate(SECTIONS):
        audio=RAW/f'voice{i}.mp3';entries=[]
        if not audio.exists() or '--voice' in __import__('sys').argv:
            voice=edge_tts.Communicate(text,'zh-CN-XiaoxiaoNeural',rate=['-2%','+1%','-1%','-2%','+1%'][i],boundary='SentenceBoundary')
            with audio.open('wb') as f:
                async for ch in voice.stream():
                    if ch['type']=='audio':f.write(ch['data'])
                    elif ch['type'] in ('SentenceBoundary','WordBoundary'):entries.append(ch)
            (RAW/f'voice{i}.json').write_text(json.dumps(entries,ensure_ascii=False),encoding='utf-8')
        else:entries=json.loads((RAW/f'voice{i}.json').read_text(encoding='utf-8'))
        dur=duration(audio)+1.1;a=ev[title];b=ev[SECTIONS[i+1][0]] if i+1<len(SECTIONS) else ev['结束']
        factor=dur/(b-a)
        (RAW/f'chapter{i}.txt').write_text(f'归序  /  {i+1:02d} {title}    ·    模拟业务数据  本地规则模式  本轮真实运行',encoding='utf-8')
        filt=f"setpts={factor:.8f}*(PTS-STARTPTS),fps=24,scale=1440:1080,pad=1440:1220:0:60:color=0x112f36,drawtext=fontfile='C\\:/Windows/Fonts/msyh.ttc':textfile=chapter{i}.txt:fontsize=25:fontcolor=white:x=30:y=18"
        run(['-ss',str(a),'-t',str(b-a),'-i','raw.webm','-i',f'voice{i}.mp3','-vf',filt,'-af','apad','-t',str(dur),'-c:v','libx264','-preset','fast','-crf','20','-c:a','aac','-b:a','160k',f'clip{i}.mp4'])
        for e in entries:
            start=offset+e['offset']/1e7;end=start+e['duration']/1e7;sentence=e['text'];lines=[]
            while len(sentence)>42:
                cut=max(sentence.rfind(c,0,42)+1 for c in '，。；');cut=cut if cut>12 else 36;lines.append(sentence[:cut]);sentence=sentence[cut:]
            lines.append(sentence);captions.append([start,end,'\n'.join(lines)])
        timeline.append(dict(chapter=title,source_start=a,source_end=b,output_start=round(offset,3),output_duration=round(dur,3),source_to_output_ratio=round(factor,5)));offset+=dur
    (RAW/'concat.txt').write_text('\n'.join(f"file 'clip{i}.mp4'" for i in range(len(SECTIONS))),encoding='utf-8')
    run(['-f','concat','-safe','0','-i','concat.txt','-c','copy','joined.mp4'])
    for i in range(len(captions)-1):captions[i][1]=min(captions[i][1],captions[i+1][0]-.01)
    (DEMO/'演示字幕.srt').write_text('\n\n'.join(f'{i+1}\n{stamp(a)} --> {stamp(b)}\n{t}' for i,(a,b,t) in enumerate(captions)),encoding='utf-8')
    header='''[Script Info]
ScriptType: v4.00+
PlayResX: 1440
PlayResY: 1220
WrapStyle: 2
[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,Microsoft YaHei,26,&H00FFFFFF,&H00FFFFFF,&H00362F11,&H00362F11,0,0,0,0,100,100,0,0,1,1,0,2,30,30,20,1
[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
'''
    ass=header+'\n'.join(f'Dialogue: 0,{astamp(a)},{astamp(b)},Default,,0,0,0,,{t.replace(chr(10),chr(92)+"N")}' for a,b,t in captions)
    (DEMO/'演示字幕.ass').write_text(ass,encoding='utf-8');(RAW/'captions.ass').write_text(ass,encoding='utf-8')
    run(['-i','joined.mp4','-vf','ass=captions.ass','-c:v','libx264','-preset','fast','-crf','19','-c:a','copy','-movflags','+faststart',str(ROOT/'delivery/归序演示.mp4')])
    (DEMO/'演示解说词.txt').write_text('\n\n'.join(f'{i+1:02d} {title}\n{text}' for i,(title,text) in enumerate(SECTIONS)),encoding='utf-8')
    actual=json.loads((DEMO/'本轮操作结果.json').read_text(encoding='utf-8'))
    report=dict(version='1.0.0',duration_s=duration(ROOT/'delivery/归序演示.mp4'),resolution='1440x1220',video='H.264 24fps',audio='AAC 160k Chinese zh-CN-XiaoxiaoNeural',
        source='Actual Chromium workbench recording through Playwright; mouse movements, clicks, typing, scrolls and downloads executed against the real application. Cursor overlay follows actual browser mouse events. No mock responses or fabricated result frames.',
        synthetic_data=True,input_count=len(actual['state']['items']),archived_count=len(actual['ledger']),human_provided_date='2026-04-09 for missing-date report',chapters=timeline)
    (DEMO/'视频实测记录.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report,ensure_ascii=False,indent=2))
if __name__=='__main__':asyncio.run(main())
