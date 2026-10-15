/* Records actual browser interactions against the running application.
   OA_NODE_MODULES selects the bundled runtime; no results are fabricated. */
const {chromium}=require(process.env.OA_NODE_MODULES?process.env.OA_NODE_MODULES+'/playwright':'playwright');
const fs=require('fs');const path=require('path');
(async()=>{
 const base=path.resolve('.private/video');fs.mkdirSync(base,{recursive:true});const events=[];let started=Date.now();
 const mark=(name)=>{events.push({name,t:(Date.now()-started)/1000});fs.writeFileSync(base+'/events.json',JSON.stringify(events,null,2));console.log(name,events.at(-1).t)};
 const browser=await chromium.launch({channel:'chrome',headless:true});
 const context=await browser.newContext({viewport:{width:1440,height:1080},deviceScaleFactor:1,recordVideo:{dir:base,size:{width:1440,height:1080}},acceptDownloads:true});
 const page=await context.newPage();started=Date.now();await page.goto('http://127.0.0.1:8765');await page.waitForTimeout(1200);
 await page.addStyleTag({content:'html{scroll-behavior:smooth}#demo-pointer{position:fixed;left:0;top:0;width:15px;height:20px;background:#123b35;clip-path:polygon(0 0,0 100%,35% 72%,62% 100%,77% 91%,52% 62%,95% 60%);z-index:9999;pointer-events:none;filter:drop-shadow(0 0 2px white)}'});
 await page.evaluate(()=>{let p=document.createElement('div');p.id='demo-pointer';document.body.append(p);document.addEventListener('mousemove',e=>{p.style.transform=`translate(${e.clientX}px,${e.clientY}px)`});document.addEventListener('mousedown',()=>p.style.background='#d76d34');document.addEventListener('mouseup',()=>p.style.background='#123b35')});
 const wait=ms=>page.waitForTimeout(ms);
 async function click(locator){let box=await locator.boundingBox();if(!box)throw Error('Target not visible');await page.mouse.move(box.x+box.width/2,box.y+box.height/2,{steps:22});await wait(250);await locator.click();await wait(650)}
 mark('业务问题');await page.mouse.move(600,240,{steps:25});await wait(7000);
 mark('输入资料');await click(page.locator('[data-view=settings]'));await page.locator('[name=template_hint]').fill('归口单位代表所属部门；制发日期代表生成日期；材料类别代表文件类型。');await click(page.getByRole('button',{name:'保存本机设置'}));await wait(1200);await click(page.locator('[data-view=work]'));
 await click(page.locator('#folder'));await page.locator('#folder').pressSequentially('C:/Users/Public/OAArchiveDemo/模拟资料',{delay:50});await click(page.locator('#import'));await wait(2500);await page.screenshot({path:'docs/导入界面.png'});
 mark('操作处理');await click(page.locator('#start'));await wait(1800);mark('模型等待开始');
 let deadline=Date.now()+1100000;
 while(true){let x=await page.evaluate(async()=>await(await fetch('/api/state')).json());if(!x.busy&&x.done===6)break;if(Date.now()>deadline)throw Error('Live recognition exceeded demo deadline');await wait(1000)}
 await wait(1200);mark('模型等待结束');
 let result=await page.evaluate(async()=>await(await fetch('/api/state')).json());fs.writeFileSync('demo/识别原始结果.json',JSON.stringify(result.items,null,2));
 if(result.items.some(x=>!['可归档','待复核','已存在'].includes(x.status)))throw Error('Recognition failed; do not publish this recording');
 mark('查看结果');await page.mouse.move(1000,650,{steps:25});await wait(4000);
 await click(page.locator('#input-table tbody tr').nth(3).getByRole('button'));await wait(5000);await page.screenshot({path:'docs/OCR实际结果.png'});await click(page.locator('#close-review'));
 for(let i=0;i<result.items.length;i++){
   let x=result.items[i];if(x.status!=='待复核')continue;
   await click(page.locator('#input-table tbody tr').nth(i).getByRole('button'));
   if(!x.date){await page.locator('#review-form [name=date]').fill('2026-04-09');await wait(2500)}
   await click(page.locator('#review-form button[type=submit]'));await wait(800);
 }
 await click(page.locator('#archive'));await wait(2200);await page.screenshot({path:'docs/归档实际结果.png'});await page.mouse.wheel(0,450);await wait(1800);await page.mouse.wheel(0,-450);await wait(1200);
 mark('搜索与导出');await click(page.locator('[data-view=ledger]'));await wait(3000);await page.screenshot({path:'docs/台账实际结果.png'});
 await click(page.locator('#search'));await page.locator('#search').pressSequentially('法务部',{delay:250});await click(page.locator('#do-search'));await wait(1600);
 const filtered=page.waitForEvent('download');await click(page.locator('#export'));await (await filtered).saveAs(path.resolve('demo/筛选台账_法务部.xlsx'));await wait(1500);
 await page.locator('#search').fill('');await click(page.locator('#do-search'));await wait(1500);
 const excel=page.waitForEvent('download');await click(page.locator('#export'));await (await excel).saveAs(path.resolve('demo/本轮文件目录台账.xlsx'));
 const zip=page.waitForEvent('download');await click(page.locator('#package'));await (await zip).saveAs(path.resolve('demo/本轮归档成果.zip'));await wait(4500);
 let final=await page.evaluate(async()=>({state:await(await fetch('/api/state')).json(),ledger:await(await fetch('/api/ledger')).json()}));
 fs.writeFileSync('demo/本轮操作结果.json',JSON.stringify(final,null,2));mark('结束');
 const v=page.video();await context.close();await v.saveAs(base+'/raw.webm');await browser.close();
})().catch(e=>{console.error(e);process.exitCode=1});
