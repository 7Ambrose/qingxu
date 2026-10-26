import fs from 'node:fs/promises';
import {spawnSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {basename,dirname,join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {SpreadsheetFile,FileBlob} from '@oai/artifact-tool';

const path=process.argv[2],python=process.argv[3];let input='';for await(const c of process.stdin)input+=c;
const q=JSON.parse(input),stat=await fs.stat(path,{bigint:true});
if(String(stat.mtimeNs)!==String(q.version))throw Error('检测到 Excel 在读取后已被修改。本次同步已取消，没有覆盖 Excel；本地 JSON 修改仍保留，请核对后再重试。');
const digestOf=async file=>createHash('sha256').update(await fs.readFile(file)).digest('hex');
const baseDigest=await digestOf(path);
if(q.digest&&baseDigest!==String(q.digest))throw Error('检测到 Excel 内容与读取时不一致（所在磁盘的时间戳精度不足时常见）。本次同步已取消，没有覆盖 Excel；本地 JSON 修改仍保留，请核对后再重试。');
const pruneBackups=async file=>{const folder=dirname(file),prefix=basename(file)+'.qingxu-backup-';const names=(await fs.readdir(folder)).filter(name=>name.startsWith(prefix)&&name.endsWith('.xlsx')).sort();for(const stale of names.slice(0,-8))await fs.rm(join(folder,stale),{force:true});};
const linksScript=fileURLToPath(new URL('./excel_links.py',import.meta.url));
const runLinks=(args,body)=>{const result=spawnSync(python,[ '-X','utf8',linksScript,...args],{input:body?JSON.stringify(body):undefined,encoding:'utf8',windowsHide:true,maxBuffer:16*1024*1024});if(result.status!==0)throw Error('Excel 联动读取失败，请保存并关闭 Excel 后重试。');const parsed=JSON.parse(result.stdout);if(parsed.error)throw Error(parsed.error);return parsed;};
const linkSeed=runLinks([path]);
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const value=(s,a)=>s.getRange(a).values[0][0],set=(s,a,v)=>{s.getRange(a).values=[[v??null]];};
const applyLinkPatches=plan=>{for(const patch of plan.patches){set(wb.worksheets.getItem(patch.sheet),patch.cell,patch.value);linkSeed.data[patch.sheet][patch.cell]=patch.value;}};
const firstLinks=runLinks(['--resolve'],{data:linkSeed.data,baseline:linkSeed.baseline});applyLinkPatches(firstLinks);
const linkBaseline=runLinks(['--snapshot'],{data:linkSeed.data,previous:firstLinks.state});
const timeValue=v=>{if(v===undefined||v===null||v==='')return null;const m=String(v).match(/^(\d{1,2}):(\d{2})$/);if(!m)throw Error('时间请使用 HH:mm 格式。');const h=Number(m[1]),n=Number(m[2]);if(h>23||n>59)throw Error('时间不正确。');return (h*60+n)/1440;};
const serialOf=d=>Math.round((Date.parse(d+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000);
const mappings={};
for(const op of q.ops||[]){
 const parts=(op.id||'new').split(':'),type=parts[0];let s,row,cols;
 const serial=serialOf(op.date);if(!Number.isFinite(serial)||serial<46313||serial>46404)throw Error('日期须在2026年10月18日至2027年1月17日内。');
 if(type==='daily'||type==='new'||type==='local'){
  s=wb.worksheets.getItem('学习记录');row=Number(parts[1]);cols={title:'D',subject:'C',planned:'E',actual:'F',done:'G',highlight:'K',start:'L',order:'B'};
  if(type==='new'||type==='local'){
   const rows=s.getRange('A15:D599').values;let i=rows.findIndex(r=>r[0]===serial&&!r[3]);
   if(i<0)i=rows.findIndex((r,j)=>j>=273&&!r[3]);if(i<0)throw Error('学习记录已满，请先整理。');row=15+i;
   set(s,`A${row}`,serial);set(s,`B${row}`,rows.filter(r=>r[0]===serial&&r[3]).length+1);
   if(type==='local')mappings[op.id]=`daily:${row}`;
  }
 }else if(type==='class'){s=wb.worksheets.getItem('课内安排');row=Number(parts[1]);cols={title:'F',subject:'G',planned:'H',actual:'I',done:'J',highlight:'M'};}
 else if(type==='weekly'){s=wb.worksheets.getItem(parts[1]);row=Number(parts[2]);cols={title:'B',done:'E',highlight:'I',scheduleDate:'J',start:'K',planned:'L',order:'M'};}
 else if(type==='semester'){s=wb.worksheets.getItem('学期总清单');row=Number(parts[1]);cols={title:'A',done:'B',highlight:'G'};}
 else if(type==='event'){s=wb.worksheets.getItem('日程事项');row=Number(parts[1]);cols={title:'E',scheduleDate:'A',start:'C',planned:'D',done:'H',highlight:'I',order:'J'};}
 else if(type==='habit'){s=wb.worksheets.getItem(parts[1]);row=Number(parts[3]);cols={title:parts[2],done:parts[2],highlight:parts[2],start:parts[2],planned:parts[2],order:parts[2]};}
 else throw Error('任务类型不支持：'+op.id);
 if(!Number.isInteger(row)||row<5||row>599)throw Error('任务位置无效。');
 if(type==='habit'&&(!['B','C','D','E','F','G','H'].includes(parts[2])||![67,68].includes(row)))throw Error('每日项目位置无效。');
 if(op.fields.delete!==true&&op.fields.title!==undefined&&!String(op.fields.title).trim())throw Error('请填写任务名称。');
 if(op.fields.delete===true){if(type==='habit'){set(s,`${parts[2]}${row}`,'已删除');for(const clearRow of [row===67?69:72,row===67?73:74,row+3,row+8,row+10])set(s,`${parts[2]}${clearRow}`,null);}else{const clearColumns=type==='class'?['F','G','H','I','J','M']:type==='weekly'?['B','C','D','E','F','G','H','I','J','K','L','M']:type==='semester'?['A','B','D','G']:type==='event'?['A','B','C','D','E','F','G','H','I','J']:['A','B','C','D','E','F','G','H','I','J','K','L'];for(const column of clearColumns)set(s,`${column}${row}`,null);}continue;}
 const eventStart=type==='event'?(op.fields.start!==undefined?timeValue(op.fields.start):value(s,`C${row}`)):null;
 const eventMinutes=type==='event'?(op.fields.planned!==undefined&&op.fields.planned!==''?Number(op.fields.planned):(Number(value(s,`D${row}`))-Number(value(s,`C${row}`)))*1440):null;
 for(const [key,v]of Object.entries(op.fields)){
  if(key==='incomplete'){set(s,`${cols.done}${row}`,v?'☒ 未完成':op.fields.done===true?'☑ 已完成':'☐ 待完成');continue;}
  if(!cols[key])continue;let val=v,r=row;
  if(type==='habit'&&key==='title'){r=row+10;if(!value(s,'A'+r))set(s,'A'+r,row===67?'背单词名称':'跑步打卡名称');}
  if(type==='habit'&&key==='order'){r=row+8;if(!value(s,'A'+r))set(s,'A'+r,row===67?'背单词顺序':'跑步打卡顺序');}
  if(type==='habit'&&key==='start'){r=Number(parts[3])===67?69:72;if(!value(s,'A'+r))set(s,'A'+r,Number(parts[3])===67?'背单词开始时间':'跑步打卡开始时间');}
  if(type==='habit'&&key==='planned'){r=Number(parts[3])===67?73:74;if(!value(s,'A'+r))set(s,'A'+r,Number(parts[3])===67?'背单词时长（分钟）':'跑步时长（分钟）');}
  if(key==='done')val=op.fields.incomplete===true?'☒ 未完成':v?'☑ 已完成':'☐ 待完成';
  if(key==='highlight'){val=v?'强调':'普通';if(type==='habit')r+=3;}
  if(key==='start')val=timeValue(v);
  if(key==='order'){val=Number(v);if(!Number.isInteger(val)||val<1)throw Error('任务顺序无效。');if(type==='weekly'&&!value(s,'M49'))set(s,'M49','显示顺序');}
  if(key==='scheduleDate'){
   val=v?serialOf(v):null;if(type==='weekly'&&val!==null){const weekStart=Number(value(s,'B5'));if(!Number.isFinite(val)||val<weekStart||val>weekStart+6)throw Error('周任务只能安排在所属周的周一至周日。');}if(type==='event'&&(val===null||!Number.isFinite(val)||val<46313||val>46404))throw Error('请填写学期范围内的事项日期。');
  }
  if(['planned','actual'].includes(key)){val=v===''||v===null?null:Number(v);if(val!==null&&(!Number.isFinite(val)||val<0||val>1440))throw Error('分钟须在0—1440之间。');}
  set(s,`${cols[key]}${r}`,val);
  if(type==='habit'&&key==='start')s.getRange(`${cols[key]}${r}`).format.numberFormat='hh:mm';
 }
 if(type==='event'){
  set(s,`C${row}`,eventStart);const finish=eventStart===null?null:Number(eventStart)+Math.max(1,Number(eventMinutes)||60)/1440;if(finish!==null&&finish>1)throw Error('结束时间不能超过当天24:00。');set(s,`D${row}`,finish);
  if(!value(s,'H4'))set(s,'H4','完成状态');if(!value(s,'I4'))set(s,'I4','强调');if(!value(s,'J4'))set(s,'J4','显示顺序');
 }
}
for(const [sheet,cells]of Object.entries(linkSeed.data)){const source=wb.worksheets.getItem(sheet);const end=Object.keys(cells).reduce((max,cell)=>Math.max(max,Number(cell.replace(/[A-Z]/g,''))),80);const rows=source.getRange(`A1:M${end}`).values;for(const cell of Object.keys(cells)){const match=cell.match(/^([A-Z]+)(\d+)$/);let column=0;for(const letter of match[1])column=column*26+letter.charCodeAt(0)-64;cells[cell]=rows[Number(match[2])-1]?.[column-1]??null;}}
 const eventRows=wb.worksheets.getItem('日程事项').getRange('A1:J199').values;for(let row=4;row<200;row++)for(const col of 'ABCDEFGHIJ')linkSeed.data['日程事项'][`${col}${row}`]=eventRows[row-1]?.[col.charCodeAt(0)-65]??null;
const finalLinks=runLinks(['--resolve'],{data:linkSeed.data,baseline:linkBaseline});applyLinkPatches(finalLinks);
const finalLinkState=runLinks(['--snapshot'],{data:linkSeed.data,previous:finalLinks.state});
wb.recalculate();const temp=path+'.pending.xlsx';
try{
 await(await SpreadsheetFile.exportXlsx(wb)).save(temp);
 const current=await fs.stat(path,{bigint:true});if(current.mtimeNs!==stat.mtimeNs)throw Error('保存期间 Excel 有更新，已取消本次写入，请重新同步。');
 const check=spawnSync(python,['-X','utf8',fileURLToPath(new URL('./read_plan.py',import.meta.url)),'--check','--book',path],{encoding:'utf8',windowsHide:true});
 const result=JSON.parse(check.stdout);if(result.error)throw Error(result.error);
 const savedStat=await fs.stat(temp,{bigint:true});
 if(await digestOf(path)!==baseDigest)throw Error('保存期间 Excel 有更新，已取消本次写入，请重新同步。');
 const backup=path+'.qingxu-backup-'+Date.now()+'.xlsx';
 await fs.copyFile(path,backup);
 try{await fs.rename(temp,path);}catch(error){await fs.rm(backup,{force:true});throw error;}
 await pruneBackups(path);
 const stateFile=path+'.agenda-links.json';await fs.writeFile(stateFile+'.pending',JSON.stringify(finalLinkState));await fs.rename(stateFile+'.pending',stateFile);console.log(JSON.stringify({ok:true,mappings,version:String(savedStat.mtimeNs),digest:await digestOf(path)}));
}finally{await fs.rm(temp,{force:true});await fs.rm(temp+'.inspect.ndjson',{force:true});}
