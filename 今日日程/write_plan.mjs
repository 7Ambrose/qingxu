import fs from 'node:fs/promises';
import {spawnSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {basename,dirname,join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {SpreadsheetFile,FileBlob} from '@oai/artifact-tool';
const path=process.argv[2],python=process.argv[3];let input='';for await(const c of process.stdin)input+=c;
const q=JSON.parse(input),stat=await fs.stat(path,{bigint:true});
if(String(stat.mtimeNs)!==String(q.version))throw Error('Excel 已更新，请刷新后重试，避免覆盖修改。');
const digestOf=async file=>createHash('sha256').update(await fs.readFile(file)).digest('hex');
const baseDigest=await digestOf(path);
if(q.digest&&baseDigest!==String(q.digest))throw Error('Excel 内容与读取时不一致，已取消本次写入，请刷新后重试。');
const pruneBackups=async file=>{const folder=dirname(file),prefix=basename(file)+'.qingxu-backup-';const names=(await fs.readdir(folder)).filter(name=>name.startsWith(prefix)&&name.endsWith('.xlsx')).sort();for(const stale of names.slice(0,-8))await fs.rm(join(folder,stale),{force:true});};
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path));
const value=(s,a)=>s.getRange(a).values[0][0];const set=(s,a,v)=>{s.getRange(a).values=[[v??null]];};
const timeValue=(v)=>{if(v===undefined||v===null||v==='')return null;const m=String(v).match(/^(\d{1,2}):(\d{2})$/);if(!m)throw Error('时间请使用 HH:mm 格式。');const h=Number(m[1]),n=Number(m[2]);if(h>23||n>59)throw Error('时间不正确。');return (h*60+n)/1440;};
const parts=(q.id||'new').split(':'),type=parts[0];let s,row,cols;
const serial=Math.round((Date.parse(q.date+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000);
if(!Number.isFinite(serial)||serial<46313||serial>46404)throw Error('日期须在2026年10月18日至2027年1月17日内。');
if(type==='daily'||type==='new'){
 s=wb.worksheets.getItem('学习记录');row=Number(parts[1]);cols={title:'D',subject:'C',planned:'E',actual:'F',done:'G',highlight:'K',start:'L',order:'B'};
 if(type==='new'){
  const rows=s.getRange('A15:D599').values;let i=rows.findIndex(r=>r[0]===serial&&!r[3]);
  if(i<0)i=rows.findIndex((r,j)=>j>=273&&!r[3]);if(i<0)throw Error('学习记录已满，请先整理。');row=15+i;
  set(s,`A${row}`,serial);set(s,`B${row}`,rows.filter(r=>r[0]===serial&&r[3]).length+1);
 }
}else if(type==='class'){s=wb.worksheets.getItem('课内安排');row=Number(parts[1]);cols={title:'F',subject:'G',planned:'H',actual:'I',done:'J',highlight:'M'};
}else if(type==='weekly'){s=wb.worksheets.getItem(parts[1]);row=Number(parts[2]);cols={title:'B',done:'E',highlight:'I',scheduleDate:'J',start:'K',planned:'L',order:'M'};
}else if(type==='semester'){s=wb.worksheets.getItem('学期总清单');row=Number(parts[1]);cols={title:'A',done:'B',highlight:'G'};
}else if(type==='habit'){s=wb.worksheets.getItem(parts[1]);row=Number(parts[3]);cols={title:parts[2],done:parts[2],highlight:parts[2],start:parts[2],planned:parts[2],order:parts[2]};
}else throw Error('任务类型不支持。');
if(!Number.isInteger(row)||row<5||row>599)throw Error('任务位置无效。');
if(type==='habit'&&(!['B','C','D','E','F','G','H'].includes(parts[2])||![67,68].includes(row)))throw Error('每日项目位置无效。');
if(q.fields.delete!==true&&q.fields.title!==undefined&&!String(q.fields.title).trim())throw Error('请填写任务名称。');
if(q.fields.delete===true){if(type==='habit'){set(s,`${parts[2]}${row}`,'已删除');for(const clearRow of [row===67?69:72,row===67?73:74,row+3,row+8,row+10])set(s,`${parts[2]}${clearRow}`,null);}else{const clearColumns=type==='class'?['F','G','H','I','J','M']:type==='weekly'?['B','C','D','E','F','G','H','I','J','K','L','M']:type==='semester'?['A','B','D','G']:['A','B','C','D','E','F','G','H','I','J','K','L'];for(const column of clearColumns)set(s,`${column}${row}`,null);}}
else for(const [key,v]of Object.entries(q.fields)){
 if(key==='incomplete'){set(s,`${cols.done}${row}`,v?'☒ 未完成':q.fields.done===true?'☑ 已完成':'☐ 待完成');continue;}
 if(!cols[key])continue;
 let val=v,r=row;
 if(type==='habit'&&key==='title'){r=row+10;if(!value(s,'A'+r))set(s,'A'+r,row===67?'背单词名称':'跑步打卡名称');}
 if(type==='habit'&&key==='order'){r=row+8;if(!value(s,'A'+r))set(s,'A'+r,row===67?'背单词顺序':'跑步打卡顺序');}
 if(type==='habit'&&key==='start'){r=Number(parts[3])===67?69:72;if(!value(s,'A'+r))set(s,'A'+r,Number(parts[3])===67?'背单词开始时间':'跑步打卡开始时间');}
 if(type==='habit'&&key==='planned'){r=Number(parts[3])===67?73:74;if(!value(s,'A'+r))set(s,'A'+r,Number(parts[3])===67?'背单词时长（分钟）':'跑步时长（分钟）');}
 if(key==='done')val=q.fields.incomplete===true?'☒ 未完成':v?'☑ 已完成':'☐ 待完成';
 if(key==='highlight'){val=v?'强调':'普通';if(type==='habit')r+=3;}
 if(key==='start')val=timeValue(v);
 if(key==='order'){val=Number(v);if(!Number.isInteger(val)||val<1)throw Error('任务顺序无效。');if(type==='weekly'&&!value(s,'M49'))set(s,'M49','显示顺序');}
 if(key==='scheduleDate'){
  val=v?Math.round((Date.parse(v+'T00:00:00Z')-Date.UTC(1899,11,30))/86400000):null;
  if(val!==null){const weekStart=Number(value(s,'B5'));if(!Number.isFinite(val)||val<weekStart||val>weekStart+6)throw Error('周任务只能安排在所属周的周一至周日。');}
 }
 if(['planned','actual'].includes(key)){
  val=v===''||v===null?null:Number(v);if(val!==null&&(!Number.isFinite(val)||val<0||val>1440))throw Error('分钟须在0—1440之间。');
 }
 set(s,`${cols[key]}${r}`,val);
 if(type==='habit'&&key==='start')s.getRange(`${cols[key]}${r}`).format.numberFormat='hh:mm';
}
wb.recalculate();const temp=path+'.pending.xlsx';
try{
 await(await SpreadsheetFile.exportXlsx(wb)).save(temp);
 const current=await fs.stat(path,{bigint:true});if(current.mtimeNs!==stat.mtimeNs)throw Error('保存期间 Excel 有更新，已取消本次写入，请刷新后重试。');
 const check=spawnSync(python,['-X','utf8',fileURLToPath(new URL('./read_plan.py',import.meta.url)),'--check','--book',path],{encoding:'utf8',windowsHide:true});
 const result=JSON.parse(check.stdout);if(result.error)throw Error(result.error);
 if(await digestOf(path)!==baseDigest)throw Error('保存期间 Excel 有更新，已取消本次写入，请刷新后重试。');
 const backup=path+'.qingxu-backup-'+Date.now()+'.xlsx';
 await fs.copyFile(path,backup);
 try{await fs.rename(temp,path);}catch(error){await fs.rm(backup,{force:true});throw error;}
 await pruneBackups(path);
 console.log(JSON.stringify({ok:true,id:type==='new'?`daily:${row}`:q.id,digest:await digestOf(path)}));
}finally{await fs.rm(temp,{force:true});await fs.rm(temp+'.inspect.ndjson',{force:true});}
