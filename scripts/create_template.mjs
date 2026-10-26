import fs from 'node:fs/promises';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

const output = process.argv[2];
const previewDir = process.argv[3];
const workbook = Workbook.create();
const base = Date.UTC(1899,11,30);
const first = (Date.UTC(2026,9,19)-base)/86400000;
const names = ['月历','学习记录','课内安排','日程事项','学期总清单',...Array.from({length:13},(_,index)=>index<3?'秋'+(index+6)+'周':'冬'+(index-2)+'周')];
for(const name of names){
    const sheet=workbook.worksheets.add(name);
    sheet.getRange('A1:M80').format.font={name:'Microsoft YaHei',size:11};
    sheet.getRange('A1:M80').format.columnWidth=17;
    sheet.getRange('A1:M80').format.rowHeight=23;
    sheet.getRange('A1').values=[[name+' · 空白模板']];
    sheet.getRange('A1:M1').format.fill='#DCEEE2';
}
const headers=(name,row,values)=>workbook.worksheets.getItem(name).getRangeByIndexes(row-1,0,1,values.length).values=[values];
headers('学习记录',14,['日期','顺序','科目','任务','计划分钟','实际分钟','状态','','','','强调','开始时间']);
headers('课内安排',5,['日期','','课程','开始时间','结束时间','课内任务','科目','计划分钟','实际分钟','状态','','','强调']);
headers('日程事项',4,['日期','类型','开始时间','结束时间','事项','备注','来源','状态','强调','顺序']);
headers('学期总清单',4,['任务','状态','','安排日期','','','强调']);
for(const [name,range] of [['学习记录','A15:A599'],['课内安排','A6:A399'],['日程事项','A5:A199'],['学期总清单','D5:D7']])workbook.worksheets.getItem(name).getRange(range).format.numberFormat='yyyy-mm-dd';
for(const [name,range] of [['学习记录','L15:L599'],['课内安排','D6:E399'],['日程事项','C5:D199']])workbook.worksheets.getItem(name).getRange(range).format.numberFormat='hh:mm';
for(let index=0;index<13;index++){
    const sheet=workbook.worksheets.getItem(names[index+5]);
    sheet.getRange('B5:H5').values=[Array.from({length:7},(_,day)=>first+index*7+day)];
    sheet.getRange('B5:H5').format.numberFormat='mm/dd';
    sheet.getRange('B6:H6').values=[['周一','周二','周三','周四','周五','周六','周日']];
    sheet.getRange('B7:H7').values=[Array(7).fill('08:15起床')];
    sheet.getRange('A8:A38').values=Array.from({length:31},(_,slot)=>{const minutes=420+slot*30;return [String(Math.floor(minutes/60)).padStart(2,'0')+':'+String(minutes%60).padStart(2,'0')];});
    sheet.getRange('B49:M49').values=[['每周任务','科目','计划分钟','状态','原日期','备注','实际分钟','强调','安排日期','开始时间','时长分钟','顺序']];
    sheet.getRange('A67:A68').values=[['背单词'],['跑步打卡']];
    sheet.getRange('B67:H68').values=[Array(7).fill('☐ 待完成'),Array(7).fill('☐ 待完成')];
    sheet.getRange('J50:J62').format.numberFormat='yyyy-mm-dd';
    sheet.getRange('K50:K62').format.numberFormat='hh:mm';
    sheet.getRange('B5:H6').format.fill='#EAF4ED';
}
const calendar=workbook.worksheets.getItem('月历');
for(let week=0;week<14;week++){
    const row=4+week*2;
    calendar.getRange(`B${row}:H${row}`).values=[Array.from({length:7},(_,day)=>first-7+week*7+day)];
    calendar.getRange(`B${row}:H${row}`).format.numberFormat='mm/dd';
    calendar.getRange(`B${row}:H${row}`).format.fill='#EAF4ED';
}
workbook.recalculate();
console.log((await workbook.inspect({kind:'region',sheetId:'秋6周',range:'A5:H9',maxChars:1000})).ndjson);
console.log((await workbook.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
await(await SpreadsheetFile.exportXlsx(workbook)).save(output);
if(previewDir){
    await fs.mkdir(previewDir,{recursive:true});
    for(const [index,name] of names.entries()){
        const range=index>=5?'A5:H9':name==='学习记录'?'A14:G16':name==='课内安排'?'A5:F7':name==='日程事项'?'A4:F6':name==='学期总清单'?'A4:G6':'A4:H7';
        const preview=await workbook.render({sheetName:name,range,scale:1,format:'png'});
        await fs.writeFile(previewDir+'/'+index+'.png',new Uint8Array(await preview.arrayBuffer()));
    }
}
