"""Local Excel agenda reader and guarded task update bridge."""
import argparse, datetime as dt, json, pathlib, zipfile, xml.etree.ElementTree as ET, sys, subprocess, ctypes, os, shutil, hashlib

ROOT=pathlib.Path(__file__).resolve().parent.parent
SETTINGS_PATH=ROOT/'qingxu.local.json'
SETTINGS=json.loads(SETTINGS_PATH.read_text(encoding='utf-8-sig')) if SETTINGS_PATH.exists() else {}
BOOK=pathlib.Path(os.environ.get('AGENDA_BOOK') or SETTINGS.get('bookPath') or ROOT/'outputs'/'study-plan'/'计划.xlsx').expanduser()
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
BASE=dt.date(1899,12,30)
_DIGESTS={}

def file_digest(path,version,refresh=False):
    """内容摘要，与 mtime 版本一起用于判断工作簿是否被改动（低时间戳精度的磁盘上 mtime 不可靠）。"""
    key=(str(path),str(version))
    if not refresh and key in _DIGESTS: return _DIGESTS[key]
    try:
        hasher=hashlib.sha256()
        with open(path,'rb') as stream:
            for block in iter(lambda:stream.read(1<<20),b''): hasher.update(block)
    except (TypeError,ValueError): return ''  # 非文件系统路径（单元测试注入的内存工作簿）不参与校验
    if str(path.stat().st_mtime_ns)!=str(version): raise RuntimeError('读取期间 Excel 有更新，请稍后刷新。')
    _DIGESTS.clear();_DIGESTS[key]=hasher.hexdigest()
    return _DIGESTS[key]

def sheets(path):
    with zipfile.ZipFile(path) as z:
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            shared=[''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        book=ET.fromstring(z.read('xl/workbook.xml'))
        rel={e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        result={}
        for sh in book.find('s:sheets',NS):
            target=rel[sh.attrib['{'+NS['r']+'}id']]
            name=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            cells={}
            for c in ET.fromstring(z.read(name)).findall('.//s:sheetData/s:row/s:c',NS):
                typ=c.attrib.get('t'); v=c.find('s:v',NS)
                if typ=='inlineStr':
                    inline=c.find('s:is',NS); value=''.join(inline.itertext()) if inline is not None else ''
                elif v is None: value=''
                elif typ=='s':
                    try: index=int(v.text)
                    except (TypeError,ValueError): index=-1
                    value=shared[index] if 0<=index<len(shared) else ''
                elif typ in ('str','e'): value=v.text or ''
                else:
                    try: value=float(v.text)
                    except (TypeError,ValueError): value=v.text or ''
                cells[c.attrib['r']]=value
            result[sh.attrib['name']]=cells
        return result

def text(cells,cell):
    v=cells.get(cell,'')
    return '' if v is None else str(v).strip()

def minute_label(v):
    try: return str(int(round(float(v)*1440//1)))
    except (ValueError,TypeError): return ''

def time_label(value):
    try:
        minutes=round(float(value)*1440)
        return f'{minutes//60:02d}:{minutes%60:02d}'
    except (ValueError,TypeError): return ''

def add_minutes(label, minutes):
    try:
        hour, minute = map(int, label.split(':'))
        total = hour * 60 + minute + max(1, int(float(minutes or 60)))
        return f'{total//60:02d}:{total%60:02d}'
    except (ValueError,TypeError): return label

def course_index(data):
    inside=data.get('课内安排',{});result={}
    for row in range(6,400):
        serial=inside.get(f'A{row}')
        if isinstance(serial,(int,float)) and text(inside,f'C{row}'):
            result.setdefault(serial,[]).append(row)
    return result

def agenda(path=BOOK,date=None,data=None,linked_events=None,course_rows=None):
    day=date or dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    serial=(day-BASE).days; data=sheets(path) if data is None else data; day_index=(day-dt.date(2026,10,19)).days
    if linked_events is None:
        from excel_links import resolve,load_state
        links=resolve(data,load_state(path));linked_events=links['events'];data={name:dict(values) for name,values in data.items()}
        for patch in links['patches']:
            if patch['sheet']=='课内安排': data.setdefault('课内安排',{})[patch['cell']]=patch['value']
    version=str(path.stat().st_mtime_ns)
    result={'date':day.isoformat(),'title':day.strftime('%m月%d日')+' 星期'+'一二三四五六日'[day.weekday()], 'schedule':[],'tasks':[],'habits':[],'weekly':[],'weekSchedule':[],'pendingWeekly':[],'source':str(path),'error':'','version':version,'digest':file_digest(path,version)}
    week_name=(f'秋{6+day_index//7}周' if day_index//7<3 else f'冬{day_index//7-2}周') if 0<=day_index<91 else None
    if week_name and week_name in data:
        week=data[week_name]; col=chr(66+day.weekday())
        result['schedule'].append({'time':text(week,f'{col}7')[:5],'title':'起床','detail':'','kind':'fixed'})
        for row,label in [(67,'背单词'),(68,'跑步打卡')]:
            state=text(week,f'{col}{row}')
            if '已删除' in state: continue
            start_row=69 if row==67 else 72
            duration_row=73 if row==67 else 74
            start=time_label(week.get(f'{col}{start_row}'))
            planned=week.get(f'{col}{duration_row}',30) or 30
            routine={'id':f'habit:{week_name}:{col}:{row}','time':start+'—'+add_minutes(start,planned) if start else '','start':start,'planned':planned,'title':text(week,f'{col}{row+10}') or text(week,f'A{row}') or label,'order':week.get(f'{col}{row+8}',row-69),'done':'已完成' in state,'highlight':text(week,f'{col}{row+3}')=='强调','detail':start+'—'+add_minutes(start,planned) if start else ''}
            result['habits'].append(routine)
            if start:
                routine['kind']='task'
                result['schedule'].append(routine)
            else:
                result['tasks'].append(routine)
    inside=data.get('课内安排',{})
    for r in range(6,400):
        if inside.get(f'A{r}')!=serial: continue
        course=text(inside,f'C{r}')
        detail=text(inside,f'F{r}')
        if detail: detail='课内：'+detail
        result['schedule'].append({'id':f'class:{r}','time':time_label(inside.get(f'D{r}'))+'—'+time_label(inside.get(f'E{r}')),'title':course,'taskTitle':text(inside,f'F{r}'),'subject':text(inside,f'G{r}'),'planned':inside.get(f'H{r}',''),'actual':inside.get(f'I{r}',''),'done':'已完成' in text(inside,f'J{r}'),'highlight':text(inside,f'M{r}')=='强调','detail':detail,'kind':'class'})
    for event in SETTINGS.get('weeklyEvents',[]):
        if week_name and day.weekday()==event['weekday']:
            result['schedule'].append({'time':event['time'],'title':event['title'],'detail':'','kind':'fixed'})
    event_data=data.get('日程事项',{})
    for identity,event in linked_events.items():
        if event['date']!=day.isoformat(): continue
        r=int(identity.split(':')[1]);span=event['start']+'—'+event['end']
        title=event['title']; detail=text(event_data,f'F{r}') if '暂定' in title else ''
        existing=next((x for x in result['schedule'] if x['time']==span),None)
        if existing: existing['title']=title; existing['detail']=detail or existing['detail']
        else:
            duration=(int(event['end'][:2])*60+int(event['end'][3:])-int(event['start'][:2])*60-int(event['start'][3:])) if event['start'] and event['end'] else ''
            item={'id':identity,'time':span if event['start'] else '', 'title':title,'detail':detail,'kind':'task','start':event['start'],'planned':duration,'plannedDate':event['date'],'order':event_data.get(f'J{r}',0),'done':'已完成' in text(event_data,f'H{r}'),'highlight':text(event_data,f'I{r}')=='强调'}
            result['schedule' if event['start'] else 'tasks'].append(item)
    records=data.get('学习记录',{})
    for r in range(15,600):
        if records.get(f'A{r}')!=serial or not text(records,f'D{r}') or text(records,f'G{r}')=='— 取消': continue
        title=text(records,f'D{r}');subject=text(records,f'C{r}');planned=records.get(f'E{r}')
        detail=''
        start_time=time_label(records.get(f'L{r}'))
        item={'id':f'daily:{r}','title':title,'subject':subject,'planned':planned,'actual':records.get(f'F{r}',''),'detail':detail,'done':'已完成' in text(records,f'G{r}'),'highlight':text(records,f'K{r}')=='强调','start':start_time,'order':records.get(f'B{r}',r)}
        if start_time:
            item['time']=start_time+'—'+add_minutes(start_time,planned)
            item['kind']='task'
            result['schedule'].append(item)
        else:
            item['order']=records.get(f'B{r}',r)
            result['tasks'].append(item)
    daily_tasks=[item for item in result['tasks'] if item.get('id','').startswith(('daily:','habit:'))]
    daily_tasks.sort(key=lambda item:float(item.get('order') or 0))
    daily_iter=iter(daily_tasks)
    result['tasks']=[next(daily_iter) if item.get('id','').startswith(('daily:','habit:')) else item for item in result['tasks']]
    if week_name and week_name in data:
        week=data[week_name]
        for r in range(50,63):
            if week.get(f'F{r}')==serial and text(week,f'B{r}') and text(week,f'E{r}')!='— 本周无任务' and not isinstance(week.get(f'J{r}'),(int,float)):
                result['tasks'].append({'id':f'weekly:{week_name}:{r}','title':text(week,f'B{r}'),'detail':'每周清单','done':'已完成' in text(week,f'E{r}'),'highlight':text(week,f'I{r}')=='强调'})
    selected_week=week_name or ('秋6周' if day_index<0 else '冬10周')
    weekly=data.get(selected_week,{})
    start_serial=int(weekly.get('B5',46314));week_start=BASE+dt.timedelta(days=start_serial)
    result['weekName']=selected_week
    result['weekTitle']=selected_week+' · '+week_start.strftime('%m.%d')+' — '+(week_start+dt.timedelta(days=6)).strftime('%m.%d')
    result['weeklyDate']=week_start.isoformat()
    weekdays='一二三四五六日'
    course_rows=course_index(data) if course_rows is None else course_rows
    for offset in range(7):
        course_date=week_start+dt.timedelta(days=offset)
        course_serial=start_serial+offset
        for r in course_rows.get(course_serial,[]):
            course=text(inside,f'C{r}')
            if not course: continue
            result['weekSchedule'].append({'id':f'class:{r}','date':course_date.isoformat(),'weekday':'周'+weekdays[offset],'time':time_label(inside.get(f'D{r}'))+'—'+time_label(inside.get(f'E{r}')),'title':course})
    result['weekSchedule'].sort(key=lambda item:(item['date'],item['time'].split('—')[0]))
    for r in range(50,63):
        title=text(weekly,f'B{r}')
        if title:
            planned_day=weekly.get(f'J{r}')
            planned_start=time_label(weekly.get(f'K{r}'))
            duration=weekly.get(f'L{r}',60) or 60
            scheduled_date=(BASE+dt.timedelta(days=int(planned_day))).isoformat() if isinstance(planned_day,(int,float)) else ''
            item={'id':f'weekly:{selected_week}:{r}','title':title,'detail':('安排：'+(BASE+dt.timedelta(days=int(planned_day))).strftime('%m月%d日 星期'+'一二三四五六日'[(BASE+dt.timedelta(days=int(planned_day))).weekday()])) if isinstance(planned_day,(int,float)) else '本周未安排','done':'已完成' in text(weekly,f'E{r}'),'highlight':text(weekly,f'I{r}')=='强调','plannedDate':scheduled_date,'weekStart':week_start.isoformat(),'start':planned_start,'planned':duration,'order':weekly.get(f'M{r}',r)}
            result['weekly'].append(item)
            if planned_start and isinstance(planned_day,(int,float)) and int(planned_day)==serial:
                item['time']=planned_start+'—'+add_minutes(planned_start,duration)
                item['kind']='task'
                result['schedule'].append(item)
            elif isinstance(planned_day,(int,float)) and int(planned_day)==serial:
                result['tasks'].append(item)
                result['pendingWeekly'].append(item)
            elif not planned_day:
                result['pendingWeekly'].append(item)
    result['weekly'].sort(key=lambda item:float(item.get('order') or 0))
    weekly_tasks=[item for item in result['tasks'] if item.get('id','').startswith('weekly:')]
    weekly_tasks.sort(key=lambda item:float(item.get('order') or 0))
    weekly_iter=iter(weekly_tasks)
    result['tasks']=[next(weekly_iter) if item.get('id','').startswith('weekly:') else item for item in result['tasks']]
    sem=data.get('学期总清单',{})
    for r in range(5,8):
        if sem.get(f'D{r}')==serial: result['tasks'].append({'id':f'semester:{r}','title':text(sem,f'A{r}'),'detail':'学期任务','done':'已完成' in text(sem,f'B{r}'),'highlight':text(sem,f'G{r}')=='强调'})
    result['schedule'].sort(key=lambda x:(not bool(x.get('time')),x.get('time','').split('—')[0]))
    for group in ('schedule','tasks','habits','weekly'):
        for item in result[group]:
            parts=item.get('id','').split(':');status=''
            if len(parts)==2 and parts[0] in ('daily','class','semester','event'):
                sheet,column={'daily':('学习记录','G'),'class':('课内安排','J'),'semester':('学期总清单','B'),'event':('日程事项','H')}[parts[0]]
                status=text(data.get(sheet,{}),column+parts[1])
            elif parts[0]=='weekly' and len(parts)==3: status=text(data.get(parts[1],{}),'E'+parts[2])
            elif parts[0]=='habit' and len(parts)==4: status=text(data.get(parts[1],{}),parts[2]+parts[3])
            item['incomplete']='未完成' in status
    result['updated']=dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).strftime('%H:%M:%S')
    return result

def read_range(path,date):
    from excel_links import resolve,load_state,save_state
    version=str(path.stat().st_mtime_ns);digest=file_digest(path,version,refresh=True)
    data=sheets(path);baseline=load_state(path);links=resolve(data,baseline)
    for patch in links['patches']:
        if patch['sheet']=='课内安排': data.setdefault('课内安排',{})[patch['cell']]=patch['value']
    rows=course_index(data);anchor=dt.date.fromisoformat(date)
    parts=[agenda(path,anchor+dt.timedelta(days=offset),data,links['events'],rows) for offset in range(3)]
    if str(path.stat().st_mtime_ns)!=version or file_digest(path,version,refresh=True)!=digest:
        raise RuntimeError('读取期间 Excel 有更新，请稍后刷新。')
    if links['state']!=baseline: save_state(path,links['state'])
    return {'parts':parts,'version':version,'digest':digest}

def serve(path):
    for line in sys.stdin:
        try:
            request=json.loads(line)
            result=read_range(path,request['date'])
        except Exception as exc:
            result={'error':str(exc)}
        print(json.dumps(result,ensure_ascii=False),flush=True)

def check_available(path):
    if (path.parent/('~$'+path.name)).exists(): raise RuntimeError('请先保存并关闭 Excel／WPS 中的计划，再在小窗修改。')
    if os.name=='nt':
        kernel=ctypes.WinDLL('kernel32',use_last_error=True)
        kernel.CreateFileW.restype=ctypes.c_void_p
        kernel.CreateFileW.argtypes=[ctypes.c_wchar_p,ctypes.c_uint32,ctypes.c_uint32,ctypes.c_void_p,ctypes.c_uint32,ctypes.c_uint32,ctypes.c_void_p]
        handle=kernel.CreateFileW(str(path),0x80000000|0x40000000,0,None,3,0,None)
        if handle==ctypes.c_void_p(-1).value: raise RuntimeError('计划正在使用中，请先保存并关闭 Excel／WPS，再在小窗修改。')
        kernel.CloseHandle.argtypes=[ctypes.c_void_p];kernel.CloseHandle(handle)

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--date');parser.add_argument('--range',action='store_true');parser.add_argument('--serve',action='store_true');parser.add_argument('--write',action='store_true');parser.add_argument('--batch-write',action='store_true');parser.add_argument('--review-sync',action='store_true');parser.add_argument('--check',action='store_true');parser.add_argument('--stamp',action='store_true');parser.add_argument('--book',type=pathlib.Path,default=BOOK);args=parser.parse_args()
    try:
        if args.serve:
            serve(args.book)
        elif args.check:
            check_available(args.book); print(json.dumps({'ok':True}))
        elif args.stamp:
            version=str(args.book.stat().st_mtime_ns)
            print(json.dumps({'version':version,'digest':file_digest(args.book,version)}))
        elif args.range:
            print(json.dumps(read_range(args.book,args.date),ensure_ascii=False))
        elif args.review_sync:
            from sync_review import review_sync
            check_available(args.book)
            print(json.dumps(review_sync(args.book,json.load(sys.stdin)['ops']),ensure_ascii=False))
        elif args.write:
            check_available(args.book)
            node=os.environ.get('AGENDA_NODE') or shutil.which('node.exe') or shutil.which('node')
            if not node: raise RuntimeError('未找到 Node.js，请设置 AGENDA_NODE 环境变量。')
            p=subprocess.run([node,str(pathlib.Path(__file__).with_name('write_plan.mjs')),str(args.book),sys.executable],input=sys.stdin.read(),capture_output=True,text=True,encoding='utf-8')
            if p.returncode: raise RuntimeError(p.stderr.strip() or '保存失败，原计划未更改。')
            print(p.stdout.strip().splitlines()[-1])
        elif args.batch_write:
            check_available(args.book)
            node=os.environ.get('AGENDA_NODE') or shutil.which('node.exe') or shutil.which('node')
            if not node: raise RuntimeError('未找到 Node.js，请设置 AGENDA_NODE 环境变量。')
            p=subprocess.run([node,str(pathlib.Path(__file__).with_name('write_batch.mjs')),str(args.book),sys.executable],input=sys.stdin.read(),capture_output=True,text=True,encoding='utf-8')
            if p.returncode:
                details=p.stderr.strip()
                if '读取后已被修改' in details or '保存期间 Excel 有更新' in details:
                    print(json.dumps({'error':'Excel 已有更新，尚未写入。请核对两边修改后再合并。','code':'VERSION_CONFLICT'},ensure_ascii=False));sys.exit(0)
                message=next((line.partition('Error: ')[2] for line in details.splitlines() if line.startswith('Error: ')), '同步失败，请检查运行依赖及工作簿；原计划未更改。')
                raise RuntimeError(message)
            print(p.stdout.strip().splitlines()[-1])
        else: print(json.dumps(agenda(args.book,dt.date.fromisoformat(args.date) if args.date else None),ensure_ascii=False))
    except Exception as exc: print(json.dumps({'error':str(exc)},ensure_ascii=False))
