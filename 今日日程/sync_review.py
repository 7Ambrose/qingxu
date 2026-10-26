import datetime as dt
import json
import hashlib

from read_plan import agenda


LABELS = {'title': '名称', 'done': '完成', 'incomplete': '明确未完成', 'highlight': '强调', 'start': '开始时间',
          'planned': '时长（分钟）', 'actual': '实际用时', 'subject': '科目',
          'order': '顺序', 'scheduleDate': '安排日期', 'delete': '删除'}


def display(fields):
    lines = []
    for key, value in fields.items():
        if key == 'deleteScope':
            continue
        if isinstance(value, bool):
            value = '是' if value else '否'
        if value is None or value == '':
            value = '未填写'
        lines.append(f'{LABELS.get(key, key)}：{value}')
    return '\n'.join(lines)


def review_sync(path, operations):
    version = str(path.stat().st_mtime_ns)
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    dates = {}
    entries = []
    for index, operation in enumerate(operations):
        date = operation['date']
        if date not in dates:
            dates[date] = agenda(path, dt.date.fromisoformat(date))
        data = dates[date]
        identity = operation['id']
        item = next((item for group in ('schedule', 'tasks', 'weekly')
                     for item in data[group] if item.get('id') == identity), None)
        is_new = identity.startswith(('local:', 'new'))
        fields = operation['fields']
        current = {}
        for key in fields:
            source = 'taskTitle' if key == 'title' and identity.startswith('class:') else key
            if key == 'scheduleDate':
                source = 'plannedDate'
            current[key] = item.get(source) if item else None
        if fields.get('delete') and item:
            current = {key: item.get(key) for key in ('title', 'done', 'start', 'planned')}
        entries.append({'index': index, 'id': identity, 'date': date,
                        'title': (item or {}).get('title') or fields.get('title') or identity,
                        'canApply': is_new or item is not None,
                        'excel': '尚未创建' if is_new else display(current) if item else '任务已删除或位置已改变，请保留 Excel 版本。',
                        'local': display(fields)})
    if str(path.stat().st_mtime_ns) != version or hashlib.sha256(path.read_bytes()).hexdigest() != digest:
        raise RuntimeError('核对期间 Excel 又有更新，请关闭 Excel 后重试。')
    return {'version': version, 'digest': digest, 'entries': entries}
