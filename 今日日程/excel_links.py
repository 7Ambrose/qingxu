import datetime as dt
import json
import pathlib
import re
import sys
import math

EPOCH = dt.date(1899, 12, 30)
SPAN = re.compile(r'^(\d{1,2}:\d{2})\s*[—–-]\s*(\d{1,2}:\d{2})\s+(.+)$')
WEEK = re.compile(r'^(秋|冬)\d+周$')


def text(value):
    return '' if value is None else str(value).strip()


def same_value(left, right):
    if left in (None, '') and right in (None, ''):
        return True
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(left, right, rel_tol=0, abs_tol=1e-10)
    return left == right


def clock(value):
    if not isinstance(value, (int, float)):
        return ''
    minute = round(value * 1440)
    return f'{minute // 60:02d}:{minute % 60:02d}'


def serial(date):
    return (dt.date.fromisoformat(date) - EPOCH).days


def minutes(label):
    hour, minute = map(int, label.split(':'))
    return hour * 60 + minute


def stamp(value):
    return (EPOCH + dt.timedelta(days=int(value))).isoformat()


def location_key(sheet, cell):
    return sheet + '!' + cell


def identity_row(identity, low=5, high=599):
    """从 event:12 / class:8 这类标识里取行号，格式或范围不合法时返回 None（不再直接抛异常）。"""
    parts = str(identity).split(':')
    if len(parts) != 2 or not parts[1].isdigit():
        return None
    row = int(parts[1])
    return row if low <= row <= high else None


def merge_changes(title, old, sources, fields):
    removed = [name for name, value in sources.items() if value is None and old.get('sources', {}).get(name) is not None]
    changed = [name for name, value in sources.items() if value is not None and value != old.get('sources', {}).get(name) and name in old.get('sources', {})]
    if removed:
        if changed:
            raise ValueError(f'Excel 中“{title}”同时被删除和编辑，未自动覆盖。')
        return None
    effective = dict(old.get('effective') or next((value for value in sources.values() if value), {}))
    for field in fields:
        updates = [value.get(field) for name, value in sources.items() if value is not None and name in old.get('sources', {}) and value.get(field) != (old['sources'].get(name) or {}).get(field)]
        distinct = {json.dumps(value, sort_keys=True) for value in updates}
        if len(distinct) > 1:
            raise ValueError(f'Excel 中“{title}”的同一个字段在多处被改成了不同内容，未自动覆盖。')
        if updates:
            effective[field] = updates[0]
    return effective


def occurrences(data):
    result = []
    for name, cells in data.items():
        if not WEEK.match(name) and name != '月历':
            continue
        for cell, value in cells.items():
            if not isinstance(value, str) or ':' not in value:
                continue
            match = re.fullmatch(r'([B-H])(\d+)', cell)
            if not match or not isinstance(value, str):
                continue
            column, row = match[1], int(match[2])
            day = cells.get(column + '5') if WEEK.match(name) else cells.get(column + str(row - 1))
            if not isinstance(day, (int, float)) or day < 46313 or day > 46404:
                continue
            if WEEK.match(name) and not 8 <= row <= 47:
                continue
            for line in value.splitlines():
                timed = SPAN.match(line.strip())
                if timed:
                    result.append({'sheet': name, 'cell': cell, 'line': line,
                                   'value': {'date': stamp(day), 'start': timed[1].zfill(5),
                                             'end': timed[2].zfill(5), 'title': timed[3].strip()}})
    return result


def load_state(path):
    state_path = pathlib.Path(str(path) + '.agenda-links.json')
    return json.loads(state_path.read_text(encoding='utf-8-sig')) if state_path.exists() else {'schema': 1, 'events': {}}


def save_state(path, state):
    state_path = pathlib.Path(str(path) + '.agenda-links.json')
    temporary = pathlib.Path(str(state_path) + '.pending')
    temporary.write_text(json.dumps(state, ensure_ascii=False), encoding='utf-8')
    temporary.replace(state_path)


def resolve(data, baseline=None):
    baseline = baseline or {'schema': 1, 'events': {}}
    from course_links import resolve_courses
    original_data = data
    courses = resolve_courses(data, baseline)
    data = {name: dict(values) for name, values in data.items()}
    for patch in courses['patches']:
        data.setdefault(patch['sheet'], {})[patch['cell']] = patch['value']
    records = data.get('日程事项', {})
    found = occurrences(data)
    old_events = baseline.get('events', {})
    rows = {row for row in (identity_row(key) for key in old_events) if row is not None}
    rows.update(row for row in range(5, 200) if text(records.get(f'E{row}')))
    known_titles = {text(records.get(f'E{row}')) for row in rows}
    events, state, patches = {}, {'schema': 1, 'events': {}, 'classes': courses['state']}, {(patch['sheet'], patch['cell']): patch['value'] for patch in courses['patches']}

    def get(sheet, cell):
        return patches.get((sheet, cell), data.get(sheet, {}).get(cell))

    def put(sheet, cell, value):
        if not same_value(get(sheet, cell), value):
            patches[sheet, cell] = value

    for row in sorted(rows):
        identity = f'event:{row}'
        old = old_events.get(identity, {})
        record = None
        if isinstance(records.get(f'A{row}'), (int, float)) and text(records.get(f'E{row}')):
            record = {'date': stamp(records[f'A{row}']), 'start': clock(records.get(f'C{row}')),
                      'end': clock(records.get(f'D{row}')), 'title': text(records[f'E{row}'])}
        title = (record or old.get('effective') or {}).get('title', '')
        old_locations = {location_key(item['sheet'], item['cell']) for item in old.get('locations', [])}
        matched = [item for item in found if item['value']['title'] == title or
                   (location_key(item['sheet'], item['cell']) in old_locations and
                    item['value']['title'] not in known_titles and
                    sum(other['sheet'] == item['sheet'] and other['cell'] == item['cell'] for other in found) == 1)]
        sources = {'record': record}
        for source in ('week', 'calendar'):
            candidates = [item for item in matched if (item['sheet'] != '月历') == (source == 'week')]
            unique = {json.dumps(item['value'], sort_keys=True) for item in candidates}
            if len(unique) == 1:
                sources[source] = candidates[0]['value']
            elif candidates:
                previous = old.get('sources', {}).get(source)
                changed = [item['value'] for item in candidates if item['value'] != previous]
                if previous and len({json.dumps(value, sort_keys=True) for value in changed}) == 1:
                    sources[source] = changed[0]
                else:
                    raise ValueError(f'Excel 中“{title}”有多个不同的{source}安排，未自动覆盖。')
            elif not candidates and source in old.get('sources', {}):
                sources[source] = None
        if old:
            effective = merge_changes(title or identity, old, sources, ('date', 'start', 'end', 'title'))
        else:
            effective = sources.get('week') or record or sources.get('calendar')
        state['events'][identity] = {'sources': sources, 'effective': effective, 'locations': matched}
        if effective:
            events[identity] = dict(effective)
            put('日程事项', f'A{row}', serial(effective['date']))
            put('日程事项', f'C{row}', sum(int(part) * factor for part, factor in zip(effective['start'].split(':'), (60, 1))) / 1440 if effective['start'] else None)
            put('日程事项', f'D{row}', sum(int(part) * factor for part, factor in zip(effective['end'].split(':'), (60, 1))) / 1440 if effective['end'] else None)
            put('日程事项', f'E{row}', effective['title'])
        else:
            for column in 'ABCDEFGHIJ':
                put('日程事项', f'{column}{row}', None)
        for item in matched:
            remaining = [line for line in text(get(item['sheet'], item['cell'])).splitlines() if line.strip() != item['line'].strip()]
            put(item['sheet'], item['cell'], '\n'.join(remaining) or None)
        if not effective or not effective['start'] or not effective['end']:
            continue
        line = effective['start'] + '—' + effective['end'] + ' ' + effective['title']
        date_serial = serial(effective['date'])
        for name, cells in data.items():
            if WEEK.match(name):
                columns = [column for column in 'BCDEFGH' if cells.get(column + '5') == date_serial]
                if not columns:
                    continue
                prior = next((item for item in matched if item['sheet'] == name), None)
                old_row = ''.join(character for character in prior['cell'] if character.isdigit()) if prior else ''
                slots = [(abs(minutes(text(cells.get(f'A{slot}'))[:5]) - minutes(effective['start'])), slot)
                         for slot in range(8, 48) if re.match(r'^\d{2}:\d{2}', text(cells.get(f'A{slot}')))]
                row_number = int(old_row) if old_row else min(slots)[1] if slots else 8
                targets = [columns[0] + str(row_number)]
                existing_course = text(get(name, targets[0])).splitlines()
                if existing_course and (existing_course[0] == effective['title'] or '小测' in existing_course[0] and '小测' in effective['title']):
                    continue
            elif name == '月历':
                headers = [cell for cell, value in cells.items() if value == date_serial and re.match(r'^[B-H]\d+$', cell)]
                if not headers:
                    continue
                targets = [header[0] + str(int(header[1:]) + 1) for header in headers]
            else:
                continue
            for target in targets:
                existing = text(get(name, target))
                put(name, target, existing + '\n' + line if existing else line)
    return {'events': events, 'state': state,
            'patches': [{'sheet': sheet, 'cell': cell, 'value': value} for (sheet, cell), value in patches.items()
                        if original_data.get(sheet, {}).get(cell) != value]}


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] in ('--resolve', '--snapshot'):
        try:
            body = json.load(sys.stdin)
            result = resolve(body['data'], body.get('baseline'))
            if sys.argv[1] == '--snapshot':
                for identity, event in body.get('previous', {}).get('events', {}).items():
                    if identity not in result['state']['events'] and event.get('effective') is None:
                        result['state']['events'][identity] = {'sources': {'record': None, 'week': None, 'calendar': None}, 'effective': None, 'locations': []}
                result = result['state']
            print(json.dumps(result, ensure_ascii=False))
        except (ValueError, KeyError) as error:
            print(json.dumps({'error': str(error)}, ensure_ascii=False))
    else:
        from read_plan import sheets
        path = pathlib.Path(sys.argv[1])
        data = sheets(path)
        selected = {name: values for name, values in data.items() if WEEK.match(name) or name in ('月历', '日程事项', '课内安排')}
        print(json.dumps({'data': selected, 'baseline': load_state(path)}, ensure_ascii=False))
