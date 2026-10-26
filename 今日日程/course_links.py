import re

from excel_links import WEEK, SPAN, clock, identity_row, location_key, merge_changes, minutes, serial, stamp, text, same_value


def resolve_courses(data, baseline):
    records = data.get('课内安排', {})
    old_classes = baseline.get('classes', {})
    rows = {row for row in (identity_row(key, 6, 399) for key in old_classes) if row is not None}
    rows.update(row for row in range(6, 400) if isinstance(records.get(f'A{row}'), (int, float)) and text(records.get(f'C{row}')))
    originals = {}
    for row in rows:
        if not isinstance(records.get(f'A{row}'), (int, float)) or not text(records.get(f'C{row}')):
            originals[row] = None
            continue
        originals[row] = {'date': stamp(records[f'A{row}']), 'start': clock(records.get(f'D{row}')),
                          'end': clock(records.get(f'E{row}')), 'title': text(records[f'C{row}']),
                          'taskTitle': text(records.get(f'F{row}'))}
    names = {value['title'] for value in originals.values() if value}
    prior_locations = {location_key(item['sheet'], item['cell']) for value in old_classes.values() for item in value.get('locations', [])}
    candidates = []
    for name, cells in data.items():
        if not WEEK.match(name):
            continue
        for column, row in ((column, row) for column in 'BCDEFGH' for row in range(8, 48)):
            cell = column + str(row)
            value = cells.get(cell)
            if not text(value):
                continue
            lines = text(value).splitlines()
            explicit = SPAN.match(lines[0].strip())
            title = explicit[3].strip() if explicit else lines[0].strip()
            if title not in names and location_key(name, cell) not in prior_locations:
                continue
            date_serial = cells.get(column + '5')
            start = explicit[1] if explicit else text(cells.get('A' + str(row)))[:5]
            if not isinstance(date_serial, (int, float)) or not re.fullmatch(r'\d{2}:\d{2}', start):
                continue
            task_title = next((re.sub(r'^课内[：:]\s*', '', line).strip() for line in lines[1:] if re.match(r'^课内[：:]', line)), '')
            candidates.append({'sheet': name, 'cell': cell, 'lines': lines, 'date': stamp(date_serial),
                               'start': start.zfill(5), 'end': explicit[2] if explicit else '', 'title': title, 'taskTitle': task_title})
    by_location, by_exact, by_title = {}, {}, {}
    for item in candidates:
        by_location[location_key(item['sheet'], item['cell'])] = item
        by_exact.setdefault((item['date'], item['start'], item['title']), []).append(item)
        by_title.setdefault(item['title'], []).append(item)
    matched, used = {}, set()
    for row, record in originals.items():
        old = old_classes.get(f'class:{row}', {})
        keys = {location_key(item['sheet'], item['cell']) for item in old.get('locations', [])}
        exact = [by_location[key] for key in keys if key in by_location]
        if not exact and record:
            exact = by_exact.get((record['date'], record['start'], record['title']), [])
        if len(exact) == 1 and location_key(exact[0]['sheet'], exact[0]['cell']) not in used:
            matched[row] = exact[0]
            used.add(location_key(exact[0]['sheet'], exact[0]['cell']))
    for row, record in originals.items():
        if row in matched or not record:
            continue
        remaining = [item for item in by_title.get(record['title'], []) if abs(serial(item['date']) - serial(record['date'])) <= 6 and location_key(item['sheet'], item['cell']) not in used]
        if len(remaining) == 1:
            matched[row] = remaining[0]
            used.add(location_key(remaining[0]['sheet'], remaining[0]['cell']))
    patches, state = {}, {}
    dates, slots_by_sheet = {}, {}
    for name, cells in data.items():
        if not WEEK.match(name):
            continue
        for column in 'BCDEFGH':
            if isinstance(cells.get(column + '5'), (int, float)):
                dates.setdefault(cells[column + '5'], (name, column))
        slots_by_sheet[name] = [(minutes(text(cells.get(f'A{slot}'))[:5]), slot) for slot in range(8, 48) if re.match(r'^\d{2}:\d{2}', text(cells.get(f'A{slot}')))]

    def get(sheet, cell):
        return patches.get((sheet, cell), data.get(sheet, {}).get(cell))

    def put(sheet, cell, value):
        if not same_value(get(sheet, cell), value):
            patches[sheet, cell] = value

    for row, record in sorted(originals.items()):
        identity = f'class:{row}'
        old = old_classes.get(identity, {})
        candidate = matched.get(row)
        sources = {'record': record}
        if candidate:
            duration_record = record or old.get('effective') or {}
            duration = minutes(duration_record['end']) - minutes(duration_record['start']) if duration_record.get('start') and duration_record.get('end') else 95
            end = candidate['end'] or f"{(minutes(candidate['start']) + duration) // 60:02d}:{(minutes(candidate['start']) + duration) % 60:02d}"
            sources['week'] = {key: candidate[key] for key in ('date', 'start', 'title', 'taskTitle')}
            sources['week']['end'] = end
        elif 'week' in old.get('sources', {}):
            sources['week'] = None
        effective = merge_changes((record or old.get('effective') or {}).get('title', identity), old, sources, ('date', 'start', 'end', 'title', 'taskTitle')) if old else record
        if not old and candidate and record:
            effective = dict(record)
            for key in ('date', 'start', 'end'):
                effective[key] = sources['week'][key]
        state[identity] = {'sources': sources, 'effective': effective,
                           'locations': [{'sheet': candidate['sheet'], 'cell': candidate['cell']}] if candidate else []}
        if not candidate and not old:
            continue
        if not effective:
            for column in 'ABCDEFGHIJKLM':
                put('课内安排', f'{column}{row}', None)
            if candidate:
                put(candidate['sheet'], candidate['cell'], None)
            continue
        for column, value in [('A', serial(effective['date'])), ('C', effective['title']), ('D', minutes(effective['start']) / 1440),
                              ('E', minutes(effective['end']) / 1440), ('F', effective['taskTitle'])]:
            put('课内安排', f'{column}{row}', value)
        if not candidate and not old.get('locations'):
            continue
        destination = dates.get(serial(effective['date']))
        if not destination:
            continue
        target_sheet, target_column = destination
        cells = data[target_sheet]
        start_minute = minutes(effective['start'])
        slots = [(abs(minute - start_minute), slot) for minute, slot in slots_by_sheet[target_sheet]]
        if not slots:
            continue
        slot = min(slots)[1]
        if candidate and candidate['sheet'] == target_sheet and candidate['start'] == effective['start']:
            slot = int(re.sub(r'\D', '', candidate['cell']))
        target_cell = target_column + str(slot)
        old_key = location_key(candidate['sheet'], candidate['cell']) if candidate else None
        if old_key != location_key(target_sheet, target_cell) and text(get(target_sheet, target_cell)) not in ('', '课间'):
            raise ValueError(f'“{effective["title"]}”在{target_sheet}!{target_cell}的新时间格已有其他内容，未自动覆盖。')
        if candidate and old_key != location_key(target_sheet, target_cell):
            put(candidate['sheet'], candidate['cell'], None)
        first = effective['title']
        if text(cells.get('A' + str(slot)))[:5] != effective['start']:
            first = effective['start'] + '—' + effective['end'] + ' ' + first
        extra = [line for line in (candidate['lines'][1:] if candidate else []) if not re.match(r'^课内[：:]', line)]
        lines = [first] + extra + (['课内：' + effective['taskTitle']] if effective['taskTitle'] else [])
        put(target_sheet, target_cell, '\n'.join(lines))
    return {'state': state, 'patches': [{'sheet': sheet, 'cell': cell, 'value': value} for (sheet, cell), value in patches.items()
                                       if data.get(sheet, {}).get(cell) != value]}
