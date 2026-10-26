import datetime as dt
import json
import pathlib
import subprocess
import sys
from unittest.mock import patch

root = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(root / '今日日程'))
from read_plan import BOOK, agenda, sheets

dotnet = root / '.tools/dotnet/dotnet.exe'
assembly = root / 'Qingxu.Tests/bin/Release/net10.0-windows/Qingxu.Tests.dll'
result = subprocess.run([str(dotnet), str(assembly), '--dump-range', str(BOOK)], capture_output=True, text=True, encoding='utf-8', check=True)
native = json.loads(result.stdout)
data = sheets(BOOK)
count = 0
with patch('read_plan.sheets', return_value=data):
    for actual in native:
        expected = agenda(BOOK, dt.date.fromisoformat(actual['date']))
        for collection in ('schedule', 'tasks', 'weekly'):
            def normalize(items):
                return sorted((item.get('id', ''), item.get('title', ''), item.get('taskTitle', ''),
                               bool(item.get('done')), bool(item.get('highlight')),
                               item.get('start', ''), item.get('time', '') if collection != 'weekly' else '') for item in items)
            expected_items = normalize(expected[collection])
            actual_items = normalize(actual[collection])
            assert expected_items == actual_items, (actual['date'], collection, expected_items, actual_items)
            count += len(actual_items)
print(f'PASS: {len(native)} days / {count} entries match legacy reader identity, course/task names, completion, highlight and scheduled times.')
