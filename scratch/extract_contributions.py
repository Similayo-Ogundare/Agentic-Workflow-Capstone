from pathlib import Path
from datetime import date
import json
import re
import openpyxl

root = Path(__file__).resolve().parents[1]
files = []
records = []
roster = None
for path in root.glob('Marvellous Church/1st Quarter/*/*.xlsx'):
    if path.name.startswith('~$'):
        continue  # Excel lock files are not contribution spreadsheets.
    month_name, year = path.parent.name.split()
    month = {'January': 1, 'February': 2, 'March': 3, 'April': 4}[month_name]
    day = int(re.search(r'(\d+)\.xlsx$', path.name).group(1))
    service_date = date(int(year), month, day)
    assert service_date.weekday() == 6, path.name
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    sheet = wb['Sheet1']
    rows = list(sheet.values)
    assert [str(v).strip() for v in rows[0]] == ['First Name', 'Last Name', 'Tithe', 'Offering']
    names = [' '.join(str(v).strip() for v in row[:2]) for row in rows[1:]]
    assert len(set(names)) == len(names)
    if roster is None:
        roster = names
    else:
        assert names == roster, path.name
    tithe_total = 0
    offering_total = 0
    for row_number, row in enumerate(rows[1:], 2):
        first, last, tithe, offering = row
        assert tithe is None or tithe == 'N/A' or isinstance(tithe, (int, float))
        assert offering is None or isinstance(offering, (int, float))
        if isinstance(tithe, (int, float)):
            assert tithe > 0 and day <= 7
            tithe_total += tithe
        if isinstance(offering, (int, float)):
            assert offering > 0
            offering_total += offering
        records.append({'date': service_date.isoformat(), 'month': service_date.replace(day=1).isoformat(),
                        'member': f'{str(first).strip()} {str(last).strip()}',
                        'tithe': tithe, 'offering': offering, 'file': path.name, 'row': row_number})
    files.append({'date': service_date.isoformat(), 'month': service_date.replace(day=1).isoformat(),
                  'file': path.name, 'rows': len(names), 'tithe': tithe_total, 'offering': offering_total})
    wb.close()
files.sort(key=lambda x: x['date'])
roster_order = {name: i for i, name in enumerate(roster)}
records.sort(key=lambda x: (x['date'], roster_order[x['member']]))
(root / 'scratch' / 'contributions.json').write_text(json.dumps({'files': files, 'records': records, 'roster': roster}, indent=2), encoding='utf-8')
print(f'Extracted {len(records)} records from {len(files)} files for {len(roster)} members.')
