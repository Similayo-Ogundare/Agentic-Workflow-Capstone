from pathlib import Path
from collections import Counter, defaultdict
from datetime import date
import calendar
import hashlib
import json
import re
import openpyxl

source = Path('Marvellous Church/1st Quarter')
month_numbers = {'January': 1, 'February': 2, 'March': 3, 'April': 4}
rosters = []
all_records = []
source_stats = []
abnormalities = []
for path in sorted(source.rglob('*.xlsx')):
    month_label, year_str = path.parent.name.split()
    month = month_numbers[month_label]
    day = int(re.search(r'(\d+)\.xlsx$', path.name).group(1))
    service_date = date(int(year_str), month, day)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    if rows[0] != ('First Name', 'Last Name', 'Tithe ', 'Offering '):
        abnormalities.append({'header': str(path), 'value': rows[0]})
    roster = []
    cells = defaultdict(Counter)
    filetotals = defaultdict(int)
    for n, row in enumerate(rows[1:], start=2):
        name = f'{row[0]} {row[1]}'
        roster.append(name)
        record = {'date': service_date.isoformat(), 'month': month_label, 'name': name, 'source': str(path), 'row': n}
        for kind, raw in zip(('tithe', 'offering'), row[2:4]):
            if raw is None or raw == '':
                status = 'blank'
                amount = 0
            elif isinstance(raw, (int, float)):
                status = 'positive' if raw > 0 else 'zero' if raw == 0 else 'negative'
                amount = raw
            elif str(raw).strip().upper() == 'N/A':
                status = 'N/A'
                amount = 0
            else:
                status = 'unexpected'
                amount = 0
            cells[kind][status] += 1
            filetotals[kind] += amount
            record[kind] = {'raw': raw, 'status': status, 'amount': amount}
            if status in ('negative', 'unexpected'):
                abnormalities.append(record)
            if kind == 'tithe' and ((day <= 7 and status == 'N/A') or (day > 7 and status != 'N/A')):
                abnormalities.append({'tithe_schedule': record})
        all_records.append(record)
    rosters.append(roster)
    duplicates = [name for name, count in Counter(roster).items() if count > 1]
    if duplicates:
        abnormalities.append({'duplicates': duplicates, 'file': str(path)})
    if service_date.weekday() != 6:
        abnormalities.append({'non_sunday': str(path)})
    source_stats.append({'file': str(path), 'date': service_date.isoformat(), 'members': len(roster), 'cells': {k: dict(v) for k,v in cells.items()}, 'totals': dict(filetotals), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
    wb.close()

roster_consistent = all(roster == rosters[0] for roster in rosters)
monthly = {}
for month_label, month in month_numbers.items():
    records = [record for record in all_records if record['month'] == month_label]
    present_dates = sorted({record['date'] for record in records})
    expected_dates = [date(2025, month, day).isoformat() for day in range(1,calendar.monthrange(2025,month)[1]+1) if date(2025,month,day).weekday() == 6]
    totals = {kind: sum(record[kind]['amount'] for record in records) for kind in ('tithe','offering')}
    statuses = {kind: dict(Counter(record[kind]['status'] for record in records)) for kind in ('tithe','offering')}
    member_stats = []
    for name in rosters[0]:
        member_records = [record for record in records if record['name'] == name]
        counts = {}
        for kind in ('tithe', 'offering'):
            relevant = [record for record in member_records if kind == 'offering' or date.fromisoformat(record['date']).day <= 7]
            counts[kind] = {'total': sum(record[kind]['amount'] for record in member_records), 'paid': sum(record[kind]['status'] == 'positive' for record in relevant), 'missed': sum(record[kind]['status'] in ('blank','zero') for record in relevant), 'opportunities': len(relevant)}
        counts['both_paid'] = sum(record['tithe']['status'] == 'positive' and record['offering']['status'] == 'positive' for record in member_records)
        counts['both_missed_first_sunday'] = sum(date.fromisoformat(record['date']).day <= 7 and record['tithe']['status'] in ('blank','zero') and record['offering']['status'] in ('blank','zero') for record in member_records)
        counts['total'] = counts['tithe']['total']+counts['offering']['total']
        member_stats.append({'name': name, **counts})
    monthly[month_label] = {'dates': present_dates, 'expected_dates':expected_dates, 'complete':present_dates == expected_dates, 'totals':totals, 'statuses':statuses, 'members':member_stats}

quarter = {'totals':{}, 'statuses':{}, 'counts':{}, 'members':[]}
quarter_records = [record for record in all_records if record['month'] in ('January','February','March')]
for kind in ('tithe','offering'):
    quarter['totals'][kind] = sum(record[kind]['amount'] for record in quarter_records)
    quarter['statuses'][kind] = dict(Counter(record[kind]['status'] for record in quarter_records))
    relevant = [record for record in quarter_records if kind == 'offering' or date.fromisoformat(record['date']).day <= 7]
    quarter['counts'][kind] = {'paid':sum(record[kind]['status']=='positive' for record in relevant), 'missed':sum(record[kind]['status'] in ('blank','zero') for record in relevant),'opportunities':len(relevant)}
for name in rosters[0]:
    member = {'name':name}
    for kind in ('tithe','offering'):
        stats = [m for month_label in ('January','February','March') for m in monthly[month_label]['members'] if m['name'] == name]
        member[kind] = {field:sum(s[kind][field] for s in stats) for field in ('total','paid','missed','opportunities')}
    member['both_paid'] = sum(s['both_paid'] for s in stats)
    member['total'] = member['tithe']['total']+member['offering']['total']
    quarter['members'].append(member)
quarter['totals']['combined'] = sum(quarter['totals'].values())
result = {'files':source_stats, 'roster_consistent':roster_consistent, 'roster':rosters[0], 'abnormalities':abnormalities, 'monthly':monthly, 'quarter':quarter}
Path('scratch/independent_audit/results.json').write_text(json.dumps(result, indent=2), encoding='utf-8')

def flat_metrics(records):
    first_sundays = [r for r in records if date.fromisoformat(r['date']).day <= 7]
    paid_both = sum(r['tithe']['status'] == 'positive' and r['offering']['status'] == 'positive' for r in first_sundays)
    any_sundays = sum(r['tithe']['status'] == 'positive' or r['offering']['status'] == 'positive' for r in records)
    return {
        'tithe': sum(r['tithe']['amount'] for r in records),
        'offering': sum(r['offering']['amount'] for r in records),
        'total': sum(r[k]['amount'] for r in records for k in ('tithe','offering')),
        'tithe_paid': sum(r['tithe']['status'] == 'positive' for r in first_sundays),
        'tithe_missed': sum(r['tithe']['status'] in ('blank','zero') for r in first_sundays),
        'offering_paid': sum(r['offering']['status'] == 'positive' for r in records),
        'offering_missed': sum(r['offering']['status'] in ('blank','zero') for r in records),
        'both_paid': paid_both,
        'both_incomplete': len(first_sundays) - paid_both,
        'any_sundays': any_sundays,
        'no_sundays': len(records)-any_sundays,
    }

root_schema = {
    'roster': rosters[0],
    'months': {f'{month} 2025': flat_metrics([r for r in all_records if r['month']==month]) for month in month_numbers},
    'members': {name: {**{f'{month} 2025':flat_metrics([r for r in all_records if r['month']==month and r['name']==name]) for month in month_numbers}, 'Q1':flat_metrics([r for r in quarter_records if r['name']==name])} for name in rosters[0]},
    'quarter':flat_metrics(quarter_records),
    'file_count':len(source_stats),
    'row_count':len(all_records),
    'all_calendars_complete':all(m['complete'] for m in monthly.values()),
    'roster_consistent':roster_consistent,
    'abnormalities':abnormalities,
}
Path('scratch/independent_audit/audit.json').write_text(json.dumps(root_schema, indent=2), encoding='utf-8')
print(json.dumps({'roster_consistent':roster_consistent,'roster':rosters[0],'abnormalities':abnormalities,'monthly':{month:{k:v for k,v in stats.items() if k!='members'} for month,stats in monthly.items()}, 'quarter':quarter}, indent=2))
