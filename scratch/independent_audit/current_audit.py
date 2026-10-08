from pathlib import Path
from datetime import date
from collections import Counter
import calendar
import hashlib
import json
import re
import openpyxl

source = Path('Marvellous Church/1st Quarter')
prior = json.loads(Path('scratch/independent_audit/results.json').read_text(encoding='utf-8'))
prior_hashes = {f['file']:f['sha256'] for f in prior['files']}
month_numbers = {'January':1,'February':2,'March':3,'April':4}
files = []
records = []
rosters = []
anomalies = []
for path in sorted(source.rglob('*.xlsx')):
    month_label, year_str = path.parent.name.split()
    day = int(re.search(r'(\d+)\.xlsx$',path.name).group(1))
    service_date = date(int(year_str),month_numbers[month_label],day)
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=False)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    roster=[]
    for i,row in enumerate(rows[1:],start=2):
        first,last,tithe,offering = row
        name = f'{first} {last}'
        roster.append(name)
        record={'first':first,'last':last,'name':name,'month':month_label,'date':service_date.isoformat(),'file':str(path),'row':i}
        for kind,raw in [('tithe',tithe),('offering',offering)]:
            if isinstance(raw,(float,int)):
                status='positive' if raw>0 else 'zero' if raw==0 else 'negative'
                amount=raw
            elif raw is None or raw=='':
                status='blank'; amount=0
            elif str(raw).strip().upper()=='N/A':
                status='N/A'; amount=0
            else:
                status='unexpected'; amount=0
            record[kind]=amount
            record[kind+'_paid']=int(status=='positive')
            record[kind+'_status']=status
            if status in ('negative','unexpected'):
                anomalies.append({'file':str(path),'row':i,'kind':kind,'raw':raw})
        records.append(record)
    rosters.append(roster)
    files.append({'file':str(path),'date':service_date.isoformat(),'members':len(roster),'sha256':digest,'matches_prior':digest==prior_hashes.get(str(path))})
    workbook.close()

names=rosters[0]
def metrics(rs):
    return {'tithe':sum(r['tithe'] for r in rs),'offering':sum(r['offering'] for r in rs),'tithe_paid':sum(r['tithe_paid'] for r in rs),'offering_paid':sum(r['offering_paid'] for r in rs)}

monthly={f'{month} 2025':metrics([r for r in records if r['month']==month]) for month in month_numbers}
members={name:{**{f'{month} 2025':metrics([r for r in records if r['name']==name and r['month']==month]) for month in month_numbers},'final':metrics([r for r in records if r['name']==name])} for name in names}
rankings={}
rankings_shared={}
exact_ties={}
for kind in ('tithe','offering'):
    ranked=sorted(names,key=lambda name:(-members[name]['final'][kind+'_paid'],-members[name]['final'][kind],name.casefold()))
    rankings[kind]=[]
    rankings_shared[kind]=[]
    for rank,name in enumerate(ranked,start=1):
        members[name]['final'][kind+'_rank']=rank
        rankings[kind].append({'name':name,'paid':members[name]['final'][kind+'_paid'],'amount':members[name]['final'][kind],'rank':rank})
        own=(members[name]['final'][kind+'_paid'],members[name]['final'][kind])
        shared_rank=1+sum((members[other]['final'][kind+'_paid'],members[other]['final'][kind])>own for other in names)
        members[name]['final'][kind+'_rank_shared']=shared_rank
        rankings_shared[kind].append({'name':name,'paid':own[0],'amount':own[1],'rank':shared_rank})
    grouped={}
    for name in names:
        grouped.setdefault((members[name]['final'][kind+'_paid'],members[name]['final'][kind]),[]).append(name)
    exact_ties[kind]=[{'paid':counts,'amount':amount,'names':tie_names} for (counts,amount),tie_names in grouped.items() if len(tie_names)>1]

calendars={}
for month,month_number in month_numbers.items():
    actual=sorted({r['date'] for r in records if r['month']==month})
    expected=[date(2025,month_number,d).isoformat() for d in range(1,calendar.monthrange(2025,month_number)[1]+1) if date(2025,month_number,d).weekday()==6]
    calendars[month]={'dates':actual,'expected':expected,'complete':actual==expected}

result={'files':files,'file_count':len(files),'row_count':len(records),'roster':names,'roster_consistent':all(r==names for r in rosters),'all_fingerprints_match_prior':all(f['matches_prior'] for f in files),'calendars':calendars,'statuses':{k:dict(Counter(r[k+'_status'] for r in records)) for k in ('tithe','offering')},'monthly':monthly,'members':members,'final':metrics(records),'rankings':rankings,'ranks_shared':rankings_shared,'exact_ties':exact_ties,'anomalies':anomalies,'exact_tie_rule_for_verification':'rankings uses sequential ordinal positions with alphabetical fallback for exact ties; ranks_shared uses competition ranking with equal ranks for exact count and amount ties and skipped following positions.'}
Path('scratch/independent_audit/current_audit.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ('files','members','roster')},indent=2))
