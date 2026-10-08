from pathlib import Path
from collections import defaultdict
import hashlib
import json
import openpyxl

BASE=Path('scratch/independent_audit')
source_audit=json.loads((BASE/'current_audit.json').read_text(encoding='utf-8'))
discrepancies=[]
checks=0
members=[]
for label in ('January','February','March','April','Final'):
    path=Path(f'outputs/{label}_Summary.xlsx')
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    rows=list(wb['Summary'].iter_rows(values_only=True))[6:25]
    names=[f'{r[0]} {r[1]}' for r in rows]
    if names!=source_audit['roster']:
        discrepancies.append({'workbook':label,'type':'roster mismatch','actual':names})
    for row_number,row in enumerate(rows,start=7):
        name=f'{row[0]} {row[1]}'
        audit=source_audit['members'][name]['final' if label=='Final' else f'{label} 2025']
        fields=((2,'tithe'),(3,'offering'),(4,'tithe_paid'),(6,'offering_paid'),(5,'tithe_rank_combined'),(7,'offering_rank_combined')) if label=='Final' else ((2,'tithe'),(3,'offering'),(4,'tithe_paid'),(5,'offering_paid'))
        for column,key in fields:
            checks+=1
            if row[column]!=audit[key]:
                discrepancies.append({'workbook':label,'row':row_number,'key':key,'actual':row[column],'expected':audit[key]})
        if label=='Final':
            members.append({'name':name,'row':row_number,'tithe_rank':row[5],'offering_rank':row[7],'average':(row[5]+row[7])/2})
    wb.close()
for member in members:
    member['overall_rank']=1+sum(other['average']<member['average'] for other in members)
groups=defaultdict(list)
for member in members:
    groups[member['average']].append(member['name'])
ties=[{'average':average,'overall_rank':1+sum(m['average']<average for m in members),'members':names} for average,names in sorted(groups.items()) if len(names)>1]
result={'instructions_sha256':hashlib.sha256(Path('INSTRUCTIONS.md').read_bytes()).hexdigest(),'source_files':source_audit['file_count'],'source_hashes_unchanged':source_audit['all_fingerprints_match_prior'],'source_calendars_complete':all(m['complete'] for m in source_audit['calendars'].values()),'existing_output_metric_checks':checks,'existing_output_discrepancies':discrepancies,'members':members,'ties':ties,'rank_rule':'Average=(tithe consistency rank+offering consistency rank)/2; Overall Rank=1+number of members with a strictly lower average (ascending competition/RANK.EQ)','baseline_path':str((BASE/'rerun_before_snapshot.json').resolve())}
(BASE/'average_expected.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps(result,indent=2))
