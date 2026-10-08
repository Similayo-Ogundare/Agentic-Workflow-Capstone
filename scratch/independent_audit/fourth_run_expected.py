from pathlib import Path
from collections import defaultdict
import hashlib
import json
import openpyxl

BASE=Path('scratch/independent_audit')
audit=json.loads((BASE/'current_audit.json').read_text(encoding='utf-8'))
members=[]
for row,name in enumerate(audit['roster'],start=7):
    m=audit['members'][name]['final']
    members.append({'name':name,'row':row,'tithe':m['tithe'],'offering':m['offering'],'combined':m['tithe']+m['offering'],'tithe_rank':m['tithe_rank_combined'],'offering_rank':m['offering_rank_combined'],'average':(m['tithe_rank_combined']+m['offering_rank_combined'])/2})
for m in members:
    m['overall_rank']=1+sum(other['average']<m['average'] or (other['average']==m['average'] and other['combined']>m['combined']) for other in members)
groups=defaultdict(list)
for m in members:
    groups[(m['average'],m['combined'])].append(m['name'])
residual_ties=[{'average':a,'combined':total,'names':names} for (a,total),names in groups.items() if len(names)>1]
wb=openpyxl.load_workbook('outputs/Final_Summary.xlsx',read_only=True,data_only=True)
rows=list(wb['Summary'].iter_rows(values_only=True))
changes=[]
discrepancies=[]
for m in members:
    row=rows[m['row']-1]
    for col,key in ((2,'tithe'),(3,'offering'),(5,'tithe_rank'),(7,'offering_rank'),(8,'average')):
        if row[col]!=m[key]:
            discrepancies.append({'name':m['name'],'key':key,'actual':row[col],'expected':m[key]})
    m['previous_overall_rank']=row[9]
    if row[9]!=m['overall_rank']:
        changes.append({'name':m['name'],'previous_rank':row[9],'new_rank':m['overall_rank'],'average':m['average'],'combined':m['combined']})
wb.close()
inventory=sorted(str(p.resolve()) for p in Path('outputs').rglob('*.xlsx') if not p.name.startswith('~$'))
result={'instructions_sha256':hashlib.sha256(Path('INSTRUCTIONS.md').read_bytes()).hexdigest(),'source_files':audit['files'],'source_totals':audit['final'],'member_records':audit['row_count'],'source_fingerprints_unchanged':audit['all_fingerprints_match_prior'],'all_sundays_present':all(m['complete'] for m in audit['calendars'].values()),'members':members,'existing_averages_category_ranks_amount_discrepancies':discrepancies,'rank_changes':changes,'unresolved_ties':residual_ties,'unique_overall_1_to_19':sorted(m['overall_rank'] for m in members)==list(range(1,20)),'rank_rule':'Ascending average of category consistency ranks; for equal averages, descending combined tithe+offering total. Competition ranking only if both criteria are still equal.','baseline_path':str((BASE/'fourth_run_before_snapshot.json').resolve()),'output_workbook_inventory':inventory}
(BASE/'fourth_run_expected.json').write_text(json.dumps(result,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in result.items() if k not in ('source_files','members','output_workbook_inventory')},indent=2))
