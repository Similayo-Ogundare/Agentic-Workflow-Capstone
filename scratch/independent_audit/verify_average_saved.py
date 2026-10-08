from pathlib import Path
from datetime import date,datetime
import hashlib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
import openpyxl

BASE=Path('scratch/independent_audit')
baseline=json.loads((BASE/'rerun_before_snapshot.json').read_text(encoding='utf-8'))
expected=json.loads((BASE/'average_expected.json').read_text(encoding='utf-8'))
source_audit=json.loads((BASE/'current_audit.json').read_text(encoding='utf-8'))
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS={'s':S}
issues=[]
checks=0
def canon(n):
    return(n.tag,tuple(sorted(n.attrib.items())),n.text or '',tuple(canon(c) for c in n))
def normalize_relationships(root,part,parts):
    relpart=part if part.endswith('.rels') else str(Path(part).parent/'_rels'/f'{Path(part).name}.rels').replace('\\','/')
    if relpart not in parts:
        return
    rels=ET.fromstring(parts[relpart])
    mapping={n.attrib['Id']:json.dumps({k:v for k,v in n.attrib.items() if k!='Id'},sort_keys=True) for n in rels}
    if part.endswith('.rels'):
        for n in root:
            n.attrib['Id']=mapping[n.attrib['Id']]
    else:
        for n in root.iter():
            for k in n.attrib:
                if k.startswith('{'+R+'}') and n.attrib[k] in mapping:
                    n.attrib[k]=mapping[n.attrib[k]]
def json_value(v):
    return v.isoformat() if isinstance(v,(date,datetime)) else v

monthly_hashes=[]
for label in ('January','February','March','April'):
    digest=hashlib.sha256(Path(baseline[label]['file']).read_bytes()).hexdigest()
    unchanged=digest==baseline[label]['sha256']
    monthly_hashes.append({'month':label,'byte_identical':unchanged})
    if not unchanged:
        issues.append({'type':'monthly workbook changed','month':label})

path=Path('outputs/Final_Summary.xlsx')
old=baseline['Final']
wf=openpyxl.load_workbook(path,read_only=True,data_only=False)
wv=openpyxl.load_workbook(path,read_only=True,data_only=True)
saved_pairs=[]
for sheet in wf.worksheets:
    actual={}
    cached=list(wv[sheet.title].iter_rows(values_only=True))
    for ri,row in enumerate(sheet.iter_rows(),start=1):
        for ci,c in enumerate(row,start=1):
            if c.value is not None:
                actual[c.coordinate]={'value':json_value(c.value),'cached_value':json_value(cached[ri-1][ci-1]),'data_type':c.data_type,'style_id':c._style_id,'number_format':c.number_format}
    for coord,cell in old['sheets'][sheet.title]['openpyxl_cells'].items():
        checks+=1
        if actual.get(coord)!=cell:
            issues.append({'type':'old nonblank cell changed','sheet':sheet.title,'cell':coord,'before':cell,'after':actual.get(coord)})
    if sheet.title=='Summary':
        for coord,title in (('I6','Average'),('J6','Overall Rank')):
            checks+=1
            if actual[coord]['value']!=title:
                issues.append({'type':'new header wrong','cell':coord})
        for member in expected['members']:
            row=member['row']
            for col,formula,key in (('I',f'=AVERAGE(F{row},H{row})','average'),('J',f'=1+COUNTIF($I$7:$I$25,"<"&I{row})','overall_rank')):
                cell=actual[col+str(row)]
                checks+=2
                if cell['value']!=formula:
                    issues.append({'type':'new formula wrong','cell':col+str(row),'actual':cell['value'],'expected':formula})
                if cell['cached_value']!=member[key]:
                    issues.append({'type':'new cached value wrong','cell':col+str(row),'actual':cell['cached_value'],'expected':member[key]})
            saved_pairs.append({'name':member['name'],'average':actual['I'+str(row)]['cached_value'],'overall_rank':actual['J'+str(row)]['cached_value']})
wf.close();wv.close()

with zipfile.ZipFile(path) as z:
    new_parts={name:z.read(name) for name in z.namelist()}
old_parts={name:data['xml'].encode('utf-8') for name,data in old['parts'].items() if 'xml' in data}
if set(new_parts)!=set(old['parts']):
    issues.append({'type':'Final ZIP native-part inventory changed'})
serialization_only=[]
native_additions=[]
for part,data in old['parts'].items():
    if part not in new_parts:
        continue
    if hashlib.sha256(new_parts[part]).hexdigest()==data['sha256']:
        continue
    if 'xml' not in data:
        issues.append({'type':'old binary native part changed','part':part});continue
    a=ET.fromstring(data['xml']);b=ET.fromstring(new_parts[part])
    normalize_relationships(a,part,old_parts);normalize_relationships(b,part,new_parts)
    if canon(a)==canon(b):
        serialization_only.append(part);continue
    if part=='xl/styles.xml':
        prefix_ok=True
        if [n.tag for n in a]!=[n.tag for n in b]:
            prefix_ok=False
        for ac,bc in zip(a,b):
            old_attrs={k:v for k,v in ac.attrib.items() if k!='count'}
            new_attrs={k:v for k,v in bc.attrib.items() if k!='count'}
            if old_attrs!=new_attrs or len(bc)<len(ac) or any(canon(x)!=canon(y) for x,y in zip(ac,bc)):
                prefix_ok=False
        if prefix_ok:
            native_additions.append({'part':part,'type':'new style definitions appended; every existing style definition preserved'});continue
    if part==old['sheets']['Summary']['part']:
        # Appended columns I/J, including their blank styled cells, are authorized.
        for root in (a,b):
            for row in root.findall('s:sheetData/s:row',NS):
                for cell in list(row):
                    coord=cell.attrib.get('r','')
                    letters=re.match(r'[A-Z]+',coord)
                    if (letters and letters.group() in ('I','J')) or coord=='A32':
                        row.remove(cell)
            cols=root.find('s:cols',NS)
            if cols is not None:
                for column in list(cols):
                    if int(column.attrib['min'])>=9:
                        cols.remove(column)
        if canon(a)==canon(b):
            native_additions.append({'part':part,'type':'only styled I/J columns and A32 note appended; all existing cells, row/column dimensions and features preserved'});continue
    issues.append({'type':'unexpected Final native XML change','part':part})

source_changes=[f['file'] for f in source_audit['files'] if hashlib.sha256(Path(f['file']).read_bytes()).hexdigest()!=f['sha256']]
issues.extend({'type':'source file changed','file':p} for p in source_changes)
report={'pass':not issues,'checks':checks,'issues':issues,'monthly_workbooks':monthly_hashes,'source_hash_changes':source_changes,'new_results':saved_pairs,'ties':expected['ties'],'serialization_only_parts':serialization_only,'authorized_native_additions':native_additions,'instruction_sha256':hashlib.sha256(Path('INSTRUCTIONS.md').read_bytes()).hexdigest()}
(BASE/'average_saved_qa.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2))
