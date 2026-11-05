from artifact_tool import Blob,SpreadsheetFile
import json,pathlib
src=pathlib.Path(__file__).parent/'data'/'고3더프성적26년.xlsx'
w=SpreadsheetFile.import_xlsx(Blob.load(str(src)))
months=['3월','5월','7월','8월','9월','10월']
records={}; roster={}
for month in months:
 s=w.worksheets.get_item(month)
 # Wide source, each month same documented ordinal columns
 rows=s.get_range('A4:CU150').values
 for r in rows:
  name=r[4] if len(r)>4 else None
  if not isinstance(name,str) or not name.strip():continue
  name=name.strip(); number=str(r[3] or '').strip(); uid=f'{number}|{name}'
  roster[uid]={'id':uid,'student_number':number,'name':name,'school':'','grade':'고3'}
  def score(i):
   x=r[i] if len(r)>i else None
   if isinstance(x,(int,float)) and 0<=x<=100:return int(x)
   return None
  def grade(i):
   x=r[i] if len(r)>i else None
   try:
    z=int(x);return z if 1<=z<=9 else None
   except:return None
  records.setdefault(uid,{})[month]={
    '국어':{'choice':r[7], 'score':score(11),'grade':grade(14)},
    '수학':{'choice':r[21], 'score':score(25),'grade':grade(28)},
    '영어':{'choice':'영어','score':score(35),'grade':grade(36)},
    '한국사':{'choice':'한국사','score':score(75),'grade':grade(76)},
    '탐구1':{'choice':r[53], 'score':score(54),'grade':grade(57)},
    '탐구2':{'choice':r[64], 'score':score(65),'grade':grade(68)}
   }
# choose latest choice defaults from 10月 to earlier
for uid,v in roster.items():
 for subject,key in [('국어','korean'),('수학','math'),('탐구1','explore1'),('탐구2','explore2')]:
  choices=[records[uid][m][subject]['choice'] for m in reversed(months) if m in records[uid] and records[uid][m][subject]['choice']]
  v[key]=choices[0] if choices else ''
path=pathlib.Path(__file__).parent/'data'
(path/'history.json').write_text(json.dumps({'roster':list(roster.values()),'history':records},ensure_ascii=False,indent=2),encoding='utf8')
print('students',len(roster),'history rows',sum(map(len,records.values())))
