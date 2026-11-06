"""2026 11월 대성 THE PREMIUM | OMR·월별 등급 성적표"""
import io, json, os, re, zipfile, csv, base64, copy
from pathlib import Path
from datetime import datetime
import numpy as np
import streamlit as st
import fitz
import cv2
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from template_omr import read_omr

ROOT=Path(__file__).resolve().parent
REFERENCE_ZIP=ROOT/'reference'/'november-answers.zip'
DATA=ROOT/'data'; DATA.mkdir(exist_ok=True)
STORAGE=DATA/'work.json'
MONTHS=['3월','5월','7월','8월','9월','10월','11월']
CATS=['국어','수학','영어','한국사','탐구1','탐구2']
LABEL={'탐구1':'탐구Ⅰ','탐구2':'탐구Ⅱ'}
OPTIONS={
'국어':['화법과 작문','언어와 매체'],
'수학':['확률과 통계','미적분','기하'],
'영어':['영어'],'한국사':['한국사'],
'탐구':['생활과 윤리','윤리와 사상','한국지리','세계지리','동아시아사','세계사','경제','정치와 법','사회·문화','물리학Ⅰ','화학Ⅰ','생명과학Ⅰ','지구과학Ⅰ','물리학Ⅱ','화학Ⅱ','생명과학Ⅱ','지구과학Ⅱ']}
ALIAS={'언어와매체':'언어와 매체','언매':'언어와 매체','화법과작문':'화법과 작문','화작':'화법과 작문','확률과통계':'확률과 통계','확통':'확률과 통계','생윤':'생활과 윤리','윤사':'윤리와 사상','한지':'한국지리','세지':'세계지리','동사':'동아시아사','세사':'세계사','정법':'정치와 법','사문':'사회·문화','물리1':'물리학Ⅰ','화학1':'화학Ⅰ','생명1':'생명과학Ⅰ','지구1':'지구과학Ⅰ','물리2':'물리학Ⅱ','화학2':'화학Ⅱ','생명2':'생명과학Ⅱ','지구2':'지구과학Ⅱ'}
# 2026년 10월 대성 THE PREMIUM 실제 원점수 기준표 — 첨부된 10월 표 기준. 9등급 하한은 0.
CUTS={
'국어|화법과 작문':[100,95,91,86,79,71,61,48,0],
'국어|언어와 매체':[95,91,87,81,74,65,55,42,0],
'수학|확률과 통계':[96,90,86,80,73,64,53,34,0],
'수학|미적분':[95,92,84,75,68,56,47,28,0],
'수학|기하':[95,95,88,78,71,62,51,32,0],
'영어|영어':[90,80,70,60,50,40,30,20,0],
'한국사|한국사':[40,35,30,25,20,15,10,5,0],
'탐구|생활과 윤리':[50,47,44,41,36,31,25,17,0],
'탐구|윤리와 사상':[48,47,43,39,36,31,24,16,0],
'탐구|한국지리':[50,48,47,44,39,32,24,16,0],
'탐구|세계지리':[50,48,45,41,36,29,18,12,0],
'탐구|동아시아사':[50,47,44,39,32,23,14,0,0],
'탐구|세계사':[50,48,45,41,36,29,18,12,0],
'탐구|경제':[47,42,38,33,28,23,17,12,0],
'탐구|정치와 법':[48,47,44,40,36,30,22,15,0],
'탐구|사회·문화':[50,47,45,44,41,38,33,25,0],
'탐구|물리학Ⅰ':[47,44,41,38,34,29,20,11,0],
'탐구|화학Ⅰ':[50,47,45,39,34,27,18,12,0],
'탐구|생명과학Ⅰ':[45,42,39,36,32,28,22,15,0],
'탐구|지구과학Ⅰ':[50,47,44,40,35,29,22,14,0],
'탐구|물리학Ⅱ':[50,47,43,39,34,28,16,9,0],
'탐구|화학Ⅱ':[48,45,40,35,29,22,16,10,0],
'탐구|생명과학Ⅱ':[47,44,40,36,31,26,17,10,0],
'탐구|지구과학Ⅱ':[50,47,44,39,33,22,10,8,0]}
# 국어·수학 공통/선택 문항 분할: 국어 1~34 + 선택 35~45 / 수학 1~22 + 선택 23~30
N={'국어':45,'수학':30,'영어':45,'한국사':20,'탐구':20}
NCHOICES={'국어':5,'수학':5,'영어':5,'한국사':5,'탐구':5}
# ReportLab 기본 Helvetica는 한글을 지원하지 않습니다.
# 운영체제 폰트가 없는 Streamlit Cloud에서는 오픈 라이선스 나눔고딕을
# 공식 Google Fonts 저장소에서 최초 1회 다운로드하여 서버 임시 캐시에 등록합니다.
# 폰트 파일은 GitHub에 올리거나 프로그램 ZIP에 포함하지 않습니다.
def ensure_korean_pdf_font():
 import urllib.request
 candidates=[
  Path('C:/Windows/Fonts/malgun.ttf'),
  Path('/usr/share/fonts/truetype/nanum/NanumGothic.ttf'),
  Path('/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf'),
  Path('/usr/share/fonts/truetype/nanum/NanumSquareR.ttf'),
 ]
 cache=Path('/tmp/thepremium-font-cache/NanumGothic-Regular.ttf')
 candidates.append(cache)
 errors=[]
 for p in candidates:
  if p.is_file() and p.stat().st_size > 100_000:
   try:
    pdfmetrics.registerFont(TTFont('Korean',str(p)))
    return 'Korean'
   except Exception as e: errors.append(str(e))
 if not cache.is_file() or cache.stat().st_size < 100_000:
  cache.parent.mkdir(parents=True,exist_ok=True)
  url='https://raw.githubusercontent.com/google/fonts/main/ofl/nanumgothic/NanumGothic-Regular.ttf'
  try:
   with urllib.request.urlopen(url,timeout=25) as response:
    font_bytes=response.read(6_000_000)
   if len(font_bytes)<100_000 or font_bytes[:4] not in (b'\x00\x01\x00\x00',b'true'):
    raise RuntimeError('다운로드한 파일이 올바른 TTF 글꼴이 아닙니다.')
   cache.write_bytes(font_bytes)
  except Exception as e:
   raise RuntimeError('한글 PDF 글꼴을 준비하지 못했습니다. Streamlit 서버의 외부 네트워크를 확인해 주세요: '+str(e)) from e
 try:
  pdfmetrics.registerFont(TTFont('Korean',str(cache)))
  return 'Korean'
 except Exception as e:
  raise RuntimeError('한글 PDF 글꼴 등록 실패: '+str(e)) from e

FONTNAME='Korean'

def norm(s):
 s=str(s or '').strip();return ALIAS.get(s,s)
def which(subject):return '탐구' if subject.startswith('탐구') else subject
def grade_of(subject,choice,score,cuts):
 if score is None:return None
 key=which(subject)+'|'+norm(choice)
 vals=cuts.get(key)
 if not vals:return None
 for i,v in enumerate(vals):
  if score>=v:return i+1
 return 9

def starting_state():
 history=json.loads((DATA/'history.json').read_text(encoding='utf8')) if (DATA/'history.json').exists() else {'roster':[],'history':{}}
 roster=history['roster']
 return {'roster':roster,'history':history['history'],'results':{},'keys':{},'key_verified':{},'cuts':CUTS,'layout':{},'logos':{k:base64.b64encode((ROOT/'assets'/v).read_bytes()).decode() for k,v in {'횡성인재육성관':'hoengseong.png','대성학력개발연구소':'daesung.png','천명의선택입시연구소':'choice1000.png'}.items() if (ROOT/'assets'/v).exists()},'notes':{},'settings':{'title':'대성THE프리미엄 모의고사 성적 [11월]','subtitle':'자체 정산-10월 성적대비등급'}}

def get_state():
 if 'db' not in st.session_state:
  try:db=json.loads(STORAGE.read_text(encoding='utf8')) if STORAGE.exists() else starting_state()
  except:db=starting_state()
  st.session_state.db=db
 return st.session_state.db

def save(db):
 # 테스트 모드는 서버 파일을 변경하지 않습니다. 테스트 데이터는 현재 로그인 세션의 메모리에서만 관리됩니다.
 if st.session_state.get('test_mode', False):
  return
 STORAGE.write_text(json.dumps(db,ensure_ascii=False,indent=2),encoding='utf8')


def encode_page(b,kind='pdf',page=0,dpi=160):
 if kind=='pdf':
  doc=fitz.open(stream=b,filetype='pdf'); p=doc[page];px=p.get_pixmap(matrix=fitz.Matrix(dpi/72,dpi/72),alpha=False)
  img=np.frombuffer(px.samples,dtype=np.uint8).reshape(px.h,px.w,px.n)[:,:,:3]
  doc.close();return img.copy()
 return np.array(Image.open(io.BytesIO(b)).convert('RGB'))

def interpret_omr(image,conf,answer_count,choices):
 """Per-question bubble-grid scan; returns answer, confidence, cause. Coordinates normalized 0..1."""
 gray=cv2.cvtColor(image,cv2.COLOR_RGB2GRAY);height,width=gray.shape
 x=float(conf['x']);y=float(conf['y']);w=float(conf['w']);h=float(conf['h'])
 cols=max(1,int(conf.get('columns',1)));gap=float(conf.get('colgap',0.0));
 rows=int(conf.get('rows',int(np.ceil(answer_count/cols))))
 chosen=[];stats=[]
 for q in range(answer_count):
  group=q//rows; row=q%rows
  xleft=x+group*(w+gap)
  yy=y+(row+.5)*h/rows
  values=[]
  for k in range(choices):
   xx=xleft+(k+.5)*w/choices
   xc=int(xx*width);yc=int(yy*height)
   rx=max(2,int(w*width/choices*float(conf.get('radius',.27))))
   ry=max(2,int(h*height/rows*float(conf.get('radius',.27))))
   box=gray[max(0,yc-ry):min(height,yc+ry+1),max(0,xc-rx):min(width,xc+rx+1)]
   values.append(float(np.mean(box<145)) if box.size else 0.)
  top=int(np.argmax(values));best=values[top];other=sorted(values,reverse=True)[1] if len(values)>1 else 0
  threshold=float(conf.get('threshold',.25)); margin=float(conf.get('margin',.09))
  status='정상' if best>=threshold and best-other>=margin else '미표기' if best<threshold else '복수/불확실'
  chosen.append(str(top+1) if status=='정상' else '')
  stats.append({'question':q+1,'values':[round(v,3) for v in values],'status':status,'confidence':round(best-other,3)})
 return chosen,stats

def outline(img,conf,n,choices):
 out=img.copy();h,w=out.shape[:2]
 for i in range(n):
  rows=int(conf.get('rows',n));k=i//rows;r=i%rows
  xx=float(conf['x'])+k*(float(conf['w'])+float(conf.get('colgap',0)))
  yy=float(conf['y'])+(r+.5)*float(conf['h'])/rows
  for z in range(choices):
   cx=int((xx+(z+.5)*float(conf['w'])/choices)*w);cy=int(yy*h)
   cv2.circle(out,(cx,cy),max(2,int(min(w,h)*.004)),(0,160,220),1)
  if i%5==0:cv2.putText(out,str(i+1),(int(xx*w)-22,int(yy*h)+3),cv2.FONT_HERSHEY_SIMPLEX,.34,(255,0,0),1)
 return out

def answer_key(db,subject,choice):
 k=f'{which(subject)}|{norm(choice)}'; return db['keys'].get(k,[])
def scored(db,uid,subject):
 roster=next((x for x in db['roster'] if x['id']==uid),None)
 if roster is None:return None
 prop={'국어':'korean','수학':'math','탐구1':'explore1','탐구2':'explore2','영어':None,'한국사':None}[subject]
 choice=norm(roster.get(prop,'')) if prop else subject
 rec=db['results'].get(uid,{}).get(subject,{})
 answers=rec.get('answers',[]); key=answer_key(db,subject,choice)
 n=N[which(subject)]
 if (subject in ('국어','수학','탐구1','탐구2') and not choice) or len(key)!=n or not db.get('key_verified',{}).get(which(subject)+'|'+norm(choice),False) or any(str(k.get('answer','')).strip()=='' or not isinstance(k.get('points'),int) or k['points']<=0 for k in key):
  return {'choice':choice,'answers':answers,'correct':[],'score':None,'grade':None,'ready':False,'review':rec.get('review','미채점'),'key':key}
 if not answers or len(answers)!=n or rec.get('review')!='확정':
  return {'choice':choice,'answers':answers,'correct':[],'score':None,'grade':None,'ready':False,'review':rec.get('review','미채점'),'key':key}
 correct=[];total=0
 for a,k in zip(answers,key):
  ok=str(a).strip()==str(k['answer']).strip() if str(a).strip() else False
  correct.append('○' if ok else '×' if a else '—')
  if ok:total+=k['points']
 return {'choice':choice,'answers':answers,'correct':correct,'score':total,'grade':grade_of(subject,choice,total,db['cuts']),'ready':True,'review':'확정','key':key}

def combined(db,uid,month,subject):
 if month=='11월':return scored(db,uid,subject)
 return db['history'].get(uid,{}).get(month,{}).get(subject)

def small_chart(c,values,x,y,width,height,subject):
 vs=[(i,v) for i,v in enumerate(values) if isinstance(v,(int,float)) and 1<=v<=9]
 c.setStrokeColor(HexColor('#C9D8E2'));c.setLineWidth(.6)
 for g in [1,3,5,7,9]:
  yy=y+height-(g-1)/8*height;c.line(x,yy,x+width,yy)
 c.setFont(FONTNAME,6);c.setFillColor(HexColor('#44566C'))
 for i,m in enumerate(MONTHS): c.drawCentredString(x+i*width/6,y-11,m)
 c.setFont(FONTNAME,7);c.drawString(x,y+height+8,f'{subject} 등급 추이 (1등급이 상위)')
 if len(vs)>=2:
  c.setStrokeColor(HexColor('#166C90'));c.setLineWidth(1.5)
  for (i,a),(j,b) in zip(vs,vs[1:]):
   if j==i+1:c.line(x+i*width/6,y+height-(a-1)/8*height,x+j*width/6,y+height-(b-1)/8*height)
 for i,g in vs:
  px=x+i*width/6;py=y+height-(g-1)/8*height;c.setFillColor(HexColor('#166C90'));c.circle(px,py,2.2,fill=1,stroke=0)
  c.setFillColor(HexColor('#111827'));c.drawCentredString(px,py+5,str(g))

def render_pdf(db,uid,include_answers=True,test_mode=False):
 ensure_korean_pdf_font()
 stu=next(z for z in db['roster'] if z['id']==uid)
 buff=io.BytesIO();c=canvas.Canvas(buff,pagesize=(595.28,841.89))
 c.setTitle(f"{stu['name']} 11월 더프 성적표")
 def text(x,y,s,size=9,bold=False):
  c.setFont(FONTNAME,size);c.setFillColor(HexColor('#20384E'));c.drawString(x,y,str(s))
 def line(y):c.setStrokeColor(HexColor('#D5DFE5'));c.line(36,y,558,y)
 def logo_or_text(key,x,y,w,h,txt):
  if key in db.get('logos',{}):
   try:
    img=ImageReader(io.BytesIO(base64.b64decode(db['logos'][key])));c.drawImage(img,x,y,width=w,height=h,preserveAspectRatio=True,anchor='c',mask='auto');return
   except:pass
  text(x,y+8,txt,9)
 def header():
  logo_or_text('횡성인재육성관',38,779,175,34,'횡성인재육성관')
  text(365,805,'2026년 11월 / 고3',8);line(775)
  title=db['settings'].get('title','대성THE프리미엄 모의고사 성적 [11월]')
  fs=min(17, max(11, 480/max(1,pdfmetrics.stringWidth(title,FONTNAME,1))))
  c.setFont(FONTNAME,fs);c.setFillColor(HexColor('#20384E'));c.drawCentredString(297.64,745,title)
  c.setFont(FONTNAME,10);c.drawCentredString(297.64,727,'(자체 정산 · 10월 성적 대비 잠정등급)')
  text(40,696,f"학생명  {stu['name']}     학교  {stu.get('school') or '미입력'}     학년  {stu.get('grade') or '고3'}",10)
  line(684)
  if test_mode:
   c.setFont(FONTNAME,9)
   c.setFillColor(HexColor('#B45309'))
   c.drawRightString(554,805,'[테스트용 · 실제 성적 아님]')
 def footer():
  line(63)
  logo_or_text('대성학력개발연구소',167,20,185,38,'대성  |  대성학력개발연구소')
  logo_or_text('천명의선택입시연구소',408,18,140,42,'천명의선택입시연구소')
  text(38,13,'※ 비공식 성적: 11월 원점수 / 10월 원점수 기준 잠정등급',7)
 header()
 text(42,662,'교과별 11월 자체 채점 결과',11)
 columns=[42,98,239,319,398,483]
 for xx,s in zip(columns,['영역','선택과목','원점수','잠정등급','10월등급','증감']):text(xx,640,s,8)
 line(634)
 for j,sub in enumerate(CATS):
  yy=615-j*29; r=scored(db,uid,sub);old=db['history'].get(uid,{}).get('10월',{}).get(sub,{})
  oldg=old.get('grade');newg=r.get('grade');diff=f'{oldg-newg:+d}' if isinstance(oldg,int) and isinstance(newg,int) else '-'
  text(columns[0],yy,LABEL.get(sub,sub),9);text(columns[1],yy,(r.get('choice') or '미선택')[:13],8)
  text(columns[2],yy,str(r['score']) if r['score'] is not None else '미확정',8)
  text(columns[3],yy,str(newg) if newg else '-',9)
  text(columns[4],yy,str(oldg) if oldg else '-',9);text(columns[5],yy,diff,9)
  line(yy-10)
 text(42,425,'대성THE프리미엄모의고사 응시 월별 등급 변화',11)
 for j,sub in enumerate(['국어','수학','영어','탐구1','탐구2']):
  x=50+(j%2)*265;y=345-(j//2)*94
  values=[]
  for m in MONTHS:
   r=combined(db,uid,m,sub) or {};values.append(r.get('grade'))
  small_chart(c,values,x,y,235,52,LABEL.get(sub,sub))
 text(42,78,'※ 탐구 선택과목이 바뀌어도 탐구Ⅰ·Ⅱ 영역 단위로 비교합니다.',7)
 footer()
 if include_answers:
  # 정오답은 과목별 별도 페이지: 가독성을 위해 5문항씩 두 줄 표기
  for sub in CATS:
   r=scored(db,uid,sub)
   if not r.get('answers'):continue
   c.showPage();header();text(42,660,f'{LABEL.get(sub,sub)} | {r.get("choice") or "미선택"}  · 답안/정오답',12)
   text(42,640,'문항 아래에 학생 답안, 그 아래에 정오답을 배치했습니다.',8)
   answers=r['answers'];correct=r.get('correct',[]);k=r.get('key',[])
   for start in range(0,len(answers),10):
    yy=602-(start//10)*88
    text(44,yy,f'{start+1}~{min(start+10,len(answers))}번',8)
    for row,values in [(0,[str(x) for x in range(start+1,min(start+10,len(answers))+1)]),(1,answers[start:start+10]),(2,correct[start:start+10] if correct else ['?']*len(answers[start:start+10]))]:
     for t,z in enumerate(values):
      xx=126+t*40
      if row==0:c.setFillColor(HexColor('#DDE8F0'));c.rect(xx-5,yy-6,34,20,fill=1,stroke=0)
      text(xx,yy-21*row,str(z or '-'),8)
   text(44,110,'※ 빈칸 및 복수 판독은 반드시 확인하고 답안을 확정하십시오.',8)
   footer()
 c.save();return buff.getvalue()

def export_xlsx(db,test_mode=False):
 wb=Workbook();ss=wb.active;ss.title='11월성적'
 if test_mode:
  note=wb.create_sheet('테스트안내',0)
  note['A1']='테스트 모드 출력: 실제 성적이 아닙니다. 학생 배포 금지'
  note.column_dimensions['A'].width=70
 ss.append(['학생번호','학생명','학교','학년','국어선택','국어점수','국어잠정등급','수학선택','수학점수','수학잠정등급','영어점수','영어등급','한국사등급','탐구1등급','탐구2등급'])
 for stu in db['roster']:
  uid=stu['id'];r={q:scored(db,uid,q) for q in CATS}
  ss.append([stu.get('student_number'),stu['name'],stu.get('school',''),stu.get('grade','고3'),r['국어']['choice'],r['국어']['score'],r['국어']['grade'],r['수학']['choice'],r['수학']['score'],r['수학']['grade'],r['영어']['score'],r['영어']['grade'],r['한국사']['grade'],r['탐구1']['grade'],r['탐구2']['grade']])
 hs=wb.create_sheet('월별등급변화');hs.append(['학생번호','학생명','교과',*MONTHS])
 for stu in db['roster']:
  for sub in CATS:
   vals=[(combined(db,stu['id'],month,sub) or {}).get('grade') for month in MONTHS]
   hs.append([stu.get('student_number'),stu['name'],sub,*vals])
 for sheet in [ss,hs]:
  sheet.freeze_panes='D2';sheet.auto_filter.ref=sheet.dimensions
  for cell in sheet[1]:
   cell.fill=PatternFill('solid',fgColor='17324D');cell.font=Font(color='FFFFFF',bold=True)
   cell.alignment=Alignment(horizontal='center')
  for col in sheet.columns:
   letter=col[0].column_letter
   sheet.column_dimensions[letter].width=18
 out=io.BytesIO();wb.save(out);return out.getvalue()


# 기존 3~10월 성적 엑셀 재등록 — 공개 GitHub 저장소에는 학생 자료를 넣지 않습니다.
def import_history_excel(content):
 from openpyxl import load_workbook
 from collections import defaultdict
 wb=load_workbook(io.BytesIO(content),read_only=True,data_only=True)
 missing=[m for m in MONTHS[:-1] if m not in wb.sheetnames]
 if missing:raise ValueError('필수 시트 누락: '+', '.join(missing))
 roster={};history={};id_names=defaultdict(set);month_count={}
 for month in MONTHS[:-1]:
  ws=wb[month];count=0
  for row in ws.iter_rows(min_row=4,max_col=99,values_only=True):
   if len(row)<78:continue
   name=row[4]
   num=row[3]
   if not isinstance(name,str) or not name.strip() or num in (None,''):continue
   name=name.strip();num=str(int(num)) if isinstance(num,(int,float)) and float(num).is_integer() else str(num).strip()
   uid=f'{num}|{name}'  # 번호 충돌 시 별개 학생 보호
   school=str(row[5] or '').strip()
   id_names[num].add(name)
   if uid not in roster:roster[uid]={'id':uid,'student_number':num,'name':name,'school':school,'grade':'고3'}
   elif school and not roster[uid].get('school'):roster[uid]['school']=school
   def score(i):
    v=row[i] if i<len(row) else None
    if isinstance(v,(float,int)) and 0<=v<=100:return int(v) if float(v).is_integer() else float(v)
    return None
   def grade(i):
    v=row[i] if i<len(row) else None
    try:
     z=int(v);return z if 1<=z<=9 else None
    except:return None
   history.setdefault(uid,{})[month]={
    '국어':{'choice':row[8],'score':score(12),'grade':grade(15)},
    '수학':{'choice':row[22],'score':score(26),'grade':grade(29)},
    '영어':{'choice':'영어','score':score(36),'grade':grade(37)},
    '한국사':{'choice':'한국사','score':score(76),'grade':grade(77)},
    '탐구1':{'choice':row[54],'score':score(55),'grade':grade(58)},
    '탐구2':{'choice':row[65],'score':score(66),'grade':grade(69)}
   }
   count+=1
  month_count[month]=count
 for uid,stu in roster.items():
  for sub,field in [('국어','korean'),('수학','math'),('탐구1','explore1'),('탐구2','explore2')]:
   picks=[history[uid][m][sub].get('choice') for m in reversed(MONTHS[:-1]) if m in history[uid]]
   stu[field]=next((norm(p) for p in picks if p and str(p).strip()),'')
 collision={num:sorted(names) for num,names in id_names.items() if len(names)>1}
 return {'roster':list(roster.values()),'history':history},collision,month_count

def install_history(db,record):
 # 과거 OMR 검수 결과는 학생번호+이름이 정확히 일치하는 인원에만 유지
 allowed={s['id'] for s in record['roster']}
 db['results']={uid:r for uid,r in db.get('results',{}).items() if uid in allowed}
 db['roster']=record['roster'];db['history']=record['history']
 return len(allowed)

st.set_page_config(page_title='11월 더프 | 성적처리 NAVI',page_icon='📊',layout='wide')
# 배포된 앱은 비밀번호가 없으면 학생 개인정보 접근을 차단합니다.
try:
 APP_PASSWORD=st.secrets.get('APP_PASSWORD','')
except Exception:
 APP_PASSWORD=''
if not APP_PASSWORD:
 st.error('관리자 비밀번호가 설정되지 않았습니다. Streamlit Cloud App settings → Secrets에 APP_PASSWORD를 설정하십시오.')
 st.stop()
if not st.session_state.get('authorized',False):
 st.title('대성 THE PREMIUM 성적처리 NAVI')
 entered=st.text_input('관리자 비밀번호',type='password')
 if st.button('로그인',type='primary'):
  if entered==APP_PASSWORD:
   st.session_state.authorized=True;st.rerun()
  else:st.error('비밀번호를 확인하십시오.')
 st.stop()
# 운영/테스트 데이터 완전 분리: 테스트에서는 운영 DB를 복제만 하고 저장하지 않습니다.
production_db=get_state()
mode=st.sidebar.radio('운영 모드', ['실제 운영', '테스트 모드'], key='operating_mode')
test_mode=(mode=='테스트 모드')
if st.session_state.get('previous_operating_mode') != mode:
 for k in ('pdf_zip','result_xlsx','all_student_omr'):
  st.session_state.pop(k,None)
 st.session_state.previous_operating_mode=mode
st.session_state.test_mode=test_mode
if test_mode:
 if 'test_db' not in st.session_state:
  st.session_state.test_db=copy.deepcopy(production_db)
 db=st.session_state.test_db
 st.sidebar.warning('테스트 전용: 변경·채점 내용은 실제 성적에 저장되지 않습니다.')
 if st.sidebar.button('테스트 데이터 전체 초기화', use_container_width=True, type='secondary'):
  # 저장되지 않은 테스트 답안, 테스트 정답, 판독설정과 PDF 캐시를 모두 삭제합니다.
  st.session_state.test_db=copy.deepcopy(production_db)
  for k in ('pdf_zip','result_xlsx','all_student_omr'):
   st.session_state.pop(k,None)
  st.rerun()
else:
 db=production_db
 st.sidebar.success('실제 운영 데이터: 테스트 기록과 분리됨')
if db.get('settings',{}).get('title')=='11월 대성THE프리미엄 모의고사 성적':
 db['settings']['title']='대성THE프리미엄 모의고사 성적 [11월]'
st.title('11월 대성 THE PREMIUM | OMR 성적처리 NAVI')
if test_mode:
 st.error('🧪 테스트 모드 — 샘플 OMR을 판독하고 PDF까지 확인할 수 있으나, 실제 11월 성적에는 반영되지 않습니다. 결과는 세션 종료 시 사라집니다.')
else:
 st.info('실제 운영 모드 — 확정된 답안과 정답, 성적을 운영 데이터에 저장합니다.')
st.caption('Streamlit Cloud 배포판 · 2026년 11월 고3 · OMR 판독 / 검수 / 10월 기준 잠정 등급 / 월별 변화 / PDF 인쇄 및 파일 저장')
st.warning('11월 전국 공동 채점 결과가 아닙니다. 10월 원점수 등급컷을 11월에 적용한 비공식 참고 성적이며 시험 난이도 차이는 보정하지 않습니다.')
page=st.sidebar.radio('메뉴',['11월 정답·배점','학생 명부','OMR 판독·검수','학생 성적표','일괄 인쇄·저장','설정·로고·백업'])
if st.sidebar.button('현재 수정사항 저장',use_container_width=True):
 if test_mode:st.sidebar.info('테스트 중에는 서버 저장을 하지 않습니다. 입력값은 현재 세션에서만 유지됩니다.')
 else:save(db);st.sidebar.success('운영 작업 파일에 저장했습니다.')

if page=='학생 명부':
 st.subheader('3월~10월 성적 엑셀 연결 명부')
 st.markdown('**수정된 학생 명부와 성적 엑셀 불러오기**')
 st.caption('학생 개인정보가 포함된 엑셀은 GitHub에 올리지 마십시오. 로그인 후 이 화면에서만 올리십시오. 학생번호와 이름을 함께 대조합니다.')
 new_excel=st.file_uploader('3월·5월·7월·8월·9월·10월 시트가 있는 엑셀(.xlsx)',type=['xlsx'],key='history_xlsx')
 if new_excel and st.button('성적 엑셀 검증 및 학생 이력 갱신',type='primary'):
  try:
   parsed,collision,month_count=import_history_excel(new_excel.getvalue())
   count=install_history(db,parsed);save(db)
   st.success(f'고유 학생 {count}명 / 월별 기록 {sum(month_count.values())}건 갱신')
   st.write('월별 응시 기록 수:',month_count)
   if collision:st.warning('학생번호가 같은데 이름이 다른 기록이 있습니다. 임의 병합하지 않았습니다.');st.json(collision)
   st.rerun()
  except Exception as e:st.error(f'엑셀 불러오기 실패: {e}')

 st.write(f"학생 {len(db['roster'])}명 · 월별 기존 성적이 있는 학생을 통합 관리합니다.")
 if not db['roster']:st.info('위에서 수정 엑셀을 먼저 업로드하십시오.');st.stop()
 names=[f"{a['student_number']} | {a['name']}" for a in db['roster']]
 selected=st.selectbox('수정할 학생',list(range(len(names))),format_func=lambda i:names[i])
 stu=db['roster'][selected]
 a,b,c=st.columns(3)
 stu['name']=a.text_input('학생명',stu['name']);stu['school']=b.text_input('학교',stu.get('school',''));stu['grade']=c.text_input('학년',stu.get('grade','고3'))
 for k,label,opts in [('korean','국어',OPTIONS['국어']),('math','수학',OPTIONS['수학']),('explore1','탐구Ⅰ',OPTIONS['탐구']),('explore2','탐구Ⅱ',OPTIONS['탐구'])]:
  previous=norm(stu.get(k,''));options=['']+opts
  stu[k]=st.selectbox(label,options,index=options.index(previous) if previous in options else 0,key='choice_'+k)
 if st.button('학생 정보 저장'):save(db);st.success('저장했습니다.')
 st.dataframe([{'학생번호':x.get('student_number'),'이름':x['name'],'학교':x.get('school',''),'국어':norm(x.get('korean')),'수학':norm(x.get('math')),'탐구Ⅰ':norm(x.get('explore1')),'탐구Ⅱ':norm(x.get('explore2'))} for x in db['roster']],use_container_width=True,hide_index=True)
 st.info('수정 엑셀의 학생번호·이름·학교를 대조해 이력을 연결합니다. 서로 다른 이름이 동일 학생번호를 쓴 경우 별도 인원으로 구분합니다.')

elif page=='11월 정답·배점':
 st.subheader('① 11월 정답 원본 확인')
 st.info('앞서 제공받은 11월 정답 이미지 16장을 원본 대조용으로 포함했습니다. 이미지를 보고 정답·배점을 검수한 뒤 아래 표를 확정해야 채점됩니다. 원본 이미지에서 정답이 자동 추출·검증된 상태는 아닙니다.')
 source=st.file_uploader('추가 정답 원본 ZIP 또는 PDF·이미지 올리기',type=['zip','pdf','png','jpg','jpeg'],key='key_reference_upload')
 images=[]
 try:
  if source and source.name.lower().endswith('.zip'):
   with zipfile.ZipFile(io.BytesIO(source.getvalue())) as z:
    images=[(n,z.read(n)) for n in z.namelist() if n.lower().endswith(('.jpg','.jpeg','.png')) and not n.startswith('__MACOSX/')]
  elif source:
   if source.name.lower().endswith('.pdf'):
    with fitz.open(stream=source.getvalue(),filetype='pdf') as d:
     for i in range(len(d)):
      pix=d[i].get_pixmap(matrix=fitz.Matrix(1.4,1.4),alpha=False)
      images.append((f'{source.name} / {i+1}쪽',pix.tobytes('png')))
   else:images=[(source.name,source.getvalue())]
  elif REFERENCE_ZIP.exists():
   with zipfile.ZipFile(REFERENCE_ZIP) as z:
    images=[(n,z.read(n)) for n in z.namelist() if n.lower().endswith(('.jpg','.jpeg','.png'))]
 except Exception as e: st.error(f'원본 이미지 열기 실패: {e}')
 if images:
  st.caption(f'정답 원본 이미지 {len(images)}장 · 원본의 정답 및 배점 수치를 직접 대조하십시오.')
  imgidx=st.selectbox('확인할 정답 원본',range(len(images)),format_func=lambda i:f'{i+1:02d} · {images[i][0]}',key='answer_reference_idx')
  st.image(images[imgidx][1],caption=images[imgidx][0],use_container_width=True)
 else:st.warning('정답 원본이 없습니다. 상단에 정답 ZIP을 올리십시오.')
 st.divider()
 st.subheader('② 과목·선택과목별 정답 및 배점 입력·검증')
 sub=st.selectbox('영역',['국어','수학','영어','한국사','탐구'])
 choice=st.selectbox('선택과목',OPTIONS[sub]);k=sub+'|'+choice;n=N[sub]
 current=db['keys'].get(k,[{'question':i+1,'answer':'','points':None} for i in range(n)])
 if len(current)!=n:current=[{'question':i+1,'answer':'','points':None} for i in range(n)]
 st.caption('국어 1~34번·수학 1~22번 공통문항도 포함해 각 선택과목의 전체 문항을 입력하십시오. 수학 단답형은 정수를 입력합니다.')
 raw='\n'.join(f"{p['question']},{p.get('answer','')},{p.get('points') or ''}" for p in current)
 csv_file=st.file_uploader('정답·배점 CSV 가져오기 (선택 사항)',type=['csv'],key=f'keycsv_{k}')
 if csv_file:
  try:raw=csv_file.getvalue().decode('utf-8-sig').strip()
  except UnicodeDecodeError:st.error('CSV를 UTF-8로 저장해 주십시오.')
 st.download_button('빈 정답 CSV 양식 다운로드',('\n'.join(f'{i},,' for i in range(1,n+1))).encode('utf-8-sig'),file_name=f'thepremium-key-{sub}.csv',mime='text/csv')
 text_csv=st.text_area('문항번호,정답,배점 — 한 줄에 한 문항',raw,height=300,key=f'key_editor_{k}')
 total=st.number_input('만점 확인',min_value=1,max_value=100,value=50 if sub=='탐구' else 100,key=f'key_fullscore_{k}')
 st.caption(f'현재 확정 여부: {"확정" if db.get("key_verified",{}).get(k) else "미확정"} · 정답 원본과 재확인 후 저장하십시오.')
 if st.button('문항수·정답·배점 검증 후 확정 저장',type='primary'):
  try:
   result=[]
   for line in text_csv.splitlines():
    parts=[x.strip() for x in line.split(',')]
    if len(parts)!=3:raise ValueError('각 행은 문항번호,정답,배점 3개여야 합니다.')
    q=int(parts[0]);ans=parts[1];points=int(parts[2])
    if points<=0 or not ans:raise ValueError(f'{q}번 정답 또는 배점이 비어 있습니다.')
    if sub=='수학' and 23<=q<=30:
     if not ans.isdecimal():raise ValueError(f'{q}번 수학 단답은 숫자로 입력하십시오.')
    elif ans not in ('1','2','3','4','5'):
     raise ValueError(f'{q}번 객관식 답안은 1~5 중 하나여야 합니다.')
    result.append({'question':q,'answer':ans,'points':points})
   if [z['question'] for z in result]!=list(range(1,n+1)):raise ValueError(f'1~{n}번 문항번호가 중복·누락·순서 오류입니다.')
   if sum(z['points'] for z in result)!=total:raise ValueError(f'배점 합계 {sum(z["points"] for z in result)}점 ≠ 설정 만점 {total}점')
   # 공통문항이 기존 선택과목과 다르면 오입력 가능성을 알려 주고 확정을 막음.
   if sub in ('국어','수학'):
    common=34 if sub=='국어' else 22
    for other,otherkey in db['keys'].items():
     if other.startswith(sub+'|') and other!=k and len(otherkey)==n and otherkey[:common]!=result[:common]:
      raise ValueError(f'{other}와 1~{common}번 공통 정답·배점이 일치하지 않습니다. 원본을 대조해 주십시오.')
   db['keys'][k]=result;db.setdefault('key_verified',{})[k]=True;save(db)
   st.success(f'{choice}: {n}문항, {total}점 정답 등록 완료')
  except Exception as e:st.error(f'저장 불가: {e}')
 st.divider()
 st.subheader('③ 과목별 정답 확정 현황')
 st.dataframe([{'영역·과목':name,'확정 상태':'확정' if db.get('key_verified',{}).get(name) else '미확정','문항수':len(entries),'총 배점':sum((e.get('points') or 0) for e in entries)} for name,entries in db['keys'].items()],hide_index=True,use_container_width=True)
 st.download_button('확정 정답 백업 JSON 다운로드',json.dumps({'keys':db['keys'],'key_verified':db.get('key_verified',{})},ensure_ascii=False,indent=2).encode('utf-8'),file_name='thepremium-answer-keys.json',mime='application/json')
 st.divider()
 st.subheader('10월 등급컷 검토·수정')
 cut=db['cuts'].get(k)
 if cut:
  cutoff=st.text_input('1~9등급의 원점수 하한(쉼표 구분)',','.join(map(str,cut)))
  if st.button('등급컷 저장'):
   try:
    vals=[int(x.strip()) for x in cutoff.split(',')]
    assert len(vals)==9 and vals==sorted(vals,reverse=True) and vals[-1]==0
    db['cuts'][k]=vals;save(db);st.success('저장했습니다.')
   except:st.error('9개 정수, 내림차순, 마지막 값 0으로 입력하십시오.')

elif page=='OMR 판독·검수':
 if not db['roster']:
  st.warning('학생 명부 메뉴에서 수정 엑셀을 먼저 업로드하십시오.');st.stop()
 st.subheader('① 교과별 OMR PDF 일괄 업로드')
 st.info('시험 종료 후 교과별 PDF를 각각 올리십시오. 예: 국어 PDF 20쪽 = 학생 20명의 국어 답안지. PDF의 모든 페이지를 한 번에 판독합니다. 학생명은 페이지별로 확인하여 연결합니다.')
 subject_group=st.selectbox('업로드할 교과 영역',['국어','수학','영어','한국사','탐구'],key='subject_batch_group')
 st.caption('탐구는 한 페이지의 탐구Ⅰ·탐구Ⅱ를 모두 판독합니다. 수학 단답형은 별도 확인이 필요합니다.')
 uploaded=st.file_uploader(f'{subject_group} OMR 묶음 PDF (여러 학생의 답안지가 들어 있는 파일)',type=['pdf','png','jpg','jpeg'],key='all_student_omr')
 if uploaded is not None:
  content=uploaded.getvalue()
  kind='pdf' if uploaded.name.lower().endswith('.pdf') else 'image'
  if kind=='pdf':
   with fitz.open(stream=content,filetype='pdf') as doc:pages=len(doc)
  else:pages=1
  import hashlib
  fingerprint=hashlib.sha256(content).hexdigest()
  batch=st.session_state.get('omr_batch')
  if batch and (batch.get('fingerprint')!=fingerprint or batch.get('mode')!=mode or batch.get('subject_group')!=subject_group):
   st.session_state.pop('omr_batch',None);batch=None
  st.caption(f'파일: {uploaded.name} · {subject_group} · {pages}페이지')
  st.subheader('② 전체 페이지 자동판독')
  if st.button(f'{subject_group} {pages}페이지 전체 자동판독',type='primary',use_container_width=True):
   results=[]
   bar=st.progress(0,text='교과별 OMR 판독 준비')
   for pi in range(pages):
    subkeys=['탐구1','탐구2'] if subject_group=='탐구' else [subject_group]
    try:
     img=encode_page(content,kind,pi,dpi=160)
     for sub in subkeys:
      try:
       answers,diagnostics,preview,meta=read_omr(img,sub)
       results.append({'page':pi+1,'bundle':pi,'subject':sub,'answers':answers,
                       'stats':diagnostics,'rotation':meta.get('rotation',0),
                       'error':'','reviewed':False,'student_id':'','suspect':meta.get('suspect',False),'detected':meta.get('detected',0)})
      except Exception as exc:
       results.append({'page':pi+1,'bundle':pi,'subject':sub,'answers':['']*N[which(sub)],
                       'stats':[],'rotation':None,'error':str(exc),'reviewed':False,'student_id':'','suspect':True,'detected':0})
    except Exception as exc:
     for sub in subkeys:
      results.append({'page':pi+1,'bundle':pi,'subject':sub,'answers':['']*N[which(sub)],
                      'stats':[],'rotation':None,'error':'PDF 변환 실패: '+str(exc),'reviewed':False,'student_id':'','suspect':True,'detected':0})
    bar.progress((pi+1)/pages,text=f'{pi+1}/{pages}페이지 처리 완료')
   st.session_state.omr_batch={'fingerprint':fingerprint,'mode':mode,'rows':results,
                               'subject_group':subject_group,'file':uploaded.name}
   st.rerun()
  batch=st.session_state.get('omr_batch')
  if batch and batch.get('fingerprint')==fingerprint and batch.get('mode')==mode and batch.get('subject_group')==subject_group:
   rows=batch['rows']
   good=sum(not r['error'] for r in rows)
   st.success(f'일괄 판독 시도 완료: {pages}페이지 / 답안 영역 {len(rows)}건 / 정렬 성공 {good}건')
   st.caption('자동 판독 성공은 실제 마킹의 정확성을 보증하지 않습니다. 학생·답안을 검수한 결과만 반영합니다.')
   summary=[]
   for i,r in enumerate(rows):
    ok=sum(x.get('상태')=='정상' for x in r['stats'])
    summary.append({'번호':i+1,'PDF 페이지':r['page'],'교과':LABEL.get(r['subject'],r['subject']),
                    '자동 판독':f'{ok}/{len(r["answers"])}',
                    '확인 필요':len(r['answers'])-ok,
                    '상태':'실패' if r['error'] else ('재판독/확인 필요' if r.get('suspect') else ('검수완료' if r['reviewed'] else '검수대기'))})
   st.dataframe(summary,use_container_width=True,hide_index=True)
   st.subheader('③ PDF 페이지별 학생 연결 · 답안 검수')
   roster=db['roster']
   labels=['학생 선택 필요']+[f"{a['student_number']} | {a['name']} | {a.get('school','')}" for a in roster]
   chosen_page=st.selectbox('검수할 PDF 페이지',range(1,pages+1),format_func=lambda j:f'{j}페이지 · {subject_group} OMR')
   indices=[i for i,r in enumerate(rows) if r['page']==chosen_page]
   existing=next((rows[i]['student_id'] for i in indices if rows[i].get('student_id')),'')
   selected_index=next((j+1 for j,v in enumerate(roster) if v['id']==existing),0)
   selected=st.selectbox(f'{chosen_page}페이지 학생',range(len(labels)),index=selected_index,
                         format_func=lambda j:labels[j],key=f'subject_student_{chosen_page}_{fingerprint[:8]}_{subject_group}')
   selected_uid=roster[selected-1]['id'] if selected else ''
   # 페이지 하나는 학생 한 명의 OMR: 탐구 두 선택과목에도 같은 학생 ID를 적용합니다.
   for i in indices:
    if rows[i]['student_id']!=selected_uid:rows[i]['reviewed']=False
    rows[i]['student_id']=selected_uid
   if selected_uid:
    dupe_pages=sorted({r['page'] for r in rows if r['student_id']==selected_uid and r['page']!=chosen_page})
    if dupe_pages:st.warning(f'같은 PDF의 {dupe_pages}페이지에도 해당 학생이 연결되어 있습니다. 중복 여부를 확인하십시오.')
   st.caption('PDF 페이지 순서와 학생명은 자동으로 확정되지 않습니다. 원본에 표시된 학생번호·이름을 확인하여 연결하십시오.')
   if st.checkbox('해당 OMR 원본 미리보기',key=f'subject_preview_{chosen_page}_{fingerprint[:8]}'):
    st.image(encode_page(content,kind,chosen_page-1,dpi=110),width=700)
   for i in indices:
    r=rows[i]
    st.markdown(f'**{chosen_page}페이지 · {LABEL.get(r["subject"],r["subject"])}**')
    if r['error']:st.error('자동판독 실패: '+r['error'])
    if r.get('suspect') and not r['error']:
     st.error('자동 판독한 답안 수가 지나치게 적습니다. OMR 정렬 또는 채점 원 위치가 틀릴 가능성이 있으므로 이 결과를 그대로 확정할 수 없습니다. 아래 정렬 미리보기를 확인해 주십시오.')
    if not r['error']:
     if st.checkbox('자동 정방향·판독 원 위치 미리보기',value=bool(r.get('suspect')),key=f'aligned_preview_{i}_{fingerprint[:8]}'):
      try:
       im=encode_page(content,kind,r['page']-1,dpi=160)
       _,_,overlay,info=read_omr(im,r['subject'])
       st.caption(f'자동회전 {info["rotation"]}° / 정렬 대응점 {info["inliers"]}개 · 초록 원=자동판독, 주황/파랑 원=검수 필요')
       st.image(overlay,use_container_width=True)
      except Exception as exc:st.error('정렬 미리보기 생성 실패: '+str(exc))
    if st.button('이 답안 영역 다시 판독',key=f'subject_retry_{i}_{fingerprint[:8]}_{subject_group}'):
     try:
      im=encode_page(content,kind,r['page']-1)
      ans,stats,_,meta=read_omr(im,r['subject'])
      r.update(answers=ans,stats=stats,rotation=meta.get('rotation'),error='',reviewed=False,suspect=meta.get('suspect',False),detected=meta.get('detected',0))
      st.rerun()
     except Exception as e:st.error(str(e))
    if r['stats']:
     uncertain=[x for x in r['stats'] if x.get('상태')!='정상']
     if uncertain:
      st.warning(f'재확인이 필요한 문항 {len(uncertain)}개 (수학 단답형 포함)')
      st.dataframe(uncertain,use_container_width=True,hide_index=True)
     else:st.caption('자동 판독 후보는 모두 정상 판정되었습니다. 실제 OMR과 대조하십시오.')
    edited=st.text_area('판독 답안 (쉼표로 구분, 미표기=빈칸)',value=','.join(map(str,r['answers'])),
                        key=f'subject_edit_{i}_{r["subject"]}_{fingerprint[:8]}',height=95)
    parsed=[x.strip() for x in edited.split(',')]
    correct_count=N[which(r['subject'])]
    if len(parsed)!=correct_count:
     st.error(f'답안 수 {len(parsed)}개 / 필요 {correct_count}개')
     r['reviewed']=False
    else:
     if parsed!=r['answers']:r['reviewed']=False
     r['answers']=parsed
     if r.get('suspect'):
      st.checkbox('원본과 1번부터 마지막 문항까지 직접 대조했고 필요한 답안을 수정했습니다.',key=f'manual_verified_{i}_{fingerprint[:8]}')
     if st.button('이 교과 검수 완료',key=f'subject_review_{i}_{fingerprint[:8]}'):
      if not selected_uid:st.error('학생을 먼저 선택하십시오.')
      elif r['error']:st.error('자동판독 오류를 해결하거나 재판독하십시오.')
      elif r.get('suspect') and not st.session_state.get(f'manual_verified_{i}_{fingerprint[:8]}',False):
       st.error('오검출 가능성이 커서 원본 마킹을 직접 대조했다는 확인이 필요합니다. 아래 체크박스를 먼저 선택하십시오.')
      else:r['reviewed']=True;st.success('검수 완료 — 아직 저장되지 않았습니다.')
   st.divider()
   st.subheader('④ 검수 완료 답안 일괄 확정')
   ready=[r for r in rows if r['reviewed'] and r['student_id'] and not r['error']]
   st.write(f'검수 완료 **{len(ready)} / {len(rows)}건**')
   # 같은 학생/교과 중복 저장으로 서로 다른 답안이 뒤집히지 않게 방지합니다.
   keys=[(r['student_id'],r['subject']) for r in ready]
   duplicates_in_upload=len(keys)!=len(set(keys))
   if duplicates_in_upload:st.error('동일 학생·교과가 여러 페이지에 중복 연결되었습니다. 학생 매칭을 수정하십시오.')
   st.warning('확정하면 해당 학생·교과의 기존 11월 답안이 교체될 수 있습니다. 정답·배점이 확정된 경우에만 성적을 계산할 수 있습니다.')
   confirmed=st.checkbox('원본의 학생명·교과·마킹 답안을 검수했습니다.',key='subject_batch_commit_confirm')
   if st.button('검수 완료 교과 답안 일괄 저장',disabled=not ready or not confirmed or duplicates_in_upload,type='primary'):
    newresults=copy.deepcopy(db['results'])
    existing_count=sum(bool(db['results'].get(r['student_id'],{}).get(r['subject'])) for r in ready)
    for r in ready:
     newresults.setdefault(r['student_id'],{})[r['subject']]={
      'answers':list(r['answers']),'review':'확정','source':uploaded.name,'page':r['page']}
    db['results']=newresults;save(db)
    st.success(f'{subject_group} 답안 {len(ready)}건 저장 완료 (기존 기록 대체 {existing_count}건).')
    if test_mode:st.info('테스트 모드 — 실제 운영 데이터에는 반영되지 않았습니다.')
    for r in ready:r['reviewed']=False
    st.info('다음 교과 PDF를 처리하려면 현재 파일을 제거하고 상단에서 다음 교과를 선택하여 업로드하십시오.')
  else:st.caption('「전체 자동판독」 버튼을 클릭하십시오.')
 else:st.caption('교과를 선택한 뒤 PDF를 업로드하십시오.')

elif page=='학생 성적표':
 if not db['roster']:st.warning('학생 명부 메뉴에서 수정 엑셀을 먼저 업로드하십시오.');st.stop()
 st.subheader('성적표 조회 · 인쇄')
 opts=[f"{x['student_number']} | {x['name']}" for x in db['roster']]
 i=st.selectbox('학생 선택',range(len(opts)),format_func=lambda x:opts[x]);stu=db['roster'][i];uid=stu['id']
 rows=[]
 for sub in CATS:
  r=scored(db,uid,sub);rows.append({'영역':sub,'선택과목':r.get('choice'),'11월 원점수':r.get('score'),'11월 잠정등급':r.get('grade'),'검수':r.get('review')})
 st.dataframe(rows,hide_index=True,use_container_width=True)
 st.subheader('대성THE프리미엄모의고사 응시 월별 등급 변화')
 history=[]
 for month in MONTHS:
  row={'월':month}
  for sub in CATS:
   r=combined(db,uid,month,sub) or {};row[sub]=r.get('grade')
  history.append(row)
 st.dataframe(history,hide_index=True,use_container_width=True)
 # native chart uses numerical missing values
 st.line_chart({sub:[r[sub] for r in history] for sub in ['국어','수학','영어','탐구1','탐구2']},height=260)
 show_ans=st.checkbox('성적표에 교과별 답안/정오답 추가 페이지 포함',value=True)
 b=render_pdf(db,uid,show_ans,test_mode=test_mode)
 st.download_button('학생별 A4 PDF 저장 · 인쇄',b,file_name=f"thepremium-{'TEST-' if test_mode else ''}{stu['name']}.pdf",mime='application/pdf',type='primary')
 st.caption('PDF를 다운로드한 뒤 PDF 뷰어에서 Ctrl+P로 바로 인쇄하실 수 있습니다.')

elif page=='일괄 인쇄·저장':
 st.subheader('전체 성적표 및 엑셀 다운로드')
 only=st.checkbox('11월에 검수 확정된 과목이 하나라도 있는 학생만',value=True)
 include=st.checkbox('정오답 상세 페이지 포함',value=True)
 filtered=[s for s in db['roster'] if not only or any(db['results'].get(s['id'],{}).get(k,{}).get('review')=='확정' for k in CATS)]
 st.write(f'출력 대상: **{len(filtered)}명**')
 if st.button('전체 성적표 PDF 묶음 생성'):
  if not filtered:st.error('출력할 학생이 없습니다. 검수된 학생을 추가하거나 필터를 해제하십시오.')
  else:
   z=io.BytesIO()
   with zipfile.ZipFile(z,'w',zipfile.ZIP_DEFLATED) as archive:
    for s in filtered:archive.writestr(f"{s['student_number']}_{s['name']}.pdf",render_pdf(db,s['id'],include,test_mode=test_mode))
   st.session_state['pdf_zip']=z.getvalue()
 if st.session_state.get('pdf_zip'):
  st.download_button('학생별 PDF ZIP 저장',st.session_state['pdf_zip'],file_name='thepremium-test-reports.zip' if test_mode else 'thepremium-reports.zip',mime='application/zip')
 if st.button('전체 결과 엑셀 생성'):
  try:
   b=export_xlsx(db,test_mode=test_mode);st.session_state['result_xlsx']=b
  except Exception as e:st.error(f'엑셀 생성 실패: {e}')
 if st.session_state.get('result_xlsx'):
  st.download_button('성적 집계 + 월별 등급 엑셀 저장',st.session_state['result_xlsx'],file_name='thepremium-test-results.xlsx' if test_mode else 'thepremium-results.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

elif page=='설정·로고·백업':
 st.subheader('기관 로고 및 성적표 머리말·꼬리말')
 st.write('상단 왼쪽: 횡성인재육성관 / 하단 가운데: 대성 로고+대성학력개발연구소 / 하단 오른쪽: 천명의선택입시연구소')
 for org in ['횡성인재육성관','대성학력개발연구소','천명의선택입시연구소']:
  f=st.file_uploader(org+' 로고',type=['png','jpg','jpeg'],key='logo_'+org)
  if f:
   db['logos'][org]=base64.b64encode(f.getvalue()).decode();save(db);st.success(org+' 로고 등록 완료')
 db['settings']['title']=st.text_input('성적표 제목',db['settings'].get('title',''))
 if st.button('제목 저장'):save(db);st.success('저장했습니다.')
 st.divider();st.subheader('백업 및 복원')
 if test_mode:
  st.info('테스트 모드의 데이터는 운영 백업에 포함되지 않습니다. 운영 백업·복원은 실제 운영 모드에서만 수행하십시오.')
 else:
  st.download_button('모든 결과·명부·정답·설정 JSON 백업',json.dumps(db,ensure_ascii=False,indent=2).encode(),file_name='thepremium-backup.json',mime='application/json')
 restore=st.file_uploader('기존 JSON 백업 불러오기',type=['json'],disabled=test_mode)
 if restore and not test_mode and st.button('백업 복원 실행'):
  try:
   new=json.loads(restore.getvalue().decode());assert all(k in new for k in ['roster','history','results','keys','cuts'])
   st.session_state.db=new;save(new);st.success('복원 완료. 화면을 새로고침하십시오.')
  except Exception as e:st.error(f'복원 실패: {e}')
 st.warning('Streamlit Community Cloud의 data/work.json은 서버 재시작 시 사라질 수 있으며 안전한 영구 저장소가 아닙니다. 매 작업 후 JSON 백업을 내려받고 개인정보 취급 규정을 준수하십시오.')
