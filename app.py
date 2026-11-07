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
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
 rec=db['results'].get(uid,{}).get(subject,{})
 choice=norm(rec.get('choice') or (roster.get(prop,'') if prop else subject))
 answers=rec.get('answers',[]); key=answer_key(db,subject,choice)
 n=N[which(subject)]
 if rec.get('manual_score') is not None and rec.get('review')=='확정':
  manual=int(rec['manual_score'])
  ceiling=50 if subject.startswith('탐구') or subject=='한국사' else 100
  if 0<=manual<=ceiling:
   return {'choice':choice,'answers':answers,'correct':[], 'score':manual, 'grade':grade_of(subject,choice,manual,db['cuts']), 'ready':True,'review':'확정(점수수정)','key':key}
 if (subject in ('국어','수학','탐구1','탐구2') and not choice) or len(key)!=n or not db.get('key_verified',{}).get(which(subject)+'|'+norm(choice),False) or any(str(k.get('answer','')).strip()=='' or not isinstance(k.get('points'),int) or k['points']<=0 for k in key):
  return {'choice':choice,'answers':answers,'correct':[],'score':None,'grade':None,'ready':False,'review':rec.get('review','미채점'),'key':key}
 if not answers or len(answers)!=n or rec.get('review')!='확정':
  return {'choice':choice,'answers':answers,'correct':[],'score':None,'grade':None,'ready':False,'review':rec.get('review','미채점'),'key':key}
 correct=[];total=0
 for a,k in zip(answers,key):
  ok=str(a).strip()==str(k['answer']).strip() if str(a).strip() else False
  correct.append('○' if ok else '×')  # 미표기 역시 오답·0점
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
 ss.append(['학생번호','학생명','전화번호','11월 수험번호','학교','학년','국어선택','국어점수','국어잠정등급','수학선택','수학점수','수학잠정등급','영어점수','영어등급','한국사등급','탐구1등급','탐구2등급'])
 for stu in db['roster']:
  uid=stu['id'];r={q:scored(db,uid,q) for q in CATS}
  ss.append([stu.get('student_number'),stu['name'],stu.get('phone',''),stu.get('exam_number',''),stu.get('school',''),stu.get('grade','고3'),r['국어']['choice'],r['국어']['score'],r['국어']['grade'],r['수학']['choice'],r['수학']['score'],r['수학']['grade'],r['영어']['score'],r['영어']['grade'],r['한국사']['grade'],r['탐구1']['grade'],r['탐구2']['grade']])
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



# 하나의 시트가 11월 정정의 유일한 입력원입니다. 다른 시트는 열람만 가능합니다.
EXCEL_COLUMNS=['적용(Y)','내부학생ID','학생명','전화번호','수험번호','학교','영역','11월 응시과목','검수상태','원점수(참고)','점수직접수정(선택)','수정사유','판독파일','PDF쪽']+[f'{i}번' for i in range(1,46)]
def phone_digits(v):return re.sub(r'\D','',str(v or ''))
def text_id(v):
 if v is None:return ''
 if isinstance(v,(int,float)) and float(v).is_integer():return str(int(v))
 return str(v).strip()
def editable_workbook(db,test_mode=False):
 wb=Workbook();ws=wb.active;ws.title='답안수정'
 ws.append(EXCEL_COLUMNS)
 for stu in db['roster']:
  uid=stu['id']
  for sub in CATS:
   r=db.get('results',{}).get(uid,{}).get(sub,{})
   sr=scored(db,uid,sub)
   answers=list(r.get('answers',[]))
   answers=(answers+['']*45)[:45]
   ws.append(['Y' if r.get('review')=='확정' else '',uid,stu.get('name',''),stu.get('phone',''),stu.get('exam_number') or stu.get('student_number',''),stu.get('school',''),sub,r.get('choice') or sr.get('choice') or '',r.get('review','미등록'),sr.get('score'),r.get('manual_score'),r.get('manual_reason',''),r.get('source',''),r.get('page','')]+answers)
 ws.freeze_panes='O2';ws.auto_filter.ref=ws.dimensions
 ws.row_dimensions[1].height=33
 for c in ws[1]:c.fill=PatternFill('solid',fgColor='123A56');c.font=Font(name='맑은 고딕',size=10,bold=True,color='FFFFFF');c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
 widths={'A':12,'B':22,'C':14,'D':20,'E':17,'F':18,'G':13,'H':22,'I':15,'J':17,'K':21,'L':25,'M':23,'N':10}
 for key,value in widths.items():ws.column_dimensions[key].width=value
 for col in range(15,60):ws.column_dimensions[ws.cell(1,col).column_letter].width=8
 for row in ws.iter_rows(min_row=2):
  for cell in row:
   cell.font=Font(name='맑은 고딕',size=10)
   if cell.column in (1,8,11,12) or cell.column>=15:cell.fill=PatternFill('solid',fgColor='FFF6D8')
  row[3].number_format='@';row[4].number_format='@'
 guide=wb.create_sheet('사용안내');instructions=[
  'thepremium · 답안 수정 업로드 안내',
  '1. 수정할 시트는 오직 「답안수정」입니다. 다른 시트는 편집하지 마십시오.',
  '2. OMR 판독 완료 시 엑셀 업로드는 필요하지 않습니다.',
  '3. 수정할 학생·영역 행의 적용(Y) 칸에 Y를 입력합니다. 빈칸 행은 가져오지 않습니다.',
  '4. 1번~45번 칸의 답안을 수정합니다. 미표기는 빈칸을 유지하며 0점(오답) 처리됩니다.',
  '5. 수학 단답형은 해당 문항에 숫자 전체를 입력합니다.',
  '6. 11월 응시과목이 이전 과목과 달라도 「11월 응시과목」을 올바르게 선택해 정산합니다.',
  '7. 원점수 자체를 변경해야 하는 예외 상황에는 점수직접수정(선택)에 수치를 적고 수정사유를 필수 입력합니다.',
  '8. 점수직접수정에 값이 있으면 답안 채점점수 대신 입력 점수로 10월 컷 잠정등급을 계산합니다. 정오답 표시는 직접 수정점수와 별개입니다.',
  '9. 점수직접수정을 해제하려면 빈칸으로 만들고 답안·정답 기준 재채점을 합니다.',
  '10. 학생명은 반드시 일치해야 하며 전화번호/수험번호는 보조 식별키입니다. 동명이인·번호 충돌은 거부합니다.',
  '11. 3~10월 기록은 수정되지 않습니다. 11월 성적만 변경됩니다.',
  '12. 업로드 전 JSON 백업을 내려받고, 검증 결과 확인 후 적용하십시오.',
  '13. 이 파일에는 개인정보가 포함될 수 있으므로 공개 GitHub에 업로드하지 마십시오.',
 ]
 for item in instructions:guide.append([item])
 guide.column_dimensions['A'].width=125
 guide['A1'].font=Font(name='맑은 고딕',size=14,bold=True,color='123A56')
 for i in range(2,len(instructions)+1):guide.cell(i,1).font=Font(name='맑은 고딕',size=11)
 if test_mode:guide.append(['테스트 모드에서 내보낸 파일은 실제 운영용이 아닙니다.'])
 out=io.BytesIO();wb.save(out);return out.getvalue()

def _find_student(db,uid,name,phone,exam):
 name=str(name or '').strip();phone=phone_digits(phone);exam=text_id(exam)
 if not name:raise ValueError('학생명 누락')
 students=[s for s in db['roster'] if str(s.get('name','')).strip()==name]
 if not students:raise ValueError(f'명부에 없는 학생명: {name}')
 if uid:
  students=[s for s in students if s['id']==uid]
  if not students:raise ValueError(f'학생 ID/이름 불일치: {name} ({uid})')
 if phone:
  with_phone=[s for s in students if phone_digits(s.get('phone'))==phone]
  if with_phone:students=with_phone
  elif any(phone_digits(s.get('phone')) for s in students):raise ValueError(f'전화번호 불일치: {name}')
 if exam:
  with_exam=[s for s in students if exam in (text_id(s.get('exam_number')),text_id(s.get('student_number')))]
  if with_exam:students=with_exam
  elif any(text_id(s.get('exam_number')) for s in students):raise ValueError(f'수험번호 불일치: {name}')
 if len(students)!=1:raise ValueError(f'동명이인/식별정보 충돌: {name} (확인 필요 {len(students)}명)')
 return students[0]

def parse_corrections(db,content):
 wb=load_workbook(io.BytesIO(content),read_only=True,data_only=True)
 if '답안수정' not in wb.sheetnames:raise ValueError('「답안수정」 시트가 없습니다.')
 ws=wb['답안수정'];headers=[str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1,max_row=1))]
 if headers[:len(EXCEL_COLUMNS)]!=EXCEL_COLUMNS:raise ValueError('열 제목 또는 순서가 바뀌었습니다. 원본 수정용 엑셀을 사용하십시오.')
 changes=[];errors=[];seen=set()
 for idx,row in enumerate(ws.iter_rows(min_row=2,values_only=True),start=2):
  if str(row[0] or '').strip().upper() not in ('Y','YES','1','적용'):continue
  try:
   uid,name,phone,exam,school,sub,choice=(row[i] if i<len(row) else None for i in (1,2,3,4,5,6,7))
   stu=_find_student(db,str(uid or '').strip(),name,phone,exam)
   sub=str(sub or '').strip()
   if sub not in CATS:raise ValueError('교과/영역이 잘못되었습니다')
   key=(stu['id'],sub)
   if key in seen:raise ValueError('동일 학생·영역이 두 번 입력되었습니다')
   seen.add(key)
   choice=norm(choice)
   if sub not in ('영어','한국사') and choice not in OPTIONS['탐구' if sub.startswith('탐구') else sub]:raise ValueError('11월 선택과목이 유효하지 않습니다')
   if sub in ('영어','한국사'):choice=sub
   n=N[which(sub)];answers=[]
   for q in range(n):
    raw=row[14+q] if 14+q<len(row) else None
    val=text_id(raw)
    if sub=='수학' and q+1 in list(range(16,23))+[29,30]:
     if val and (not val.isdigit() or len(val)>3):raise ValueError(f'{q+1}번 단답형은 0~999 숫자만 입력')
    elif val and val not in ('1','2','3','4','5'):raise ValueError(f'{q+1}번 객관식 답안은 1~5 또는 빈칸')
    answers.append(val)
   manual=row[10] if len(row)>10 else None
   reason=str(row[11] or '').strip() if len(row)>11 else ''
   if manual not in (None,''):
    try:
     mm=float(manual)
     if not mm.is_integer():raise ValueError()
     manual=int(mm)
    except:raise ValueError('점수직접수정은 정수만 가능합니다')
    maxscore=50 if sub.startswith('탐구') or sub=='한국사' else 100
    if not 0<=manual<=maxscore:raise ValueError(f'점수직접수정 범위는 0~{maxscore}점')
    if not reason:raise ValueError('직접 점수 수정 시 수정사유 필수')
   else:manual=None
   existing=db.get('results',{}).get(stu['id'],{}).get(sub,{})
   record=dict(existing,answers=answers,choice=choice,review='확정',source='엑셀 답안수정',manual_score=manual,manual_reason=reason)
   changes.append({'학생':name,'학교':stu.get('school',''),'영역':sub,'11월선택':choice,'수정문항':sum(a!=b for a,b in zip(answers,(existing.get('answers',[])+['']*n)[:n])),'점수직접수정':manual if manual is not None else '-','uid':stu['id'],'record':record})
  except Exception as ex:errors.append(f'{idx}행: {ex}')
 return changes,errors

# 원본 정답 통합 엑셀: 국어/수학 공통과 선택을 결합한 후 채점기에 등록합니다.
KEY_FILE=ROOT/'reference'/'thepremium-answers-11.xlsx'
def parse_november_keys(contents):
 workbook=load_workbook(io.BytesIO(contents),read_only=True,data_only=True)
 if '정답통합' not in workbook.sheetnames:raise ValueError('「정답통합」 시트가 없습니다.')
 rows=list(workbook['정답통합'].iter_rows(min_row=2,values_only=True))
 groups={}
 for vals in rows:
  if not vals or vals[0] is None:continue
  sub=str(vals[0]).strip();choice=str(vals[1]).strip();q=int(vals[3]);ans=text_id(vals[4]);points=int(vals[5]);
  if not ans or points<=0:raise ValueError(f'{sub} {choice} {q}번 정답/배점 오류')
  groups.setdefault((sub,choice),{})[q]={'question':q,'answer':ans,'points':points}
 results={}
 for sub,options in [('국어',OPTIONS['국어']),('수학',OPTIONS['수학'])]:
  common=groups.get((sub,'공통'),{})
  for choice in options:
   selection=groups.get((sub,choice),{})
   combined={**common,**selection};n=N[sub]
   if sorted(combined)!=list(range(1,n+1)):
    raise ValueError(f'{sub} {choice} 문항번호 누락/중복')
   key=[combined[i] for i in range(1,n+1)]
   if sum(x['points'] for x in key)!=100:raise ValueError(f'{sub} {choice} 배점합계 오류')
   results[f'{sub}|{choice}']=key
 for sub,choice in groups:
  if sub in ('국어','수학'):continue
  target='탐구' if sub in ('탐구','사회탐구','과학탐구') else sub
  if target=='탐구':name=norm(choice)
  elif target in ('영어','한국사'):name=target
  else:continue
  keydata=groups[(sub,choice)];n=N[target]
  if sorted(keydata)!=list(range(1,n+1)):raise ValueError(f'{target} {name} 문항번호 오류')
  key=[keydata[i] for i in range(1,n+1)]
  limit=50 if target in ('한국사','탐구') else 100
  if sum(x['points'] for x in key)!=limit:raise ValueError(f'{target} {name} 배점합계 오류')
  results[f'{target}|{name}']=key
 return results

# OMR 검수는 단일 엑셀 파일의 「판독검수」 시트에서만 수행합니다.
REVIEW_FIELDS=['검수확인(Y)','PDF쪽','영역','OMR이름후보','확정학생명','전화번호','수험번호','OMR선택과목후보','11월확정선택과목','판독상태','원본파일','처리(Y/제외)']+[f'{q}번' for q in range(1,46)]+['선택과목정정사유']

def batch_review_excel(rows,source_name):
 wb=Workbook();ws=wb.active;ws.title='판독검수';ws.append(REVIEW_FIELDS)
 for r in rows:
  n=N[which(r['subject'])];candidate=norm(r.get('explore_name') if r['subject'].startswith('탐구') else r.get('selection_choice'))
  if r['subject'] in ('영어','한국사'):candidate=r['subject']
  vals=[str(v or '') for v in (r.get('answers',[]) + ['']*45)[:45]]
  ws.append(['',r['page'],r['subject'],r.get('name_candidate',''),r.get('name_candidate',''),'','',candidate,candidate,
    '판독실패: '+r['error'] if r.get('error') else '확인필요' if r.get('suspect') else '자동판독',source_name,'Y']+vals+[''])
 ws.freeze_panes='M2';ws.auto_filter.ref=ws.dimensions
 ws.row_dimensions[1].height=32
 for cell in ws[1]:cell.fill=PatternFill('solid',fgColor='17324D');cell.font=Font(name='맑은 고딕',bold=True,color='FFFFFF');cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
 for letter in 'ABCDEFGHIJKL':ws.column_dimensions[letter].width=19 if letter in 'DEGHIJK' else 14
 ws.column_dimensions['K'].width=34
 for col in range(13,58):ws.column_dimensions[ws.cell(1,col).column_letter].width=8
 ws.column_dimensions['BF'].width=35
 for row in ws.iter_rows(min_row=2):
  for cell in row:
   cell.font=Font(name='맑은 고딕',size=10)
   if cell.column in (1,5,6,7,9,12) or cell.column>=13:cell.fill=PatternFill('solid',fgColor='FFF2C7')
  row[5].number_format='@';row[6].number_format='@'
 guide=wb.create_sheet('검수안내')
 for line in [
  'thepremium · OMR 전체 PDF 판독 후 엑셀 검수',
  '수정 시트: 판독검수 (이 시트만 수정)',
  '1. 원본 PDF와 각 행의 PDF쪽·교과·판독답안을 대조하십시오.',
  '2. 확정학생명에는 원본의 이름을 정확히 입력하고, 전화번호·수험번호는 가능한 경우 입력하십시오.',
  '3. 11월확정선택과목은 OMR 마킹으로 인식한 과목을 확인하십시오. 이전 선택과목은 적용하지 않습니다.',
  '4. 1번~45번의 답안을 확인/수정하십시오. 객관식은 1~5, 수학 단답형은 0~999, 미표기는 빈칸입니다.',
  '5. OMR 과목 판독이 잘못됐다면 11월확정선택과목을 정정하고 마지막 열 「선택과목정정사유」를 입력하십시오.',
  '6. 검수가 완료된 행의 검수확인(Y) 칸에 Y를 입력하십시오. 이 표시가 없는 행은 최종 반영할 수 없습니다.',
  '6. 처리(Y/제외)를 제외로 입력한 행은 저장하지 않습니다. 판독에 실패한 행은 제외 후 재스캔하십시오.',
  '7. 11월 정답·배점이 검증 확정된 과목만 점수를 산출합니다. 미표기는 오답 0점입니다.',
  '8. 같은 학생·영역이 중복된 PDF쪽 또는 엑셀 행에 있으면 자동 확정하지 않습니다.',
  '9. 업로드 후 검증 결과를 확인하고 최종 확정합니다. 실제 운영 중에는 전체 JSON 백업을 권장합니다.',
  '10. 학생정보가 포함되므로 이 엑셀은 공개 GitHub에 올리지 마십시오.'
 ]:guide.append([line])
 guide.column_dimensions['A'].width=120;guide['A1'].font=Font(name='맑은 고딕',size=15,bold=True,color='17324D')
 out=io.BytesIO();wb.save(out);return out.getvalue()

def verify_batch_review(db,content,rows,source_name):
 wb=load_workbook(io.BytesIO(content),read_only=True,data_only=True)
 if '판독검수' not in wb.sheetnames:raise ValueError('「판독검수」 시트가 없습니다.')
 ws=wb['판독검수']; data=list(ws.iter_rows(values_only=True))
 if not data or [str(v or '').strip() for v in data[0][:len(REVIEW_FIELDS)]]!=REVIEW_FIELDS:
  raise ValueError('검수 엑셀의 열 제목/순서가 변경되었습니다.')
 expected={(r['page'],r['subject']) for r in rows}
 changes=[];errors=[];seen=set();processed=set()
 for idx,record in enumerate(data[1:],2):
  if not any(v is not None for v in record):continue
  try:
   page=int(record[1]);sub=str(record[2] or '').strip();token=(page,sub)
   if token not in expected:raise ValueError('현재 PDF에서 판독한 페이지·영역이 아닙니다')
   if token in processed:raise ValueError('페이지·영역 중복 행')
   processed.add(token)
   original=next(r for r in rows if (r['page'],r['subject'])==token)
   if str(record[11] or '').strip().lower() in ('제외','skip','n','no'):continue
   if original.get('error'):raise ValueError('판독 실패. 재스캔 또는 이 행을 제외하십시오')
   if str(record[0] or '').strip().upper()!='Y':raise ValueError('원본 대조 후 검수확인(Y)을 입력해야 합니다')
   name=str(record[4] or '').strip();phone=str(record[5] or '').strip();exam=text_id(record[6])
   stu=_find_student(db,'',name,phone,exam)
   choice=norm(record[8])
   if sub in ('영어','한국사'):choice=sub
   elif choice not in OPTIONS['탐구' if sub.startswith('탐구') else sub]:raise ValueError('11월 선택과목 미확정/잘못된 과목')
   # PDF 원본에서 인식된 선택과목이 명확하다면 임의로 다른 과목으로 바꾸는 것을 막음
   recognized=norm(original.get('explore_name') if sub.startswith('탐구') else original.get('selection_choice'))
   confident=original.get('explore_code_confident') if sub.startswith('탐구') else original.get('selection_confident')
   if confident and recognized and choice!=recognized and not str((record[57] if len(record)>57 else '') or '').strip():raise ValueError(f'OMR 판독 과목({recognized})과 엑셀({choice}) 불일치: 마지막 열에 선택과목정정사유를 입력하십시오')
   n=N[which(sub)];answers=[]
   for q in range(1,n+1):
    raw=record[11+q] if 11+q<len(record) else None
    val=text_id(raw)
    if sub=='수학' and (16<=q<=22 or q in (29,30)):
     if val and (not val.isdigit() or len(val)>3):raise ValueError(f'{q}번: 수학 단답은 0~999 정수')
    elif val and val not in ('1','2','3','4','5'):raise ValueError(f'{q}번: 객관식은 1~5 또는 공란')
    answers.append(val)
   key=(stu['id'],sub)
   if key in seen:raise ValueError('동일 학생·영역 중복 반영')
   seen.add(key)
   changes.append({'학생':stu['name'],'학교':stu.get('school',''),'영역':sub,'11월선택':choice,'PDF쪽':page,
     '미표기':sum(not x for x in answers),'uid':stu['id'],
     'record':{'answers':answers,'choice':choice,'review':'확정','source':source_name,'page':page}})
  except Exception as ex:errors.append(f'{idx}행: {ex}')
 if len(processed)!=len(expected):errors.append('PDF 판독 결과 일부 페이지/영역이 검수 엑셀에서 누락되었습니다')
 return changes,errors

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
   if uid not in roster:roster[uid]={'id':uid,'student_number':num,'exam_number':num,'phone':'','name':name,'school':school,'grade':'고3'}
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
 old={x['id']:x for x in db.get('roster',[])}
 # 11월 신규 등록자는 기존 3~10월 자료를 갱신해도 유지합니다.
 newcomers=[copy.deepcopy(x) for x in db.get('roster',[]) if x.get('november_new')]
 for st_ in record['roster']:
  prev=old.get(st_['id'],{})
  for field in ('phone','exam_number'):
   if prev.get(field):st_[field]=prev[field]
 used={s['id'] for s in record['roster']}
 for newstu in newcomers:
  if newstu['id'] not in used:record['roster'].append(newstu)
 allowed={s['id'] for s in record['roster']}
 db['results']={uid:r for uid,r in db.get('results',{}).items() if uid in allowed}
 db['roster']=record['roster'];db['history']=record['history']
 return len(allowed)

# 11월 신규 응시자 등록: 3~10월 이력/OMR 결과를 덮어쓰지 않는 추가 명부.
NEW_ROSTER_COLUMNS=['학생명','학교','학년','11월 수험번호','전화번호','국어 선택','수학 선택','탐구Ⅰ 선택','탐구Ⅱ 선택']
def november_roster_template():
 wb=Workbook();ws=wb.active;ws.title='11월신규명부';ws.append(NEW_ROSTER_COLUMNS)
 for _ in range(25):ws.append(['']*len(NEW_ROSTER_COLUMNS))
 ws.freeze_panes='A2';ws.auto_filter.ref=f'A1:I26'
 for cell in ws[1]:
  cell.fill=PatternFill('solid',fgColor='123A56');cell.font=Font(name='맑은 고딕',bold=True,color='FFFFFF');cell.alignment=Alignment(horizontal='center')
 for col in 'ABCDEFGHI':ws.column_dimensions[col].width=20
 ws.column_dimensions['A'].width=16;ws.column_dimensions['B'].width=23
 for row in ws.iter_rows(min_row=2):
  row[3].number_format='@';row[4].number_format='@'
 helpws=wb.create_sheet('사용안내')
 for line in ['신규 응시자 명부는 「11월신규명부」 시트에서만 작성하십시오.',
 '학생명·학교·11월 수험번호/전화번호 중 하나 이상을 입력하십시오. 식별정보는 개인정보이므로 외부 공개 금지.',
 '기존 3~10월 명부와 동일 학생은 별도 신규로 생성하지 않고 연결됩니다.',
 '동명이인 또는 서로 다른 학생이 동일 전화번호/수험번호를 쓰는 경우 등록하지 않고 오류를 알려줍니다.',
 '과목명 칸은 11월 채점과 매칭에 사용하지 않습니다. 학생명·전화번호·수험번호로 식별하고 OMR 선택과목을 적용합니다.',
 '기존 월별 성적은 바뀌지 않으며 신규학생의 3~10월은 미응시로 표시합니다.']:
  helpws.append([line])
 helpws.column_dimensions['A'].width=115
 out=io.BytesIO();wb.save(out);return out.getvalue()

def parse_november_roster(db,content):
 wb=load_workbook(io.BytesIO(content),read_only=True,data_only=True)
 if '11월신규명부' not in wb.sheetnames:raise ValueError('「11월신규명부」 시트가 없습니다.')
 ws=wb['11월신규명부'];headers=[str(c.value or '').strip() for c in next(ws.iter_rows(max_row=1))]
 if headers[:9]!=NEW_ROSTER_COLUMNS:raise ValueError('신규명부의 열 제목 또는 순서가 다릅니다.')
 roster=copy.deepcopy(db['roster']);changed=[];errors=[]
 for lineno,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
  vals=list(row[:9])+[None]*max(0,9-len(row))
  if not any(v is not None and str(v).strip() for v in vals):continue
  try:
   name,school,grade,exam,phone,korean,math,explore1,explore2=vals[:9]
   name=str(name or '').strip();school=str(school or '').strip()
   grade=str(grade or '고3').strip();exam=text_id(exam);phone=phone_digits(phone)
   if not name:raise ValueError('학생명 필수')
   if not exam and not phone:raise ValueError('수험번호 또는 전화번호를 하나 이상 입력하십시오')
   # 명부 업로드에서는 선택과목을 읽거나 검사하지 않습니다.
   # 11월 선택과목은 OMR 마킹(국어/수학)과 탐구 과목코드로만 판별합니다.
   incoming={}
   # 번호가 다른 학생에게 연결된 경우 절대 덮어쓰지 않습니다.
   for other in roster:
    if other['name']!=name and ((phone and phone==phone_digits(other.get('phone'))) or (exam and exam in (text_id(other.get('exam_number')), text_id(other.get('student_number'))))):
     raise ValueError(f'전화번호 또는 수험번호가 다른 학생({other["name"]})과 중복됩니다')
   same=[x for x in roster if x['name']==name]
   if same:
    if len(same)>1:
     same=[x for x in same if (phone and phone==phone_digits(x.get('phone'))) or (exam and exam in (text_id(x.get('exam_number')),text_id(x.get('student_number'))))]
     if len(same)!=1:raise ValueError('동명이인: 전화번호 또는 수험번호로 학생을 명확히 구분하십시오')
    stu=same[0]
    for field,key in [('phone',phone),('exam_number',exam)]:
     current=phone_digits(stu.get(field)) if field=='phone' else text_id(stu.get(field))
     # 3~10월 수험번호를 기본값으로 가진 기존 학생은 새로운 11월 수험번호로 갱신 가능.
     legacy_exam=(field=='exam_number' and current==text_id(stu.get('student_number')))
     if key and current and current!=key and not legacy_exam:
      raise ValueError(f'기존 {name}의 {field}과 다릅니다. 명부에서 직접 확인하십시오')
    if phone:stu['phone']=phone
    if exam:stu['exam_number']=exam
    if school:stu['school']=school
    stu['grade']=grade
    # 11월 선택과목은 OMR 판독 후 별도로 저장합니다.
    changed.append({'행':lineno,'학생명':name,'처리':'기존 학생 연결','id':stu['id']})
   else:
    import hashlib
    seed=f'{name}|{exam}|{phone}|{school}'
    uid='november|'+hashlib.sha256(seed.encode()).hexdigest()[:18]
    if any(x['id']==uid for x in roster):raise ValueError('신규 ID 중복')
    stu={'id':uid,'student_number':exam or uid,'exam_number':exam,'phone':phone,'name':name,'school':school,'grade':grade,'november_new':True,**incoming}
    roster.append(stu)
    changed.append({'행':lineno,'학생명':name,'처리':'신규 등록','id':uid})
  except Exception as e:errors.append(f'{lineno}행: {e}')
 return roster,changed,errors

# 모든 응시자(기존 + 11월 신규)를 하나의 편집 가능한 파일로 관리합니다.
MASTER_COLUMNS=['학생ID(기존유지)','학생명','학교','학년','수험번호','전화번호','국어 선택','수학 선택','탐구Ⅰ 선택','탐구Ⅱ 선택']
def complete_roster_workbook(db):
 wb=Workbook();ws=wb.active;ws.title='전체응시자명부';ws.append(MASTER_COLUMNS)
 for stu in db['roster']:
  ws.append([stu['id'],stu['name'],stu.get('school',''),stu.get('grade','고3'),
             text_id(stu.get('exam_number') or stu.get('student_number')),phone_digits(stu.get('phone')),
             norm(stu.get('korean')),norm(stu.get('math')),norm(stu.get('explore1')),norm(stu.get('explore2'))])
 for row in ws.iter_rows(min_row=2):
  row[4].number_format='@';row[5].number_format='@'
 for c in ws[1]:
  c.fill=PatternFill('solid',fgColor='123A56');c.font=Font(name='맑은 고딕',bold=True,color='FFFFFF')
 for col in 'ABCDEFGHIJ':ws.column_dimensions[col].width=21
 ws.column_dimensions['A'].width=29
 ws.freeze_panes='B2';ws.auto_filter.ref=f'A1:J{max(2,ws.max_row)}'
 info=wb.create_sheet('사용안내')
 for message in ['수정은 오직 「전체응시자명부」 시트에서 합니다.',
 '기존 학생의 학생ID(기존유지)는 수정하지 마십시오. 3~10월 누적 성적과 연결되는 고유키입니다.',
 '11월 신규 응시자는 맨 아래 빈 행에 추가하고 학생ID 칸을 비워두십시오.',
 '기존 학생의 이름을 변경하더라도 학생ID가 같으면 3~10월 성적이 유지됩니다.',
 '전화번호·수험번호는 앞의 0이 보존되도록 텍스트로 입력하십시오.',
 '명부의 선택과목 입력은 무시됩니다. 11월 응시과목은 실제 OMR 마킹에서만 판별합니다.',
 '이 파일은 민감한 학생정보가 있으므로 공개 GitHub에 올리지 마십시오.',
 '기존 학생을 시트에서 삭제해도 프로그램에서는 삭제되지 않습니다.']:info.append([message])
 info.column_dimensions['A'].width=118
 out=io.BytesIO();wb.save(out);return out.getvalue()

def preview_complete_roster(db, content):
 wb=load_workbook(io.BytesIO(content),read_only=True,data_only=True)
 if '전체응시자명부' not in wb.sheetnames:raise ValueError('「전체응시자명부」 시트를 찾을 수 없습니다.')
 ws=wb['전체응시자명부'];headers=[str(c.value or '').strip() for c in next(ws.iter_rows(max_row=1))]
 if headers[:len(MASTER_COLUMNS)]!=MASTER_COLUMNS:raise ValueError('전체응시자명부 열 이름·순서를 변경하지 마십시오.')
 original={str(item['id']):item for item in db['roster']}
 roster=copy.deepcopy(db['roster']); byid={str(v['id']):v for v in roster}
 errors=[];summary=[];seen=set()
 for line,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
  values=(list(row[:10])+[None]*10)[:10]
  if not any(v is not None and str(v).strip() for v in values):continue
  try:
   uid,name,school,grade,exam,phone,korean,math,ex1,ex2=values
   uid=str(uid or '').strip();name=str(name or '').strip();school=str(school or '').strip()
   grade=str(grade or '고3').strip();exam=text_id(exam);phone=phone_digits(phone)
   if not name:raise ValueError('학생명 필수')
   if not exam and not phone:raise ValueError('전화번호 또는 수험번호 입력 필요')
   if uid and uid not in original:raise ValueError('등록되지 않은 학생ID. 신규 학생은 학생ID 칸을 비워두십시오')
   if uid and uid in seen:raise ValueError('같은 학생ID가 파일에 중복되어 있습니다')
   # 학생명·전화번호·수험번호만 명부 매칭에 사용합니다.
   # 엑셀에 적힌 과목명은 검증 대상이 아니며 기존 선택기록도 덮어쓰지 않습니다.
   incoming={}
   if not uid:
    samephone=[p for p in roster if phone and phone==phone_digits(p.get('phone'))]
    samename=[p for p in roster if p['name']==name]
    if samephone:
     if len(samephone)!=1 or samephone[0]['name']!=name:raise ValueError('전화번호가 다른 학생명과 중복됩니다')
     uid=str(samephone[0]['id'])
    elif samename:
     if len(samename)!=1:raise ValueError('동명이인입니다. 기존 학생ID를 확인하십시오')
     candidate=samename[0]
     existing_phone=phone_digits(candidate.get('phone'))
     if existing_phone and phone and existing_phone!=phone:raise ValueError('학생명 동일·전화번호 불일치: 수동 확인 필요')
     uid=str(candidate['id'])
    else:
     import hashlib
     uid='november|'+hashlib.sha256(f'{name}|{phone}|{exam}|{school}'.encode()).hexdigest()[:18]
     if uid in byid:raise ValueError('학생ID 중복')
     new={'id':uid,'student_number':exam or uid,'name':name,'november_new':True}
     roster.append(new);byid[uid]=new
   if uid in seen:raise ValueError('동일 학생에 대해 두 행 이상 작성되었습니다')
   for other in roster:
    if str(other['id'])==uid:continue
    if phone and phone==phone_digits(other.get('phone')):raise ValueError(f'전화번호 중복: {other["name"]}')
    other_exam=text_id(other.get('exam_number'))
    if exam and other_exam and exam==other_exam:raise ValueError(f'수험번호 중복: {other["name"]}')
   item=byid[uid]
   before=(item.get('name'),item.get('phone'),item.get('exam_number'))
   item.update(name=name,school=school,grade=grade,phone=phone,exam_number=exam,**incoming)
   seen.add(uid)
   summary.append({'엑셀행':line,'학생명':name,'처리':'신규 학생' if uid not in original else ('기존 학생 정보수정' if before!=(name,phone,exam) else '기존 학생 연결')})
  except Exception as exc:errors.append(f'{line}행: {exc}')
 return roster,summary,errors

def preview_roster_excel_auto(db, content):
 """Accept both master edit and November roster Excel uploads without guessing sheet names.
 Existing students are updated by stable ID, then by unambiguous name+identifiers.
 Unlisted students and all 3~10월 / 11월 marks are preserved.
 """
 workbook=load_workbook(io.BytesIO(content),read_only=True,data_only=True)
 sheets=set(workbook.sheetnames)
 workbook.close()
 if '전체응시자명부' in sheets:
  roster,log,errors=preview_complete_roster(db,content)
  return roster,log,errors,'전체응시자명부', '기존 학생 정보 덮어쓰기 및 신규 등록'
 if '11월신규명부' in sheets:
  roster,log,errors=parse_november_roster(db,content)
  return roster,log,errors,'11월신규명부', '기존 학생 매칭·갱신 및 신규 응시자 추가'
 raise ValueError('명부 시트를 찾지 못했습니다. 「전체응시자명부」 또는 「11월신규명부」 시트가 있는 파일을 업로드하십시오. 원본 3~10월 성적 엑셀은 위쪽 「성적 엑셀 검증 및 학생 이력 갱신」에서 등록하십시오.')

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
page=st.sidebar.radio('메뉴',['11월 정답·배점','학생 명부','OMR 판독·검수','엑셀 답안수정','학생 성적표','일괄 인쇄·저장','설정·로고·백업'])
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

 st.divider()
 st.subheader('전체 응시자 명부 · 엑셀 덮어쓰기/추가')
 st.info('아래 한 곳에서 「전체응시자명부」와 「11월신규명부」 두 가지 엑셀 파일을 자동 구분합니다. 기존 학생은 이름·전화번호·수험번호 또는 고유 학생ID로 연결하고, 신규 학생은 추가합니다. 3~10월 성적과 11월 확정 답안은 삭제하지 않습니다.')
 st.download_button('전체 응시자 명부 엑셀 다운로드 (수정용)',complete_roster_workbook(db),file_name='thepremium-roster-all.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',key='master_roster_download')
 master_file=st.file_uploader('전체 명부 덮어쓰기 또는 11월 응시자 명부 업로드 (.xlsx)',type=['xlsx'],key='master_roster_upload')
 if master_file:
  try:
   suggested,log,problems,sheet_name,method=preview_roster_excel_auto(db,master_file.getvalue())
   st.caption(f'인식된 시트: {sheet_name} · 처리 방식: {method}')
   if problems:
    st.error(f'명부 검사 오류 {len(problems)}건. 기존 명부와 성적은 전혀 변경하지 않았습니다.')
    for reason in problems:st.write(reason)
   else:
    st.success(f'검증 완료: {len(log)}건 / 반영 후 총 {len(suggested)}명')
    st.dataframe(log,hide_index=True,use_container_width=True)
    st.warning('기존 학생 명부의 정보가 변경됩니다. 업로드 파일에 없는 학생은 삭제하지 않으며, 기존 3~10월·11월 성적은 그대로 유지됩니다.')
    import hashlib
    signature=hashlib.sha256(master_file.getvalue()).hexdigest()[:16]
    if st.checkbox('명부 변경 내용과 기존 학생 매칭을 확인했습니다.',key='master_roster_confirm_'+signature):
     if st.button('검증한 엑셀로 학생 명부 덮어쓰기·추가',type='primary',key='master_roster_apply_'+signature):
      db['roster']=suggested
      save(db)
      st.success('학생 명부 갱신 완료. 누적 성적과 확정 OMR 결과는 보존했습니다.')
      st.rerun()
  except Exception as exc:st.error('전체 명부 업로드 오류: '+str(exc))
 st.divider()
 st.subheader('11월 신규 응시자 · 엑셀 명부 일괄 등록')
 st.caption('기존 학생명과 대조하여 연결하고, 신규 학생만 추가합니다. 과거 성적 기록은 변경하지 않습니다.')
 st.download_button('thepremium-roster.xlsx 양식 다운로드',data=november_roster_template(),file_name='thepremium-roster.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
 added_file=st.file_uploader('11월신규명부 시트가 작성된 엑셀(.xlsx) 업로드',type=['xlsx'],key='november_roster_file')
 if added_file:
  try:
   proposed,summary,issues=parse_november_roster(db,added_file.getvalue())
   if issues:
    st.error(f'명부 확인 필요 {len(issues)}건. 오류가 있으면 아무도 등록하지 않습니다.')
    for issue in issues:st.write(issue)
   else:
    st.success(f'명부 검증 통과 · 신규 및 연결 {len(summary)}명')
    st.dataframe(summary,use_container_width=True,hide_index=True)
    import hashlib
    key='apply_roster_'+hashlib.sha256(added_file.getvalue()).hexdigest()[:12]
    if st.checkbox('학생명·전화번호·수험번호를 대조했습니다.',key=key):
     if st.button('11월 학생 명부 등록/갱신',type='primary'):
      db['roster']=proposed;save(db);st.success('11월 명부가 반영되었습니다.');st.rerun()
  except Exception as e:st.error('신규 명부 오류: '+str(e))
 st.write(f"학생 {len(db['roster'])}명 · 기존 이력 및 11월 신규 응시자를 통합 관리합니다.")
 if not db['roster']:st.info('위에서 수정 엑셀을 먼저 업로드하십시오.');st.stop()
 names=[f"{a['student_number']} | {a['name']}" for a in db['roster']]
 selected=st.selectbox('수정할 학생',list(range(len(names))),format_func=lambda i:names[i])
 stu=db['roster'][selected]
 a,b,c=st.columns(3)
 stu['name']=a.text_input('학생명',stu['name']);stu['school']=b.text_input('학교',stu.get('school',''));stu['grade']=c.text_input('학년',stu.get('grade','고3'))
 a,b=st.columns(2)
 stu['phone']=a.text_input('전화번호(학생 식별용)',stu.get('phone',''),key='phone_'+stu['id'])
 stu['exam_number']=b.text_input('11월 수험번호(학생 식별용)',str(stu.get('exam_number') or stu.get('student_number') or ''),key='exam_'+stu['id'])
 st.caption('선택과목은 학생 명부에서 수정하지 않습니다. 11월 OMR 마킹을 자동 판독하여 정답표를 연결합니다.')
 if st.button('학생 정보 저장'):save(db);st.success('저장했습니다.')
 st.dataframe([{'학생번호':x.get('student_number'),'이름':x['name'],'학교':x.get('school',''),'국어':norm(x.get('korean')),'수학':norm(x.get('math')),'탐구Ⅰ':norm(x.get('explore1')),'탐구Ⅱ':norm(x.get('explore2'))} for x in db['roster']],use_container_width=True,hide_index=True)
 st.info('수정 엑셀의 학생번호·이름·학교를 대조해 이력을 연결합니다. 서로 다른 이름이 동일 학생번호를 쓴 경우 별도 인원으로 구분합니다.')

elif page=='11월 정답·배점':
 st.subheader('11월 정답 통합 엑셀 불러오기 · 정산 기준 등록')
 st.caption('국어·수학은 공통+OMR 선택과목 결합, 영어·한국사·탐구는 해당 원본 과목 정답을 읽습니다. 정답 자료를 등록해도 학생은 선택과목을 직접 고르지 않습니다.')
 key_upload=st.file_uploader('검증한 11월 정답·배점 엑셀 업로드 (.xlsx)',type=['xlsx'],key='key_xlsx_import')
 if KEY_FILE.exists():st.download_button('thepremium-answers-11.xlsx 받기',data=KEY_FILE.read_bytes(),file_name='thepremium-answers-11.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
 key_bytes=key_upload.getvalue() if key_upload else (KEY_FILE.read_bytes() if KEY_FILE.exists() else None)
 if key_bytes:
  try:
   verified_keys=parse_november_keys(key_bytes)
   st.success(f'정답표 구조 검사 통과 · {len(verified_keys)}개 과목/선택과목 구성')
   st.dataframe([{'과목':k,'문항':len(v),'배점':sum(t['points'] for t in v)} for k,v in verified_keys.items()],hide_index=True,use_container_width=True)
   correct_key=st.checkbox('원본 정답·배점을 직접 대조했고 이 정답 엑셀을 채점 기준으로 확정합니다.',key='key_xlsx_confirm')
   if st.button('11월 정답 엑셀 검증 후 등록',disabled=not correct_key,type='primary'):
    db.setdefault('keys',{}).update(verified_keys)
    db.setdefault('key_verified',{}).update({k:True for k in verified_keys})
    save(db);st.success('정답·배점 등록 완료 — 이제 검수 엑셀 업로드로 점수/잠정등급이 자동 정산됩니다.')
  except Exception as ex:st.error('정답 엑셀 오류: '+str(ex))
 st.divider()
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
 st.info('교과별 PDF 전체 자동판독 후 학생명·수험번호·전화번호로 매칭합니다. OMR 필기값을 신뢰할 수 없으면 화면에서 확인한 정보를 입력해 연결합니다.')
 subject_group=st.selectbox('업로드할 교과 영역',['국어','수학','영어','한국사','탐구'],key='subject_batch_group')
 st.caption('선택과목은 OMR 마킹을 우선 판독하며, 과거 선택과목이 바뀌어도 11월 현재 응시과목의 정답을 사용합니다. 탐구는 한 페이지의 탐구Ⅰ·탐구Ⅱ를 별도 판독합니다. 수학 단답형은 세 자리 숫자 마킹을 자동 판독하여 후보를 제시하고, 원본 검수 후 확정합니다. 모든 교과의 미표기는 0점입니다.')
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
                       'error':'','reviewed':False,'student_id':'','suspect':meta.get('suspect',False),'detected':meta.get('detected',0),'explore_code':meta.get('explore_code'),'explore_name':meta.get('explore_name',''),'explore_code_confident':meta.get('explore_code_confident',False),
                       'selection_choice':meta.get('selection_choice',''),
                       'selection_confident':meta.get('selection_confident',False),
                       'selection_diagnostics':meta.get('selection_diagnostics',{}),
                       'math_track':meta.get('math_track',''),
                       'math_track_confident':meta.get('math_track_confident',False),
                       'name_candidate':meta.get('name_candidate',''),
                       'name_confident':meta.get('name_confident',False),
                       'name_components':meta.get('name_components',[])})
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
    summary.append({'번호':i+1,'PDF 페이지':r['page'],'교과':LABEL.get(r['subject'],r['subject'])+(' / '+r.get('explore_name','') if r.get('explore_code_confident') else ''),
                    '자동 판독':f'{ok}/{len(r["answers"])}',
                    '확인 필요':len(r['answers'])-ok,
                    '상태':'실패' if r['error'] else ('재판독/확인 필요' if r.get('suspect') else ('검수완료' if r['reviewed'] else '검수대기'))})
   st.dataframe(summary,use_container_width=True,hide_index=True)
   st.divider()
   st.subheader('③ 전체 OMR 판독 결과 엑셀 검수 → 업로드 → 정산')
   st.info('OMR 페이지별 수동 입력 대신 「판독검수」 엑셀에서 모든 학생·과목·답안을 검수하십시오. 인쇄용 성적표는 기존 양식을 유지합니다.')
   st.download_button('thepremium-omr-review.xlsx 다운로드',data=batch_review_excel(rows,uploaded.name),file_name='thepremium-omr-review.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary',key='omr_batch_review_export')
   reviewed_file=st.file_uploader('검수가 끝난 엑셀 업로드 · 판독검수 시트',type=['xlsx'],key=f'omr_excel_review_{fingerprint[:12]}_{mode}')
   if reviewed_file is not None:
    try:
     changes,errors=verify_batch_review(db,reviewed_file.getvalue(),rows,uploaded.name)
     if errors:
      st.error(f'검수 오류 {len(errors)}건 — 저장하지 않았습니다.')
      for err in errors[:40]:st.write(err)
     else:
      st.success(f'검수 확인 완료: 반영 대상 {len(changes)}건 / 미반영 {len(rows)-len(changes)}건')
      if changes:st.dataframe([{k:v for k,v in c.items() if k not in ('uid','record')} for c in changes],hide_index=True,use_container_width=True)
      checked=st.checkbox('원본 OMR과 엑셀을 대조했습니다. 기존 11월 해당 학생·교과 답안이 교체될 수 있음을 확인했습니다.',key=f'confirm_excel_review_{fingerprint[:12]}')
      if st.button('검수 엑셀 확정 업로드 · 11월 자동 정산',disabled=not changes or not checked,type='primary',key=f'commit_excel_review_{fingerprint[:12]}'):
       updated=copy.deepcopy(db['results'])
       for c in changes:updated.setdefault(c['uid'],{})[c['영역']]=c['record']
       db['results']=updated;save(db)
       st.success(f'{len(changes)}건 최종 저장 완료. 확정된 정답·배점이 있는 과목은 자동 정산됩니다. 「학생 성적표」에서 바로 PDF를 출력할 수 있습니다.')
       if test_mode:st.info('테스트 모드 — 실제 성적에는 반영되지 않습니다.')
    except Exception as exc:st.error('엑셀 검수 실패: '+str(exc))
   st.stop()
   st.subheader('③ PDF 페이지별 학생 연결 · 답안 검수')
   roster=db['roster']
   labels=['학생 선택 필요']+[f"{a['student_number']} | {a['name']} | {a.get('school','')}" for a in roster]
   chosen_page=st.selectbox('검수할 PDF 페이지',range(1,pages+1),format_func=lambda j:f'{j}페이지 · {subject_group} OMR')
   indices=[i for i,r in enumerate(rows) if r['page']==chosen_page]
   existing=next((rows[i]['student_id'] for i in indices if rows[i].get('student_id')),'')
   selected_index=next((j+1 for j,v in enumerate(roster) if v['id']==existing),0)
   st.caption('학생명은 우선 확인하며, 동명이인은 전화번호 또는 수험번호로 구별합니다. 필기된 연락처를 자동으로 읽었다고 가정하지 않습니다.')
   match_a,match_b,match_c=st.columns(3)
   first_candidate=next((rows[i].get('name_candidate','') for i in indices if rows[i].get('name_confident') and rows[i].get('name_candidate')), '')
   if first_candidate:
    st.info(f'한글 초성·중성·종성 자동 조합 후보: **{first_candidate}** — 반드시 원본과 대조하여 확인하십시오.')
   else:
    st.warning('OMR 한글 이름의 마킹이 불명확하거나 일부 자모를 판독하지 못했습니다. 학생 이름을 직접 확인하십시오.')
   if st.checkbox('학생 이름 마킹 조합 과정 보기',key=f'name_parts_{chosen_page}_{fingerprint[:8]}'):
    for row_i in indices:
     if rows[row_i].get('name_components'):
      st.dataframe(rows[row_i]['name_components'],hide_index=True,use_container_width=True)
   entered_name=match_a.text_input('OMR 학생명 (후보 수정 가능)',value=first_candidate,key=f'match_name_{chosen_page}_{fingerprint[:8]}')
   if first_candidate:
    suggested_names=[v for v in roster if v.get('name')==first_candidate]
    if len(suggested_names)==1:st.caption('명부 일치 후보: '+first_candidate+' · 해당 학생을 확인 후 연결하십시오.')
    elif len(suggested_names)>1:st.warning('동명이인 후보입니다. 전화번호 또는 수험번호 확인이 필요합니다.')
   entered_phone=match_b.text_input('전화번호(선택)',key=f'match_phone_{chosen_page}_{fingerprint[:8]}')
   entered_exam=match_c.text_input('수험번호(선택)',key=f'match_exam_{chosen_page}_{fingerprint[:8]}')
   widget_key=f'subject_student_{chosen_page}_{fingerprint[:8]}_{subject_group}'
   if st.button('입력한 이름·전화번호·수험번호로 학생 찾기',key=f'match_btn_{chosen_page}_{fingerprint[:8]}'):
    try:
     matched=_find_student(db,'',entered_name,entered_phone,entered_exam)
     selected_index=next(j+1 for j,item in enumerate(roster) if item['id']==matched['id'])
     st.session_state[widget_key]=selected_index
     st.success(f'학생 매칭: {matched["name"]} / {matched.get("school", "")}')
    except Exception as ex:st.error('매칭 실패: '+str(ex))
   selected=st.selectbox(f'{chosen_page}페이지 학생',range(len(labels)),index=selected_index,
                         format_func=lambda j:labels[j],key=widget_key)
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
    if r['subject'] in ('탐구1','탐구2'):
     field='explore1' if r['subject']=='탐구1' else 'explore2'
     student=next((a for a in roster if a['id']==selected_uid),None)
     roster_choice=norm(student.get(field,'')) if student else ''
     auto_name=r.get('explore_name','') if r.get('explore_code_confident') else ''
     st.info(f'탐구 과목: OMR 코드 {r.get("explore_code") if auto_name else "판독 불확실"} / 자동 후보 {auto_name or "확인 필요"} / 학생 명부 {roster_choice or "미연결"}')
     if auto_name and roster_choice and auto_name!=roster_choice:
      st.info('이전 선택과목과 다릅니다. 11월 응시과목을 우선 적용하며 채점을 차단하지 않습니다.')
     if auto_name and not roster_choice:
      st.info('OMR에서 읽은 11월 선택과목을 사용할 수 있습니다.')
    # 11월 응시과목은 이전 3~10월 선택과목과 독립적으로 보관합니다.
    old_student=next((a for a in roster if a['id']==selected_uid),None)
    pickfield={'국어':'korean','수학':'math','탐구1':'explore1','탐구2':'explore2'}.get(r['subject'])
    oldchoice=norm(old_student.get(pickfield,'')) if old_student and pickfield else r['subject']
    detected_choice=(norm(r.get('explore_name','')) if r['subject'].startswith('탐구') and r.get('explore_code_confident')
                     else norm(r.get('selection_choice','')) if r.get('selection_confident') else '')
    if r['subject'] in ('국어','수학'):
     if r.get('selection_confident'):
      st.success(f'OMR 선택과목 자동판독: {detected_choice}'+(f' / 계열: {r.get("math_track")}' if r.get('math_track_confident') else ''))
     else:
      st.warning('OMR 선택과목 마킹이 없거나 판독이 불확실합니다. 검수 엑셀에서 원본과 대조하여 확인해야 하며, 과거 명부의 과목으로 대체하지 않습니다.')
    # 학생 명부에 적힌 과목명과 관계없이 11월 OMR 마킹이 유일한 채점 기준입니다.
    # 불확실한 과목은 임의 보완하지 않고 엑셀 검수에서 처리합니다.
    r['choice']=(r['subject'] if r['subject'] in ('영어','한국사') else detected_choice)
    if r['subject'] not in ('영어','한국사'):
     st.caption('11월 OMR 판독 선택과목: '+(r['choice'] or '판독 불확실 · 검수 필요'))
     if oldchoice and r['choice'] and oldchoice!=r['choice']:
      st.caption(f'과거 명부 표기 {oldchoice}와 달라도 OMR의 {r["choice"]}로 채점합니다.')
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
      r.update(answers=ans,stats=stats,rotation=meta.get('rotation'),error='',reviewed=False,suspect=meta.get('suspect',False),detected=meta.get('detected',0),explore_code=meta.get('explore_code'),explore_name=meta.get('explore_name',''),explore_code_confident=meta.get('explore_code_confident',False),
               selection_choice=meta.get('selection_choice',''),selection_confident=meta.get('selection_confident',False),
               selection_diagnostics=meta.get('selection_diagnostics',{}),
               math_track=meta.get('math_track',''),math_track_confident=meta.get('math_track_confident',False),
               name_candidate=meta.get('name_candidate',''),
               name_confident=meta.get('name_confident',False),
               name_components=meta.get('name_components',[]))
      st.rerun()
     except Exception as e:st.error(str(e))
    if r['stats']:
     uncertain=[x for x in r['stats'] if x.get('상태')!='정상']
     if uncertain:
      st.warning(f'재확인이 필요한 문항 {len(uncertain)}개 (수학 단답형 포함)')
      st.dataframe(uncertain,use_container_width=True,hide_index=True)
     else:st.caption('자동 판독 후보는 모두 정상 판정되었습니다. 실제 OMR과 대조하십시오.')
     if any(str(t.get('상태','')).startswith('미표기') for t in r['stats']):st.caption('전 교과 공통: 실제 미표기는 해당 문항 오답·0점입니다. 답안 수정란에서는 빈칸으로 유지하며, 특히 수학의 숫자 0 마킹과 혼동하지 않습니다.')
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
     math_short_checked = True
     if r['subject']=='수학':
      math_short_checked=st.checkbox('수학 단답형 16~22번·29~30번의 자리별 마킹을 원본과 대조했습니다.',key=f'math_digits_checked_{i}_{fingerprint[:8]}')
     if st.button('이 교과 검수 완료',key=f'subject_review_{i}_{fingerprint[:8]}'):
      if not selected_uid:st.error('학생을 먼저 선택하십시오.')
      elif r['error']:st.error('자동판독 오류를 해결하거나 재판독하십시오.')
      elif r['subject'] not in ('영어','한국사') and not r.get('choice'):
       st.error('11월 응시 선택과목을 지정하십시오.')
      elif not math_short_checked:st.error('수학 단답형 9문항을 원본과 대조한 후 검수 완료 처리하십시오.')
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
      'answers':list(r['answers']),'review':'확정','choice':r.get('choice') or r['subject'],'source':uploaded.name,'page':r['page']}
    db['results']=newresults;save(db)
    st.success(f'{subject_group} 답안 {len(ready)}건 저장 완료 (기존 기록 대체 {existing_count}건).')
    if test_mode:st.info('테스트 모드 — 실제 운영 데이터에는 반영되지 않았습니다.')
    for r in ready:r['reviewed']=False
    st.info('다음 교과 PDF를 처리하려면 현재 파일을 제거하고 상단에서 다음 교과를 선택하여 업로드하십시오.')
  else:st.caption('「전체 자동판독」 버튼을 클릭하십시오.')
 else:st.caption('교과를 선택한 뒤 PDF를 업로드하십시오.')

elif page=='엑셀 답안수정':
 st.subheader('11월 OMR 판독 결과 · 엑셀 선택 수정')
 st.info('기본은 OMR 자동판독입니다. 답안·점수를 정정해야 하는 경우에만 이 메뉴에서 엑셀을 내려받아 수정하고 다시 업로드하십시오. 기존 성적표 양식은 변경되지 않습니다.')
 st.markdown('**수정할 시트: `답안수정` (한 곳에서만 수정)**')
 if not db['roster']:
  st.warning('학생 명부가 없습니다. 먼저 3~10월 이력 엑셀을 등록하십시오.');st.stop()
 st.download_button('thepremium-답안수정.xlsx 내려받기',data=editable_workbook(db,test_mode=test_mode),file_name='thepremium-answer-edit.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
 st.caption('시트의 노란 칸만 수정하십시오. 변경할 행의 「적용(Y)」에 Y를 입력해야 업로드 시 반영됩니다. 빈칸 행은 변경되지 않습니다.')
 uploaded_fix=st.file_uploader('수정한 엑셀 업로드 (.xlsx)',type=['xlsx'],key='excel_answer_correction')
 if uploaded_fix:
  try:
   import hashlib
   token=hashlib.sha256(uploaded_fix.getvalue()).hexdigest()
   changes,errors=parse_corrections(db,uploaded_fix.getvalue())
   if errors:
    st.error(f'검증 오류 {len(errors)}건 — 수정 사항을 적용하지 않습니다.')
    for e in errors[:40]:st.write(e)
   else:
    st.success(f'검증 통과 · 정산 예정 {len(changes)}개 학생·교과 기록')
    st.dataframe([{k:v for k,v in change.items() if k not in ('uid','record')} for change in changes],hide_index=True,use_container_width=True)
    st.caption('3~10월 등급은 그대로 두며, 11월 해당 교과만 갱신합니다. 이전 선택과목과 달라도 현재 응시과목을 적용합니다.')
    checked=st.checkbox('원본 OMR과 수정 답안을 대조했으며, 해당 11월 결과만 덮어쓰는 데 동의합니다.',key='excel_apply_check_'+token[:10])
    if st.button('검증 완료 · 11월 답안/성적 다시 정산',disabled=not checked or not changes,type='primary'):
     newresults=copy.deepcopy(db['results'])
     for change in changes:newresults.setdefault(change['uid'],{})[change['영역']]=change['record']
     db['results']=newresults;save(db)
     st.success(f'{len(changes)}개 교과 기록 갱신 완료. 성적표와 월별 등급 변화가 새 값으로 계산됩니다.')
  except Exception as e:st.error(f'수정 엑셀 검사 실패: {e}')
 st.warning('운영 데이터는 Streamlit 서버 재시작 시 사라질 수 있으니 적용 후 「설정·로고·백업」에서 JSON 백업을 반드시 내려받으십시오. 학생정보가 담긴 엑셀을 공개 GitHub에 올리지 마십시오.')

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
