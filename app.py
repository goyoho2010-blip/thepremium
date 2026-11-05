"""2026 11월 대성 THE PREMIUM | OMR·월별 등급 성적표"""
import io, json, os, re, zipfile, csv, base64
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

ROOT=Path(__file__).resolve().parent
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
FONT='C:/Windows/Fonts/malgun.ttf'
for p in ['C:/Windows/Fonts/malgun.ttf','C:/Windows/Fonts/gulim.ttc','/usr/share/fonts/truetype/nanum/NanumGothic.ttf','/usr/share/fonts/truetype/nanum/NanumBarunGothic.ttf','/usr/share/fonts/truetype/nanum/NanumSquareR.ttf']:
 if Path(p).exists():FONT=p;break
try:pdfmetrics.registerFont(TTFont('Korean',FONT)); FONTNAME='Korean'
except:FONTNAME='Helvetica'

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
 history=json.loads((DATA/'history.json').read_text(encoding='utf8'))
 roster=history['roster']
 return {'roster':roster,'history':history['history'],'results':{},'keys':{},'cuts':CUTS,'layout':{},'logos':{k:base64.b64encode((ROOT/'assets'/v).read_bytes()).decode() for k,v in {'횡성인재육성관':'hoengseong.png','대성학력개발연구소':'daesung.png','천명의선택입시연구소':'choice1000.png'}.items() if (ROOT/'assets'/v).exists()},'notes':{},'settings':{'title':'11월 대성THE프리미엄 모의고사 성적','subtitle':'자체 정산-10월 성적대비등급'}}

def get_state():
 if 'db' not in st.session_state:
  try:db=json.loads(STORAGE.read_text(encoding='utf8')) if STORAGE.exists() else starting_state()
  except:db=starting_state()
  st.session_state.db=db
 return st.session_state.db

def save(db):
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
 if (subject in ('국어','수학','탐구1','탐구2') and not choice) or len(key)!=n or any(str(k.get('answer','')).strip()=='' or not isinstance(k.get('points'),int) or k['points']<=0 for k in key):
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

def render_pdf(db,uid,include_answers=True):
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
  title=db['settings']['title'];text(76,745,title,17);text(136,727,'(자체 정산 · 10월 성적 대비 잠정등급)',10)
  text(40,696,f"학생명  {stu['name']}     학교  {stu.get('school') or '미입력'}     학년  {stu.get('grade') or '고3'}",10)
  line(684)
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
 text(42,425,'월별 등급 변화 · 3월—11월',11)
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

def export_xlsx(db):
 wb=Workbook();ss=wb.active;ss.title='11월성적'
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
db=get_state()
st.title('11월 대성 THE PREMIUM | OMR 성적처리 NAVI')
st.caption('Streamlit Cloud 배포판 · 2026년 11월 고3 · OMR 판독 / 검수 / 10월 기준 잠정 등급 / 월별 변화 / PDF 인쇄 및 파일 저장')
st.warning('11월 전국 공동 채점 결과가 아닙니다. 10월 원점수 등급컷을 11월에 적용한 비공식 참고 성적이며 시험 난이도 차이는 보정하지 않습니다.')
page=st.sidebar.radio('메뉴',['학생 명부','11월 정답·배점','OMR 판독·검수','학생 성적표','일괄 인쇄·저장','설정·로고·백업'])
if st.sidebar.button('현재 수정사항 저장',use_container_width=True):save(db);st.sidebar.success('작업 파일에 저장했습니다.')

if page=='학생 명부':
 st.subheader('3월~10월 성적 엑셀 연결 명부')
 st.write(f"학생 {len(db['roster'])}명 · 월별 기존 성적이 있는 학생을 통합 관리합니다.")
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
 st.info('기존 엑셀의 월별 성적은 포함되어 있습니다. 기존 성적은 이름과 학생번호 결합 ID로 연결하므로, 이름만 변경하더라도 과거 이력은 유지됩니다.')

elif page=='11월 정답·배점':
 st.subheader('정답표 등록 및 검증')
 sub=st.selectbox('영역',['국어','수학','영어','한국사','탐구'])
 choice=st.selectbox('선택과목',OPTIONS[sub]);k=sub+'|'+choice;n=N[sub]
 current=db['keys'].get(k,[{'question':i+1,'answer':'','points':None} for i in range(n)])
 if len(current)!=n:current=[{'question':i+1,'answer':'','points':None} for i in range(n)]
 st.caption('11월 정답 원본과 문항별 배점을 반드시 대조하십시오. 정답/배점이 모두 입력되고 학생 검수가 확정되기 전에는 점수가 산출되지 않습니다.')
 # CSV paste avoids heavy spreadsheet deps
 raw='\n'.join(f"{p['question']},{p.get('answer','')},{p.get('points') or ''}" for p in current)
 text_csv=st.text_area('문항번호, 정답, 배점 (한 줄에 한 문항)',raw,height=350)
 total=st.number_input('만점 확인',min_value=1,max_value=100,value=50 if sub=='탐구' else 100)
 if st.button('정답·배점 검증 후 저장'):
  try:
   result=[]
   for line in text_csv.splitlines():
    parts=[x.strip() for x in line.split(',')];assert len(parts)==3
    q=int(parts[0]);v=parts[1];pt=int(parts[2]);assert pt>0 and v!=''
    result.append({'question':q,'answer':v,'points':pt})
   assert [z['question'] for z in result]==list(range(1,n+1)),'문항번호 또는 문항수 오류'
   assert sum(z['points'] for z in result)==total,'문항 배점 합계 오류'
   db['keys'][k]=result;save(db);st.success(f'{choice} 정답 {n}문항 저장 완료 (총 {total}점)')
  except Exception as e:st.error(f'저장 불가: {e}')
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
 st.subheader('OMR PDF / 이미지 업로드 → 마킹 판독 → 확인 → 확정')
 st.info('OMR 용지 종류에 따라 답안란 위치가 다릅니다. 처음 한 번 좌표를 조정한 뒤 미리보기로 마킹 중심을 확인하십시오. PDF 텍스트 추출이 아닌 실제 마킹의 명암을 분석합니다.')
 opt=[f"{x['student_number']} | {x['name']}" for x in db['roster']]
 idx=st.selectbox('학생',range(len(opt)),format_func=lambda i:opt[i]);uid=db['roster'][idx]['id']
 subject=st.selectbox('교과 영역',CATS);subject_name=which(subject);n=N[subject_name]
 choice=(db['roster'][idx].get({'국어':'korean','수학':'math','탐구1':'explore1','탐구2':'explore2'}.get(subject,''),'') if subject not in ('영어','한국사') else subject)
 st.caption(f'선택과목: {norm(choice) or "미설정"} / 총 {n}문항')
 uploaded=st.file_uploader('학생 OMR (PDF·PNG·JPG)',type=['pdf','png','jpg','jpeg'])
 if uploaded:
  content=uploaded.getvalue();kind='pdf' if uploaded.name.lower().endswith('.pdf') else 'image'
  pages=len(fitz.open(stream=content,filetype='pdf')) if kind=='pdf' else 1
  page_idx=st.number_input('PDF 페이지',min_value=1,max_value=pages,value=1)-1
  img=encode_page(content,kind,page_idx)
  base=db['layout'].get(subject,{'x':.22,'y':.16,'w':.50,'h':.72,'columns':1,'colgap':.02,'rows':n,'radius':.27,'threshold':.25,'margin':.09})
  with st.expander('OMR 위치 보정',expanded=True):
   st.caption('x/y: 첫 번째 답안 묶음 왼쪽 위, w/h: 한 묶음의 너비/높이. 여러 열은 열당 문항수와 열간 간격을 설정합니다.')
   c1,c2,c3,c4=st.columns(4)
   conf=dict(base)
   for param,col in [('x',c1),('y',c2),('w',c3),('h',c4)]:conf[param]=col.slider(param,0.01,0.98,float(base[param]),0.005)
   c1,c2,c3,c4=st.columns(4)
   conf['columns']=c1.number_input('문항 묶음 열 수',1,6,int(base.get('columns',1)))
   conf['rows']=c2.number_input('열당 문항 수',1,50,int(base.get('rows',n)))
   conf['colgap']=c3.slider('열 사이 간격',0.,.3,float(base.get('colgap',.02)),.005)
   conf['radius']=c4.slider('판독 반경',.1,.45,float(base.get('radius',.27)),.01)
   c1,c2=st.columns(2)
   conf['threshold']=c1.slider('마킹 검출 민감도',.05,.9,float(base.get('threshold',.25)),.01)
   conf['margin']=c2.slider('중복 마킹 차이값',.01,.5,float(base.get('margin',.09)),.01)
  try:
   result,stats=interpret_omr(img,conf,n,NCHOICES[subject_name]);st.image(outline(img,conf,n,NCHOICES[subject_name]),caption='파란 원 = 각 답안 후보의 판독 중심점 (설정값을 조정하여 실제 마킹에 맞추십시오)',use_container_width=True)
   alerts=[x for x in stats if x['status']!='정상']
   st.write(f'1차 자동판독: 정상 {n-len(alerts)}개 / **검수 필요 {len(alerts)}개**')
   if alerts:st.dataframe(alerts,use_container_width=True)
   db['layout'][subject]=conf
   if st.button('판독 결과 검수표로 가져오기'):
    db['results'].setdefault(uid,{})[subject]={'answers':result,'review':'검수중','source':uploaded.name,'page':int(page_idx+1)}
    save(db);st.success('검수표로 가져왔습니다. 바로 아래에서 답안을 확인하고 확정하십시오.')
  except Exception as e:st.error(f'판독 실패: {e}')
 rec=db['results'].setdefault(uid,{}).setdefault(subject,{})
 original=rec.get('answers',['']*n)
 if len(original)!=n:original=['']*n
 st.subheader('학생 답안 검수 · 정오답')
 st.caption('문항번호 순서대로 학생 기입값을 입력하십시오. 미표기/복수 표기는 빈칸으로 남기고 검수합니다. 수학 단답형에는 숫자를 입력할 수 있습니다.')
 editable=st.text_area('학생답안 (쉼표로 구분, 빈칸 허용)',','.join(str(v) for v in original),height=105,key=f'edit_{uid}_{subject}_{len(original)}')
 if st.button('검수 완료 · 답안 확정',type='primary'):
  v=[x.strip() for x in editable.split(',')]
  if len(v)!=n:st.error(f'{n}개 답안이 필요합니다. 현재 {len(v)}개입니다.')
  else:
   rec.update({'answers':v,'review':'확정','source':rec.get('source','직접 검수')});save(db);st.success('답안 확정 및 자동 재채점 완료')
 r=scored(db,uid,subject)
 if r['ready']:
  st.metric('원점수',f"{r['score']}점",f"10월 대비 기준 {r['grade']}등급" if r['grade'] else '등급컷 미등록')
  def chunks(seq,n=10):return [seq[i:i+n] for i in range(0,len(seq),n)]
  for p,ans in enumerate(chunks(r['answers'])):
   o=p*10;cor=r['correct'][o:o+10];st.table([dict(zip(['구분']+[str(i) for i in range(o+1,o+len(ans)+1)],['학생답안']+ans)),dict(zip(['구분']+[str(i) for i in range(o+1,o+len(ans)+1)],['정오답']+cor))])
 else:st.caption('답안 검수 확정 및 해당 과목의 검증된 정답·배점 등록 후 점수와 등급을 표시합니다.')

elif page=='학생 성적표':
 st.subheader('성적표 조회 · 인쇄')
 opts=[f"{x['student_number']} | {x['name']}" for x in db['roster']]
 i=st.selectbox('학생 선택',range(len(opts)),format_func=lambda x:opts[x]);stu=db['roster'][i];uid=stu['id']
 rows=[]
 for sub in CATS:
  r=scored(db,uid,sub);rows.append({'영역':sub,'선택과목':r.get('choice'),'11월 원점수':r.get('score'),'11월 잠정등급':r.get('grade'),'검수':r.get('review')})
 st.dataframe(rows,hide_index=True,use_container_width=True)
 st.subheader('월별 등급 변화 (3·5·7·8·9·10·11월)')
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
 b=render_pdf(db,uid,show_ans)
 st.download_button('학생별 A4 PDF 저장 · 인쇄',b,file_name=f"11월더프_성적표_{stu['name']}.pdf",mime='application/pdf',type='primary')
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
    for s in filtered:archive.writestr(f"{s['student_number']}_{s['name']}.pdf",render_pdf(db,s['id'],include))
   st.session_state['pdf_zip']=z.getvalue()
 if st.session_state.get('pdf_zip'):
  st.download_button('학생별 PDF ZIP 저장',st.session_state['pdf_zip'],file_name='thepremium-reports.zip',mime='application/zip')
 if st.button('전체 결과 엑셀 생성'):
  try:
   b=export_xlsx(db);st.session_state['result_xlsx']=b
  except Exception as e:st.error(f'엑셀 생성 실패: {e}')
 if st.session_state.get('result_xlsx'):
  st.download_button('성적 집계 + 월별 등급 엑셀 저장',st.session_state['result_xlsx'],file_name='thepremium-results.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

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
 st.download_button('모든 결과·명부·정답·설정 JSON 백업',json.dumps(db,ensure_ascii=False,indent=2).encode(),file_name='더프11월_전체백업.json',mime='application/json')
 restore=st.file_uploader('기존 JSON 백업 불러오기',type=['json'])
 if restore and st.button('백업 복원 실행'):
  try:
   new=json.loads(restore.getvalue().decode());assert all(k in new for k in ['roster','history','results','keys','cuts'])
   st.session_state.db=new;save(new);st.success('복원 완료. 화면을 새로고침하십시오.')
  except Exception as e:st.error(f'복원 실패: {e}')
 st.info('로컬 실행: 모든 데이터는 이 노트북의 data/work.json에 저장됩니다. 주기적으로 JSON 백업 파일을 외장 저장장치에 보관하십시오.')
