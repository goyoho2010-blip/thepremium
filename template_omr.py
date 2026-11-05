"""대성 THE PREMIUM 5종 기준 OMR: 이미지 템플릿 자동 정렬 및 답안 후보 추출.
판독은 자동 확정하지 않음. 결과의 검수/확정은 app.py에서 수행.
"""
from pathlib import Path
import cv2
import numpy as np

TEMPLATE_DIR=Path(__file__).resolve().parent/'reference'/'omr'
FORM={'국어':1,'수학':2,'영어':3,'한국사':4,'탐구1':5,'탐구2':5,'탐구':5}

# 기준 이미지 1030x796 (가로 x 세로; 90도 회전 정방향)
# row grid values are center positions, in reference pixels.
def bubble_centers(subject):
    out={}
    def put(q,x0,y0,dx=15.5):
        out[q]=[(float(x0+dx*j),float(y0)) for j in range(5)]
    if subject=='국어':
        for q in range(1,21):put(q,643,158+31.25*(q-1),15.5)
        for q in range(21,35):put(q,768,158+31.25*(q-21),15.5)
        for q in range(35,46):put(q,908,158+31.25*(q-35),15.5)
    elif subject=='영어':
        for q in range(1,21):put(q,628,158+31.25*(q-1),15.5)
        for q in range(21,41):put(q,768,158+31.25*(q-21),15.5)
        for q in range(41,46):put(q,908,158+31.25*(q-41),15.5)
    elif subject=='한국사':
        for q in range(1,21):put(q,660,158+31.25*(q-1),15.5)
    elif subject.startswith('탐구'):
        right=subject=='탐구2'
        for q in range(1,21):put(q,898 if right else 613,111+31.25*(q-1),15.5)
    elif subject=='수학':
        for q in range(1,9):put(q,408,110+31.25*(q-1),15.5)
        for q in range(9,16):put(q,532,110+31.25*(q-9),15.5)
        # 단답형 16~22, 29~30은 숫자 전사 필요 -> 판독 보류
        for q in range(23,29):put(q,800,469+31.25*(q-23),15.5)
    return out

def _ref(subject):
    f=TEMPLATE_DIR/f'form_{FORM[subject]}.png'
    im=cv2.imread(str(f),cv2.IMREAD_COLOR)
    if im is None:raise RuntimeError('OMR 기준 양식 이미지가 없습니다: '+str(f))
    return cv2.cvtColor(im,cv2.COLOR_BGR2RGB)

def _features_align(img,ref):
    """각도 및 위치를 ORB 특징점으로 자동 보정; 신뢰도 낮으면 실패 반환."""
    orb=cv2.ORB_create(nfeatures=3500,fastThreshold=9)
    target_gray=cv2.cvtColor(ref,cv2.COLOR_RGB2GRAY)
    rt=cv2.resize(target_gray,None,fx=.85,fy=.85)
    kt,dt=orb.detectAndCompute(rt,None)
    if dt is None:return None,None
    matcher=cv2.BFMatcher(cv2.NORM_HAMMING)
    best=None
    for turns in range(4):
        rotate=np.rot90(img,turns).copy()
        h,w=rotate.shape[:2]
        ratio=min(1.0,1100/max(w,h))
        small=cv2.resize(rotate,(int(w*ratio),int(h*ratio)))
        gs=cv2.cvtColor(small,cv2.COLOR_RGB2GRAY)
        ks,ds=orb.detectAndCompute(gs,None)
        if ds is None:continue
        matched=matcher.knnMatch(ds,dt,k=2)
        good=[m for pair in matched if len(pair)==2 for m,n in [pair] if m.distance<.72*n.distance]
        if len(good)<18:continue
        a=np.float32([ks[m.queryIdx].pt for m in good]).reshape(-1,1,2)
        b=np.float32([kt[m.trainIdx].pt for m in good]).reshape(-1,1,2)
        H,mask=cv2.findHomography(a,b,cv2.RANSAC,4.5)
        if H is None or mask is None:continue
        inliers=int(mask.sum())
        if inliers<17:continue
        H=np.diag([1/.85,1/.85,1])@H@np.diag([ratio,ratio,1])
        corners=np.float32([[0,0],[w,0],[w,h],[0,h]]).reshape(-1,1,2)
        warped_corners=cv2.perspectiveTransform(corners,H).reshape(4,2)
        poly_area=cv2.contourArea(warped_corners.astype('float32'))
        expected=ref.shape[0]*ref.shape[1]
        if not (.30*expected<poly_area<2.0*expected):continue
        score=inliers
        if best is None or score>best[0]:best=(score,rotate,H,turns,len(good))
    if best is None:return None,None
    _,rotate,H,turns,good=best
    aligned=cv2.warpPerspective(rotate,H,(ref.shape[1],ref.shape[0]),borderValue=(255,255,255))
    return aligned,{'matched':good,'inliers':best[0],'rotation':turns*90}

def _bubble_ink(image,pt):
    x,y=map(int,map(round,pt))
    g=cv2.cvtColor(image,cv2.COLOR_RGB2GRAY)
    roi=g[max(0,y-4):y+5,max(0,x-4):x+5]
    if roi.shape!=(9,9):return 0.
    yy,xx=np.ogrid[-4:5,-4:5]
    interior=roi[xx*xx+yy*yy<=7]
    # Printed light-color circles have bright centers; black pencil fills have low lightness.
    return float(np.mean(interior<110))

def read_omr(image,subject):
    """Return answers, diagnostic rows, aligned preview, metadata. Blank/uncertain => ''."""
    ref=_ref(subject)
    aligned,meta=_features_align(image,ref)
    if aligned is None:
        raise ValueError('답안지 자동 정렬에 실패했습니다. 전체 OMR이 보이도록 스캔한 PDF/이미지를 사용하고, 촬영 사진은 가능한 한 정면에서 촬영하십시오.')
    centers=bubble_centers(subject)
    n=45 if subject in ('국어','영어') else 30 if subject=='수학' else 20
    answers=['']*n; diagnostics=[]
    preview=aligned.copy()
    for q in range(1,n+1):
        if q not in centers:
            diagnostics.append({'문항':q,'판독답안':'','상태':'단답형 수동 확인','최고마킹':0.0,'차이':0.0})
            continue
        values=[_bubble_ink(aligned,p) for p in centers[q]]
        ranking=sorted(range(5),key=lambda i:values[i],reverse=True)
        top,second=ranking[:2]
        strongest=values[top];gap=strongest-values[second]
        status='정상' if strongest>=.42 and gap>=.22 else '미표기' if strongest<.42 else '중복/불확실'
        if status=='정상': answers[q-1]=str(top+1)
        diagnostics.append({'문항':q,'판독답안':answers[q-1],'상태':status,'최고마킹':round(strongest,2),'차이':round(gap,2)})
        for k,pt in enumerate(centers[q]):
            x,y=map(lambda v:int(round(v)),pt)
            cv2.circle(preview,(x,y),5,(10,190,10) if k==top and status=='정상' else (0,160,240),1)
    return answers,diagnostics,preview,meta
