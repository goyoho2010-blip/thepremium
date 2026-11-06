"""THE PREMIUM OMR: rotation, perspective alignment and conservative bubble detection.

Returns *candidates*, not grades. Uncertain readings must be reviewed against the
original scan. No calibration controls are required for the operator.
"""
from pathlib import Path
import cv2
import numpy as np

TEMPLATE_DIR = Path(__file__).resolve().parent / 'reference' / 'omr'
FORM = {'국어':1,'수학':2,'영어':3,'한국사':4,'탐구1':5,'탐구2':5,'탐구':5}

def bubble_centers(subject):
    out = {}
    def put(q,x,y,dx=15.5):
        out[q] = [(float(x+dx*i),float(y)) for i in range(5)]
    if subject == '국어':
        for q in range(1,21):put(q,643,158+31.25*(q-1))
        for q in range(21,35):put(q,768,158+31.25*(q-21))
        for q in range(35,46):put(q,908,158+31.25*(q-35))
    elif subject == '영어':
        for q in range(1,21):put(q,628,158+31.25*(q-1))
        for q in range(21,41):put(q,768,158+31.25*(q-21))
        for q in range(41,46):put(q,908,158+31.25*(q-41))
    elif subject == '한국사':
        for q in range(1,21):put(q,660,158+31.25*(q-1))
    elif subject.startswith('탐구'):
        for q in range(1,21):put(q,898 if subject=='탐구2' else 613,111+31.25*(q-1))
    elif subject == '수학':
        for q in range(1,9):put(q,408,110+31.25*(q-1))
        for q in range(9,16):put(q,532,110+31.25*(q-9))
        for q in range(23,29):put(q,800,469+31.25*(q-23))
    return out

def _ref(subject):
    f = TEMPLATE_DIR / f'form_{FORM[subject]}.png'
    im = cv2.imread(str(f),cv2.IMREAD_COLOR)
    if im is None:raise RuntimeError('기준 OMR 템플릿 누락: '+str(f))
    return cv2.cvtColor(im,cv2.COLOR_BGR2RGB)

def _features_align(img,ref):
    orb=cv2.ORB_create(nfeatures=4800,fastThreshold=8)
    targ=cv2.cvtColor(ref,cv2.COLOR_RGB2GRAY)
    rt=cv2.resize(targ,None,fx=.85,fy=.85)
    kt,dt=orb.detectAndCompute(rt,None)
    if dt is None:return None,None
    matcher=cv2.BFMatcher(cv2.NORM_HAMMING)
    best=None
    for turns in range(4):
        rotated=np.rot90(img,turns).copy()
        hh,ww=rotated.shape[:2]
        ratio=min(1.,1250/max(ww,hh))
        small=cv2.resize(rotated,(round(ww*ratio),round(hh*ratio)))
        gray=cv2.cvtColor(small,cv2.COLOR_RGB2GRAY)
        ks,ds=orb.detectAndCompute(gray,None)
        if ds is None:continue
        matches=matcher.knnMatch(ds,dt,k=2)
        good=[m for pair in matches if len(pair)==2 for m,n in [pair] if m.distance<.72*n.distance]
        if len(good)<22:continue
        a=np.float32([ks[m.queryIdx].pt for m in good]).reshape(-1,1,2)
        b=np.float32([kt[m.trainIdx].pt for m in good]).reshape(-1,1,2)
        H,mask=cv2.findHomography(a,b,cv2.RANSAC,3.5)
        if H is None or mask is None:continue
        inliers=int(mask.sum())
        if inliers<20:continue
        H=np.diag([1/.85,1/.85,1.])@H@np.diag([ratio,ratio,1.])
        corners=np.float32([[0,0],[ww,0],[ww,hh],[0,hh]]).reshape(-1,1,2)
        wc=cv2.perspectiveTransform(corners,H).reshape(4,2)
        area=cv2.contourArea(wc.astype('float32'))
        expected=ref.shape[0]*ref.shape[1]
        if not (.50*expected<area<1.60*expected):continue
        # Guard against upside-down / mirrored homographies.
        score=inliers/max(len(good),1)
        candidate=(inliers,score,rotated,H,turns,len(good))
        if best is None or candidate[:2]>best[:2]:best=candidate
    if best is None:return None,None
    inliers,score,rotated,H,turns,matched=best
    aligned=cv2.warpPerspective(rotated,H,(ref.shape[1],ref.shape[0]),borderValue=(255,255,255))
    return aligned,{'matched':matched,'inliers':inliers,'rotation':turns*90,'inlier_ratio':round(score,3)}

def _score_at(rgb, gray, ref_gray, x, y):
    """Assess a filled bubble with tolerance for imperfect scanner registration.

    Uses a central disk (not the printed red oval edge), scans +/- 3 px,
    and requires neutral dark ink or local darkening relative to the blank form.
    """
    h, w = gray.shape
    cx, cy = int(round(x)), int(round(y))
    if not (9 <= cx < w-9 and 9 <= cy < h-9):
        return 0.0
    # On this form, the five bubble centers are ~15 px apart. The search
    # radius must remain small enough not to touch a neighboring answer.
    best = 0.0
    yy, xx = np.ogrid[-3:4, -3:4]
    disk = (xx*xx + yy*yy) <= 10
    for dy in (-3, -1, 0, 1, 3):
        for dx in (-3, -1, 0, 1, 3):
            y0, x0 = cy+dy, cx+dx
            pix = rgb[y0-3:y0+4, x0-3:x0+4][disk].astype(np.float32)
            lum = gray[y0-3:y0+4, x0-3:x0+4][disk].astype(np.float32)
            ref = ref_gray[cy-3:cy+4, cx-3:cx+4][disk].astype(np.float32)
            # Dark pencil is nearly neutral. The printed red rings have a
            # strong chromatic component and are not counted as pencil fill.
            neutral = (pix.max(axis=1)-pix.min(axis=1)) < 55
            dark = np.mean((lum < 155) & neutral)
            delta = np.mean((ref-lum > 52) & (lum < 185))
            score = max(dark, 0.85*delta)
            best = max(best, float(score))
    return best

def read_omr(image,subject):
    ref=_ref(subject)
    aligned,meta=_features_align(image,ref)
    if aligned is None:
        raise ValueError('OMR 자동 정렬 실패: 페이지 전체가 보이도록 스캔하거나 사진을 다시 촬영하십시오.')
    centers=bubble_centers(subject)
    n=45 if subject in ('국어','영어') else 30 if subject=='수학' else 20
    gray=cv2.cvtColor(aligned,cv2.COLOR_RGB2GRAY)
    ref_gray=cv2.cvtColor(ref,cv2.COLOR_RGB2GRAY)
    answers=['']*n
    diagnostics=[]
    preview=aligned.copy()
    for q in range(1,n+1):
        if q not in centers:
            diagnostics.append({'문항':q,'판독답안':'','상태':'단답형 수동 확인','최고마킹':0.,'차이':0.})
            continue
        vals=[_score_at(aligned,gray,ref_gray,*p) for p in centers[q]]
        order=sorted(range(5),key=lambda i:vals[i],reverse=True)
        top,second=order[:2]
        strongest,gap=vals[top],vals[top]-vals[second]
        # Dark-ink evidence is explicit; ambiguous marks never become answers.
        if strongest>=.30 and gap>=.13:
            status='정상';answers[q-1]=str(top+1)
        elif strongest<.18:
            status='미표기'
        else:
            status='중복/불확실'
        diagnostics.append({'문항':q,'판독답안':answers[q-1],'상태':status,
                            '최고마킹':round(strongest,2),'차이':round(gap,2)})
        for k,(x,y) in enumerate(centers[q]):
            center=(round(x),round(y))
            color=(0,175,0) if status=='정상' and k==top else (235,75,25) if status!='정상' else (0,165,235)
            cv2.circle(preview,center,5,color,1,cv2.LINE_AA)
    meta['detected']=sum(bool(a) for a in answers)
    meta['total_auto']=len(centers)
    meta['uncertain']=sum(row['상태'] not in ('정상','단답형 수동 확인') for row in diagnostics)
    # A whole page of blank answers is treated as a FAILED reading candidate,
    # not as an accepted empty response sheet.
    meta['suspect']=meta['detected']==0 or (meta['detected']/max(len(centers),1)<.55)
    meta['aligned_preview']=True
    return answers,diagnostics,preview,meta
