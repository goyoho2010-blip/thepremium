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
        for q in range(1,21):put(q,908 if subject=='탐구2' else 613,111+31.25*(q-1))
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
    """Sample a whole answer bubble using local template subtraction.

    Prefer dark neutral graphite over printed red rings; use small offset
    tolerance but keep each measurement isolated from adjoining bubbles.
    """
    h,w=gray.shape;cx,cy=round(x),round(y)
    if not (9<cx<w-9 and 9<cy<h-9):return 0.
    yy,xx=np.ogrid[-4:5,-4:5]
    mask=(xx*xx+yy*yy<=13)
    best=0.
    for dy in (-2,0,2):
        for dx in (-2,0,2):
            px=rgb[cy+dy-4:cy+dy+5,cx+dx-4:cx+dx+5][mask].astype(np.float32)
            lum=gray[cy+dy-4:cy+dy+5,cx+dx-4:cx+dx+5][mask].astype(np.float32)
            blank=ref_gray[cy-4:cy+5,cx-4:cx+5][mask].astype(np.float32)
            neutral=px.max(1)-px.min(1)<58
            dark=np.mean((lum<156)&neutral)
            new_ink=np.mean((blank-lum>46)&(lum<190)&neutral)
            best=max(best,float(max(dark,.92*new_ink)))
    return best

# Mathematics short-answer fields: q16–22,29–30. On this OMR each
# three-digit field has hundreds/tens/ones columns. A blank digit column
# is NOT the same as a marked zero. All extracted numbers require review.
_MATH_DIGIT_X = {
    # 수학 실제 단답형 양식: 문항당 세 자리(백/십/일) 원의 중심.
    # 상단 16~20과 하단 21~22, 29~30은 인쇄 위치가 각각 다릅니다.
    16: (704, 720, 735), 17: (767, 782, 798),
    18: (830, 846, 861), 19: (894, 910, 925),
    20: (957, 973, 989), 21: (657, 673, 688),
    22: (720, 735, 751), 29: (896, 911, 927),
    30: (958, 974, 990),
}

def _math_digit_centers(question):
    # Digit centers at 0 through 9. The upper row is y=109;
    # the lower set (21,22,29,30) starts at y=468.
    y0 = 109 if question <= 20 else 468
    return [[(x, y0 + k*31.2) for k in range(10)]
            for x in _MATH_DIGIT_X[question]]

def _digit_ink(rgb, gray, ref_gray, x, y):
    """Candidate strength for one 0–9 digit, with scan drift tolerance."""
    cx,cy=round(x),round(y);h,w=gray.shape
    if cx<9 or cy<9 or cx>=w-9 or cy>=h-9:return 0.
    yy,xx=np.ogrid[-4:5,-4:5];mask=xx*xx+yy*yy<=11
    best=0.
    for dy in (-2,0,2):
        for dx in (-2,0,2):
            lum=gray[cy+dy-4:cy+dy+5,cx+dx-4:cx+dx+5][mask].astype('float32')
            base=ref_gray[cy-4:cy+5,cx-4:cx+5][mask].astype('float32')
            rgb_roi=rgb[cy+dy-4:cy+dy+5,cx+dx-4:cx+dx+5][mask].astype('float32')
            neutral=(rgb_roi.max(1)-rgb_roi.min(1))<60
            new=np.mean((base-lum>38)&(lum<180)&neutral)
            dark=np.mean((lum<137)&neutral)
            best=max(best,float(max(new,.70*dark)))
    return best

def _digit_candidate(rgb, gray, ref_gray, question):
    columns = []
    for centers in _math_digit_centers(question):
        vals = [_digit_ink(rgb, gray, ref_gray, x, y) for x,y in centers]
        idx = np.argsort(vals)[::-1]
        high, low = float(vals[idx[0]]),float(vals[idx[1]])
        if high >= .27 and high-low >= .12:
            columns.append((str(int(idx[0])), '후보'))
        elif high < .16:
            columns.append(('', '공란'))
        else:
            columns.append(('', '불확실'))
    if any(state=='불확실' for _,state in columns):
        return '', '단답형 판독 불확실'
    digits=''.join(value for value,_ in columns)
    if not digits:
        return '', '단답형 미표기'
    # A gap between marked columns is ambiguous, and a marked 0 is distinct
    # from an empty column. Reject malformed digit alignment.
    started=False;gap=False
    for value,_ in columns:
        if value: 
            if gap:return '', '단답형 자리검수 필요'
            started=True
        elif started:gap=True
    return str(int(digits)), '단답형 자동후보·검수 필수'

# THE PREMIUM 탐구 과목코드: 20 물리Ⅰ, 21 화학Ⅰ, 22 생명과학Ⅰ,
# 23 지구과학Ⅰ. Do not label a candidate as verified without clear marks.
SCIENCE_CODE={11:'생활과 윤리',12:'윤리와 사상',13:'한국지리',14:'세계지리',15:'동아시아사',16:'세계사',17:'경제',18:'정치와 법',19:'사회·문화',20:'물리학Ⅰ',21:'화학Ⅰ',22:'생명과학Ⅰ',23:'지구과학Ⅰ',
              24:'물리학Ⅱ',25:'화학Ⅱ',26:'생명과학Ⅱ',27:'지구과학Ⅱ'}

def exploration_subject_candidate(aligned, subject):
    """Read the two-digit 탐구 과목코드 with conservative confidence.

    The selector is *not* the 1–20 multiple-choice bubble block. Returns
    (code, label, confident), or (None,'',False) if incomplete/ambiguous.
    """
    if subject not in ('탐구1','탐구2'):return None,'',False
    gray=cv2.cvtColor(aligned,cv2.COLOR_RGB2GRAY)
    refgray=cv2.cvtColor(_ref(subject),cv2.COLOR_RGB2GRAY)
    # selector's two-digit lanes left of each answer block, 0...9 vertically
    xs=(568,584) if subject=='탐구1' else (846,862)
    out=[]
    for x in xs:
        values=[_digit_ink(aligned,gray,refgray,x,266+31.2*k) for k in range(10)]
        inds=np.argsort(values)[::-1];a,b=int(inds[0]),int(inds[1])
        if values[a]<.29 or values[a]-values[b]<.16:return None,'',False
        out.append(str(a))
    code=int(''.join(out))
    return code,SCIENCE_CODE.get(code,''),bool(code in SCIENCE_CODE)

# 선택과목 OMR 전용 영역. 공통문항 답란과 별개이며, 이전 월 선택과목으로 대체하지 않습니다.
SUBJECT_BUBBLES = {
    '국어': [('화법과 작문',(363,329)), ('언어와 매체',(363,361))],
    '수학': [('확률과 통계',(362,391)), ('미적분',(362,423)), ('기하',(362,454))],
}
TRACK_BUBBLES = [('인문계',(362,571)), ('자연계',(362,605))]

def _choice_marks(rgb,subject,positions):
    """Template-relative dark neutral marks; never force an ambiguous result."""
    gray=cv2.cvtColor(rgb,cv2.COLOR_RGB2GRAY)
    refgray=cv2.cvtColor(_ref(subject),cv2.COLOR_RGB2GRAY)
    vals=[_score_at(rgb,gray,refgray,*pt) for _,pt in positions]
    order=np.argsort(vals)[::-1]
    first=int(order[0]); second=int(order[1])
    high=float(vals[first]); gap=high-float(vals[second])
    confident=high>=.29 and gap>=.13
    return (positions[first][0] if confident else ''),confident,{
        'strengths':{label:round(float(v),3) for (label,_),v in zip(positions,vals)},
        'highest':round(high,3),'margin':round(gap,3)
    }

def selection_subject_candidate(aligned,subject):
    if subject not in SUBJECT_BUBBLES:
        return {'choice':'','confident':False,'diagnostics':{}}
    choice,ok,stats=_choice_marks(aligned,subject,SUBJECT_BUBBLES[subject])
    result={'choice':choice,'confident':ok,'diagnostics':stats}
    if subject=='수학':
        track,track_ok,track_stats=_choice_marks(aligned,subject,TRACK_BUBBLES)
        result.update(track=track,track_confident=track_ok,track_diagnostics=track_stats)
    return result

# 학생명: 각 글자는 [초성, 중성, 종성] 3열, 최대 4글자.
# 인쇄 순서는 현대 한글 자모의 표준 음절 조합 순서. 종성의 일부 복합자모는
# 추가 칸 배열이 있으므로 미검증 후보를 자동 확정하지 않습니다.
NAME_CHO=['ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']
NAME_JUNG=['ㅏ','ㅐ','ㅑ','ㅒ','ㅓ','ㅔ','ㅕ','ㅖ','ㅗ','ㅘ','ㅙ','ㅚ','ㅛ','ㅜ','ㅝ','ㅞ','ㅟ','ㅠ','ㅡ','ㅢ','ㅣ']
# OMR 인쇄에서 첫 세 종성 칸은 ㄱ,ㄲ,ㄴ 순서 (ㄳ은 추가 칸).
NAME_JONG=['ㄱ','ㄲ','ㄴ','ㄵ','ㄶ','ㄷ','ㄹ','ㄺ','ㄻ','ㄼ','ㄽ','ㄾ','ㄿ','ㅀ','ㅁ','ㅂ','ㅄ','ㅅ','ㅆ','ㅇ','ㅈ']
HANGUL_JONG=['','ㄱ','ㄲ','ㄳ','ㄴ','ㄵ','ㄶ','ㄷ','ㄹ','ㄺ','ㄻ','ㄼ','ㄽ','ㄾ','ㄿ','ㅀ','ㅁ','ㅂ','ㅄ','ㅅ','ㅆ','ㅇ','ㅈ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']

def _name_component(rgb,gray,refgray,x,y,count):
    vals=[_score_at(rgb,gray,refgray,x,y+15.5*j) for j in range(count)]
    order=np.argsort(vals)[::-1]
    high=float(vals[order[0]])
    margin=high-float(vals[order[1]]) if len(vals)>1 else high
    # Be deliberately conservative: the result is for human confirmation.
    if high<.24:return None,'빈칸',round(high,2)
    if high<.36 or margin<.13:return None,'불확실',round(high,2)
    return int(order[0]),'판독후보',round(high,2)

def decode_name_candidate(aligned,subject):
    """Name candidate (not verified identity) from printed Hangul OMR.

    Returns confidence false if a component is inconsistent/ambiguous.
    This cannot replace confirmation against student roster and exam number.
    """
    gray=cv2.cvtColor(aligned,cv2.COLOR_RGB2GRAY)
    refgray=cv2.cvtColor(_ref(subject),cv2.COLOR_RGB2GRAY)
    parts=[];statuses=[];ambiguous=False
    for syll in range(4):
        x0=425.5+46.7*syll
        rows=[]
        for col,count in [(0,19),(1,21),(2,len(NAME_JONG))]:
            rows.append(_name_component(aligned,gray,refgray,x0+15.55*col,433.5,count))
        (ci,cs,_),(vi,vs,_),(fi,fs,_)=rows
        if cs=='빈칸' and vs=='빈칸' and fs=='빈칸':
            # A trailing unmarked syllable column is valid.
            statuses.append({'글자':syll+1,'초성':'','중성':'','종성':'','상태':'미사용'})
            continue
        if ci is None or vi is None or fs=='불확실':
            ambiguous=True
            statuses.append({'글자':syll+1,'초성':NAME_CHO[ci] if ci is not None else '?',
                             '중성':NAME_JUNG[vi] if vi is not None else '?',
                             '종성':NAME_JONG[fi] if fi is not None else '', '상태':'확인필요'})
            continue
        jong=NAME_JONG[fi] if fi is not None else ''
        cp=0xAC00+(ci*21+vi)*28+HANGUL_JONG.index(jong)
        parts.append(chr(cp))
        statuses.append({'글자':syll+1,'초성':NAME_CHO[ci],'중성':NAME_JUNG[vi],
                         '종성':jong or '없음','상태':'판독후보'})
    if not parts or ambiguous:return {'candidate':'','confident':False,'parts':statuses}
    # Require continuous syllables rather than filled column after a gap.
    used=[x['상태']!='미사용' for x in statuses]
    contiguous=not any(used[j] and not used[j-1] for j in range(1,4))
    return {'candidate':''.join(parts) if contiguous else '',
            'confident':contiguous,'parts':statuses}

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
            if subject=='수학' and q in _MATH_DIGIT_X:
                candidate,status=_digit_candidate(aligned,gray,ref_gray,q)
                answers[q-1]=candidate
                for digit_col in _math_digit_centers(q):
                    for x,y in digit_col:
                        cv2.circle(preview,(round(x),round(y)),3,(255,150,0),1,cv2.LINE_AA)
                diagnostics.append({'문항':q,'판독답안':candidate,'상태':status,
                                    '최고마킹':0.,'차이':0.})
            else:
                diagnostics.append({'문항':q,'판독답안':'','상태':'판독영역 없음',
                                    '최고마킹':0.,'차이':0.})
            continue
        vals=[_score_at(aligned,gray,ref_gray,*p) for p in centers[q]]
        order=sorted(range(5),key=lambda i:vals[i],reverse=True)
        top,second=order[:2]
        strongest,gap=vals[top],vals[top]-vals[second]
        if strongest>=.30 and gap>=.13:
            status='정상';answers[q-1]=str(top+1)
        elif strongest<.18:
            status='미표기(0점)' 
        else:
            status='중복/불확실'
        diagnostics.append({'문항':q,'판독답안':answers[q-1],'상태':status,
                            '최고마킹':round(strongest,2),'차이':round(gap,2)})
        for k,(x,y) in enumerate(centers[q]):
            center=(round(x),round(y))
            color=(0,175,0) if status=='정상' and k==top else (235,75,25) if status!='정상' else (0,165,235)
            cv2.circle(preview,center,5,color,1,cv2.LINE_AA)
    nm=decode_name_candidate(aligned,subject)
    meta['name_candidate']=nm['candidate']
    meta['name_confident']=nm['confident']
    meta['name_components']=nm['parts']
    if subject in ('국어','수학'):
        sel=selection_subject_candidate(aligned,subject)
        meta['selection_choice']=sel['choice']
        meta['selection_confident']=sel['confident']
        meta['selection_diagnostics']=sel['diagnostics']
        if subject=='수학':
            meta['math_track']=sel['track']
            meta['math_track_confident']=sel['track_confident']
    if subject in ('탐구1','탐구2'):
        code,label,verified=exploration_subject_candidate(aligned,subject)
        meta['explore_code']=code
        meta['explore_name']=label
        meta['explore_code_confident']=verified
    meta['detected']=sum(bool(a) for a in answers)
    meta['total_auto']=len(centers)
    meta['uncertain']=sum(row['상태']!='정상' for row in diagnostics)
    meta['suspect']=sum(row['상태']=='정상' for row in diagnostics)/max(len(centers),1)<.55
    meta['aligned_preview']=True
    meta['math_short_review_required']= subject=='수학'
    return answers,diagnostics,preview,meta
