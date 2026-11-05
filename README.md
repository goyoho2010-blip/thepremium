# thepremium — GitHub / Streamlit Cloud 배포판

## 배포
1. GitHub에서 **Private** 저장소 `thepremium` 생성
2. 이 ZIP 내부의 `app.py`, `requirements.txt`, `data/`, `assets/`, `.streamlit/`을 저장소 **루트**에 업로드
3. https://share.streamlit.io/ → Create app → GitHub 저장소 연결 → Main file path `app.py` → Deploy
4. Streamlit Cloud 앱 Settings → Secrets에 아래 값을 등록한 후 재실행:

```toml
APP_PASSWORD = "여기에_길고_고유한_비밀번호"
```

5. 앱에서 로그인 후 정답/배점 입력 → OMR 판독·검수 → 성적표 PDF / XLSX 다운로드

## 매우 중요한 주의
- 이 배포본의 `data/history.json`과 `data/고3더프성적26년.xlsx`에는 학생 식별정보가 포함될 수 있으므로 **공개 저장소를 사용하지 마십시오.** 공유하기 전에 개인정보 처리를 승인받고 필요 시 가명처리하십시오.
- 앱 비밀번호 하나는 기본적인 접근 제한일 뿐, 학교 개인정보 처리에 필요한 완전한 보안체계는 아닙니다. 사용자별 인증, 접근기록, 보관기간, 백업·삭제 정책은 추가 설계가 필요합니다.
- Streamlit Community Cloud의 로컬 파일 저장(`data/work.json`)은 **영구 저장되지 않습니다.** 매 작업 후 `설정·로고·백업` 메뉴에서 **JSON 백업을 다운로드**하고, 재시작 후 **JSON 백업 불러오기**로 복원하십시오.
- 11월 정답/배점이 미등록이고, OMR 좌표는 실제 OMR로 보정/검수해야 합니다. 10월 기준 잠정등급은 11월 공식 등급이 아닙니다.
- GitHub에 `data/work.json`이나 `.streamlit/secrets.toml`을 올리지 마십시오.
