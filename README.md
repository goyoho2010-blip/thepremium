# 대성 THE PREMIUM – 11월 9일 누락 정답 보완판

- **정치와 법** 20문항, 50점 / **사회·문화** 20문항, 50점: 사용자가 제공한 11월 9일 이미지 원본 전사.
- `reference/november_additional_keys.json`을 읽어, 앱 실행 시 **기존 서버 JSON 결과는 그대로 유지하고** 누락된 두 정답표만 추가합니다.
- 기존 11월 OMR 확정 답안은 **재업로드 불필요**, 학생 성적표를 다시 조회하면 채점됩니다.
- 다른 교과 원점수, 학생별 기록, 3~10월 이력, PDF 출력 양식 및 3개년 수능 참고등급 산정은 그대로 유지합니다.
- 원래 정답표를 관리자가 별도로 20문항 확정한 경우 내장 자료가 덮어쓰지 않습니다.
- 정답 엑셀도 `reference/thepremium-answers-11.xlsx`에 두 과목을 통합해 새로 포함했습니다.

## GitHub 반영 순서
1. 먼저 현재 Streamlit `설정·로고·백업`에서 **전체 JSON 백업**을 받습니다.
2. ZIP 내 `app.py`, `reference/november_additional_keys.json`, `reference/thepremium-answers-11.xlsx`를 GitHub 저장소의 동일 경로에 업로드/교체합니다. JSON/엑셀은 `reference` 폴더에 넣어야 합니다.
3. Commit changes → Streamlit 재배포 → `학생 성적표`에서 사회·문화, 정치와 법 학생을 조회합니다.
4. 성적을 확인한 후 다시 JSON 백업을 받습니다. Community Cloud의 로컬 파일 저장은 영구 저장을 보장하지 않습니다.

## 테스트
- 문항수 20 + 배점 50 각각 확인, 정답값 1~5 확인, 코드 문법 확인.
- 합성 학생 OMR (20문항 모두 정답)은 각각 50점, 전체 오답은 0점으로 계산.
- 실제 서버 저장 데이터와 학생 명부의 교차 검증은 배포 후 확인이 필요합니다.
