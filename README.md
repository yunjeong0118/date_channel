# 채널별 편성표 → 요일별 PDF 생성기

두 가지 버전이 들어있습니다.

| 버전 | 파일 | 배포 방식 | 팀원이 필요한 것 |
|---|---|---|---|
| 웹앱 | `app.py` | Streamlit Community Cloud | 브라우저만 있으면 됨 |
| Windows exe | `channel_schedule_gui.py` | GitHub Actions → exe | Excel 설치된 PC |

## 웹앱으로 배포하기 (추천 — "웹배포")

1. 이 폴더를 그대로(= `app.py`, `requirements.txt`, `packages.txt` 포함) GitHub 저장소에 push
2. [share.streamlit.io](https://share.streamlit.io) 접속 → GitHub 계정으로 로그인
3. **New app** 클릭
   - Repository: `yunjeong0118/NO.1` 선택
   - Branch: `main`
   - Main file path: `app.py`
4. **Deploy** 클릭 → 1~3분 후 `https://xxxxx.streamlit.app` 같은 링크 생성
5. 그 링크를 팀원에게 공유하면 끝. 각자 브라우저에서 파일 업로드 → 요일 체크 → PDF 다운로드.

**주의할 점**
- `packages.txt`에 `libreoffice`, `fonts-nanum`이 들어있어야 서버에서 한글 PDF 변환이 정상 동작합니다 (이미 넣어뒀습니다).
- Streamlit Community Cloud 무료 플랜의 앱은 **링크를 아는 사람은 누구나 접속 가능**합니다(저장소가 Private이어도 배포된 앱 자체는 공개 URL). 완전히 접근을 제한하려면 Streamlit의 유료 플랜(뷰어 제한 기능) 또는 사내 서버 배포가 필요해요. 사내용으로 링크만 안 퍼뜨리는 정도로 충분하면 무료 플랜으로 충분합니다.
- 첫 실행 시 LibreOffice가 서버에 설치되느라 1~2분 정도 더 걸릴 수 있어요(이후엔 빠름).
- 코드를 수정해서 push하면 배포된 앱도 자동으로 다시 빌드됩니다.

## Windows exe로 배포하기

`channel_schedule_gui.py`, GitHub Actions 워크플로우 관련 안내는 이전 답변을 참고하세요.
(Excel이 설치된 팀원 PC용 — 웹앱이 안 맞는 경우의 대안입니다.)

## 공통 로직

- 채널 인식: 시트탭 이름/파일명 안의 키워드로 판단 (`detect_channel_name`)
- 입력 파일 전제: SPOTV 주간 편성표 표준 양식 (첫 시트 = 이번 주, 6행 헤더, 7~201행 데이터)
- 두 버전 모두 동일한 취합 로직을 쓰고, PDF 변환 엔진만 다릅니다
  (웹앱 = LibreOffice, exe = Excel COM)
