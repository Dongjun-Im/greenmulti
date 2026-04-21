# greenmulti (초록멀티)

소리샘 시각장애인 사이트를 **방향키·Enter만으로** 이용할 수 있게 해주는
데스크톱 유틸리티 v1.2. 초록등대 동호회의 일반 회원 이상이면 사용 가능.

## 진입점·흐름
`main.py` → wx.App → green_auth 로그인 → `_auto_detect_menus` (소리샘 메인
페이지에서 메뉴 URL 자동 감지 → MenuManager 저장) → `main_frame.MainFrame`

## 주요 모듈
- `main_frame.py` (~2.2k줄): 메뉴 / 하위메뉴 / 게시판 / 게시물 뷰 상태 머신
- `page_parser.py` (~1k줄): BeautifulSoup 기반 게시판·게시물·하위메뉴 파싱
- `post_dialog.py` (~1k줄): 게시물 본문 보기·댓글·첨부 저장·URL 목록
- `write_dialog.py`: 게시물 작성·수정
- `theme.py`: 저시력 지원 8가지 테마 + 글꼴 크기 9~40 조절
- `menu_manager.py`, `menu_detector.py`: 메뉴 URL 영구 저장/감지
- `screen_reader.py`: NVDA/센스리더 TTS
- `credentials.py`, `login_dialog.py`, `authenticator.py`: 인증
  (green_auth 패키지와 별도로 한 번 더 존재 — 번들 구성 흔적)
- `config.py`: 경로, 소리샘 URL, 앱 메타 (v1.2.0, 2026-04-18)
- `green_auth/`: 공용 인증 패키지 **번들 복사본** (원본 리포는 별도)

## 화면 단축키 (`data/manual.txt` 참조)
- 메뉴 탐색: ↑↓, Enter, Backspace/Esc(뒤로), Home/End
- 이동: Alt+Home, Alt+G(바로가기), Ctrl+G(페이지), PageUp/Down, Ctrl+F(검색), F5(게시판 새로고침)
- 게시물: W(작성), Alt+M(수정), Alt+D/Delete(삭제), Alt+R(답변)
- 본문: B(TXT저장), Alt+S(첨부), C(댓글 작성), D/M(댓글 삭제/수정), Ctrl+U(URL 목록)
- 테마/글꼴: F7(선택), F6/Shift+F6(다음/이전), Ctrl++/-/0

## Git 워크플로우
원격: https://github.com/Dongjun-Im/greenmulti.git
main 브랜치에서 직접 작업.

사용자가 커밋·푸시 요청 시:
1. `git status` 변경 확인
2. `git diff` 내용 검토
3. `git add -A`
4. `git commit -m "<메시지>"`
5. `git push`

## 민감 파일 (`.gitignore` 로 제외됨)
- `data/credentials.ini`, `data/download_dir.txt` (사용자 로컬 설정)
- `build/`, `dist/`, `__pycache__/`, `.venv*/`

## green_auth 동기화 주의
이 리포의 `green_auth/` 하위는 `\\mac\Home\Downloads\My program\green_auth`
리포의 복사본. 원본이 업데이트되면 여기도 수동으로 맞춰야 한다.
