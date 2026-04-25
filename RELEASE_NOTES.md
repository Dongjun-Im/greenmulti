## 변경 사항

- 자동 업데이트로 받는 자산 파일 이름을 깔끔한 ASCII 이름으로 정리했습니다.
  - 설치형: `chorokmulti_v1.6_setup.exe`
  - 무설치(포터블): `chorokmulti_v1.6.3.zip` — 압축을 풀면 폴더명은 `초록멀티 v1.6` 으로 그대로 유지됩니다.
  - 체크섬: `chorokmulti_v1.6_setup.exe.sha256`
- 그동안 `_v1.6_setup.exe`, `v1.6.zip` 처럼 한글이 잘려 어색하게 표시되던 문제 해결 (GitHub 가 릴리스 자산 파일명에 비-ASCII 문자를 허용하지 않아, 한글 대신 안전한 영문 표기 `chorokmulti` 를 접두사로 채택).
