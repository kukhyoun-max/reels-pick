# 요즘 본 릴스 알려줘

전주대학교 학생들이 실제로 반응한 릴스·쇼츠를 모으는 제보 폼입니다.

- `index.html` — 학생용 폼 (GitHub Pages로 배포)
- `Code.gs` — 비공개 구글 시트에 저장하는 Apps Script

## 연결 방법

1. 비공개 구글 시트 → 확장 프로그램 → Apps Script → `Code.gs` 내용 붙여넣기
2. 배포 → 새 배포 → 웹 앱 (실행 계정: 나 / 액세스: 모든 사용자) → `/exec` URL 복사
3. `index.html`의 `GAS_URL` 값에 그 URL 붙여넣고 커밋
4. 학생에게 GitHub Pages 주소 공유

시트는 공유하지 않아 소유자만 볼 수 있으며, 제보는 Apps Script가 소유자 권한으로 대신 기록합니다.
