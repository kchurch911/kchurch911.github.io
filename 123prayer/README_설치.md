# 123 Prayer (사역자 기도 언약) — 설치 안내

## 구성
| 파일 | 용도 | 주소 (배포 후) |
|---|---|---|
| `index.html` | 사역자용 기도 제목 제출 폼 (한/영) | https://kchurch911.com/123prayer/ |
| `pray.html` | 김성수 기도방 — 목록 보기, "기도함" 기록, 상태 변경 (비밀번호) | https://kchurch911.com/123prayer/pray.html |
| `Apps_Script_123Prayer.gs` | Google Sheet 저장 백엔드 (Apps Script 웹 앱) | — |

## 1. Apps Script 배포 (한 번만, 5분)
1. sheets.google.com → 새 스프레드시트 → 이름 **123 사역자 기도**
2. 메뉴 **확장 프로그램 → Apps Script** → 코드 칸 비우고 `Apps_Script_123Prayer.gs` 전체 붙여넣기
3. 코드 위쪽 `PRAY_KEY = 'kingdom123'` 를 **원하는 비밀번호**로 바꾸기 (기도방 입장 비밀번호)
   - 새 기도 제목이 올 때 이메일 알림을 받으려면 `NOTIFY_EMAIL = ''` 에 이메일 입력
4. Ctrl+S → **배포 → 새 배포** → 톱니바퀴 → **웹 앱** → 실행: **나** / 액세스: **모든 사용자** → **배포** → 권한 허용(고급 → 이동 → 허용)
5. **웹 앱 URL**(…/exec) 복사

## 2. 페이지에 URL 넣기
`index.html` 과 `pray.html` 두 파일 모두에서
```
const API_URL = '';
```
→ `const API_URL = 'https://script.google.com/macros/s/…/exec';`

(Claude에게 URL을 붙여넣어 주면 넣고 배포까지 해 드립니다.)

## 3. 배포
`01_Kchurch911_Coaching` 폴더에서 git commit → push (GitHub Pages 자동 반영, 1~2분)

## 4. 사역자들에게 보낼 링크
https://kchurch911.com/123prayer/  — 카카오톡·문자·이메일로 전달

## 사용
- 사역자: 이름·섬기는 자리·기도 제목 1~2개·긴급 여부·공개 범위·123 언약 체크 → 보내기
- 김성수: `pray.html` → 비밀번호 → "기도 필요" 탭에 이번 주 2회가 안 된 사역자가 먼저 보임 → **기도함 — 3분** 버튼 (하루 1회, 주 2회 채우면 "이번 주 완료") → 상태(기도중/응답됨/부분응답/종료) 변경 가능 → 인쇄 버튼으로 기도 목록 출력
- 기도 기록·상태는 Google Sheet에 저장되어 휴대폰·노트북 어디서든 같은 화면
