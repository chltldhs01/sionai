# 바이맘 네이버 SA 주간보고서 생성기

네이버 SA CSV와 운영 메모를 입력하면 광고주 공유용 엑셀 주간보고서를 생성하는 로컬 웹앱입니다.

## 실행 방법

1. [start-report-generator.bat](/C:/Users/user/Documents/Codex/2026-05-19/https-github-com-coreyhaines31-marketingskills-vscode/start-report-generator.bat) 실행
2. 브라우저에서 `http://localhost:3030` 접속
3. 브랜드명, 보고 기간, CSV 파일, 운영 메모 입력
4. `주간보고서 생성` 버튼 클릭
5. 생성 완료 후 `엑셀 다운로드`

## 생성되는 시트

- 광고주 요약
- 전체 요약
- 일자별 추이
- 파워링크 상세
- 쇼핑검색 상세
- 키워드 상세
- 운영 메모
- 원본 데이터

## 입력 파일

- 네이버 SA에서 다운로드한 CSV 파일
- 권장: 파워링크, 쇼핑검색 데이터가 포함된 기간별 계정 보고서

## 현재 결과 예시

- [Bymom_주간보고서_20260514_20260518.xlsx](/C:/Users/user/Documents/Codex/2026-05-19/https-github-com-coreyhaines31-marketingskills-vscode/outputs/Bymom_주간보고서_20260514_20260518.xlsx)

## 주요 파일

- [start-report-generator.bat](/C:/Users/user/Documents/Codex/2026-05-19/https-github-com-coreyhaines31-marketingskills-vscode/start-report-generator.bat)
- [server.mjs](/C:/Users/user/Documents/Codex/2026-05-19/https-github-com-coreyhaines31-marketingskills-vscode/server.mjs)
- [public/index.html](/C:/Users/user/Documents/Codex/2026-05-19/https-github-com-coreyhaines31-marketingskills-vscode/public/index.html)
- [public/app.js](/C:/Users/user/Documents/Codex/2026-05-19/https-github-com-coreyhaines31-marketingskills-vscode/public/app.js)
- [src/reporting.mjs](/C:/Users/user/Documents/Codex/2026-05-19/https-github-com-coreyhaines31-marketingskills-vscode/src/reporting.mjs)
