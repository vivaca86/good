# 배포 체크리스트

## 1) Apps Script 배포
1. `apps_script.gs` 코드를 스프레드시트에 연결된 Apps Script 프로젝트에 반영
2. 웹 앱으로 배포(Deploy as web app)
   - Execute as: **Me**
   - Who has access: **Anyone**(또는 조직 정책에 맞는 최소 권한)
3. 발급된 `/exec` URL 확보
4. `index.html`의 `SCRIPT_URL`을 최신 URL로 갱신

## 2) 시트 구조 확인
- `SLOTS` 시트: `zone, board, slot_no, code, name, spec, qty` 컬럼 순서
- `USAGE_LOG` 시트 존재
- `DB` 시트 존재(코드=B, 품명=C, 재고=F 컬럼 기준)
- (선택) `APP_ERROR_LOG` 시트 존재

## 3) 정적 파일 배포
- `index.html`
- `manifest.webmanifest`
- `service-worker.js`
- `icons/icon.svg`

같은 경로 기준으로 호스팅해야 하며, 루트 기준 상대경로(`./`)가 유지되어야 함.

## 4) 브라우저/캐시 주의
- `service-worker.js` 변경 시 `CACHE_NAME` 버전을 증가시켜 강제 갱신
- 기존 서비스워커가 남아 있을 수 있으므로, 초기 배포 직후 강력 새로고침 권장

## 5) 최종 스모크 테스트
1. 페이지 로드 후 1~4구역 데이터 표시 확인
2. 편집 모드 진입 → 값 변경 → 완료(저장) 확인
3. DB구역 진입 후 비교 결과 표시 확인
4. 검색 입력 시 행 필터링 확인
5. PWA 설치 버튼 노출/동작 확인(지원 브라우저)
