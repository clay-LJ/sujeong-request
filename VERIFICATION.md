# 검증 기록

2026-11-19, macOS의 실제 Playwright Chromium에서 확인했습니다. 사용자 실제 샘플은 제공되지 않았으며 합성 자료만 사용했습니다. 실제 업무용 PPTX의 PowerPoint 대비 동일성 검증을 의미하지 않습니다.

## 결과

| 확인 항목 | 결과 |
| --- | --- |
| 프로덕션 빌드 | 통과 |
| 로컬 실제 브라우저 검증 | 7 / 7 통과 |
| 공개 Pages 실제 브라우저 검증 | 7 / 7 통과 |
| PPTX 3장 직접 가져오기 | 통과: 한글 텍스트·표·도형 |
| PDF 2페이지 / 이미지 가져오기 | 통과 |
| 페이지 이동·번호·요청 입력·수정·삭제 | 통과 |
| 200% 확대 후 화면 폭 변경 시 상대 좌표 | 통과: 원본 x=30%, y=55% 유지 |
| 모바일 390px 레이아웃 | 통과: 문서 가로 넘침 없음, 입력·탐색 사용 가능 |
| 원본·번호·한글 요청을 포함한 PDF 다운로드 | 통과 |
| 긴 한국어 요청의 추가 페이지 출력 | 통과 |
| PDF에서 다른 파일로 교체 | 통과 |
| 구형 .ppt / 잘못된 PPTX 오류 안내 | 통과 |
| 외부 연결 이미지가 포함된 PPTX | 외부 요청 0건 |
| 전체 흐름의 네트워크 | 원본·요청 내용 전송 0건, 같은 출처 GET 및 내부 blob/data만 사용 |
| npm audit | 전체 및 운영 의존성 취약점 0건 |

출력 PDF를 PDF.js로 다시 렌더링해 한글·표·도형·번호·요청 내용과 페이지 수를 확인하고, 출력 이미지와 모바일 화면을 시각적으로 확인했습니다. 한글은 이미지 기반으로 보존되며 텍스트 검색·선택은 지원하지 않습니다. 기본 PDF 이미지 해상도와 PPTX 미리보기의 글꼴·배치에는 원본 대비 차이가 있을 수 있습니다.

## 재현

```sh
npm ci
npx playwright install chromium
npm run dev
# 별도 터미널
npm test
BASE_URL=https://clay-lj.github.io/sujeong-request/ npm test
```

테스트 입력은 `public/samples/`의 합성 파일입니다. 외부 미디어 차단 시험은 테스트 실행 중 합성 PPTX에 가짜 외부 이미지 관계를 넣어 수행합니다. 외부로 보내는 테스트 API는 없습니다. 테스트 스크린샷·PDF·네트워크 기록은 git에서 제외된 `test-results/`에 저장됩니다.

로컬 개발 서버에서는 `tests/pdf-inspect.js`를 통해 PDF를 다시 렌더링합니다. 공개 배포 검증은 실제 다운로드와 페이지 수 안내·UI·네트워크를 검사하며, 테스트 보조 모듈은 사이트 빌드에 포함하지 않습니다.

## 게시 상태

- 사이트: https://clay-lj.github.io/sujeong-request/
- 저장소: https://github.com/clay-LJ/sujeong-request
- 최초 배포 워크플로: https://github.com/clay-LJ/sujeong-request/actions/runs/44501766199 (build / deploy 모두 성공)
- 실제 샘플 검토와 실제 iOS·Android 기기 검증은 추후 사용자 시험 대상입니다.
