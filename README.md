# 한익스프레스 스마트오더 라벨 생성기

판매채널(**CU / GS / 이마트24**)의 스마트오더 RAW 파일을 업로드하면,
**코드 기반 매핑**으로 대월직매장 부착양식(A4) 라벨을 생성하는 웹앱입니다.

## 특징
- **코드 기반 매핑** (명칭 변동에 강함)
  - 센터코드 → 거점센터(이고센터)  · GS는 GS25(VA코드)/GS슈퍼(숫자코드) 통합, E24는 입고센터명 키
  - 상품코드 → 대표상품명 (전 채널 공통 마스터)
- **미매핑 코드 즉시 등록**: 업로드 중 안 잡힌 센터/상품 코드를 화면에서 바로 마스터에 추가 → 다음부터 자동 적용
- **마스터 관리** UI (조회·검색·추가·수정·삭제)
- **출력**
  - 부착양식(A4) **xlsx 다운로드** — 기존 양식과 동일한 스타일(맑은 고딕, 62/42/48pt, 테두리, 노란 번호)
  - 화면에서 **TOTAL 표지** 즉시 확인 + **클립보드 복사** + 표지 인쇄
  - 화면에서 **라벨 미리보기 + 바로 인쇄**(A4, 라벨 1장=1페이지)
- 라벨은 창고 분류에 맞춰 **거점센터 → 대표상품 → 점포명** 순 정렬

## 구성
```
app.py                 Flask 라우팅
smartorder_core.py     RAW 파싱 · 코드 매핑 · xlsx 생성 엔진
bootstrap_masters.py   기존 파일에서 seed 마스터 생성(1회용)
masters/*.json         코드 기반 마스터(seed)
templates/             화면(index/upload/result/print/masters...)
deploy/                Lightsail(nginx+systemd) 배포 스크립트
```

## 로컬 실행 (Windows)
```
py -m pip install -r requirements.txt
py app.py           # 또는 실행.bat 더블클릭 → 브라우저 자동 열림
```
기본 포트 5057 · http://127.0.0.1:5057/

## AWS Lightsail 배포 (Ubuntu)
```bash
curl -fsSL https://raw.githubusercontent.com/sangjin94/hanex_smartorder/main/deploy/setup-smartorder.sh | bash
# 코드 갱신
~/hanex_smartorder/deploy/update-smartorder.sh
```
- 앱은 gunicorn(127.0.0.1:8090) + systemd(`smartorder.service`), nginx 리버스 프록시
- **마스터는 `~/hanex_smartorder_data/masters`(git 밖)** 에 저장 → `git pull` 갱신에도 사용자가 수정한 매핑이 보존됨
- hanex_tool 사이트 하위경로(`/smartorder/`)로 붙이려면 `deploy/nginx-smartorder-location.conf` 참고 (앱이 `X-Forwarded-Prefix` 지원)

## RAW 인식 규칙
헤더 이름으로 컬럼을 자동 인식합니다.
- CU: `센터 코드`,`센터명`,`상품 코드`,`상품명`,`총수량`
- GS: `센터코드`,`센터`,`상품코드`,`상품명`,`수량` (xls/xlsx 모두)
- E24: `입고센터명`,`점포명`,`상품코드`,`앱 상품명`,`주문수량`
