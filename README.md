# 오픈마켓 주문 취합 자동화

여러 오픈마켓에서 다운로드한 원본 주문서(xlsx)를 협력사별로 분류·취합하고,
각 협력사가 요구하는 엑셀 양식으로 자동 변환해주는 도구.

## 사용법

```bash
pip install -r requirements.txt
python3 main.py 주문서1.xlsx 주문서2.xlsx ... --out output
```

- 입력: 오픈마켓에서 받은 원본 주문서 xlsx 파일들 (여러 마켓 것을 한 번에 넣어도 됨)
- 출력(`output/` 폴더):
  - 협력사별로 `{협력사명}_{날짜}.xlsx` 파일 (해당 협력사 양식 그대로)
  - 어느 협력사에도 매칭되지 않았거나(미분류), 여러 협력사 키워드에 동시에 걸린(중복매칭) 주문은 `확인필요_{날짜}.xlsx`에 모아서 수동 확인하도록 함
  - `summary_{날짜}.json` — 마켓별/협력사별 주문 건수, TOP5 판매 제품 등 요약 통계
  - `dashboard_{날짜}.html` — 위 요약을 보고용 대시보드로 시각화한 페이지 (Claude와 대화 중이면 이 파일을 Artifact로 바로 열어서 보여줌)

## 동작 방식

1. **`src/normalize.py`** — 오픈마켓 원본 주문서를 읽어 표준 필드(주문번호/상품명/옵션/수량/수령인/연락처/주소 등)로 정규화한다. 마켓마다 헤더 이름이 달라서 `config/column_aliases.json`에 등록된 이름으로 헤더 행을 자동으로 찾는다. **새 오픈마켓 양식을 추가할 때는 이 파일에 헤더 이름만 추가하면 된다.**
2. **`src/classify.py`** — 상품명/옵션에 협력사의 브랜드 키워드(`config/vendors.json`의 `brand_keywords`)가 포함되어 있으면 그 협력사로 분류한다.
3. **`src/writer.py`** — `templates/{협력사}.xlsx` (실제 고객정보는 제거된 빈 양식)를 복사해 분류된 주문을 채워 넣는다. 헤더 텍스트를 표준 필드에 자동으로 매핑하므로, 새 협력사 양식을 추가할 때도 대부분 설정만으로 처리된다.
4. **`src/summary.py`, `src/dashboard.py`** — 분류 결과에서 마켓별/협력사별 고유 주문 건수와 전체 TOP5 판매 제품(판매수량 기준)을 집계하고, 보고하기 좋은 대시보드 HTML로 렌더링한다.
5. **`src/reconcile.py`** — 원본 대비 처리 건수가 일치하는지 확인하고, 배송메시지(특이사항)·정보 누락(수령인명/연락처/주소) 건을 모은다.

## 대시보드 날짜별 기록 (Artifact db)

대시보드는 [Artifact db capability](https://claude.ai/code/artifact/81d801cb-0c8f-4bce-ae90-ca9c97f268f3)를 이용해 매일의 `summary_{날짜}.json`을 누적 저장하고, "오늘/최근 7일/이번 달/이번 분기/특정 날짜"를 골라 볼 수 있다.

**매일 실행 후 이 순서로 반영한다** (Claude가 수행):
1. `python3 main.py ...` 실행 → `output/summary_{날짜}.json`, `output/dashboard_{날짜}.html` 생성
2. Artifact `write_db` (db_op: set)로 그날 summary를 `daily_summaries/{날짜}` 문서에 저장 (`file_path`로 summary json 그대로 넘기면 됨)
3. Artifact 퍼블리시로 `dashboard_{날짜}.html`을 같은 URL에 재배포 (db 접근이 없는 뷰어를 위한 폴백 겸, 최신 날짜의 기본 화면)

대시보드 페이지 자체는 로드 시 `daily_summaries` 컬렉션 전체를 읽어 기간별로 합산하므로, 2번을 빼먹지 않는 한 별도 코드 수정 없이 계속 쌓인다.

## 협력사 현황 (7개)

`config/vendors.json`에서 관리한다. `_needs_confirmation`이 달린 항목은 샘플 몇 건만 보고 추정한 것이니 실제 분류가 틀리면 이 파일의 `brand_keywords`만 고치면 된다.

| 협력사 | 브랜드 키워드 | 비고 |
|---|---|---|
| 잔망두유 | 잔망두유 | |
| 작심랩 | 작심랩 | |
| 어플러드 | 어플러드 | |
| 이플코리아 | 데이너프 | 확인 필요 |
| 해인 | 햇살듬뿍 | 발송인 정보(F01-결제조건/F02-택배운임 등) 고정값 확인 필요 |
| 안국약품 | 비타민D, 칼마디, 브이팩, 애사비젤리 | 전체 상품 목록 필요 (샘플 5건 기준이라 매우 불완전) |
| 마더네스트 | 마더네스트 | 최종 제출 양식 확인 완료 |

## 새 협력사 양식 추가하는 법

1. 협력사가 준 엑셀 양식에서 실제 고객정보(샘플 행)를 지우고 `templates/{협력사명}.xlsx`로 저장
2. `config/vendors.json`의 `vendors`에 항목 추가:
   - `brand_keywords`: 이 협력사 상품을 식별할 키워드 목록
   - `template_file`, `header_row`, `data_start_row`
   - 헤더 이름이 `config/column_aliases.json`에 없는 새로운 이름이면 그 파일에 추가
   - 발송인 정보처럼 매번 고정인 값이 있으면 `static_fields`에 등록
3. `python3 main.py 테스트파일.xlsx`로 확인

## 남은 작업

- 안국약품 전체 브랜드/상품 목록 확보
- 이플코리아/해인/마더네스트 브랜드 키워드 및 고정값 확인
- 나머지 오픈마켓(10개 이상 예정) 원본 주문서 샘플을 받는 대로 `config/column_aliases.json`에 헤더 이름 보강
