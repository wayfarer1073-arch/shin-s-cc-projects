"""처리한 주문 원본 라인을 날짜별로 영구 보관한다 (data/orders/{날짜}.json).

대시보드용 summary는 집계값만 담아서 개별 주문 상세(수취인/주소/주문번호 등)는
남지 않는다. 이후에 특정 기간에 대해 정산 리포트 같은 건별 상세가 필요한
작업을 하려면 원본 라인이 있어야 하므로, 매 실행마다 이 파일에도 함께 쌓아둔다.
"""
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
ARCHIVE_DIR = BASE_DIR / "data" / "orders"

_FIELDS = [
    "order_id", "order_date", "product_name", "option", "quantity",
    "receiver_name", "receiver_phone", "address", "zipcode",
    "delivery_message", "courier", "tracking_no", "shop_name", "payment_amount",
    "brand", "channel_product_no", "_source_file", "_source_row",
    "_ambiguous_matches",
]


def _serialize(rec, vendor, match_status):
    row = {f: rec.get(f) for f in _FIELDS}
    if row["order_date"] is not None:
        row["order_date"] = row["order_date"].isoformat()
    row["vendor"] = vendor
    row["match_status"] = match_status
    return row


def archive_orders(by_vendor, unclassified, ambiguous, run_date, archive_dir=ARCHIVE_DIR):
    """그날 처리한 주문 라인(분류 결과 포함)을 archive_dir/{run_date}.json에 저장한다.
    같은 날짜에 이미 보관된 내용이 있으면 그 위에 이어붙인다(하루에 여러 원본
    파일을 나눠 처리하는 경우 — 예: 스마트스토어분 + 카카오분 — 먼저 처리한
    내용이 사라지지 않도록). 단, 같은 원본 파일의 같은 줄(_source_file +
    _source_row)이 이미 보관돼 있으면 중복으로 또 쌓지 않는다(같은 파일을
    실수로 두 번 올려도 안전)."""
    rows = []
    for vendor, recs in by_vendor.items():
        rows.extend(_serialize(r, vendor, "classified") for r in recs)
    rows.extend(_serialize(r, None, "unclassified") for r in unclassified)
    rows.extend(_serialize(r, None, "ambiguous") for r in ambiguous)

    archive_dir = Path(archive_dir)
    archive_dir.mkdir(parents=True, exist_ok=True)
    out_path = archive_dir / f"{run_date}.json"

    existing_rows = []
    if out_path.exists():
        with open(out_path, encoding="utf-8") as f:
            existing_rows = json.load(f)
    seen = {(r.get("_source_file"), r.get("_source_row")) for r in existing_rows}
    new_rows = [r for r in rows if (r.get("_source_file"), r.get("_source_row")) not in seen]

    combined_rows = existing_rows + new_rows
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(combined_rows, f, ensure_ascii=False, indent=2)
    return out_path


def load_orders(start_date, end_date, vendor=None, brand=None, archive_dir=ARCHIVE_DIR):
    """start_date~end_date(둘 다 포함, "YYYY-MM-DD") 사이에 보관된 주문 라인을 불러온다.
    vendor를 지정하면 그 협력사로 분류된 것만, brand를 지정하면 그 브랜드 것만 반환한다."""
    archive_dir = Path(archive_dir)
    rows = []
    if not archive_dir.exists():
        return rows
    for path in sorted(archive_dir.glob("*.json")):
        run_date = path.stem
        if not (start_date <= run_date <= end_date):
            continue
        with open(path, encoding="utf-8") as f:
            day_rows = json.load(f)
        for r in day_rows:
            r["_run_date"] = run_date
            if vendor is not None and r.get("vendor") != vendor:
                continue
            if brand is not None and r.get("brand") != brand:
                continue
            rows.append(r)
    return rows
