"""기간을 지정하면 보관된 주문 원본(data/orders/)에서 해당 협력사 주문을 모아
정산 리포트 양식(templates/정산_*.xlsx)에 채워 넣는다.

원가/수수료/택배비처럼 우리 시스템이 갖고 있지 않은 값(관리상품명, 매입단가,
매입액, 이익액, 율, 택배 관련 컬럼)은 채우지 않고 공란으로 남긴다 — 사람이
템플릿에 같이 들어있는 '관리상품명'/'수수료정리' 참고 시트를 보고 나중에
채우도록 되어 있는 구조를 그대로 따른다.
"""
import json
import tempfile
from datetime import date, datetime
from pathlib import Path

import openpyxl

from src.archive import load_orders, ARCHIVE_DIR
from src.protect import protect_xlsx

BASE_DIR = Path(__file__).resolve().parent.parent

with open(BASE_DIR / "config" / "settlement_reports.json", encoding="utf-8") as f:
    _FULL_CFG = json.load(f)
    _CFG = _FULL_CFG["vendors"]
    EXPORT_PASSWORD = _FULL_CFG.get("export_password")

# 열 번호(1-based) -> 의미. templates/정산_해인.xlsx (TEMPLATE(JM)) 기준.
COL = {
    "no": 1, "year": 2, "recv_month": 3, "recv_day": 4,
    "sales_channel": 7, "revenue_channel": 8, "pay_vendor": 9,
    "return_office": 10, "ship_office": 11, "order_id": 12,
    "receiver_name": 13, "receiver_phone": 14, "address": 15, "zipcode": 16,
    "product_name": 17, "option": 18, "managed_name": 19,
    "shipment_count": 20, "quantity": 21, "revenue": 22, "settlement": 23,
    "sender_name": 28, "sender_phone": 29, "sender_address": 30,
    "delivery_message": 31, "note": 42,
}


def _parse_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date() if isinstance(s, str) else s


def generate_settlement_report(
    vendor, start_date, end_date, out_path, archive_dir=ARCHIVE_DIR,
    protect=True, password=None,
):
    """start_date, end_date: "YYYY-MM-DD" 문자열. 반환: (out_path, 채운 건수).

    protect=True(기본값)면 결과 파일에 password(기본은 config의 export_password,
    즉 조직 공용 열람 암호)로 열기암호를 걸어 out_path에 저장한다. 개인정보가
    들어있는 파일이라 권한 없는 사람이 못 열게 하려는 목적."""
    if vendor not in _CFG:
        raise ValueError(f"'{vendor}' 협력사의 정산 리포트 양식이 아직 설정되어 있지 않습니다.")
    cfg = _CFG[vendor]

    rows = load_orders(start_date, end_date, vendor=vendor, archive_dir=archive_dir)
    rows.sort(key=lambda r: (r.get("order_date") or r.get("_run_date") or "", r.get("order_id") or ""))

    wb = openpyxl.load_workbook(BASE_DIR / cfg["template_file"])
    ws = wb[cfg["template_sheet"]]

    year = _parse_date(start_date).year
    ws[cfg["title_cell"]] = cfg["title_template"].format(year=year)

    static_by_col = cfg.get("static_by_col", {})
    data_start = cfg["data_start_row"]

    seen_orders = set()
    for i, rec in enumerate(rows):
        r = data_start + i
        order_id = rec.get("order_id")
        is_first = order_id not in seen_orders if order_id else True
        if order_id:
            seen_orders.add(order_id)

        od = _parse_date(rec.get("order_date")) or _parse_date(rec.get("_run_date"))

        ws.cell(row=r, column=COL["no"], value=i + 1)
        if od:
            ws.cell(row=r, column=COL["year"], value=od.year)
            ws.cell(row=r, column=COL["recv_month"], value=od.month)
            ws.cell(row=r, column=COL["recv_day"], value=od.day)
        ws.cell(row=r, column=COL["sales_channel"], value=rec.get("shop_name"))
        ws.cell(row=r, column=COL["order_id"], value=order_id)
        ws.cell(row=r, column=COL["receiver_name"], value=rec.get("receiver_name"))
        ws.cell(row=r, column=COL["receiver_phone"], value=rec.get("receiver_phone"))
        ws.cell(row=r, column=COL["address"], value=rec.get("address"))
        ws.cell(row=r, column=COL["zipcode"], value=rec.get("zipcode"))
        ws.cell(row=r, column=COL["product_name"], value=rec.get("product_name"))
        ws.cell(row=r, column=COL["option"], value=rec.get("option"))
        ws.cell(row=r, column=COL["shipment_count"], value=1 if is_first else None)
        ws.cell(row=r, column=COL["quantity"], value=rec.get("quantity"))
        revenue = rec.get("payment_amount")
        ws.cell(row=r, column=COL["revenue"], value=revenue)
        if revenue is not None:
            ws.cell(row=r, column=COL["settlement"], value=f"=V{r}*0.85")
        ws.cell(row=r, column=COL["delivery_message"], value=rec.get("delivery_message"))

        for col_letter, value in static_by_col.items():
            col_idx = openpyxl.utils.column_index_from_string(col_letter)
            ws.cell(row=r, column=col_idx, value=value)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if not protect:
        wb.save(out_path)
        return out_path, len(rows)

    pwd = password or EXPORT_PASSWORD
    if not pwd:
        raise ValueError("protect=True인데 사용할 비밀번호가 없습니다 (config의 export_password 확인).")
    with tempfile.TemporaryDirectory(prefix="settlement-unprotected-") as tmp_dir:
        tmp_path = Path(tmp_dir) / out_path.name
        wb.save(tmp_path)
        protect_xlsx(tmp_path, out_path, pwd)
    return out_path, len(rows)
