"""정규화된 주문 라인을 협력사별 엑셀 양식(templates/*.xlsx)에 채워 넣는다."""
import json
from collections import Counter
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

# 행 강조 색상: 수량 2개 이상 / 동일 수령인+주소 중복 / 둘 다 해당 시 각각 다른 색으로 구분
QTY_FILL = PatternFill(fill_type="solid", start_color="FFF9E48B", end_color="FFF9E48B")
DUP_FILL = PatternFill(fill_type="solid", start_color="FFBFE0F5", end_color="FFBFE0F5")
BOTH_FILL = PatternFill(fill_type="solid", start_color="FFF7C592", end_color="FFF7C592")

BASE_DIR = Path(__file__).resolve().parent.parent

with open(BASE_DIR / "config" / "vendors.json", encoding="utf-8") as f:
    VENDORS = json.load(f)["vendors"]

with open(BASE_DIR / "config" / "column_aliases.json", encoding="utf-8") as f:
    _aliases_raw = json.load(f)

FIELD_TO_HEADERS = {k: v for k, v in _aliases_raw.items() if not k.startswith("_")}

_HEADER_TO_FIELD = {}
for _field, _names in FIELD_TO_HEADERS.items():
    for _name in _names:
        _HEADER_TO_FIELD[_name] = _field

_NO_HEADERS = {"NO", "No.", "No", "no", "번호"}


def _build_col_map(ws, cfg):
    """열 번호 -> ("static", 값) | ("combine",) | ("seq",) | ("year"/"month"/"day",)
    | ("shipment_flag",) | ("field", 표준필드명) 로 매핑한다. 매핑되지 않은 열은
    자동으로 채우지 않고 비워둔다(택배사/송장번호 등 나중에 수기로 채우는 열)."""
    header_row = cfg["header_row"]
    field_overrides = cfg.get("field_overrides", {})
    static_fields = cfg.get("static_fields", {})
    combine_field = cfg.get("combine_product_option_field")

    col_map = {}
    for c in range(1, ws.max_column + 1):
        header = ws.cell(row=header_row, column=c).value
        if not isinstance(header, str):
            continue
        header = header.strip()
        if not header:
            continue

        if header in static_fields:
            col_map[c] = ("static", static_fields[header])
        elif combine_field and header == combine_field:
            col_map[c] = ("combine",)
        elif header in field_overrides:
            col_map[c] = ("field", field_overrides[header])
        elif header in _NO_HEADERS:
            col_map[c] = ("seq",)
        elif header == "년":
            col_map[c] = ("year",)
        elif header == "월":
            col_map[c] = ("month",)
        elif header == "일":
            col_map[c] = ("day",)
        elif header == "배송건수":
            col_map[c] = ("shipment_flag",)
        elif header in _HEADER_TO_FIELD:
            col_map[c] = ("field", _HEADER_TO_FIELD[header])
    return col_map


def _capture_row_style(ws, row_idx, max_col):
    styles = {}
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row_idx, column=c)
        styles[c] = {
            "font": copy(cell.font),
            "fill": copy(cell.fill),
            "border": copy(cell.border),
            "alignment": copy(cell.alignment),
            "number_format": cell.number_format,
        }
    return styles


def _cell_value_for(kind_spec, rec, seq, is_first_of_order):
    kind = kind_spec[0]
    if kind == "seq":
        return seq
    if kind == "year":
        return f"{rec['order_date'].year}년" if rec.get("order_date") else None
    if kind == "month":
        return rec["order_date"].month if rec.get("order_date") else None
    if kind == "day":
        return rec["order_date"].day if rec.get("order_date") else None
    if kind == "static":
        return kind_spec[1]
    if kind == "combine":
        parts = [p for p in [rec.get("product_name"), rec.get("option")] if p]
        return " ".join(parts)
    if kind == "shipment_flag":
        return 1 if is_first_of_order else None
    if kind == "field":
        return rec.get(kind_spec[1]) or None
    return None


def _find_col_by_kind(col_map, predicate):
    for c, spec in col_map.items():
        if predicate(spec):
            return c
    return None


def _dup_key(rec):
    name = (rec.get("receiver_name") or "").strip().casefold()
    addr = (rec.get("address") or "").strip().casefold()
    if not name or not addr:
        return None
    return (name, addr)


def _highlight_fill_for(rec, dup_counts):
    is_qty = (rec.get("quantity") or 1) > 1
    key = _dup_key(rec)
    is_dup = key is not None and dup_counts[key] >= 2
    if is_qty and is_dup:
        return BOTH_FILL
    if is_qty:
        return QTY_FILL
    if is_dup:
        return DUP_FILL
    return None


def write_vendor_file(vendor_name, rows, out_path):
    """vendor_name 협력사 양식 템플릿을 복사해 rows(표준 필드 dict 리스트)를 채운다."""
    cfg = VENDORS[vendor_name]
    template_path = BASE_DIR / cfg["template_file"]
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]

    col_map = _build_col_map(ws, cfg)
    data_start = cfg["data_start_row"]
    styles = _capture_row_style(ws, data_start, ws.max_column)

    dup_counts = Counter(k for k in (_dup_key(r) for r in rows) if k is not None)
    highlight_counts = {"quantity": 0, "duplicate_address": 0, "both": 0}

    seen_orders = set()
    for i, rec in enumerate(rows):
        r = data_start + i
        order_id = rec.get("order_id")
        is_first = order_id not in seen_orders if order_id is not None else True
        if order_id is not None:
            seen_orders.add(order_id)

        for c, spec in col_map.items():
            value = _cell_value_for(spec, rec, i + 1, is_first)
            cell = ws.cell(row=r, column=c)

            if spec[0] == "field" and spec[1] == "order_date" and value is not None:
                cell.value = value
                cell.number_format = "yyyy-mm-dd"
            else:
                cell.value = value

            st = styles.get(c)
            if st:
                cell.font = st["font"]
                cell.fill = st["fill"]
                cell.border = st["border"]
                cell.alignment = st["alignment"]
                if not (spec[0] == "field" and spec[1] == "order_date"):
                    cell.number_format = st["number_format"]

        fill = _highlight_fill_for(rec, dup_counts)
        if fill is QTY_FILL:
            highlight_counts["quantity"] += 1
        elif fill is DUP_FILL:
            highlight_counts["duplicate_address"] += 1
        elif fill is BOTH_FILL:
            highlight_counts["both"] += 1
        if fill is not None:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = fill

    if cfg.get("template_family") == "consignment_v1" and "summary_row" in cfg:
        sr = cfg["summary_row"]
        shipment_col = _find_col_by_kind(col_map, lambda s: s[0] == "shipment_flag")
        qty_col = _find_col_by_kind(col_map, lambda s: s == ("field", "quantity"))
        if shipment_col:
            letter = ws.cell(row=sr, column=shipment_col).column_letter
            ws.cell(row=sr, column=shipment_col).value = f"=SUM({letter}{data_start}:{letter}1048576)"
        if qty_col:
            letter = ws.cell(row=sr, column=qty_col).column_letter
            ws.cell(row=sr, column=qty_col).value = f"=SUM({letter}{data_start}:{letter}1048576)"

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path, highlight_counts


_REVIEW_HEADERS = [
    ("_source_file", "원본 파일"),
    ("_source_row", "원본 행"),
    ("order_id", "주문번호"),
    ("order_date", "주문일자"),
    ("product_name", "상품명"),
    ("option", "옵션"),
    ("quantity", "수량"),
    ("receiver_name", "수령인명"),
    ("receiver_phone", "수령인 연락처"),
    ("address", "배송지 주소"),
    ("zipcode", "우편번호"),
    ("shop_name", "쇼핑몰명"),
]


def write_review_file(unclassified, ambiguous, out_path):
    """협력사 매칭이 안 되었거나(미분류) 여러 협력사에 동시에 걸린(중복매칭) 건을
    수동 확인용 엑셀로 저장한다."""
    wb = openpyxl.Workbook()
    header_font = Font(name="맑은 고딕", size=10, bold=True)

    def write_sheet(ws, title, records, note_field=None):
        ws.title = title
        headers = [h for _, h in _REVIEW_HEADERS] + (["매칭 후보"] if note_field else [])
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = header_font
        for r, rec in enumerate(records, start=2):
            for c, (field, _) in enumerate(_REVIEW_HEADERS, start=1):
                ws.cell(row=r, column=c, value=rec.get(field))
            if note_field:
                ws.cell(row=r, column=len(_REVIEW_HEADERS) + 1, value=", ".join(rec.get(note_field, [])))

    ws1 = wb.active
    write_sheet(ws1, "미분류", unclassified)
    ws2 = wb.create_sheet("중복매칭")
    write_sheet(ws2, "중복매칭", ambiguous, note_field="_ambiguous_matches")

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path
