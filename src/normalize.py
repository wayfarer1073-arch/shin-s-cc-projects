"""오픈마켓 원본 주문서(xlsx)를 표준 필드 스키마로 정규화한다.

마켓마다 헤더 이름이 다르므로 config/column_aliases.json에 등록된 이름으로
헤더 행을 자동으로 찾고, 각 열을 표준 필드에 매핑한다. 새 마켓 양식을 받으면
코드를 고칠 필요 없이 그 json에 헤더 이름만 추가하면 된다.
"""
import json
from pathlib import Path
from datetime import datetime, date

import openpyxl

from src.filename_meta import parse_filename

BASE_DIR = Path(__file__).resolve().parent.parent
ALIASES_PATH = BASE_DIR / "config" / "column_aliases.json"

with open(ALIASES_PATH, encoding="utf-8") as f:
    _aliases_raw = json.load(f)

CANONICAL_FIELDS = [k for k in _aliases_raw if not k.startswith("_")]

HEADER_TO_FIELD = {}
for _field, _names in _aliases_raw.items():
    if _field.startswith("_"):
        continue
    for _name in _names:
        HEADER_TO_FIELD[_name.strip()] = _field

_DATE_FORMATS = ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y년 %m월 %d일")


def _find_header_row(ws, max_scan_rows=6, min_fields=3):
    best_row, best_score = None, 0
    for r in range(1, min(max_scan_rows, ws.max_row) + 1):
        fields_found = set()
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() in HEADER_TO_FIELD:
                fields_found.add(HEADER_TO_FIELD[v.strip()])
        if len(fields_found) > best_score:
            best_row, best_score = r, len(fields_found)
    if best_row is None or best_score < min_fields:
        raise ValueError(
            "헤더 행을 찾지 못했습니다. config/column_aliases.json에 이 마켓의 "
            "헤더 이름을 추가해야 할 수 있습니다."
        )
    return best_row


def _normalize_phone(v):
    if v in (None, ""):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _normalize_amount(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", "").replace("원", "")
    try:
        return float(s) if "." in s else int(s)
    except ValueError:
        return None


def _normalize_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def read_market_file(path):
    """오픈마켓 원본 주문서 1개를 읽어 표준 필드 dict의 리스트로 반환한다.

    파일명이 "{YYMMDD} {브랜드} {매출처} 주문서.xlsx" 규칙을 따르면 그 일자/브랜드/
    매출처를 우선 사용하고, 안 맞으면(과거 방식 파일) 파일 내용 기반 추론으로
    넘어간다 — brand는 파일명에서만 얻을 수 있으므로 규칙에 안 맞으면 None이 된다.
    """
    filename_meta = parse_filename(Path(path).name)

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header_row = _find_header_row(ws)

    col_field = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip() in HEADER_TO_FIELD:
            col_field[c] = HEADER_TO_FIELD[v.strip()]

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        # 매핑되지 않은 컬럼에만 값이 있는 행도 "실제 주문 라인"으로 잡히도록,
        # 행 전체(매핑 여부 무관)를 기준으로 빈 행을 판단한다. 매핑된 필드만 보면
        # 원본에는 있는 라인이 조용히 누락될 수 있다.
        any_value = any(
            ws.cell(row=r, column=c).value not in (None, "")
            for c in range(1, ws.max_column + 1)
        )
        if not any_value:
            continue
        raw = {field: ws.cell(row=r, column=c).value for c, field in col_field.items()}

        rec = {field: raw.get(field) for field in CANONICAL_FIELDS}
        rec["receiver_phone"] = _normalize_phone(rec.get("receiver_phone"))
        rec["order_date"] = _normalize_date(rec.get("order_date")) or filename_meta["order_date"]
        rec["product_name"] = (rec.get("product_name") or "").strip()
        rec["option"] = (rec.get("option") or "").strip()
        rec["payment_amount"] = _normalize_amount(rec.get("payment_amount"))
        if filename_meta["market"]:
            rec["shop_name"] = filename_meta["market"]
        rec["brand"] = filename_meta["brand"]
        try:
            rec["quantity"] = int(rec["quantity"]) if rec.get("quantity") not in (None, "") else 1
        except (TypeError, ValueError):
            rec["quantity"] = 1
        rec["_source_file"] = Path(path).name
        rec["_source_row"] = r
        rows.append(rec)
    return rows
