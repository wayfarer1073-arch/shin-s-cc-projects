"""정규화된 주문 라인을 협력사별 엑셀 양식(templates/*.xlsx)에 채워 넣는다."""
import json
import re
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

_NO_HEADERS = {"NO", "No.", "No", "no", "번호", "순번"}


def _build_col_map(ws, cfg):
    """열 번호 -> ("static", 값) | ("combine",) | ("seq",) | ("year"/"month"/"day",)
    | ("shipment_flag",) | ("field", 표준필드명) | ("code",) | ("name_pair", "a"|"b")
    로 매핑한다. 매핑되지 않은 열은 자동으로 채우지 않고 비워둔다(택배사/송장번호
    등 나중에 수기로 채우는 열)."""
    header_row = cfg["header_row"]
    field_overrides = cfg.get("field_overrides", {})
    static_fields = cfg.get("static_fields", {})
    combine_field = cfg.get("combine_product_option_field")
    code_column = cfg.get("code_column")
    name_pair_columns = cfg.get("name_pair_columns", {})
    option_canon_file = cfg.get("option_canon_file")

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
        elif code_column and header == code_column:
            col_map[c] = ("code",)
        elif header in name_pair_columns:
            col_map[c] = ("name_pair", name_pair_columns[header])
        elif option_canon_file and (header in field_overrides and field_overrides[header] == "option"
                                     or header in _HEADER_TO_FIELD and _HEADER_TO_FIELD[header] == "option"):
            col_map[c] = ("canon_option",)
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


# ---------------------------------------------------------------------------
# 상품코드 조회: 협력사가 준 상품명/옵션명 <-> 자체 코드 매핑표(data/reference/
# codes_{협력사}.json)에서, 정규화한 이름이 주문 상품명+옵션 텍스트에 부분
# 일치하면 그 코드를 채운다. 매핑표 자체가 다른 상품군을 담고 있거나 커버리지가
# 낮을 수 있어 매치가 없으면 그냥 비워둔다(수기로 채우게).
# ---------------------------------------------------------------------------
_CODE_MATCH_MIN_LEN = 4
_code_tables = {}


def _normalize_for_code(s):
    s = str(s or "")
    s = re.sub(r'^[\(\[].*?[\)\]]', '', s)
    s = re.sub(r'[\s"\'/,.\-_★\[\]]+', '', s)
    return s.strip()


def _load_code_table(vendor_name, cfg):
    if vendor_name not in _code_tables:
        code_file = cfg.get("code_lookup_file")
        if not code_file:
            _code_tables[vendor_name] = None
        else:
            with open(BASE_DIR / code_file, encoding="utf-8") as f:
                entries = json.load(f)
            items = [(_normalize_for_code(e["name"]), e["code"]) for e in entries]
            items = [(n, c) for n, c in items if len(n) >= _CODE_MATCH_MIN_LEN]
            items.sort(key=lambda x: len(x[0]), reverse=True)
            _code_tables[vendor_name] = items
    return _code_tables[vendor_name]


def _code_for(vendor_name, cfg, rec):
    table = _load_code_table(vendor_name, cfg)
    if not table:
        return None
    text = _normalize_for_code(f"{rec.get('product_name') or ''} {rec.get('option') or ''}")
    if not text:
        return None
    for name, code in table:
        if name in text:
            return code
    return None


# ---------------------------------------------------------------------------
# 상품명/옵션명 쌍 조회: 협력사 자체 상품 목록(예: 이엑스 "상품"/"상품명" 두 컬럼
# 조합)을 data/reference/names_{협력사}.json({"a":.., "b":..} 쌍 목록)으로 갖고
# 있는 경우, 주문 상품명+옵션과 매칭되면 그 협력사 고유 표기(a, b)를 그대로 쓴다.
#
# 마켓 원본 상품명은 "카스텔크렘 포지타노 레몬 캔디 750g"처럼 협력사 내부
# 표기("레몬1 750g")와 단어 순서·구성이 달라 통짜 문자열 부분일치로는 못 찾는
# 경우가 많다. 대신 "a" 쪽에서 의미 있는 단어(한글 2글자 이상, 용량 표기 NNNg)를
# 전부 뽑아 그 단어들이 (순서 상관없이) 주문 텍스트에 전부 나타나는지로 판단한다
# — 예: "레몬1 750g" -> 필수 토큰 {"레몬","750g"} 둘 다 주문 텍스트에 있으면 매치.
# 후보가 여러 개면 필수 토큰이 더 많고(더 구체적인) 쪽을 우선한다.
# 매치가 없으면 우리 쪽 상품명을 그대로 양쪽 컬럼에 채운다(공란보다는 낫다).
# ---------------------------------------------------------------------------
_name_pair_tables = {}
_TOKEN_RE = re.compile(r'\d+[가-힣a-zA-Z]*|[가-힣]{2,}')


def _tokenize(s):
    """의미 있는 단어(한글 2글자 이상)와 단위 붙은 수량 표기(15포/750g/20티백
    등)를 뽑는다. 단위 없는 맨숫자(그냥 "1")는 어디에나 있을 수 있어 제외한다."""
    return {t for t in _TOKEN_RE.findall(s) if not t.isdigit()}


def _best_token_match(text, candidates):
    """candidates: [(tokens, payload), ...]. 필수 토큰이 전부 text에 있는 후보 중
    가장 구체적인(토큰 글자수 합이 큰) 것의 payload를 반환. 없으면 None.
    candidates는 미리 구체적인 순서로 정렬돼 있어야 한다. 공백 차이로 매칭이
    실패하지 않도록 text에서 공백을 제거하고 비교한다(토큰 쪽은 애초에 정규식이
    공백을 넘어가지 않아 공백이 섞이지 않는다)."""
    text = re.sub(r'\s+', '', text)
    for tokens, payload in candidates:
        if all(tok in text for tok in tokens):
            return payload
    return None


def _load_name_pair_table(vendor_name, cfg):
    if vendor_name not in _name_pair_tables:
        pair_file = cfg.get("name_pair_lookup_file")
        if not pair_file:
            _name_pair_tables[vendor_name] = None
        else:
            with open(BASE_DIR / pair_file, encoding="utf-8") as f:
                entries = json.load(f)
            items = []
            for e in entries:
                tokens = _tokenize(e["a"])
                if tokens:
                    items.append((tokens, (e["a"], e["b"])))
            items.sort(key=lambda x: sum(len(t) for t in x[0]), reverse=True)
            _name_pair_tables[vendor_name] = items
    return _name_pair_tables[vendor_name]


def _name_pair_match(vendor_name, cfg, rec):
    table = _load_name_pair_table(vendor_name, cfg)
    if not table:
        return None
    text = f"{rec.get('product_name') or ''} {rec.get('option') or ''}"
    if not text.strip():
        return None
    payload = _best_token_match(text, table)
    return {"a": payload[0], "b": payload[1]} if payload else None


# ---------------------------------------------------------------------------
# 옵션명 정규화: 협력사가 인정하는 옵션명 목록(예: "이플코리아 옵션명" 참고
# 시트, data/reference/options_{협력사}.json 문자열 리스트)이 있는 경우, 우리
# 쪽 옵션 텍스트가 두루뭉술해도 상품명+옵션 전체에서 그 목록의 항목과 토큰이
# 다 겹치면 그 정식 옵션명으로 바꿔 쓴다. 매치가 없으면 원래 옵션 텍스트를
# 그대로 둔다.
# ---------------------------------------------------------------------------
_option_canon_tables = {}


def _load_option_canon_table(vendor_name, cfg):
    if vendor_name not in _option_canon_tables:
        opt_file = cfg.get("option_canon_file")
        if not opt_file:
            _option_canon_tables[vendor_name] = None
        else:
            with open(BASE_DIR / opt_file, encoding="utf-8") as f:
                entries = json.load(f)
            items = [(_tokenize(e), e) for e in entries]
            items = [(t, e) for t, e in items if t]
            items.sort(key=lambda x: sum(len(t) for t in x[0]), reverse=True)
            _option_canon_tables[vendor_name] = items
    return _option_canon_tables[vendor_name]


def _canon_option_for(vendor_name, cfg, rec):
    table = _load_option_canon_table(vendor_name, cfg)
    if not table:
        return rec.get("option") or None
    text = f"{rec.get('product_name') or ''} {rec.get('option') or ''}"
    if not text.strip():
        return rec.get("option") or None
    match = _best_token_match(text, table)
    return match if match else (rec.get("option") or None)


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


def _cell_value_for(kind_spec, rec, seq, is_first_of_order, vendor_name=None, cfg=None):
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
    if kind == "code":
        return _code_for(vendor_name, cfg, rec)
    if kind == "name_pair":
        match = _name_pair_match(vendor_name, cfg, rec)
        if match:
            return match[kind_spec[1]]
        return rec.get("product_name") or None
    if kind == "canon_option":
        return _canon_option_for(vendor_name, cfg, rec)
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
            value = _cell_value_for(spec, rec, i + 1, is_first, vendor_name, cfg)
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
