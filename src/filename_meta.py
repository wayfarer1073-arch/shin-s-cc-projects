"""업로드 파일명에서 일자/브랜드/매출처(마켓)를 추출한다.

규칙: "{YYMMDD} {브랜드} {매출처} 주문서.xlsx" (예: "260908 JM 쿠팡 주문서.xlsx").
패턴에 맞지 않으면 조용히 None을 반환하고, 호출 쪽에서 파일 내용 기반 추론으로
넘어가면 된다 (이 규칙이 적용되기 전의 파일들과의 호환을 위함).
"""
import json
import re
from datetime import date
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

with open(BASE_DIR / "config" / "brands.json", encoding="utf-8") as f:
    _KNOWN_BRANDS = set(json.load(f)["brands"].keys())

_DATE_RE = re.compile(r"^(\d{2})(\d{2})(\d{2})$")


def _parse_yymmdd(token):
    m = _DATE_RE.match(token)
    if not m:
        return None
    yy, mm, dd = (int(x) for x in m.groups())
    year = 2000 + yy if yy < 80 else 1900 + yy
    try:
        return date(year, mm, dd)
    except ValueError:
        return None


def parse_filename(filename):
    """반환: {"order_date": date|None, "brand": str|None, "market": str|None} — 매치 안 되면 전부 None."""
    stem = Path(filename).stem.strip()
    tokens = stem.split()
    if len(tokens) < 3:
        return {"order_date": None, "brand": None, "market": None}

    order_date = _parse_yymmdd(tokens[0])
    if order_date is None:
        return {"order_date": None, "brand": None, "market": None}

    brand = tokens[1].upper() if tokens[1].upper() in _KNOWN_BRANDS else None
    if brand is None:
        return {"order_date": order_date, "brand": None, "market": None}

    rest = tokens[2:]
    if rest and rest[-1] in ("주문서", "주문서.xlsx"):
        rest = rest[:-1]
    market = " ".join(rest).strip() or None

    return {"order_date": order_date, "brand": brand, "market": market}
