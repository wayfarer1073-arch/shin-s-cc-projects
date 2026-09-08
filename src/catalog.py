"""브랜드별 상품 카탈로그(협력사/원가 참고용)를 불러오고 조회한다.

- JM: (주문상품명, 옵션구성) 정확히 일치하는 항목을 찾아 협력사/원가/관리상품명을 반환한다.
  과거 주문 이력에서 뽑은 표라 커버리지가 완벽하지 않다 — 없으면 None을 반환하고
  classify.py는 기존 브랜드 키워드 매칭으로 넘어간다.
- MSNA: 하위 브랜드명이 상품명에 포함되는지로 찾는다 (하위 브랜드 -> 배송처 매핑은
  '브랜드별 배송비 정책' 시트 기준이라 일부 하위 브랜드는 배송처가 아직 없음).
- 재고 기반 우선 배정(stock_override): 브랜드에 stock_override_file이 설정되어 있으면,
  카탈로그 조회보다 먼저 확인한다. 3PL 업체가 실제 보관 중인 재고 목록에 해당 상품이
  있으면 카탈로그의 과거 배송처 값과 무관하게 그 3PL 업체로 무조건 분류한다
  (사용자가 명시적으로 지정한 우선순위 규칙).
"""
import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

with open(BASE_DIR / "config" / "brands.json", encoding="utf-8") as f:
    _BRANDS_CFG = json.load(f)["brands"]

_loaded = {}
_stock_loaded = {}

_STOCK_MATCH_MIN_LEN = 8  # 정규화한 이름이 이보다 짧으면 오탐 위험이 커서 매칭에서 제외


def _normalize_name(s):
    s = str(s or "")
    s = re.sub(r'^\(.*?\)', '', s)
    s = re.sub(r'^(New|NEW|1\+1|"1\+1"|"4개 구성")\s*', '', s)
    s = re.sub(r'[\s"\'/,.\-_★]+', '', s)
    return s.strip()


def _load_stock_override(brand):
    cfg = _BRANDS_CFG.get(brand, {})
    stock_file = cfg.get("stock_override_file")
    if not stock_file:
        return None
    if brand not in _stock_loaded:
        with open(BASE_DIR / stock_file, encoding="utf-8") as f:
            items = json.load(f)
        names = [_normalize_name(it["product_name"]) for it in items]
        names = [n for n in names if len(n) >= _STOCK_MATCH_MIN_LEN]
        _stock_loaded[brand] = {"names": names, "vendor": cfg["stock_override_vendor"]}
    return _stock_loaded[brand]


def _stock_override_vendor(brand, product_name):
    stock = _load_stock_override(brand)
    if not stock or not product_name:
        return None
    norm_product = _normalize_name(product_name)
    for stock_name in stock["names"]:
        if stock_name in norm_product:
            return stock["vendor"]
    return None


def _load_jm(cfg):
    with open(BASE_DIR / cfg["catalog_file"], encoding="utf-8") as f:
        entries = json.load(f)
    by_name_option = {}
    by_name_vendors = {}
    for e in entries:
        key = (e["order_product_name"], e["option"])
        by_name_option[key] = e
        by_name_vendors.setdefault(e["order_product_name"], set()).add(e["vendor"])
    return {"by_name_option": by_name_option, "by_name_vendors": by_name_vendors, "raw": entries}


def _load_msna(cfg):
    with open(BASE_DIR / cfg["catalog_file"], encoding="utf-8") as f:
        data = json.load(f)
    # 이름이 긴 하위 브랜드부터 검사해야 짧은 이름이 다른 브랜드명의 접두어인 경우를 피할 수 있다.
    sub_brands = sorted(
        {p["sub_brand"] for p in data["products"]},
        key=len, reverse=True,
    )
    vendor_by_sub_brand = {sb: pol.get("ship_vendor") for sb, pol in data["sub_brand_policy"].items()}
    return {"sub_brands": sub_brands, "vendor_by_sub_brand": vendor_by_sub_brand}


def _get_catalog(brand):
    if brand not in _BRANDS_CFG:
        return None
    if brand not in _loaded:
        cfg = _BRANDS_CFG[brand]
        if cfg["catalog_type"] == "jm_exact":
            _loaded[brand] = ("jm_exact", _load_jm(cfg))
        elif cfg["catalog_type"] == "msna_subbrand":
            _loaded[brand] = ("msna_subbrand", _load_msna(cfg))
        else:
            _loaded[brand] = (None, None)
    return _loaded[brand]


def lookup_vendor(brand, product_name, option):
    """반환: {"vendor":..., "managed_name":..., "unit_cost":...} 또는 매치 없으면 None.

    재고 우선 배정은 원본 주문 상품명(마켓 리스팅 제목이라 수식어가 많이 붙음)뿐
    아니라, 카탈로그 매칭에 성공했을 때의 관리상품명(더 깔끔한 이름이라 매칭이
    잘 됨)에 대해서도 확인한다 — 둘 중 하나라도 재고 목록과 매칭되면 그 3PL로
    무조건 배정한다.
    """
    if not brand or not product_name:
        return None

    kind, catalog = _get_catalog(brand) or (None, None)

    entry = None
    if kind == "jm_exact":
        entry = catalog["by_name_option"].get((product_name, option or ""))
        if not entry:
            vendors = catalog["by_name_vendors"].get(product_name)
            if vendors and len(vendors) == 1:
                entry = {"vendor": next(iter(vendors)), "managed_name": None, "unit_cost": None}

    stock_vendor = _stock_override_vendor(brand, product_name)
    if not stock_vendor and entry and entry.get("managed_name"):
        stock_vendor = _stock_override_vendor(brand, entry["managed_name"])
    if stock_vendor:
        return {
            "vendor": stock_vendor,
            "managed_name": entry["managed_name"] if entry else None,
            "unit_cost": entry["unit_cost"] if entry else None,
        }

    if kind == "jm_exact":
        if not entry:
            return None
        return {"vendor": entry["vendor"], "managed_name": entry["managed_name"], "unit_cost": entry["unit_cost"]}
    if kind == "msna_subbrand":
        for sb in catalog["sub_brands"]:
            if sb in product_name:
                vendor = catalog["vendor_by_sub_brand"].get(sb)
                if vendor:
                    return {"vendor": vendor, "managed_name": None, "unit_cost": None, "sub_brand": sb}
                return None
        return None
    return None
