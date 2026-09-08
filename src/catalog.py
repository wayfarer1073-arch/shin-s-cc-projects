"""브랜드별 상품 카탈로그(협력사/원가 참고용)를 불러오고 조회한다.

- JM: (주문상품명, 옵션구성) 정확히 일치하는 항목을 찾아 협력사/원가/관리상품명을 반환한다.
  과거 주문 이력에서 뽑은 표라 커버리지가 완벽하지 않다 — 없으면 None을 반환하고
  classify.py는 기존 브랜드 키워드 매칭으로 넘어간다.
- MSNA: 하위 브랜드명이 상품명에 포함되는지로 찾는다 (하위 브랜드 -> 배송처 매핑은
  '브랜드별 배송비 정책' 시트 기준이라 일부 하위 브랜드는 배송처가 아직 없음).
"""
import json
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

with open(BASE_DIR / "config" / "brands.json", encoding="utf-8") as f:
    _BRANDS_CFG = json.load(f)["brands"]

_loaded = {}


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
    """반환: {"vendor":..., "managed_name":..., "unit_cost":...} 또는 매치 없으면 None."""
    if not brand or not product_name:
        return None
    kind, catalog = _get_catalog(brand) or (None, None)
    if kind == "jm_exact":
        entry = catalog["by_name_option"].get((product_name, option or ""))
        if entry:
            return {"vendor": entry["vendor"], "managed_name": entry["managed_name"], "unit_cost": entry["unit_cost"]}
        vendors = catalog["by_name_vendors"].get(product_name)
        if vendors and len(vendors) == 1:
            return {"vendor": next(iter(vendors)), "managed_name": None, "unit_cost": None}
        return None
    if kind == "msna_subbrand":
        for sb in catalog["sub_brands"]:
            if sb in product_name:
                vendor = catalog["vendor_by_sub_brand"].get(sb)
                if vendor:
                    return {"vendor": vendor, "managed_name": None, "unit_cost": None, "sub_brand": sb}
                return None
        return None
    return None
