"""브랜드별 상품 카탈로그(협력사/원가 참고용)를 불러오고 조회한다.

- JM: (주문상품명, 옵션구성) 정확히 일치하는 항목을 찾아 협력사/원가/관리상품명을 반환한다.
  과거 주문 이력에서 뽑은 표라 커버리지가 완벽하지 않다 — 없으면 None을 반환하고
  classify.py는 기존 브랜드 키워드 매칭으로 넘어간다.
- MSNA: 하위 브랜드명이 상품명에 포함되는지로 찾는다 (하위 브랜드 -> 배송처 매핑은
  '브랜드별 배송비 정책' 시트 기준이라 일부 하위 브랜드는 배송처가 아직 없음).
- 이플코리아 전역 우선 배정: "주문서 가공 지침" 4번("이플코리아_MH에서 확인할 수 있는
  상품들은 브랜드 구분 없이 모두 이플코리아 상품으로 분류")에 따라, 브랜드와 무관하게
  data/reference/options_이플코리아.json(이플코리아 3PL 실보관 재고 + 정식 옵션명
  목록)에 있는 상품은 카탈로그의 과거 배송처 값과 무관하게 무조건 이플코리아로
  분류한다. 브랜드별 카탈로그 조회보다 먼저 확인한다. 단, data/reference/stock_
  이플코리아.json(사용자가 제공한 3PL 재고 스냅샷)에서 재고합계가 0인 상품은 이
  전역 배정에서 제외한다(사용자 확인, 2026-09-11: 재고 0인 상품은 주문이 들어와도
  이플코리아가 아니라 다른 협력사 지침을 참조해야 함) — 그 경우 브랜드별 카탈로그
  조회(과거 배송처 이력)로 넘어가고, 거기서도 못 찾으면 미분류로 남아 수동 확인
  대상이 된다.
"""
import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

with open(BASE_DIR / "config" / "brands.json", encoding="utf-8") as f:
    _BRANDS_CFG = json.load(f)["brands"]

_loaded = {}

_MATCH_MIN_LEN = 8  # 정규화한 이름이 이보다 짧으면 오탐 위험이 커서 매칭에서 제외
_EP_KOREA_VENDOR = "이플코리아"
_EP_KOREA_OPTIONS_FILE = "data/reference/options_이플코리아.json"
_EP_KOREA_STOCK_FILE = "data/reference/stock_이플코리아.json"
_ep_korea_names = None


def _normalize_name(s):
    s = str(s or "")
    s = re.sub(r'^\(.*?\)', '', s)
    s = re.sub(r'^(New|NEW|1\+1|"1\+1"|"4개 구성")\s*', '', s)
    s = re.sub(r'[\s"\'/,.\-_★]+', '', s)
    return s.strip()


def _load_ep_korea_zero_stock_names():
    stock_path = BASE_DIR / _EP_KOREA_STOCK_FILE
    if not stock_path.exists():
        return set()
    with open(stock_path, encoding="utf-8") as f:
        items = json.load(f)
    return {_normalize_name(it["name"]) for it in items if it.get("stock_total") == 0}


def _load_ep_korea_names():
    global _ep_korea_names
    if _ep_korea_names is None:
        with open(BASE_DIR / _EP_KOREA_OPTIONS_FILE, encoding="utf-8") as f:
            items = json.load(f)
        zero_stock = _load_ep_korea_zero_stock_names()
        names = [_normalize_name(it) for it in items]
        _ep_korea_names = [
            n for n in names
            if len(n) >= _MATCH_MIN_LEN and n not in zero_stock
        ]
    return _ep_korea_names


def _ep_korea_override_vendor(product_name):
    if not product_name:
        return None
    norm_product = _normalize_name(product_name)
    for name in _load_ep_korea_names():
        if name in norm_product:
            return _EP_KOREA_VENDOR
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


def _load_catalog_of_type(catalog_type, cfg):
    if catalog_type == "jm_exact":
        return _load_jm(cfg)
    if catalog_type == "msna_subbrand":
        return _load_msna(cfg)
    return None


def _get_catalog(brand):
    if brand not in _BRANDS_CFG:
        return None
    if brand not in _loaded:
        cfg = _BRANDS_CFG[brand]
        catalog = _load_catalog_of_type(cfg["catalog_type"], cfg)
        fallback = None
        if cfg.get("fallback_catalog_type"):
            fallback_cfg = {"catalog_file": cfg["fallback_catalog_file"]}
            fallback = (cfg["fallback_catalog_type"], _load_catalog_of_type(cfg["fallback_catalog_type"], fallback_cfg))
        _loaded[brand] = (cfg["catalog_type"] if catalog else None, catalog, fallback)
    return _loaded[brand]


def _lookup_in(kind, catalog, product_name, option):
    """(kind, catalog) 하나에 대해서만 조회한다. jm_exact는 정확 매치 실패 시
    같은 주문상품명에 협력사가 하나뿐이면 그걸로 대체한다."""
    if not catalog:
        return None
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


def lookup_vendor(brand, product_name, option):
    """반환: {"vendor":..., "managed_name":..., "unit_cost":...} 또는 매치 없으면 None.

    재고 우선 배정은 원본 주문 상품명(마켓 리스팅 제목이라 수식어가 많이 붙음)뿐
    아니라, 카탈로그 매칭에 성공했을 때의 관리상품명(더 깔끔한 이름이라 매칭이
    잘 됨)에 대해서도 확인한다 — 둘 중 하나라도 재고 목록과 매칭되면 그 3PL로
    무조건 배정한다. 1순위 카탈로그에서 못 찾으면(fallback_catalog_type이 설정된
    브랜드는) 2순위 카탈로그로 넘어간다(예: MSNA 정확 매치표 -> 하위 브랜드 매칭).
    """
    if not product_name:
        return None

    # 이플코리아 전역 우선 배정은 브랜드 구분 없이 가장 먼저 확인한다.
    direct_override = _ep_korea_override_vendor(product_name)
    if direct_override:
        return {"vendor": direct_override, "managed_name": None, "unit_cost": None}

    if not brand:
        return None

    kind, catalog, fallback = _get_catalog(brand) or (None, None, None)

    entry = _lookup_in(kind, catalog, product_name, option)
    if not entry and fallback:
        fb_kind, fb_catalog = fallback
        entry = _lookup_in(fb_kind, fb_catalog, product_name, option)

    override_vendor = _ep_korea_override_vendor(product_name)
    if not override_vendor and entry and entry.get("managed_name"):
        override_vendor = _ep_korea_override_vendor(entry["managed_name"])
    if override_vendor:
        return {
            "vendor": override_vendor,
            "managed_name": entry["managed_name"] if entry else None,
            "unit_cost": entry["unit_cost"] if entry else None,
        }

    return entry
