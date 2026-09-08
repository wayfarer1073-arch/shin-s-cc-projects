"""주문 라인을 협력사에 매칭한다.

1순위: 브랜드별 상품 카탈로그(src/catalog.py)에서 (상품명, 옵션) 정확 매치 —
   과거 주문 이력으로 만든 표라 브랜드 키워드보다 신뢰도가 높다.
2순위: 상품명/옵션에 등록된 협력사 브랜드 키워드가 포함되는지 (config/vendors.json).
"""
from src.catalog import lookup_vendor


def classify_vendor(rec, vendors_cfg):
    """rec 하나에 대해 매칭되는 협력사 이름 리스트를 반환한다.

    0개면 미분류, 2개 이상이면 키워드가 겹치는 것이므로 자동 배정하지 않고
    수동 확인 대상으로 넘긴다. 카탈로그에서 협력사를 찾았지만 그 협력사의
    양식(config/vendors.json)이 아직 등록되어 있지 않으면 rec에
    "_catalog_vendor_unregistered"를 표시해두고 미분류로 취급한다.
    """
    brand = rec.get("brand")
    if brand:
        hit = lookup_vendor(brand, rec.get("product_name", ""), rec.get("option", ""))
        if hit:
            vendor = hit["vendor"]
            if vendor in vendors_cfg:
                return [vendor]
            rec["_catalog_vendor_unregistered"] = vendor
            return []

    text = f"{rec.get('product_name', '')} {rec.get('option', '')}"
    matches = []
    for vendor, cfg in vendors_cfg.items():
        for kw in cfg.get("brand_keywords", []):
            if kw and kw in text:
                matches.append(vendor)
                break
    return matches
