"""주문 라인을 상품명/옵션의 브랜드 키워드로 협력사에 매칭한다."""


def classify_vendor(rec, vendors_cfg):
    """rec 하나에 대해 매칭되는 협력사 이름 리스트를 반환한다.

    0개면 미분류, 2개 이상이면 키워드가 겹치는 것이므로 자동 배정하지 않고
    수동 확인 대상으로 넘긴다.
    """
    text = f"{rec.get('product_name', '')} {rec.get('option', '')}"
    matches = []
    for vendor, cfg in vendors_cfg.items():
        for kw in cfg.get("brand_keywords", []):
            if kw and kw in text:
                matches.append(vendor)
                break
    return matches
