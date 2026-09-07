"""원본 주문서 대조: 누락 건 확인, 특이사항(배송메시지) 및 정보 누락 건 추출."""


def check_no_omission(all_rows, by_vendor, unclassified, ambiguous):
    """정규화된 전체 행 수와 (협력사별 + 미분류 + 중복매칭) 합계가 같은지 확인한다.
    normalize.read_market_file은 원본의 내용이 있는 모든 행을 rows로 반환하므로,
    이 둘이 다르면 분류 단계에서 어딘가 행이 새어나간 것이다."""
    processed = sum(len(rows) for rows in by_vendor.values()) + len(unclassified) + len(ambiguous)
    return {
        "raw_count": len(all_rows),
        "processed_count": processed,
        "match": len(all_rows) == processed,
    }


def collect_special_notes(all_rows):
    """배송메시지(특이사항)가 있는 주문 라인을 모은다."""
    notes = []
    for rec in all_rows:
        msg = (rec.get("delivery_message") or "").strip()
        if not msg:
            continue
        notes.append({
            "order_id": rec.get("order_id"),
            "product_name": rec.get("product_name"),
            "receiver_name": rec.get("receiver_name"),
            "message": msg,
        })
    return notes


_REQUIRED_FOR_DELIVERY = [
    ("receiver_name", "수령인명"),
    ("receiver_phone", "연락처"),
    ("address", "주소"),
]


def collect_missing_info(all_rows):
    """배송에 꼭 필요한 정보(수령인명/연락처/주소)가 빠진 주문 라인을 모은다."""
    issues = []
    for rec in all_rows:
        missing = [label for field, label in _REQUIRED_FOR_DELIVERY if not str(rec.get(field) or "").strip()]
        if missing:
            issues.append({
                "order_id": rec.get("order_id"),
                "product_name": rec.get("product_name"),
                "receiver_name": rec.get("receiver_name"),
                "missing_fields": missing,
                "_source_file": rec.get("_source_file"),
                "_source_row": rec.get("_source_row"),
            })
    return issues
