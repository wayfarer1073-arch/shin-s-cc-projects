"""보고용 요약 통계 계산: 오픈마켓별/협력사별 주문 건수, top5 판매 제품 등."""
from collections import Counter, defaultdict


def _distinct_order_counts(rows_by_group):
    """group -> rows 목록에서 그룹별 '고유 주문번호 개수'를 센다.
    주문번호가 없는 라인은 라인 자체를 1건으로 취급한다."""
    counts = {}
    for group, rows in rows_by_group.items():
        order_ids = set()
        no_id_count = 0
        for r in rows:
            oid = r.get("order_id")
            if oid:
                order_ids.add(oid)
            else:
                no_id_count += 1
        counts[group] = len(order_ids) + no_id_count
    return counts


def compute_summary(all_rows, by_vendor, unclassified, ambiguous, run_date):
    by_market = defaultdict(list)
    for rec in all_rows:
        market = rec.get("shop_name") or rec.get("_source_file") or "(알수없음)"
        by_market[market].append(rec)

    market_order_counts = _distinct_order_counts(by_market)
    vendor_order_counts = _distinct_order_counts(by_vendor)

    product_qty = Counter()
    for rec in all_rows:
        key = rec.get("product_name") or "(상품명 없음)"
        product_qty[key] += rec.get("quantity") or 0
    top5_products = [{"product_name": k, "quantity": v} for k, v in product_qty.most_common(5)]

    all_order_ids = {r.get("order_id") for r in all_rows if r.get("order_id")}
    no_id_rows = sum(1 for r in all_rows if not r.get("order_id"))

    return {
        "run_date": run_date,
        "total_order_count": len(all_order_ids) + no_id_rows,
        "total_line_count": len(all_rows),
        "market_order_counts": market_order_counts,
        "vendor_order_counts": vendor_order_counts,
        "unclassified_count": len(unclassified),
        "ambiguous_count": len(ambiguous),
        "top5_products": top5_products,
    }
