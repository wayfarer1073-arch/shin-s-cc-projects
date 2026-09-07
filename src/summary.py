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


def compute_summary(
    all_rows, by_vendor, unclassified, ambiguous, run_date,
    reconciliation=None, special_notes=None, missing_info=None, highlight_totals=None,
):
    by_market = defaultdict(list)
    for rec in all_rows:
        market = rec.get("shop_name") or rec.get("_source_file") or "(알수없음)"
        by_market[market].append(rec)

    market_order_counts = _distinct_order_counts(by_market)
    vendor_order_counts = _distinct_order_counts(by_vendor)

    product_qty = Counter()
    product_revenue = Counter()
    for rec in all_rows:
        key = rec.get("product_name") or "(상품명 없음)"
        product_qty[key] += rec.get("quantity") or 0
        if rec.get("payment_amount") is not None:
            product_revenue[key] += rec["payment_amount"]

    all_order_ids = {r.get("order_id") for r in all_rows if r.get("order_id")}
    no_id_rows = sum(1 for r in all_rows if not r.get("order_id"))

    amounts = [r.get("payment_amount") for r in all_rows if r.get("payment_amount") is not None]
    total_revenue = sum(amounts)
    revenue_missing_count = len(all_rows) - len(amounts)

    top5_products = []
    for name, qty in product_qty.most_common(5):
        revenue = product_revenue.get(name, 0)
        share = round(revenue / total_revenue * 100, 1) if total_revenue else None
        top5_products.append({
            "product_name": name,
            "quantity": qty,
            "revenue": revenue,
            "revenue_share_pct": share,
        })

    # 전체 상품별 수량/매출 (top5 넘어서도 보관 -> 여러 날짜를 합쳐서 다시
    # 랭킹을 매길 때 정확한 결과가 나오도록. 일별 대시보드 DB 문서에 저장됨.
    product_breakdown = {
        name: {"quantity": qty, "revenue": product_revenue.get(name, 0)}
        for name, qty in product_qty.items()
    }

    return {
        "run_date": run_date,
        "total_order_count": len(all_order_ids) + no_id_rows,
        "total_line_count": len(all_rows),
        "market_order_counts": market_order_counts,
        "vendor_order_counts": vendor_order_counts,
        "unclassified_count": len(unclassified),
        "ambiguous_count": len(ambiguous),
        "top5_products": top5_products,
        "product_breakdown": product_breakdown,
        "total_revenue": total_revenue,
        "revenue_missing_count": revenue_missing_count,
        "reconciliation": reconciliation or {},
        "special_notes": special_notes or [],
        "missing_info": missing_info or [],
        "highlight_totals": highlight_totals or {"quantity": 0, "duplicate_address": 0, "both": 0},
    }
