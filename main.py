#!/usr/bin/env python3
"""오픈마켓 원본 주문서 여러 개를 읽어 협력사별로 분류·취합하고,
각 협력사 양식(templates/*.xlsx)에 맞춰 엑셀 파일을 생성한다.

사용법:
    python3 main.py 주문서1.xlsx 주문서2.xlsx ... [--out output]
"""
import argparse
import json
from datetime import date, datetime
from pathlib import Path

from src.normalize import read_market_file
from src.classify import classify_vendor
from src.writer import VENDORS, write_vendor_file, write_review_file
from src.summary import compute_summary
from src.dashboard import render_dashboard_html
from src.reconcile import check_no_omission, collect_special_notes, collect_missing_info
from src.archive import archive_orders, load_orders


def _parse_iso_date(s):
    return datetime.strptime(s, "%Y-%m-%d").date() if isinstance(s, str) else s


def run(input_paths, out_dir, brand_override=None):
    """input_paths를 읽어 협력사별로 분류하고 양식을 작성한다. brand_override를
    주면(원본 파일명이 "{YYMMDD} {사업부} {매출처} 주문서.xlsx" 규칙을 안 따라
    사업부를 자동으로 못 얻는 파일 등) 읽은 모든 줄의 사업부를 그 값으로
    강제 지정한다.

    오늘 이미 다른 파일을 처리해서 보관본이 있는 경우, 방금 읽은 내용만으로
    양식을 다시 쓰면 먼저 처리한 내용이 사라진다. 그래서 보관(archive_orders,
    같은 날짜 보관본에 이어붙임) 이후 그날 보관된 전체 내용을 다시 불러와
    양식 작성·요약·대시보드는 항상 "오늘 지금까지 처리한 전체"를 기준으로
    한다."""
    all_rows = []
    for path in input_paths:
        rows = read_market_file(path)
        if brand_override:
            for rec in rows:
                rec["brand"] = brand_override
        all_rows.extend(rows)
        print(f"[읽음] {Path(path).name}: {len(rows)}건" + (f" (사업부={brand_override}로 수동 지정)" if brand_override else ""))

    # 이번 처리분만(리뷰 파일의 "매칭 후보" 표시용 — 보관본을 거치면 후보
    # 목록은 안 남으므로 이번 회차 결과를 따로 보관해둔다).
    by_vendor = {}
    this_run_unclassified = []
    this_run_ambiguous = []
    for rec in all_rows:
        matches = classify_vendor(rec, VENDORS)
        if len(matches) == 1:
            by_vendor.setdefault(matches[0], []).append(rec)
        elif len(matches) == 0:
            this_run_unclassified.append(rec)
        else:
            rec["_ambiguous_matches"] = matches
            this_run_ambiguous.append(rec)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    run_date = date.today().isoformat()
    run_date_compact = run_date.replace("-", "")

    archive_path = archive_orders(by_vendor, this_run_unclassified, this_run_ambiguous, run_date)
    print(f"[보관] 이번 처리 {len(all_rows)}건 -> {archive_path}")

    # 오늘 보관된 전체(이번 처리분 포함)를 다시 불러와 양식 작성·요약·대조의
    # 기준으로 삼는다 — 오늘 여러 번 나눠 올려도 매번 "오늘 전체"가 반영됨.
    day_rows = load_orders(run_date, run_date, archive_dir=archive_path.parent)
    all_rows = []
    by_vendor = {}
    unclassified = []
    ambiguous = []
    for r in day_rows:
        r = dict(r)
        r["order_date"] = _parse_iso_date(r.get("order_date"))
        all_rows.append(r)
        if r["match_status"] == "classified":
            by_vendor.setdefault(r["vendor"], []).append(r)
        elif r["match_status"] == "unclassified":
            unclassified.append(r)
        else:
            ambiguous.append(r)

    written = []
    highlight_totals = {"quantity": 0, "duplicate_address": 0, "both": 0, "name_pair_unmatched": 0}
    for vendor, rows in by_vendor.items():
        # 브랜드별로 나눠서 각각 별도 파일로 작성한다 — 파일명 규칙(가공 지침 8번)이
        # "{YYYYMMDD}_{협력사}_{사업부(MH/JM/MSNA)}"라 한 협력사가 여러 브랜드
        # 주문을 같이 처리하는 경우(예: 이플코리아) 브랜드마다 파일을 분리해야 한다.
        # 브랜드 표기가 없는 주문(과거 파일명 규칙 이전의 JM 전용 협력사)은 JM으로 간주.
        by_brand = {}
        for rec in rows:
            by_brand.setdefault(rec.get("brand") or "JM", []).append(rec)

        for brand, brand_rows in by_brand.items():
            file_brand = VENDORS[vendor].get("file_brand_override") or brand
            out_path = out_dir / f"{run_date_compact}_{vendor}_{file_brand}.xlsx"
            _, highlight_counts = write_vendor_file(vendor, brand_rows, out_path)
            written.append(out_path)
            for k, v in highlight_counts.items():
                highlight_totals[k] += v
            marks = []
            if highlight_counts["quantity"]:
                marks.append(f"수량다수 {highlight_counts['quantity']}건")
            if highlight_counts["duplicate_address"]:
                marks.append(f"동일수령인·주소 {highlight_counts['duplicate_address']}건")
            if highlight_counts["both"]:
                marks.append(f"둘다 해당 {highlight_counts['both']}건")
            if highlight_counts["name_pair_unmatched"]:
                marks.append(f"참고표 미매칭(원본유지) {highlight_counts['name_pair_unmatched']}건")
            mark_note = f" ({', '.join(marks)} 강조표시)" if marks else ""
            print(f"[작성] {vendor} ({brand}): {len(brand_rows)}건 -> {out_path}{mark_note}")

    review_path = None
    if unclassified or ambiguous:
        review_path = out_dir / f"확인필요_{run_date}.xlsx"
        write_review_file(unclassified, ambiguous, review_path)
        print(f"[확인 필요] 미분류 {len(unclassified)}건, 중복매칭 {len(ambiguous)}건 -> {review_path}")

    # --- 원본 대조: 누락 건 및 특이사항 조사 ---
    reconciliation = check_no_omission(all_rows, by_vendor, unclassified, ambiguous)
    if reconciliation["match"]:
        print(f"[대조] 원본 {reconciliation['raw_count']}건 = 처리 {reconciliation['processed_count']}건 (누락 없음)")
    else:
        print(
            f"[대조] ⚠ 원본 {reconciliation['raw_count']}건 vs 처리 {reconciliation['processed_count']}건 "
            "불일치 — 코드 확인 필요"
        )

    special_notes = collect_special_notes(all_rows)
    missing_info = collect_missing_info(all_rows)
    if special_notes:
        print(f"[특이사항] 배송메시지 있는 주문 {len(special_notes)}건")
    if missing_info:
        print(f"[정보누락] 수령인명/연락처/주소 중 빠진 주문 {len(missing_info)}건")

    summary = compute_summary(
        all_rows, by_vendor, unclassified, ambiguous, run_date,
        reconciliation=reconciliation,
        special_notes=special_notes,
        missing_info=missing_info,
        highlight_totals=highlight_totals,
    )
    summary_path = out_dir / f"summary_{run_date}.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2, default=str)
    print(f"[요약] 전체 주문 {summary['total_order_count']}건 -> {summary_path}")

    dashboard_path = out_dir / f"dashboard_{run_date}.html"
    dashboard_path.write_text(render_dashboard_html(summary), encoding="utf-8")
    print(f"[대시보드] -> {dashboard_path}")

    return written, review_path, summary_path, dashboard_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", help="오픈마켓 원본 주문서 xlsx 경로 (여러 개 가능)")
    parser.add_argument("--out", default="output", help="출력 폴더 (기본: output)")
    args = parser.parse_args()
    run(args.inputs, args.out)


if __name__ == "__main__":
    main()
