#!/usr/bin/env python3
"""오픈마켓 원본 주문서 여러 개를 읽어 협력사별로 분류·취합하고,
각 협력사 양식(templates/*.xlsx)에 맞춰 엑셀 파일을 생성한다.

사용법:
    python3 main.py 주문서1.xlsx 주문서2.xlsx ... [--out output]
"""
import argparse
import json
from datetime import date
from pathlib import Path

from src.normalize import read_market_file
from src.classify import classify_vendor
from src.writer import VENDORS, write_vendor_file, write_review_file
from src.summary import compute_summary
from src.dashboard import render_dashboard_html
from src.reconcile import check_no_omission, collect_special_notes, collect_missing_info
from src.archive import archive_orders


def run(input_paths, out_dir):
    all_rows = []
    for path in input_paths:
        rows = read_market_file(path)
        all_rows.extend(rows)
        print(f"[읽음] {Path(path).name}: {len(rows)}건")

    by_vendor = {}
    unclassified = []
    ambiguous = []
    for rec in all_rows:
        matches = classify_vendor(rec, VENDORS)
        if len(matches) == 1:
            by_vendor.setdefault(matches[0], []).append(rec)
        elif len(matches) == 0:
            unclassified.append(rec)
        else:
            rec["_ambiguous_matches"] = matches
            ambiguous.append(rec)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    run_date = date.today().isoformat()

    written = []
    highlight_totals = {"quantity": 0, "duplicate_address": 0, "both": 0}
    for vendor, rows in by_vendor.items():
        out_path = out_dir / f"{vendor}_{run_date}.xlsx"
        _, highlight_counts = write_vendor_file(vendor, rows, out_path)
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
        mark_note = f" ({', '.join(marks)} 강조표시)" if marks else ""
        print(f"[작성] {vendor}: {len(rows)}건 -> {out_path}{mark_note}")

    review_path = None
    if unclassified or ambiguous:
        review_path = out_dir / f"확인필요_{run_date}.xlsx"
        write_review_file(unclassified, ambiguous, review_path)
        print(f"[확인 필요] 미분류 {len(unclassified)}건, 중복매칭 {len(ambiguous)}건 -> {review_path}")

    archive_path = archive_orders(by_vendor, unclassified, ambiguous, run_date)
    print(f"[보관] 주문 원본 라인 {len(all_rows)}건 -> {archive_path}")

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
