#!/usr/bin/env python3
"""오픈마켓 원본 주문서 여러 개를 읽어 협력사별로 분류·취합하고,
각 협력사 양식(templates/*.xlsx)에 맞춰 엑셀 파일을 생성한다.

사용법:
    python3 main.py 주문서1.xlsx 주문서2.xlsx ... [--out output]
"""
import argparse
from datetime import date
from pathlib import Path

from src.normalize import read_market_file
from src.classify import classify_vendor
from src.writer import VENDORS, write_vendor_file, write_review_file


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
    for vendor, rows in by_vendor.items():
        out_path = out_dir / f"{vendor}_{run_date}.xlsx"
        write_vendor_file(vendor, rows, out_path)
        written.append(out_path)
        print(f"[작성] {vendor}: {len(rows)}건 -> {out_path}")

    review_path = None
    if unclassified or ambiguous:
        review_path = out_dir / f"확인필요_{run_date}.xlsx"
        write_review_file(unclassified, ambiguous, review_path)
        print(f"[확인 필요] 미분류 {len(unclassified)}건, 중복매칭 {len(ambiguous)}건 -> {review_path}")

    return written, review_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", help="오픈마켓 원본 주문서 xlsx 경로 (여러 개 가능)")
    parser.add_argument("--out", default="output", help="출력 폴더 (기본: output)")
    args = parser.parse_args()
    run(args.inputs, args.out)


if __name__ == "__main__":
    main()
