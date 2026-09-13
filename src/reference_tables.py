"""사람이 엑셀로 관리하는 참고 자료(재고 현황, 정식 옵션명 목록, 상품코드
참고표, 이름쌍 참고표)를 새 엑셀 파일로 교체하는 기능.

각 참고 자료는 "종류(kind)"가 있고, 종류마다 엑셀에서 어떤 열을 읽어
data/reference/*.json으로 저장할지가 정해져 있다. 업로드 즉시 덮어쓰지
않고, 먼저 미리보기(파싱 결과 일부 + 건수)를 보여주고 확인을 받은 뒤에만
실제 파일을 바꾼다 — 엑셀 형식이 예상과 달라 잘못 파싱됐을 때 그대로
반영되는 사고를 막기 위함. 교체 직전에는 항상 기존 파일을 백업 폴더에
복사해둔다(되돌릴 수 있도록).
"""
import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).resolve().parent.parent
REFERENCE_DIR = BASE_DIR / "data" / "reference"
BACKUP_DIR = REFERENCE_DIR / "_backups"


def _first_matching_col(ws, header_row, candidates):
    """헤더 행에서 candidates(후보 헤더 이름들) 중 하나와 정확히 같은 열을
    찾는다. 못 찾으면 None."""
    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=header_row, column=c).value
        if isinstance(h, str) and h.strip() in candidates:
            return c
    return None


def _parse_stock(path):
    """이플코리아 재고 현황표: "상품명"/"재고합계" 헤더가 있는 열을 찾아
    {"name":.., "stock_total":..} 목록으로 만든다."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header_row = 1
    name_col = _first_matching_col(ws, header_row, {"상품명"})
    stock_col = _first_matching_col(ws, header_row, {"재고합계"})
    if name_col is None or stock_col is None:
        raise ValueError('엑셀 1행에서 "상품명"과 "재고합계" 열을 찾지 못했습니다. 헤더 이름을 확인해주세요.')
    out = []
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=name_col).value
        if name is None or str(name).strip() == "":
            continue
        total = ws.cell(row=r, column=stock_col).value
        out.append({"name": str(name).strip(), "stock_total": total if isinstance(total, (int, float)) else None})
    return out


def _parse_option_list(path):
    """정식 옵션명 목록: 첫 번째 열(헤더 이름이 있으면 "상품명"/"옵션명"/
    "정식옵션명" 열을 우선)을 그대로 문자열 목록으로 만든다."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header_row = 1
    col = _first_matching_col(ws, header_row, {"상품명", "옵션명", "정식옵션명", "정식 옵션명"})
    if col is None:
        col = 1
    out = []
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=col).value
        if v is None or str(v).strip() == "":
            continue
        out.append(str(v).strip())
    return out


def _parse_code_table(path):
    """상품코드 참고표: "상품명"류 열 + "코드"류 열을 찾아 {"name":..,
    "code":..} 목록으로 만든다. 헤더로 못 찾으면 1·2번째 열을 순서대로 쓴다."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header_row = 1
    name_col = _first_matching_col(ws, header_row, {"상품명", "이름", "품목명"})
    code_col = _first_matching_col(ws, header_row, {"코드", "관리코드", "상품코드", "상품고유코드", "품목코드"})
    if name_col is None:
        name_col = 1
    if code_col is None:
        code_col = 2
    out = []
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=name_col).value
        if name is None or str(name).strip() == "":
            continue
        code = ws.cell(row=r, column=code_col).value
        out.append({"name": str(name).strip(), "code": (str(code).strip() if code is not None else "")})
    return out


def _parse_name_pair(path):
    """이엑스 상품/상품명 쌍 같은 두 열짜리 참고표: 1·2번째 열을 그대로
    {"a":.., "b":..} 목록으로 만든다."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header_row = 1
    out = []
    for r in range(header_row + 1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if a is None or str(a).strip() == "":
            continue
        out.append({"a": str(a).strip(), "b": (str(b).strip() if b is not None else "")})
    return out


_PARSERS = {
    "stock": _parse_stock,
    "option_list": _parse_option_list,
    "code_table": _parse_code_table,
    "name_pair": _parse_name_pair,
}

# 웹 화면에서 교체할 수 있는 참고 자료 목록. id는 URL에 쓰이므로 영문/숫자만.
TABLES = [
    {
        "id": "stock_이플코리아",
        "label": "이플코리아 재고 현황",
        "file": "stock_이플코리아.json",
        "kind": "stock",
        "help": '엑셀 1행에 "상품명"과 "재고합계" 열이 있어야 합니다(3PL에서 받는 재고 현황표 그대로).',
    },
    {
        "id": "options_이플코리아",
        "label": "이플코리아 정식 옵션명 목록",
        "file": "options_이플코리아.json",
        "kind": "option_list",
        "help": "정식으로 인정하는 옵션명을 한 줄에 하나씩 적은 목록입니다(첫 번째 열 사용).",
    },
    {
        "id": "options_투데이넛",
        "label": "투데이넛 정식 옵션명 목록",
        "file": "options_투데이넛.json",
        "kind": "option_list",
        "help": "정식으로 인정하는 옵션명을 한 줄에 하나씩 적은 목록입니다(첫 번째 열 사용).",
    },
    {
        "id": "options_굿즈코리아",
        "label": "굿즈코리아 정식 옵션명 목록",
        "file": "options_굿즈코리아.json",
        "kind": "option_list",
        "help": "정식으로 인정하는 옵션명을 한 줄에 하나씩 적은 목록입니다(첫 번째 열 사용).",
    },
    {
        "id": "codes_서울문화사",
        "label": "서울문화사 관리코드 참고표",
        "file": "codes_서울문화사.json",
        "kind": "code_table",
        "help": '엑셀에 "상품명"/"코드"(또는 비슷한 이름) 열이 있어야 합니다. 못 찾으면 1·2번째 열을 씁니다.',
    },
    {
        "id": "codes_이뮨",
        "label": "이뮨 상품고유코드 참고표",
        "file": "codes_이뮨.json",
        "kind": "code_table",
        "help": '엑셀에 "상품명"/"코드"(또는 비슷한 이름) 열이 있어야 합니다. 못 찾으면 1·2번째 열을 씁니다.',
    },
    {
        "id": "codes_듀오랩",
        "label": "듀오랩 코드 참고표",
        "file": "codes_듀오랩.json",
        "kind": "code_table",
        "help": '엑셀에 "상품명"/"코드"(또는 비슷한 이름) 열이 있어야 합니다. 못 찾으면 1·2번째 열을 씁니다.',
    },
    {
        "id": "codes_제로스킨",
        "label": "제로스킨 품목코드 참고표",
        "file": "codes_제로스킨.json",
        "kind": "code_table",
        "help": '엑셀에 "상품명"/"코드"(또는 비슷한 이름) 열이 있어야 합니다. 못 찾으면 1·2번째 열을 씁니다.',
    },
    {
        "id": "names_이엑스",
        "label": "이엑스 상품/상품명 내부 표기 참고표",
        "file": "names_이엑스.json",
        "kind": "name_pair",
        "help": "1번째 열=상품(내부 짧은 표기), 2번째 열=상품명(용량/구성)으로 읽습니다.",
    },
]

_TABLES_BY_ID = {t["id"]: t for t in TABLES}


def get_table(table_id):
    return _TABLES_BY_ID.get(table_id)


def table_json_path(table):
    return REFERENCE_DIR / table["file"]


def load_current_count(table):
    path = table_json_path(table)
    if not path.exists():
        return 0, None
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    mtime = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    return len(data), mtime


def parse_upload(table, excel_path):
    """table["kind"]에 맞는 파서로 엑셀을 읽어 새 내용(리스트)을 반환한다.
    형식이 안 맞으면 ValueError를 낸다(원인 메시지 포함)."""
    parser = _PARSERS[table["kind"]]
    data = parser(excel_path)
    if not data:
        raise ValueError("엑셀에서 읽은 내용이 없습니다. 파일이 비어있거나 형식이 다를 수 있습니다.")
    return data


def build_export_workbook(table):
    """table의 현재 내용(data/reference/*.json)을 업로드용 엑셀과 똑같은
    양식(헤더/열 순서)으로 담은 워크북을 만든다. 이 파일을 내려받아 고친
    뒤 그대로 다시 올리면 parse_upload가 문제없이 읽는다."""
    path = table_json_path(table)
    data = []
    if path.exists():
        with open(path, encoding="utf-8") as f:
            data = json.load(f)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = re.sub(r'[\\/*?:\[\]]', "_", table["label"])[:31] or "참고자료"

    kind = table["kind"]
    if kind == "stock":
        ws.append(["상품명", "재고합계"])
        for row in data:
            ws.append([row.get("name", ""), row.get("stock_total")])
    elif kind == "option_list":
        ws.append(["정식옵션명"])
        for name in data:
            ws.append([name])
    elif kind == "code_table":
        ws.append(["상품명", "코드"])
        for row in data:
            ws.append([row.get("name", ""), row.get("code", "")])
    elif kind == "name_pair":
        ws.append(["상품", "상품명"])
        for row in data:
            ws.append([row.get("a", ""), row.get("b", "")])
    else:
        raise ValueError(f"알 수 없는 참고 자료 종류: {kind}")

    for col_cells in ws.columns:
        width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(width + 2, 10), 60)

    return wb


def apply_replacement(table, new_data):
    """기존 파일을 백업한 뒤 새 내용으로 교체한다. 백업 파일 경로를 반환."""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    path = table_json_path(table)
    backup_path = None
    if path.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = BACKUP_DIR / f"{table['file'].removesuffix('.json')}_{stamp}.json"
        shutil.copy2(path, backup_path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(new_data, f, ensure_ascii=False, indent=2)
    return backup_path
