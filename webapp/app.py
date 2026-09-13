"""오픈마켓 주문서 처리 웹 대시보드.

지금까지 main.py로 터미널에서 돌리던 파이프라인(읽기->분류->양식작성->보관
->요약/대시보드)을 웹 화면으로 감싼 것. 로직은 전부 src/ 모듈을 그대로
쓰고, 이 파일은 업로드 받기/결과 보여주기/파일 내려받기만 담당한다.

실행:
    uvicorn webapp.app:app --host 0.0.0.0 --port 8000 --reload
"""
import json
import tempfile
import uuid
from datetime import date
from pathlib import Path

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import main as pipeline
from src.archive import ARCHIVE_DIR
from src import reference_tables as ref

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
STAGING_DIR = BASE_DIR / "data" / "reference" / "_staging"
STAGING_DIR.mkdir(parents=True, exist_ok=True)

with open(BASE_DIR / "config" / "brands.json", encoding="utf-8") as f:
    BRAND_LABELS = list(json.load(f)["brands"].keys())

app = FastAPI(title="오픈마켓 주문서 처리")
app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))


def _list_run_dates():
    """지금까지 보관된 날짜 목록(최신순)을 반환한다."""
    if not ARCHIVE_DIR.exists():
        return []
    dates = sorted((p.stem for p in ARCHIVE_DIR.glob("*.json")), reverse=True)
    return dates


def _run_summary(run_date):
    path = OUTPUT_DIR / f"summary_{run_date}.json"
    if not path.exists():
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _run_files(run_date):
    """그 날짜에 생성된 협력사 엑셀 파일 목록(파일명, 협력사, 사업부)을 반환한다."""
    compact = run_date.replace("-", "")
    files = []
    for p in sorted(OUTPUT_DIR.glob(f"{compact}_*.xlsx")):
        stem = p.stem[len(compact) + 1:]
        if "_" in stem:
            vendor, brand = stem.rsplit("_", 1)
        else:
            vendor, brand = stem, ""
        files.append({"filename": p.name, "vendor": vendor, "brand": brand})
    return files


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    run_dates = _list_run_dates()
    recent = [{"run_date": d, "summary": _run_summary(d)} for d in run_dates[:14]]
    return templates.TemplateResponse(
        request,
        "index.html",
        {"brands": BRAND_LABELS, "recent": recent, "today": date.today().isoformat()},
    )


@app.post("/process")
async def process(request: Request, files: list[UploadFile] = File(...), brand: str = Form("")):
    with tempfile.TemporaryDirectory() as tmp:
        saved_paths = []
        for uf in files:
            if not uf.filename:
                continue
            dest = Path(tmp) / uf.filename
            content = await uf.read()
            dest.write_bytes(content)
            saved_paths.append(str(dest))

        if not saved_paths:
            return RedirectResponse(url="/", status_code=303)

        brand_override = brand or None
        written, review_path, summary_path, dashboard_path = pipeline.run(
            saved_paths, OUTPUT_DIR, brand_override=brand_override
        )

    run_date = date.today().isoformat()
    return RedirectResponse(url=f"/runs/{run_date}", status_code=303)


@app.get("/runs/{run_date}", response_class=HTMLResponse)
def run_detail(request: Request, run_date: str):
    summary = _run_summary(run_date)
    files = _run_files(run_date)
    review_exists = (OUTPUT_DIR / f"확인필요_{run_date}.xlsx").exists()
    dashboard_exists = (OUTPUT_DIR / f"dashboard_{run_date}.html").exists()
    return templates.TemplateResponse(
        request,
        "run_detail.html",
        {
            "run_date": run_date,
            "summary": summary,
            "files": files,
            "review_exists": review_exists,
            "dashboard_exists": dashboard_exists,
        },
    )


@app.get("/runs", response_class=HTMLResponse)
def run_list(request: Request):
    run_dates = _list_run_dates()
    recent = [{"run_date": d, "summary": _run_summary(d)} for d in run_dates]
    return templates.TemplateResponse(request, "run_list.html", {"recent": recent})


@app.get("/download/{filename}")
def download(filename: str):
    safe_name = Path(filename).name  # 경로 조작 방지 (디렉터리 이동 문자 제거)
    path = OUTPUT_DIR / safe_name
    if not path.exists():
        return HTMLResponse("파일을 찾을 수 없습니다.", status_code=404)
    return FileResponse(path, filename=safe_name)


# --- 참고 자료(재고 현황/정식 옵션명/상품코드/이름쌍) 관리 -----------------

@app.get("/reference", response_class=HTMLResponse)
def reference_list(request: Request):
    tables = []
    for t in ref.TABLES:
        count, updated = ref.load_current_count(t)
        tables.append({**t, "count": count, "updated": updated})
    return templates.TemplateResponse(request, "reference_list.html", {"tables": tables})


@app.post("/reference/{table_id}/preview", response_class=HTMLResponse)
async def reference_preview(request: Request, table_id: str, file: UploadFile = File(...)):
    table = ref.get_table(table_id)
    if not table:
        return HTMLResponse("알 수 없는 참고 자료입니다.", status_code=404)

    with tempfile.TemporaryDirectory() as tmp:
        dest = Path(tmp) / (file.filename or "upload.xlsx")
        dest.write_bytes(await file.read())
        try:
            new_data = ref.parse_upload(table, dest)
        except Exception as e:
            return templates.TemplateResponse(
                request, "reference_error.html", {"table": table, "error": str(e)}
            )

    token = uuid.uuid4().hex
    staging_path = STAGING_DIR / f"{token}.json"
    with open(staging_path, "w", encoding="utf-8") as f:
        json.dump(new_data, f, ensure_ascii=False, indent=2)

    old_count, _ = ref.load_current_count(table)
    return templates.TemplateResponse(
        request,
        "reference_preview.html",
        {
            "table": table,
            "token": token,
            "new_count": len(new_data),
            "old_count": old_count,
            "preview_rows": new_data[:20],
            "kind": table["kind"],
        },
    )


@app.post("/reference/{table_id}/confirm")
def reference_confirm(table_id: str, token: str = Form(...)):
    table = ref.get_table(table_id)
    if not table:
        return HTMLResponse("알 수 없는 참고 자료입니다.", status_code=404)

    staging_path = STAGING_DIR / f"{token}.json"
    if not staging_path.exists():
        return HTMLResponse("업로드 내용을 찾을 수 없습니다(시간이 지나 만료되었을 수 있습니다). 다시 업로드해주세요.", status_code=400)

    with open(staging_path, encoding="utf-8") as f:
        new_data = json.load(f)
    ref.apply_replacement(table, new_data)
    staging_path.unlink(missing_ok=True)

    return RedirectResponse(url="/reference", status_code=303)


@app.post("/reference/{table_id}/cancel")
def reference_cancel(table_id: str, token: str = Form(...)):
    staging_path = STAGING_DIR / f"{token}.json"
    staging_path.unlink(missing_ok=True)
    return RedirectResponse(url="/reference", status_code=303)
