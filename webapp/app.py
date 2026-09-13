"""오픈마켓 주문서 처리 웹 대시보드.

지금까지 main.py로 터미널에서 돌리던 파이프라인(읽기->분류->양식작성->보관
->요약/대시보드)을 웹 화면으로 감싼 것. 로직은 전부 src/ 모듈을 그대로
쓰고, 이 파일은 업로드 받기/결과 보여주기/파일 내려받기만 담당한다.

실행:
    uvicorn webapp.app:app --host 0.0.0.0 --port 8000 --reload
"""
import io
import json
import tempfile
import uuid
from datetime import date
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

import main as pipeline
from src.archive import ARCHIVE_DIR
from src import paths
from src import reference_tables as ref
from webapp import auth
from webapp import board

BASE_DIR = Path(__file__).resolve().parent.parent

# 영구 디스크가 붙어 있으면(APP_DATA_DIR 환경변수) 처음 한 번 저장소에
# 커밋되어 있던 참고자료/이력을 디스크로 복사해 넣는다. 로컬 개발 중에는
# 아무 일도 하지 않는다(seed_if_empty 안에서 자체적으로 걸러짐).
paths.seed_if_empty()

OUTPUT_DIR = paths.OUTPUT_DIR
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
STAGING_DIR = paths.REFERENCE_DIR / "_staging"
STAGING_DIR.mkdir(parents=True, exist_ok=True)

with open(BASE_DIR / "config" / "brands.json", encoding="utf-8") as f:
    BRAND_LABELS = list(json.load(f)["brands"].keys())

app = FastAPI(title="오픈마켓 주문서 처리")


# 주의: @app.middleware("http")로 등록하는 미들웨어는 "나중에 등록할수록
# 더 바깥쪽"이 되어 먼저 실행된다. 그래서 이 로그인 검사 미들웨어를 먼저
# 등록하고, SessionMiddleware를 그 다음에 등록해야 SessionMiddleware가
# 더 바깥쪽에서 먼저 request.session을 만들어준다. 순서를 바꾸면
# "SessionMiddleware must be installed" 에러가 난다.
@app.middleware("http")
async def require_login(request: Request, call_next):
    # 파일 다운로드(/files, /download)도 업무 자료라 로그인해야만 접근 가능.
    # 계정이 하나도 없으면(최초 설치) 어떤 경로든 /setup 으로 보낸다.
    path = request.url.path
    if not auth.is_public_path(path):
        if not auth.users_exist():
            return RedirectResponse(url="/setup", status_code=303)
        if not request.session.get("username"):
            return RedirectResponse(url="/login", status_code=303)
    return await call_next(request)


app.add_middleware(SessionMiddleware, secret_key=auth.get_session_secret(), same_site="lax")


app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")
templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))


def render(request: Request, name: str, context: dict | None = None, status_code: int = 200):
    context = dict(context or {})
    username = request.session.get("username")
    current_user = auth.find_user(username) if username else None
    context["current_user"] = current_user
    # 사이드바의 "최근 게시글" 미리보기는 로그인한 모든 페이지에 나오므로
    # 여기서 한 번에 채워준다(매 라우트마다 따로 넣을 필요 없게).
    if current_user:
        context["board_latest"] = board.latest_by_tag()
    return templates.TemplateResponse(request, name, context, status_code=status_code)


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
    return render(
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
    return render(
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
def run_list(request: Request, start: str = "", end: str = ""):
    run_dates = _list_run_dates()
    # run_date는 "YYYY-MM-DD" 형식이라 문자열 비교만으로 날짜 범위 필터가 된다.
    if start:
        run_dates = [d for d in run_dates if d >= start]
    if end:
        run_dates = [d for d in run_dates if d <= end]
    recent = [{"run_date": d, "summary": _run_summary(d)} for d in run_dates]
    return render(request, "run_list.html", {"recent": recent, "start": start, "end": end})


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
    return render(request, "reference_list.html", {"tables": tables})


@app.get("/reference/{table_id}/download")
def reference_download(table_id: str):
    table = ref.get_table(table_id)
    if not table:
        return HTMLResponse("알 수 없는 참고 자료입니다.", status_code=404)

    wb = ref.build_export_workbook(table)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = table["file"].removesuffix(".json") + ".xlsx"
    # 파일명이 한글이라 그대로 헤더에 넣으면 깨진다(HTTP 헤더는 라틴-1만
    # 허용) - RFC 5987 형식(filename*)으로 인코딩하고, 구형 브라우저를 위해
    # 아스키 대체 이름도 같이 준다.
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=\"reference.xlsx\"; filename*=UTF-8''{quote(filename)}"},
    )


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
            return render(
                request, "reference_error.html", {"table": table, "error": str(e)}
            )

    token = uuid.uuid4().hex
    staging_path = STAGING_DIR / f"{token}.json"
    with open(staging_path, "w", encoding="utf-8") as f:
        json.dump(new_data, f, ensure_ascii=False, indent=2)

    old_count, _ = ref.load_current_count(table)
    return render(
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


# --- 로그인 / 계정 관리 ------------------------------------------------

@app.get("/setup", response_class=HTMLResponse)
def setup_form(request: Request):
    if auth.users_exist():
        # 이미 계정이 있으면 최초 설치 화면을 다시 쓸 수 없다.
        return RedirectResponse(url="/login", status_code=303)
    return templates.TemplateResponse(request, "setup.html", {"error": None})


@app.post("/setup")
def setup_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    display_name: str = Form(""),
):
    if auth.users_exist():
        return RedirectResponse(url="/login", status_code=303)
    try:
        auth.create_user(username, password, display_name=display_name, is_admin=True)
    except ValueError as e:
        return templates.TemplateResponse(
            request, "setup.html", {"error": str(e)}, status_code=400
        )
    request.session["username"] = username.strip()
    return RedirectResponse(url="/", status_code=303)


@app.get("/login", response_class=HTMLResponse)
def login_form(request: Request, next: str = "/"):
    if not auth.users_exist():
        return RedirectResponse(url="/setup", status_code=303)
    return templates.TemplateResponse(request, "login.html", {"error": None, "next": next})


@app.post("/login")
def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    next: str = Form("/"),
):
    user = auth.verify_login(username.strip(), password)
    if not user:
        return templates.TemplateResponse(
            request,
            "login.html",
            {"error": "아이디 또는 비밀번호가 올바르지 않습니다.", "next": next},
            status_code=401,
        )
    request.session["username"] = user["username"]
    return RedirectResponse(url=next or "/", status_code=303)


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse(url="/login", status_code=303)


def _require_admin(request: Request):
    """관리자가 아니면 None을 반환(호출 쪽에서 403 처리)."""
    username = request.session.get("username")
    user = auth.find_user(username) if username else None
    if not user or not user.get("is_admin"):
        return None
    return user


@app.get("/users", response_class=HTMLResponse)
def users_list(request: Request):
    if not _require_admin(request):
        return HTMLResponse("관리자만 접근할 수 있습니다.", status_code=403)
    return render(request, "users.html", {"users": auth.list_users(), "error": None})


@app.post("/users/create")
def users_create(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    display_name: str = Form(""),
    is_admin: str = Form(""),
):
    if not _require_admin(request):
        return HTMLResponse("관리자만 접근할 수 있습니다.", status_code=403)
    try:
        auth.create_user(username, password, display_name=display_name, is_admin=bool(is_admin))
    except ValueError as e:
        return render(request, "users.html", {"users": auth.list_users(), "error": str(e)}, status_code=400)
    return RedirectResponse(url="/users", status_code=303)


@app.post("/users/{target_username}/delete")
def users_delete(request: Request, target_username: str):
    if not _require_admin(request):
        return HTMLResponse("관리자만 접근할 수 있습니다.", status_code=403)
    try:
        auth.delete_user(target_username)
    except ValueError as e:
        return render(request, "users.html", {"users": auth.list_users(), "error": str(e)}, status_code=400)
    return RedirectResponse(url="/users", status_code=303)


@app.post("/users/{target_username}/reset-password")
def users_reset_password(request: Request, target_username: str, new_password: str = Form(...)):
    if not _require_admin(request):
        return HTMLResponse("관리자만 접근할 수 있습니다.", status_code=403)
    try:
        auth.set_password(target_username, new_password)
    except ValueError as e:
        return render(request, "users.html", {"users": auth.list_users(), "error": str(e)}, status_code=400)
    return RedirectResponse(url="/users", status_code=303)


# --- 사내 게시판 (이슈/공지/잡담) ------------------------------------------

@app.get("/board", response_class=HTMLResponse)
def board_list(request: Request):
    return render(request, "board.html", {"posts": board.list_posts(), "tags": board.TAGS, "error": None})


@app.post("/board/create")
def board_create(
    request: Request,
    tag: str = Form(...),
    title: str = Form(...),
    body: str = Form(...),
):
    username = request.session.get("username")
    user = auth.find_user(username)
    try:
        board.create_post(tag, title, body, author=user["display_name"])
    except ValueError as e:
        return render(
            request, "board.html",
            {"posts": board.list_posts(), "tags": board.TAGS, "error": str(e)},
            status_code=400,
        )
    return RedirectResponse(url="/board", status_code=303)
