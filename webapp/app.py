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
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

import main as pipeline
from src.archive import load_orders
from src import paths
from src import reference_tables as ref
from src import runs as runs_store
from src.summary import compute_summary
from src.writer import VENDORS, write_vendor_file
from webapp import auth
from webapp import board

BASE_DIR = Path(__file__).resolve().parent.parent

# 영구 디스크가 붙어 있으면(APP_DATA_DIR 환경변수) 처음 한 번 저장소에
# 커밋되어 있던 참고자료/이력을 디스크로 복사해 넣는다. 로컬 개발 중에는
# 아무 일도 하지 않는다(seed_if_empty 안에서 자체적으로 걸러짐).
paths.seed_if_empty()

# 이 기능(업로드한 사람별로 실행을 구분)을 만들기 전부터 있던 이력에는
# _run_id가 없다. 그런 옛날 라인들에 날짜별로 "레거시 실행"을 한 번만
# 붙여줘서 화면에서 완전히 사라지지 않게 한다(관리자에게만 보임).
runs_store.backfill_legacy_runs()

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
        # 관리자 화면에서 "누가 올렸는지"는 아이디(로그인용)가 아니라
        # 이름으로 보여준다 - 처리 이력/결과 상세에서 쓴다.
        if current_user["is_admin"]:
            context["user_display_names"] = {u["username"]: u["display_name"] for u in auth.list_users()}
    return templates.TemplateResponse(request, name, context, status_code=status_code)


def _current_user(request: Request):
    username = request.session.get("username")
    return auth.find_user(username) if username else None


def _run_files(run_date):
    """그 날짜에 생성된 협력사 엑셀 파일 목록(파일명, 협력사, 사업부)을 반환한다.
    이 파일들은 그날 여러 사람이 올린 내용이 전부 합쳐진 것이라 관리자만 본다."""
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


def _can_view_run(user, run):
    return bool(user["is_admin"] or run["uploaded_by"] == user["username"])


def _run_rows(run):
    """실행(run) 하나에 포함된 라인만 걸러서 반환한다(그 날짜 전체가 아니라)."""
    day_rows = load_orders(run["run_date"], run["run_date"])
    return [r for r in day_rows if r.get("_run_id") == run["run_id"]]


def _run_scoped_summary(rows, run_date):
    """실행에 포함된 라인만으로 그 실행만의 요약을 계산한다(그 날짜 전체
    합산이 아니라 - 업로드한 사람 본인 몫만 정확히 보이도록)."""
    by_vendor = {}
    unclassified = []
    ambiguous = []
    for r in rows:
        if r["match_status"] == "classified":
            by_vendor.setdefault(r["vendor"], []).append(r)
        elif r["match_status"] == "unclassified":
            unclassified.append(r)
        else:
            ambiguous.append(r)
    return compute_summary(rows, by_vendor, unclassified, ambiguous, run_date)


def _run_vendor_brands(rows):
    """실행에 포함된 라인을 (협력사, 사업부)별로 묶어 건수를 센다 - 이
    실행에서 발주 파일을 몇 개(협력사×사업부 조합) 내려받을 수 있는지 보여줄 때 쓴다."""
    counts = {}
    for r in rows:
        if r["match_status"] != "classified":
            continue
        key = (r["vendor"], r.get("brand") or "JM")
        counts[key] = counts.get(key, 0) + 1
    return [{"vendor": v, "brand": b, "count": c} for (v, b), c in sorted(counts.items())]


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    user = _current_user(request)
    owner = None if user["is_admin"] else user["username"]
    recent = runs_store.list_runs(uploaded_by=owner)[:14]
    return render(
        request,
        "index.html",
        {"brands": BRAND_LABELS, "recent": recent, "today": date.today().isoformat()},
    )


@app.post("/process")
async def process(request: Request, files: list[UploadFile] = File(...), brand: str = Form("")):
    username = request.session.get("username")
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
        run_id, written, review_path, summary_path, dashboard_path = pipeline.run(
            saved_paths, OUTPUT_DIR, brand_override=brand_override, uploaded_by=username
        )

    return RedirectResponse(url=f"/runs/{run_id}", status_code=303)


@app.get("/runs/{run_id}", response_class=HTMLResponse)
def run_detail(request: Request, run_id: str):
    user = _current_user(request)
    run = runs_store.get_run(run_id)
    if not run:
        return HTMLResponse("처리 결과를 찾을 수 없습니다.", status_code=404)
    if not _can_view_run(user, run):
        return HTMLResponse("본인이 처리한 결과만 볼 수 있습니다.", status_code=403)

    rows = _run_rows(run)
    summary = _run_scoped_summary(rows, run["run_date"])
    vendor_brands = _run_vendor_brands(rows)
    is_admin = user["is_admin"]
    files = _run_files(run["run_date"]) if is_admin else []
    review_exists = is_admin and (OUTPUT_DIR / f"확인필요_{run['run_date']}.xlsx").exists()
    dashboard_exists = is_admin and (OUTPUT_DIR / f"dashboard_{run['run_date']}.html").exists()
    return render(
        request,
        "run_detail.html",
        {
            "run": run,
            "run_date": run["run_date"],
            "summary": summary,
            "vendor_brands": vendor_brands,
            "files": files,
            "review_exists": review_exists,
            "dashboard_exists": dashboard_exists,
        },
    )


@app.get("/runs/{run_id}/vendor/{vendor}/{brand}")
def run_vendor_download(request: Request, run_id: str, vendor: str, brand: str):
    """이 실행(run)에서 특정 협력사·사업부로 분류된 주문만 모아 그 자리에서
    발주 엑셀을 만들어 내려준다 - 본인이 올린 분량에 한정되므로 업로더
    본인과 관리자만 받을 수 있다(전체 날짜 합산본과는 다름)."""
    user = _current_user(request)
    run = runs_store.get_run(run_id)
    if not run:
        return HTMLResponse("처리 결과를 찾을 수 없습니다.", status_code=404)
    if not _can_view_run(user, run):
        return HTMLResponse("본인이 처리한 결과만 볼 수 있습니다.", status_code=403)
    if vendor not in VENDORS:
        return HTMLResponse("알 수 없는 협력사입니다.", status_code=404)

    rows = [
        dict(r) for r in _run_rows(run)
        if r["match_status"] == "classified" and r["vendor"] == vendor and (r.get("brand") or "JM") == brand
    ]
    if not rows:
        return HTMLResponse("이 실행에는 해당 협력사·사업부 주문이 없습니다.", status_code=404)
    for r in rows:
        if r.get("order_date"):
            r["order_date"] = datetime.strptime(r["order_date"], "%Y-%m-%d").date()

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp) / "vendor.xlsx"
        write_vendor_file(vendor, rows, tmp_path)
        content = tmp_path.read_bytes()

    filename = f"{run['run_date'].replace('-', '')}_{vendor}_{brand}.xlsx"
    return StreamingResponse(
        io.BytesIO(content),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=\"vendor.xlsx\"; filename*=UTF-8''{quote(filename)}"},
    )


@app.get("/runs", response_class=HTMLResponse)
def run_list(request: Request, start: str = "", end: str = ""):
    user = _current_user(request)
    owner = None if user["is_admin"] else user["username"]
    recent = runs_store.list_runs(uploaded_by=owner, start_date=start or None, end_date=end or None)
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
    if not _require_admin(request):
        return HTMLResponse("참고 자료 교체는 관리자만 할 수 있습니다.", status_code=403)
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
def reference_confirm(request: Request, table_id: str, token: str = Form(...)):
    if not _require_admin(request):
        return HTMLResponse("참고 자료 교체는 관리자만 할 수 있습니다.", status_code=403)
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
def reference_cancel(request: Request, table_id: str, token: str = Form(...)):
    if not _require_admin(request):
        return HTMLResponse("참고 자료 교체는 관리자만 할 수 있습니다.", status_code=403)
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
