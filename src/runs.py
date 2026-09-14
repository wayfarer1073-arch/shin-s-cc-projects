"""업로드(처리) 실행 단위 기록 - 누가, 언제, 어떤 파일로 처리했는지.

data/runs.json에 실행 목록을 저장한다("실행" = 한 번의 업로드+처리, 날짜와는
다른 개념 - 같은 날 두 번 올리면 실행은 2건). 처리 이력/처리 결과 상세
화면에서 "본인이 올린 것만" 또는(관리자는) "전체 + 누가 올렸는지"를
보여주는 데 쓴다.

data/orders/{날짜}.json에 보관되는 각 주문 라인에도 _run_id/_uploaded_by를
함께 찍어둔다(archive_orders가 처리) - 그래야 처리 결과 상세 화면에서
"이 실행에 포함된 주문만" 걸러서 통계를 낼 수 있다.
"""
import json
import time
import uuid
from pathlib import Path

from src.paths import DATA_DIR, ARCHIVE_DIR

RUNS_PATH = DATA_DIR / "runs.json"


def new_run_id():
    return uuid.uuid4().hex


def _load():
    if not RUNS_PATH.exists():
        return []
    with open(RUNS_PATH, encoding="utf-8") as f:
        return json.load(f)


def _save(runs):
    RUNS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(RUNS_PATH, "w", encoding="utf-8") as f:
        json.dump(runs, f, ensure_ascii=False, indent=2)


def record_run(run_id, run_date, uploaded_by, filenames, order_count, unclassified_count, ambiguous_count, raw_read_count=None):
    """raw_read_count: 업로드한 파일에서 실제로 읽어들인 줄 수(중복 제거 전).
    order_count(이 실행에서 새로 보관된 줄 수)가 0인데 raw_read_count는 0보다
    크면 "파일은 읽었지만 전부 이미 처리된 내용과 중복"이라는 뜻이고, raw_read_count도
    0이면 "파일에서 아예 아무 줄도 못 읽었다"는 뜻이라 원인이 다르다 - 처리 결과
    화면에서 이 둘을 구분해 보여준다."""
    runs = _load()
    runs.append({
        "run_id": run_id,
        "run_date": run_date,
        "uploaded_by": uploaded_by,
        "uploaded_at": time.strftime("%Y-%m-%d %H:%M"),
        "filenames": filenames,
        "order_count": order_count,
        "unclassified_count": unclassified_count,
        "ambiguous_count": ambiguous_count,
        "raw_read_count": raw_read_count,
    })
    _save(runs)


def list_runs(uploaded_by=None, start_date=None, end_date=None):
    """최신순. uploaded_by를 주면 그 사람 것만. start_date/end_date("YYYY-MM-DD")를
    주면 그 기간의 run_date만."""
    runs = _load()
    if uploaded_by is not None:
        runs = [r for r in runs if r.get("uploaded_by") == uploaded_by]
    if start_date:
        runs = [r for r in runs if r["run_date"] >= start_date]
    if end_date:
        runs = [r for r in runs if r["run_date"] <= end_date]
    return sorted(runs, key=lambda r: r["uploaded_at"], reverse=True)


def get_run(run_id):
    for r in _load():
        if r["run_id"] == run_id:
            return r
    return None


def backfill_legacy_runs(archive_dir=ARCHIVE_DIR):
    """이 기능을 만들기 전에 이미 보관돼 있던 주문 라인(_run_id가 없는 것)에
    날짜별로 하나씩 "레거시 실행"을 만들어 붙여준다 - 예전 이력이 화면에서
    통째로 사라지지 않도록. 업로더를 알 수 없으므로 관리자에게만 보이게
    uploaded_by는 None으로 남긴다. 이미 처리된 파일/실행은 다시 건드리지
    않는다(멱등 - 서버가 몇 번을 다시 시작해도 안전)."""
    archive_dir = Path(archive_dir)
    if not archive_dir.exists():
        return
    existing_run_ids = {r["run_id"] for r in _load()}
    for path in sorted(archive_dir.glob("*.json")):
        run_date = path.stem
        with open(path, encoding="utf-8") as f:
            rows = json.load(f)
        missing = [r for r in rows if not r.get("_run_id")]
        if not missing:
            continue
        legacy_run_id = f"legacy-{run_date}"
        for r in missing:
            r["_run_id"] = legacy_run_id
            r["_uploaded_by"] = None
        with open(path, "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
        if legacy_run_id not in existing_run_ids:
            unclassified_count = sum(1 for r in missing if r.get("match_status") == "unclassified")
            ambiguous_count = sum(1 for r in missing if r.get("match_status") == "ambiguous")
            record_run(
                legacy_run_id, run_date, None,
                filenames=[], order_count=len(missing),
                unclassified_count=unclassified_count, ambiguous_count=ambiguous_count,
            )
            existing_run_ids.add(legacy_run_id)
