"""사내 게시판(이슈/공지/잡담 공유).

data/board.json에 글 목록을 저장한다. 글마다 태그(이슈/공지/잡담),
제목(50자 이내), 본문(200자 이내), 작성자, 작성 시각을 가진다.

쓸데없이 용량을 계속 차지하지 않도록 작성한 지 RETENTION_DAYS(30일)가
지난 글은 자동으로 지운다(별도 스케줄러 없이, 글 목록을 불러올 때마다
확인해서 지운다)."""
import json
import time
import uuid
from datetime import date, timedelta

from src.paths import DATA_DIR

BOARD_PATH = DATA_DIR / "board.json"

TAGS = ["이슈", "공지", "잡담"]
TITLE_MAX = 50
BODY_MAX = 200
RETENTION_DAYS = 30


def _load():
    if not BOARD_PATH.exists():
        return []
    with open(BOARD_PATH, encoding="utf-8") as f:
        return json.load(f)


def _save(posts):
    BOARD_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(BOARD_PATH, "w", encoding="utf-8") as f:
        json.dump(posts, f, ensure_ascii=False, indent=2)


def purge_old_posts():
    """작성한 지 RETENTION_DAYS가 지난 글을 지운다. 지운 게 있을 때만
    파일을 다시 쓴다(매번 불필요하게 저장하지 않도록)."""
    cutoff = (date.today() - timedelta(days=RETENTION_DAYS)).isoformat()
    posts = _load()
    remaining = [p for p in posts if p["created_at"][:10] >= cutoff]
    if len(remaining) != len(posts):
        _save(remaining)


def list_posts(tag=None, start_date=None, end_date=None):
    """최신순으로 글 목록을 반환한다. tag를 주면 그 태그만, start_date/
    end_date("YYYY-MM-DD")를 주면 작성일이 그 기간(둘 다 포함) 안인 글만.
    부를 때마다 보관 기간이 지난 글을 먼저 정리한다."""
    purge_old_posts()
    posts = _load()
    if tag:
        posts = [p for p in posts if p.get("tag") == tag]
    if start_date:
        posts = [p for p in posts if p["created_at"][:10] >= start_date]
    if end_date:
        posts = [p for p in posts if p["created_at"][:10] <= end_date]
    return sorted(posts, key=lambda p: p["created_at"], reverse=True)


def latest_by_tag():
    """태그별 가장 최근 글 1건씩. 사이드바 미리보기에 쓴다."""
    posts = list_posts()
    result = {tag: None for tag in TAGS}
    for p in posts:
        tag = p.get("tag")
        if tag in result and result[tag] is None:
            result[tag] = p
        if all(result.values()):
            break
    return result


def create_post(tag, title, body, author, author_username):
    tag = (tag or "").strip()
    title = (title or "").strip()
    body = (body or "").strip()
    if tag not in TAGS:
        raise ValueError("태그를 선택해주세요(이슈/공지/잡담 중 하나).")
    if not title:
        raise ValueError("제목을 입력해주세요.")
    if len(title) > TITLE_MAX:
        raise ValueError(f"제목은 {TITLE_MAX}자 이내로 입력해주세요.")
    if not body:
        raise ValueError("내용을 입력해주세요.")
    if len(body) > BODY_MAX:
        raise ValueError(f"내용은 {BODY_MAX}자 이내로 입력해주세요.")

    posts = _load()
    posts.append({
        "id": uuid.uuid4().hex,
        "tag": tag,
        "title": title,
        "body": body,
        "author": author,
        "author_username": author_username,
        "created_at": time.strftime("%Y-%m-%d %H:%M"),
    })
    _save(posts)


def get_post(post_id):
    for p in _load():
        if p["id"] == post_id:
            return p
    return None


def delete_post(post_id):
    """지운 글이 있었으면 True, 없으면(이미 지워졌거나 없는 id) False."""
    posts = _load()
    remaining = [p for p in posts if p["id"] != post_id]
    if len(remaining) == len(posts):
        return False
    _save(remaining)
    return True
