"""사내 게시판(이슈/공지/잡담 공유).

data/board.json에 글 목록을 저장한다. 글마다 태그(이슈/공지/잡담),
제목(50자 이내), 본문(200자 이내), 작성자, 작성 시각을 가진다."""
import json
import time
import uuid

from src.paths import DATA_DIR

BOARD_PATH = DATA_DIR / "board.json"

TAGS = ["이슈", "공지", "잡담"]
TITLE_MAX = 50
BODY_MAX = 200


def _load():
    if not BOARD_PATH.exists():
        return []
    with open(BOARD_PATH, encoding="utf-8") as f:
        return json.load(f)


def _save(posts):
    BOARD_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(BOARD_PATH, "w", encoding="utf-8") as f:
        json.dump(posts, f, ensure_ascii=False, indent=2)


def list_posts():
    """최신순으로 전체 글 목록을 반환한다."""
    posts = _load()
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


def create_post(tag, title, body, author):
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
        "created_at": time.strftime("%Y-%m-%d %H:%M"),
    })
    _save(posts)
