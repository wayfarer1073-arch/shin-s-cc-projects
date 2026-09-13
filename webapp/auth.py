"""직원별 로그인 계정 관리.

계정 정보는 data/users.json에 저장한다(비밀번호는 원문이 아니라 bcrypt
해시만 저장). 세션은 서명된 쿠키(Starlette SessionMiddleware)로 유지하며,
서명에 쓰는 비밀키는 서버를 껐다 켜도 같은 값을 쓰도록 파일로 보관한다
(안 그러면 재시작할 때마다 전부 로그아웃됨)."""
import json
import secrets
import time
from pathlib import Path

import bcrypt

BASE_DIR = Path(__file__).resolve().parent.parent
USERS_PATH = BASE_DIR / "data" / "users.json"
SECRET_PATH = BASE_DIR / "data" / ".session_secret"


def get_session_secret():
    if SECRET_PATH.exists():
        return SECRET_PATH.read_text().strip()
    secret = secrets.token_hex(32)
    SECRET_PATH.parent.mkdir(parents=True, exist_ok=True)
    SECRET_PATH.write_text(secret)
    return secret


def _load_users():
    if not USERS_PATH.exists():
        return []
    with open(USERS_PATH, encoding="utf-8") as f:
        return json.load(f)


def _save_users(users):
    USERS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(USERS_PATH, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def users_exist():
    return len(_load_users()) > 0


def list_users():
    return [{k: v for k, v in u.items() if k != "password_hash"} for u in _load_users()]


def find_user(username):
    for u in _load_users():
        if u["username"] == username:
            return u
    return None


def verify_login(username, password):
    user = find_user(username)
    if not user:
        return None
    if bcrypt.checkpw(password.encode("utf-8"), user["password_hash"].encode("utf-8")):
        return user
    return None


def create_user(username, password, display_name="", is_admin=False):
    username = username.strip()
    if not username or not password:
        raise ValueError("아이디와 비밀번호를 모두 입력해주세요.")
    if len(password) < 4:
        raise ValueError("비밀번호는 4자 이상이어야 합니다.")
    users = _load_users()
    if any(u["username"] == username for u in users):
        raise ValueError(f'"{username}" 아이디는 이미 있습니다.')
    pw_hash = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    users.append({
        "username": username,
        "password_hash": pw_hash,
        "display_name": display_name.strip() or username,
        "is_admin": bool(is_admin),
        "created_at": time.strftime("%Y-%m-%d %H:%M"),
    })
    _save_users(users)


def set_password(username, new_password):
    if len(new_password) < 4:
        raise ValueError("비밀번호는 4자 이상이어야 합니다.")
    users = _load_users()
    for u in users:
        if u["username"] == username:
            u["password_hash"] = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            _save_users(users)
            return True
    return False


def delete_user(username):
    users = _load_users()
    remaining = [u for u in users if u["username"] != username]
    if len(remaining) == len(users):
        return False
    if not any(u["is_admin"] for u in remaining):
        raise ValueError("마지막 남은 관리자 계정은 지울 수 없습니다.")
    _save_users(remaining)
    return True


# 로그인 없이 접근 가능한 경로(로그인 화면 자체, 최초 관리자 계정 만들기,
# 브라우저가 자동으로 요청하는 파비콘 등).
PUBLIC_PATHS = {"/login", "/setup", "/favicon.ico"}


def is_public_path(path):
    return path in PUBLIC_PATHS
