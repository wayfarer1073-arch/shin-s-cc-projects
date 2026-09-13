"""서버에 배포했을 때 데이터를 어디에 저장할지 결정하는 공용 모듈.

로컬에서 돌릴 때(지금까지 해왔던 방식)는 저장소 폴더 안의 data/, output/를
그대로 쓴다. 실제 서버(예: Render)에 올릴 때는 APP_DATA_DIR 환경변수로
영구 디스크 경로(예: /var/data)를 지정하면 그쪽에 저장하도록 바꿀 수 있다.
이렇게 해야 서버가 재배포될 때(코드가 바뀌어 다시 배포될 때마다) 계정·
처리 이력·참고자료·게시판 글이 사라지지 않는다 - 코드는 매번 새로 받아오지만
데이터는 별도 디스크에 남아있기 때문이다.
"""
import os
import shutil
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

_env_dir = os.environ.get("APP_DATA_DIR", "").strip()
APP_DATA_DIR = Path(_env_dir) if _env_dir else REPO_ROOT

DATA_DIR = APP_DATA_DIR / "data"
OUTPUT_DIR = APP_DATA_DIR / "output"
REFERENCE_DIR = DATA_DIR / "reference"
ARCHIVE_DIR = DATA_DIR / "orders"


def seed_if_empty():
    """영구 디스크를 처음 붙인 시점(APP_DATA_DIR가 저장소 밖을 가리키고
    아직 비어 있을 때)에, 저장소에 이미 커밋되어 있던 참고자료·처리 이력을
    디스크로 한 번만 복사해 넣는다. 이미 디스크에 있는 파일은 절대 덮어쓰지
    않는다(그 뒤로 실제 서비스에서 쌓인 내용이 있을 수 있으므로)."""
    if APP_DATA_DIR == REPO_ROOT:
        return  # 로컬 개발이면 저장소 폴더를 그대로 쓰니 복사할 필요 없음
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    seed_src = REPO_ROOT / "data"
    if not seed_src.exists():
        return
    for item in seed_src.rglob("*"):
        if item.is_dir():
            continue
        dest = DATA_DIR / item.relative_to(seed_src)
        if dest.exists():
            continue
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(item, dest)
