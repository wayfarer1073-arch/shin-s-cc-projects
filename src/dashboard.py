"""요약 통계(src/summary.py 결과)를 보고용 대시보드 HTML로 렌더링한다."""
import json
from pathlib import Path

_TEMPLATE = Path(__file__).resolve().parent / "dashboard_template.html"


def render_dashboard_html(summary: dict) -> str:
    template = _TEMPLATE.read_text(encoding="utf-8")
    data_json = json.dumps(summary, ensure_ascii=False)
    return template.replace("__SUMMARY_JSON__", data_json)
