"""LibreOffice로 엑셀 파일에 열기 암호를 걸어 저장한다 (원본은 그대로 두고
별도 경로에 암호화된 사본을 만든다). 정산 리포트처럼 개인정보/원가가 들어있는
파일을 조직 내 권한 없는 사람이 열지 못하게 할 때 쓴다.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

_MACRO_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub ProtectAndSave()
      Dim oArgs(1) As New com.sun.star.beans.PropertyValue
      oArgs(0).Name = "FilterName"
      oArgs(0).Value = "Calc MS Excel 2007 XML"
      oArgs(1).Name = "Password"
      oArgs(1).Value = "{password}"
      ThisComponent.storeToURL(ThisComponent.getURL(), oArgs())
      ThisComponent.close(True)
    End Sub
</script:module>"""


def _soffice_env():
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    return env


def protect_xlsx(in_path, out_path, password, timeout=60):
    """in_path를 복사해 out_path 위치에 password로 암호를 건 사본을 만든다."""
    in_path = Path(in_path)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(in_path, out_path)

    with tempfile.TemporaryDirectory(prefix="protect-lo-profile-") as profile_dir:
        profile_dir = Path(profile_dir)
        profile_url = profile_dir.as_uri()

        subprocess.run(
            ["soffice", "--headless", "--terminate_after_init",
             f"-env:UserInstallation={profile_url}"],
            capture_output=True, timeout=timeout, env=_soffice_env(),
        )

        macro_dir = profile_dir / "user" / "basic" / "Standard"
        if not macro_dir.exists():
            raise RuntimeError("LibreOffice 프로필 생성에 실패했습니다 (암호화 불가)")

        escaped_password = password.replace('"', '""')
        macro = _MACRO_TEMPLATE.format(password=escaped_password)
        (macro_dir / "Module1.xba").write_text(macro, encoding="utf-8")

        cmd = [
            "soffice", "--headless", "--norestore",
            f"-env:UserInstallation={profile_url}",
            "vnd.sun.star.script:Standard.Module1.ProtectAndSave?language=Basic&location=application",
            str(out_path.resolve()),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, env=_soffice_env(), timeout=timeout)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 암호화 실패: {result.stderr[-500:]}")

    return out_path
