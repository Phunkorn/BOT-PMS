"""Read-only T.V.C window probe. This module never clicks or changes T.V.C."""

import configparser
import json
import re
import sys

from pywinauto import Desktop
from runtime_paths import resolve_runtime_paths


RESULT_PREFIX="TVC_PROBE_RESULT "
JOB_FORM_CONTROL_IDS={"ButtonX3","ListView1","Tno"}


def _has_job_form_controls(window):
    """Inspect automation IDs only; never focus, click, type, or close a window."""
    found=set()
    try:
        controls=window.descendants()
    except Exception:
        return False
    for control in controls:
        try:
            automation_id=str(
                getattr(control.element_info,"automation_id","") or ""
            )
        except Exception:
            continue
        if automation_id in JOB_FORM_CONTROL_IDS:
            found.add(automation_id)
            if found==JOB_FORM_CONTROL_IDS:
                return True
    return False


def _load_probe_config(runtime_paths=None):
    """Load the same resolved config used by the GUI and Worker."""
    paths=runtime_paths or resolve_runtime_paths()
    config_file=paths.config_file
    if not config_file.is_file():
        raise FileNotFoundError(f"ไม่พบ config.ini สำหรับ T.V.C probe: {config_file}")
    cfg=configparser.ConfigParser()
    loaded=cfg.read(config_file,encoding="utf-8")
    if not loaded:
        raise OSError(f"อ่าน config.ini สำหรับ T.V.C probe ไม่สำเร็จ: {config_file}")
    return cfg,config_file


def probe(runtime_paths=None):
    cfg,_config_file=_load_probe_config(runtime_paths)
    configured_backend=cfg.get("tvc","backend",fallback="win32")
    title_regex=cfg.get("tvc","window_title_regex",fallback=r"^เพิ่มใบงาน \(JOB\)$")
    titles=[]
    active_job_forms=[]
    errors=[]
    for backend in dict.fromkeys(("uia",configured_backend)):
        try:
            for window in Desktop(backend=backend).windows():
                try:
                    title=str(window.window_text() or "").strip()
                except Exception:
                    continue
                if title and title not in titles:
                    titles.append(title)
                is_job_title=bool(title and re.search(title_regex,title,re.I))
                if is_job_title or _has_job_form_controls(window):
                    label=title or "<active JOB form>"
                    if label not in active_job_forms:
                        active_job_forms.append(label)
        except Exception as exc:
            errors.append(f"{backend}: {exc}")
    regex_matches=[title for title in titles if re.search(title_regex,title,re.I)]
    client_matches=[title for title in titles if "T.V.C Client" in title]
    matches=list(dict.fromkeys(client_matches+regex_matches))
    login_words=("login","log in","เข้าสู่ระบบ")
    client_after_login=[
        title for title in client_matches
        if not any(word in title.lower() for word in login_words)
    ]
    return {
        "connected":bool(matches),
        # A visible configured JOB window is the strongest read-only signal that
        # the user is already inside T.V.C. A non-login Client window is accepted
        # as the weaker fallback. No window is focused, clicked, or modified.
        "login_verified":bool(regex_matches or client_after_login),
        "active_job_form":bool(active_job_forms),
        "active_job_forms":active_job_forms,
        "matches":matches,
        "errors":errors,
    }


def main():
    try:
        result=probe()
    except Exception as exc:
        result={
            "connected":False,
            "login_verified":False,
            "active_job_form":False,
            "active_job_forms":[],
            "matches":[],
            "errors":[str(exc)],
        }
    print(RESULT_PREFIX+json.dumps(result,ensure_ascii=False),flush=True)
    return 0 if result["connected"] else 1


if __name__=="__main__":
    sys.exit(main())
