"""Bounded, read-only T.V.C window probe.

The probe never focuses, clicks, types into, or closes a window. It first
matches top-level metadata and only inspects controls that belong to a process
already identified as T.V.C.
"""

import configparser
import json
import re
import sys
import time

from pywinauto import Desktop
from runtime_paths import resolve_runtime_paths
import tvc_window_locator
from utils import configure_utf8_stdio


RESULT_PREFIX="TVC_PROBE_RESULT "
JOB_FORM_CONTROL_IDS={"ButtonX3","ListView1","Tno"}
POST_LOGIN_CONTROL_TEXTS={"ใบงาน"}
POST_LOGIN_TITLE_PATTERNS=(
    ("job_list_title",re.compile(r"รายการใบงาน\s*\(JOB\)",re.I)),
    (
        "logged_in_user_title",
        re.compile(r"(?:เข้าสู่ระบบ|เข้าระบบ)\s*โดย\s*[^\s\[\]]+",re.I),
    ),
)
DEFAULT_PROBE_TIMEOUT_SECONDS=2.5
MAX_TVC_WINDOWS=20
MAX_TVC_CONTROLS=300


def _window_metadata(window):
    """Return cheap top-level metadata without walking a control tree."""
    try:
        title=str(window.window_text() or "").strip()
    except Exception:
        title=""
    try:
        class_name=str(window.class_name() or "").strip()
    except Exception:
        class_name=""
    try:
        process_id=int(getattr(window.element_info,"process_id",0) or 0)
    except Exception:
        try:
            process_id=int(window.process_id() or 0)
        except Exception:
            process_id=0
    return title,class_name,process_id


def _positive_title_signals(title):
    """Return explicit post-login title signals; absence never implies login."""
    return [name for name,pattern in POST_LOGIN_TITLE_PATTERNS if pattern.search(title)]


def _scan_tvc_controls(window,deadline,clock=None,max_controls=MAX_TVC_CONTROLS):
    """Inspect a bounded tree belonging to an already matched T.V.C process."""
    clock=clock or time.monotonic
    found_job_ids=set()
    login_signals=[]
    pending=[window]
    inspected=0
    while pending and inspected<max_controls:
        if clock()>=deadline:
            return False,login_signals,True
        control=pending.pop(0)
        inspected+=1
        try:
            automation_id=str(
                getattr(control.element_info,"automation_id","") or ""
            )
        except Exception:
            automation_id=""
        if automation_id in JOB_FORM_CONTROL_IDS:
            found_job_ids.add(automation_id)
        try:
            control_text=str(control.window_text() or "").strip()
        except Exception:
            control_text=""
        if control_text in POST_LOGIN_CONTROL_TEXTS:
            signal=f"post_login_control:{control_text}"
            if signal not in login_signals:
                login_signals.append(signal)
        if found_job_ids==JOB_FORM_CONTROL_IDS:
            if "active_job_form_controls" not in login_signals:
                login_signals.append("active_job_form_controls")
            return True,login_signals,False
        try:
            children=list(control.children())
        except Exception:
            children=[]
        remaining=max_controls-inspected-len(pending)
        if remaining>0:
            pending.extend(children[:remaining])
    return False,login_signals,clock()>=deadline


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


def _result(status,*,connected=False,login_verified=False,active_job_forms=None,
            matches=None,login_signals=None,main_window=None,errors=None,
            started=None,clock=None):
    clock=clock or time.monotonic
    elapsed=max(0.0,clock()-(started if started is not None else clock()))
    forms=list(active_job_forms or [])
    return {
        "status":status,
        "connected":bool(connected),
        "login_verified":bool(login_verified),
        "active_job_form":bool(forms),
        "active_job_forms":forms,
        "matches":list(matches or []),
        "login_signals":list(login_signals or []),
        "main_window":dict(main_window or {}),
        "errors":list(errors or []),
        "timed_out":status=="TIMEOUT",
        "duration_ms":round(elapsed*1000),
    }


def probe(runtime_paths=None,timeout_seconds=DEFAULT_PROBE_TIMEOUT_SECONDS,
          desktop_factory=None,clock=None):
    """Probe T.V.C within a deadline and never inspect unrelated controls."""
    desktop_factory=desktop_factory or Desktop
    clock=clock or time.monotonic
    started=clock()
    timeout_seconds=max(0.1,float(timeout_seconds))
    deadline=started+timeout_seconds
    cfg,_config_file=_load_probe_config(runtime_paths)
    configured_backend=cfg.get("tvc","backend",fallback="win32").strip() or "win32"
    title_regex=cfg.get(
        "tvc","window_title_regex",fallback=r"^เพิ่มใบงาน \(JOB\)$"
    )
    try:
        compiled_title=re.compile(title_regex,re.I)
    except re.error as exc:
        return _result(
            "ERROR",errors=[f"invalid window_title_regex: {exc}"],
            started=started,clock=clock,
        )

    remaining=max(0.1,deadline-clock())
    top_scan=tvc_window_locator.scan_top_level_windows(
        configured_backend,
        timeout_seconds=remaining,
        desktop_factory=desktop_factory,
        clock=clock,
    )
    if top_scan.timed_out:
        return _result(
            "TIMEOUT",errors=list(top_scan.errors),
            started=started,clock=clock,
        )
    if top_scan.errors and not top_scan.windows:
        return _result(
            "ERROR",errors=list(top_scan.errors),
            started=started,clock=clock,
        )

    job_records=[
        record for record in top_scan.windows
        if record.title and compiled_title.search(record.title)
    ]
    main_candidates=[
        record for record in top_scan.windows
        if tvc_window_locator.is_tvc_main_title(record.title)
    ]
    selected_main=main_candidates[0] if main_candidates else None
    main_window=selected_main.as_dict() if selected_main else None

    # No candidate means NOT_FOUND; controls of unrelated apps are untouched.
    if selected_main is None and not job_records:
        return _result("NOT_FOUND",started=started,clock=clock)

    regex_matches=[record.title for record in job_records]
    client_matches=[selected_main.title] if selected_main else []
    matches=list(dict.fromkeys(client_matches+regex_matches))
    active_job_forms=list(dict.fromkeys(regex_matches))
    login_signals=[]
    if regex_matches:
        login_signals.append("configured_job_title")
    for title in client_matches:
        for signal in _positive_title_signals(title):
            if signal not in login_signals:
                login_signals.append(signal)

    process_ids={
        record.process_id for record in job_records
        if record.process_id>0
    }
    if selected_main is not None and selected_main.process_id>0:
        process_ids.add(selected_main.process_id)
    process_ids=sorted(process_ids)
    errors=[]
    if process_ids and not active_job_forms and not login_signals:
        try:
            uia_desktop=desktop_factory(backend="uia")
            for process_id in process_ids:
                if clock()>=deadline:
                    return _result(
                        "TIMEOUT",connected=True,active_job_forms=active_job_forms,
                        matches=matches,login_signals=login_signals,
                        main_window=main_window,
                        errors=errors+["T.V.C control scan exceeded deadline"],
                        started=started,clock=clock,
                    )
                try:
                    candidate_windows=list(uia_desktop.windows(process=process_id))
                except Exception as exc:
                    errors.append(f"uia process {process_id}: {exc}")
                    continue
                for window in candidate_windows[:MAX_TVC_WINDOWS]:
                    has_controls,control_signals,timed_out=_scan_tvc_controls(
                        window,deadline,clock=clock
                    )
                    for signal in control_signals:
                        if signal not in login_signals:
                            login_signals.append(signal)
                    if timed_out:
                        return _result(
                            "TIMEOUT",connected=True,active_job_forms=active_job_forms,
                            matches=matches,login_signals=login_signals,
                            main_window=main_window,
                            errors=errors+["T.V.C control scan exceeded deadline"],
                            started=started,clock=clock,
                        )
                    if has_controls:
                        title,_,_=_window_metadata(window)
                        label=title or "<active JOB form>"
                        if label not in active_job_forms:
                            active_job_forms.append(label)
        except Exception as exc:
            errors.append(f"uia: {exc}")

    login_verified=bool(login_signals)
    return _result(
        "READY" if login_verified else "LOGIN_REQUIRED",
        connected=True,
        login_verified=login_verified,
        active_job_forms=active_job_forms,
        matches=matches,
        login_signals=login_signals,
        main_window=main_window,
        errors=errors,
        started=started,
        clock=clock,
    )


def main():
    configure_utf8_stdio()
    try:
        result=probe()
    except Exception as exc:
        result=_result("ERROR",errors=[str(exc)])
    # stdout is an IPC channel. Escaping non-ASCII characters prevents the
    # frozen console worker from touching cp1252 with Thai window/control text.
    print(RESULT_PREFIX+json.dumps(result,ensure_ascii=True),flush=True)
    if result["status"] in {"READY","NOT_FOUND","LOGIN_REQUIRED","FOUND"}:
        return 0
    return 2


if __name__=="__main__":
    sys.exit(main())
