"""Shared, bounded, read-only discovery for the T.V.C main window."""

from __future__ import annotations

import configparser
from dataclasses import dataclass, field
import json
import re
import sys
import time

from pywinauto import Desktop


DIAGNOSTIC_PREFIX="TVC_WINDOW_DIAGNOSTIC "
DEFAULT_LOCATOR_TIMEOUT_SECONDS=1.5
MAX_TOP_LEVEL_WINDOWS=200
TVC_MAIN_TITLE_PATTERN=re.compile(
    r"^\s*T\.V\.C\s+Client(?:\s|\[|$)",
    re.I,
)


@dataclass
class TopLevelWindow:
    title: str
    class_name: str
    handle: int
    process_id: int
    backend: str
    window: object=field(repr=False,compare=False)

    def as_dict(self):
        return {
            "title":self.title,
            "class_name":self.class_name,
            "handle":self.handle,
            "pid":self.process_id,
            "backend":self.backend,
        }


@dataclass
class WindowScanResult:
    backend: str
    windows: list[TopLevelWindow]
    errors: list[str]
    timed_out: bool
    duration_ms: int


@dataclass
class MainWindowLocatorResult:
    backend: str
    selected: TopLevelWindow | None
    candidates: list[TopLevelWindow]
    errors: list[str]
    timed_out: bool
    duration_ms: int

    @property
    def status(self):
        if self.timed_out:
            return "TIMEOUT"
        if self.errors and not self.candidates:
            return "ERROR"
        return "FOUND" if self.selected is not None else "NOT_FOUND"


def is_tvc_main_title(title):
    """Match only the verified top-level T.V.C Client caption family."""
    return bool(TVC_MAIN_TITLE_PATTERN.search(str(title or "")))


def _window_metadata(window,backend):
    try:
        title=str(window.window_text() or "").strip()
    except Exception:
        title=""
    try:
        class_name=str(window.class_name() or "").strip()
    except Exception:
        class_name=""
    try:
        handle=int(getattr(window.element_info,"handle",0) or 0)
    except Exception:
        try:
            handle=int(getattr(window,"handle",0) or 0)
        except Exception:
            handle=0
    try:
        process_id=int(getattr(window.element_info,"process_id",0) or 0)
    except Exception:
        try:
            process_id=int(window.process_id() or 0)
        except Exception:
            process_id=0
    return TopLevelWindow(
        title=title,
        class_name=class_name,
        handle=handle,
        process_id=process_id,
        backend=backend,
        window=window,
    )


def scan_top_level_windows(
    backend="win32",
    *,
    timeout_seconds=DEFAULT_LOCATOR_TIMEOUT_SECONDS,
    desktop_factory=None,
    clock=None,
    max_windows=MAX_TOP_LEVEL_WINDOWS,
):
    """Read top-level metadata once; never inspect descendants or send input."""
    desktop_factory=desktop_factory or Desktop
    clock=clock or time.monotonic
    backend=str(backend or "win32").strip() or "win32"
    started=clock()
    deadline=started+max(0.1,float(timeout_seconds))
    try:
        raw_windows=list(desktop_factory(backend=backend).windows())
    except Exception as exc:
        return WindowScanResult(
            backend=backend,
            windows=[],
            errors=[f"{backend}: {exc}"],
            timed_out=False,
            duration_ms=round(max(0.0,clock()-started)*1000),
        )
    if clock()>=deadline:
        return WindowScanResult(
            backend=backend,
            windows=[],
            errors=["top-level window enumeration exceeded deadline"],
            timed_out=True,
            duration_ms=round(max(0.0,clock()-started)*1000),
        )

    records=[]
    for window in raw_windows[:max(1,int(max_windows))]:
        if clock()>=deadline:
            return WindowScanResult(
                backend=backend,
                windows=records,
                errors=["top-level window metadata scan exceeded deadline"],
                timed_out=True,
                duration_ms=round(max(0.0,clock()-started)*1000),
            )
        records.append(_window_metadata(window,backend))
    return WindowScanResult(
        backend=backend,
        windows=records,
        errors=[],
        timed_out=False,
        duration_ms=round(max(0.0,clock()-started)*1000),
    )


def locate_tvc_main_window(
    backend="win32",
    *,
    timeout_seconds=DEFAULT_LOCATOR_TIMEOUT_SECONDS,
    desktop_factory=None,
    clock=None,
):
    """Return the same verified main-window candidate used by probe and driver."""
    scan=scan_top_level_windows(
        backend,
        timeout_seconds=timeout_seconds,
        desktop_factory=desktop_factory,
        clock=clock,
    )
    candidates=[
        record for record in scan.windows
        if is_tvc_main_title(record.title)
    ]
    return MainWindowLocatorResult(
        backend=scan.backend,
        selected=candidates[0] if candidates else None,
        candidates=candidates,
        errors=list(scan.errors),
        timed_out=scan.timed_out,
        duration_ms=scan.duration_ms,
    )


def diagnose_tvc_main_window(runtime_paths=None,desktop_factory=None):
    """Return attach metadata only. This function never focuses or clicks."""
    if runtime_paths is None:
        from runtime_paths import resolve_runtime_paths
        runtime_paths=resolve_runtime_paths()
    cfg=configparser.ConfigParser()
    loaded=cfg.read(runtime_paths.config_file,encoding="utf-8")
    if not loaded:
        raise OSError(f"cannot read config.ini: {runtime_paths.config_file}")
    backend=cfg.get("tvc","backend",fallback="win32").strip() or "win32"
    result=locate_tvc_main_window(
        backend,
        desktop_factory=desktop_factory,
    )
    return {
        "status":result.status,
        "backend":result.backend,
        "candidate_count":len(result.candidates),
        "candidates":[candidate.as_dict() for candidate in result.candidates],
        "selected":result.selected.as_dict() if result.selected else None,
        "driver_can_attach":result.selected is not None,
        "timed_out":result.timed_out,
        "duration_ms":result.duration_ms,
        "errors":list(result.errors),
    }


def diagnostic_main():
    from utils import configure_utf8_stdio
    configure_utf8_stdio()
    try:
        result=diagnose_tvc_main_window()
    except Exception as exc:
        result={
            "status":"ERROR",
            "backend":"",
            "candidate_count":0,
            "candidates":[],
            "selected":None,
            "driver_can_attach":False,
            "timed_out":False,
            "duration_ms":0,
            "errors":[str(exc)],
        }
    print(DIAGNOSTIC_PREFIX+json.dumps(result,ensure_ascii=True),flush=True)
    return 0 if result["status"] in {"FOUND","NOT_FOUND"} else 2


if __name__=="__main__":
    sys.exit(diagnostic_main())
