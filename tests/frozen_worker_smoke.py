"""Post-build smoke tests for the real frozen Worker executable.

This test-only script creates an inert Tk window as a read-only T.V.C fixture.
It never imports or invokes bot automation.
"""

import argparse
import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time


RESULT_PREFIX="TVC_PROBE_RESULT "
DIAGNOSTIC_PREFIX="TVC_WINDOW_DIAGNOSTIC "
FIXTURE_TITLE="T.V.C Client [Version TEST] หน้าหลักทดสอบ"
NEUTRAL_FIXTURE_TITLE="T.V.C Client [Version TEST]"


def fail(message):
    raise RuntimeError(message)


def parse_probe(completed,expected_status):
    combined=(completed.stdout or "")+(completed.stderr or "")
    if "UnicodeEncodeError" in combined:
        fail("frozen Worker emitted UnicodeEncodeError")
    lines=[
        line for line in (completed.stdout or "").splitlines()
        if line.startswith(RESULT_PREFIX)
    ]
    if not lines:
        fail(f"frozen Worker returned no probe JSON: {combined}")
    line=lines[-1]
    if not line.isascii():
        fail("frozen probe IPC is not ASCII-safe")
    payload=json.loads(line[len(RESULT_PREFIX):])
    if payload.get("status")!=expected_status:
        fail(f"expected {expected_status}, got {payload.get('status')}: {payload}")
    return payload


def parse_diagnostic(completed,expected_status):
    combined=(completed.stdout or "")+(completed.stderr or "")
    if "UnicodeEncodeError" in combined:
        fail("frozen Worker diagnostic emitted UnicodeEncodeError")
    lines=[
        line for line in (completed.stdout or "").splitlines()
        if line.startswith(DIAGNOSTIC_PREFIX)
    ]
    if not lines:
        fail(f"frozen Worker returned no diagnostic JSON: {combined}")
    line=lines[-1]
    if not line.isascii():
        fail("frozen diagnostic IPC is not ASCII-safe")
    payload=json.loads(line[len(DIAGNOSTIC_PREFIX):])
    if payload.get("status")!=expected_status:
        fail(
            f"expected diagnostic {expected_status}, "
            f"got {payload.get('status')}: {payload}"
        )
    return payload


def run_worker(worker,*arguments):
    environment=os.environ.copy()
    environment["PYTHONIOENCODING"]="cp1252"
    return subprocess.run(
        [str(worker),*arguments],
        cwd=str(worker.parent),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=10,
        env=environment,
    )


def assert_same_main_window(probe_payload,diagnostic_payload):
    probe_window=probe_payload.get("main_window") or {}
    diagnostic_window=diagnostic_payload.get("selected") or {}
    for field in ("title","handle","pid","backend"):
        if probe_window.get(field)!=diagnostic_window.get(field):
            fail(
                f"probe/driver selected different {field}: "
                f"{probe_window} vs {diagnostic_window}"
            )


def native_fixture_window(ready_file,title,include_button):
    """Create inert native HWNDs that the frozen UIA backend can inspect."""
    user32=ctypes.windll.user32
    user32.CreateWindowExW.restype=wintypes.HWND
    style=0x00CF0000
    child_style=0x50000000
    window=user32.CreateWindowExW(
        0,"STATIC",title,style,-32000,-32000,300,120,0,0,0,None
    )
    if not window:
        raise ctypes.WinError()
    if include_button:
        button=user32.CreateWindowExW(
            0,"BUTTON","ใบงาน",child_style,20,20,120,32,window,0,0,None
        )
        if not button:
            raise ctypes.WinError()
    user32.ShowWindow(window,4)
    user32.UpdateWindow(window)
    ready_file.write_text("ready",encoding="ascii")
    message=wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(message),0,0,0)>0:
        user32.TranslateMessage(ctypes.byref(message))
        user32.DispatchMessageW(ctypes.byref(message))


def fixture_window(ready_file):
    native_fixture_window(ready_file,FIXTURE_TITLE,include_button=True)


def neutral_fixture_window(ready_file):
    native_fixture_window(
        ready_file,NEUTRAL_FIXTURE_TITLE,include_button=False
    )


def run_fixture_probe(worker,fixture_argument,expected_status):
    with tempfile.TemporaryDirectory(prefix="tvc_frozen_smoke_") as temp_dir:
        ready_file=Path(temp_dir)/"fixture.ready"
        creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0)
        fixture=subprocess.Popen(
            [sys.executable,str(Path(__file__).resolve()),fixture_argument,str(ready_file)],
            creationflags=creationflags,
        )
        try:
            deadline=time.monotonic()+8
            while time.monotonic()<deadline and not ready_file.is_file():
                if fixture.poll() is not None:
                    fail(f"fake T.V.C fixture exited with {fixture.returncode}")
                time.sleep(0.05)
            if not ready_file.is_file():
                fail("fake T.V.C fixture did not become ready")
            result=run_worker(worker,"--probe-tvc")
            if result.returncode!=0:
                fail(f"{expected_status} probe exit code was {result.returncode}")
            probe_payload=parse_probe(result,expected_status)
            diagnostic_result=run_worker(worker,"--diagnose-tvc-window")
            if diagnostic_result.returncode!=0:
                fail(
                    f"FOUND diagnostic exit code was "
                    f"{diagnostic_result.returncode}"
                )
            diagnostic_payload=parse_diagnostic(diagnostic_result,"FOUND")
            if not diagnostic_payload.get("driver_can_attach"):
                fail(f"driver cannot attach to fixture: {diagnostic_payload}")
            return probe_payload,diagnostic_payload
        finally:
            if fixture.poll() is None:
                fixture.terminate()
                try:
                    fixture.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    fixture.kill()
                    fixture.wait(timeout=3)


def smoke(worker):
    help_result=run_worker(worker,"--help")
    if help_result.returncode!=0 or "usage:" not in (help_result.stdout or ""):
        fail(f"frozen Worker --help failed: {help_result.stdout}{help_result.stderr}")

    not_found=run_worker(worker,"--probe-tvc")
    if not_found.returncode!=0:
        fail(f"NOT_FOUND probe exit code was {not_found.returncode}")
    parse_probe(not_found,"NOT_FOUND")
    not_found_diagnostic=run_worker(worker,"--diagnose-tvc-window")
    if not_found_diagnostic.returncode!=0:
        fail(
            f"NOT_FOUND diagnostic exit code was "
            f"{not_found_diagnostic.returncode}"
        )
    diagnostic_payload=parse_diagnostic(not_found_diagnostic,"NOT_FOUND")
    if diagnostic_payload.get("driver_can_attach"):
        fail(f"NOT_FOUND diagnostic unexpectedly attachable: {diagnostic_payload}")

    neutral,neutral_diagnostic=run_fixture_probe(
        worker,"--neutral-fixture-window","LOGIN_REQUIRED"
    )
    if neutral.get("login_verified") or neutral.get("login_signals"):
        fail(f"neutral fixture was incorrectly verified: {neutral}")
    if neutral_diagnostic.get("selected",{}).get("title")!=NEUTRAL_FIXTURE_TITLE:
        fail(f"neutral diagnostic selected wrong window: {neutral_diagnostic}")
    assert_same_main_window(neutral,neutral_diagnostic)

    payload,ready_diagnostic=run_fixture_probe(
        worker,"--fixture-window","READY"
    )
    if FIXTURE_TITLE not in payload.get("matches",[]):
        fail(f"Thai fixture title did not round-trip: {payload}")
    if "post_login_control:ใบงาน" not in payload.get("login_signals",[]):
        fail(f"frozen UIA control signal was not verified: {payload}")
    if ready_diagnostic.get("selected",{}).get("title")!=FIXTURE_TITLE:
        fail(f"READY diagnostic selected wrong window: {ready_diagnostic}")
    assert_same_main_window(payload,ready_diagnostic)

    print("Frozen Worker smoke: PASS")


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--worker",type=Path)
    parser.add_argument("--fixture-window",type=Path)
    parser.add_argument("--neutral-fixture-window",type=Path)
    args=parser.parse_args()
    if args.fixture_window is not None:
        fixture_window(args.fixture_window)
        return 0
    if args.neutral_fixture_window is not None:
        neutral_fixture_window(args.neutral_fixture_window)
        return 0
    if args.worker is None or not args.worker.is_file():
        parser.error("--worker must point to the built Worker executable")
    smoke(args.worker.resolve())
    return 0


if __name__=="__main__":
    raise SystemExit(main())
