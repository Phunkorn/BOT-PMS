from __future__ import annotations

import configparser
from dataclasses import dataclass
import importlib
import json
import os
from pathlib import Path
import queue
import stat
import subprocess
import sys
import threading
import time
import traceback
import uuid

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from gui_queue import (
    ExcelQueue,
    QueueRunController,
    build_queue_summary,
    calculate_queue_progress,
    normalize_excel_path,
)
from runtime_paths import build_worker_command, resolve_runtime_paths
from single_instance import WindowsSingleInstance
from version import APP_NAME, APP_VERSION

RUNTIME_PATH_ERROR=""
try:
    RUNTIME_PATHS=resolve_runtime_paths()
except Exception as exc:
    RUNTIME_PATHS=None
    RUNTIME_PATH_ERROR=str(exc)
if RUNTIME_PATHS is None:
    BASE=(
        Path(sys.executable).resolve().parent
        if bool(getattr(sys,"frozen",False))
        else Path(__file__).resolve().parents[1]
    )
    BOT_SCRIPT=None
    BOT_PYTHON=None
    LOG_DIR=BASE/"logs"
    ASSET_DIR=BASE/"assets"
else:
    BASE=RUNTIME_PATHS.app_dir
    BOT_SCRIPT=RUNTIME_PATHS.bot_script
    BOT_PYTHON=RUNTIME_PATHS.source_python
    LOG_DIR=RUNTIME_PATHS.logs_dir
    ASSET_DIR=RUNTIME_PATHS.assets_dir
EVENT_PREFIX="TVCBOT_EVENT "
STATS_INTERVAL_MS=1500
STOP_TIMEOUT_MS=300000
RUNTIME_CHECK_TIMEOUT_SECONDS=8
TVC_PROBE_TIMEOUT_SECONDS=4
TVC_PROBE_RESULT_PREFIX="TVC_PROBE_RESULT "
REQUIRED_BOT_MODULES=("openpyxl","pywinauto","psutil")
QUEUE_STATUS_LABELS={
    "PENDING":"รอตรวจสอบ",
    "CHECKING":"กำลังตรวจ",
    "READY":"READY",
    "RUNNING":"RUNNING",
    "DONE":"DONE",
    "ERROR":"ERROR",
    "INVALID":"INVALID",
    "REVIEW_REQUIRED":"SAFETY LOCK",
    "STOPPED":"STOPPED",
}
MASCOT_FILES={
    "READY":"bot_ready.png",
    "RUNNING":"bot_running.png",
    "SUCCESS":"bot_success.png",
    "ERROR":"bot_error.png",
}

psutil=None
get_job_stats=None
get_job_errors=None
get_safety_issues=None
inspect_recovery_state=None
reconcile_process_exit=None
validate_workbook=None
probe_tvc_client=None
DIRTY_TVC_FORM_POSSIBLE=(
    "DIRTY_TVC_FORM_POSSIBLE: กรุณาตรวจสอบ/ปิดหน้าใบงาน T.V.C ก่อนเริ่มใหม่"
)
SAFETY_METADATA_HEALTHY="HEALTHY"
SAFETY_METADATA_MISSING="MISSING"
SAFETY_METADATA_CORRUPT="CORRUPT"
SAFETY_METADATA_UNREADABLE="UNREADABLE"
SAFETY_METADATA_WRITE_FAILED="WRITE_FAILED"
SAFETY_METADATA_FAILURE_STATES={
    SAFETY_METADATA_CORRUPT,
    SAFETY_METADATA_UNREADABLE,
    SAFETY_METADATA_WRITE_FAILED,
}
SAFETY_PERSISTENCE_ERROR=(
    "ไม่สามารถบันทึกข้อมูล Safety Lock ได้ กรุณาแก้ไขก่อนเริ่มใหม่"
)
STOP_FILE_PREPARATION_ERROR="ไม่สามารถเตรียมไฟล์ควบคุมการหยุดได้"
STARTUP_ERROR_TITLE="T.V.C JOB BOT - Startup Error"


@dataclass(frozen=True)
class SafetyMetadataLoadResult:
    locks: dict
    health: str
    message: str=""


class StartupDependencyError(RuntimeError):
    def __init__(self,dependency,detail=""):
        self.dependency=dependency
        self.detail=str(detail or "").strip()
        if dependency=="excel_io":
            message="โหลด excel_io ไม่สำเร็จ กรุณาซ่อม environment"
        else:
            message=f"ไม่พบ {dependency} กรุณาติดตั้ง requirements หรือซ่อม environment"
        if self.detail:
            message=f"{message}\n\nรายละเอียด: {self.detail}"
        super().__init__(message)


def initialize_gui_dependencies(import_module=importlib.import_module):
    """Lazy-load third-party modules only after Tk root exists."""
    try:
        imported_psutil=import_module("psutil")
    except Exception as exc:
        raise StartupDependencyError("psutil",exc) from exc

    try:
        import_module("openpyxl")
    except Exception as exc:
        raise StartupDependencyError("openpyxl",exc) from exc

    try:
        excel_module=import_module("excel_io")
        imported_get_job_stats=excel_module.get_job_stats
        imported_get_job_errors=excel_module.get_job_errors
        imported_get_safety_issues=excel_module.get_safety_issues
        imported_inspect_recovery=excel_module.inspect_recovery_state
        imported_reconcile=excel_module.reconcile_process_exit
        imported_validate=excel_module.validate_workbook
        imported_dirty_marker=excel_module.DIRTY_TVC_FORM_POSSIBLE
    except Exception as exc:
        raise StartupDependencyError("excel_io",exc) from exc

    try:
        probe_module=import_module("tvc_probe")
        imported_probe=probe_module.probe
    except Exception as exc:
        raise StartupDependencyError("pywinauto",exc) from exc

    global psutil,get_job_stats,get_job_errors,get_safety_issues
    global inspect_recovery_state,reconcile_process_exit,validate_workbook
    global probe_tvc_client,DIRTY_TVC_FORM_POSSIBLE
    psutil=imported_psutil
    get_job_stats=imported_get_job_stats
    get_job_errors=imported_get_job_errors
    get_safety_issues=imported_get_safety_issues
    inspect_recovery_state=imported_inspect_recovery
    reconcile_process_exit=imported_reconcile
    validate_workbook=imported_validate
    probe_tvc_client=imported_probe
    DIRTY_TVC_FORM_POSSIBLE=imported_dirty_marker
    return True


def validate_bot_runtime(python_exe: Path | None=None,runtime_paths=None):
    """Validate source Python or the adjacent frozen worker executable."""
    paths=runtime_paths or RUNTIME_PATHS
    if paths is None:
        raise RuntimeError(RUNTIME_PATH_ERROR or "ไม่พบ runtime paths")
    if paths.frozen and python_exe is None:
        worker=Path(paths.worker_executable or "")
        if not worker.is_file():
            raise RuntimeError(f"ไม่พบ Worker executable: {worker}")
        if worker.stat().st_size<=0:
            raise RuntimeError(f"Worker executable เสียหรือว่างเปล่า: {worker}")
        try:
            check=subprocess.run(
                [str(worker),"--help"],
                cwd=str(paths.app_dir),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=RUNTIME_CHECK_TIMEOUT_SECONDS,
                creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0),
            )
        except (OSError,subprocess.SubprocessError) as exc:
            raise RuntimeError(f"เรียก Worker executable ไม่สำเร็จ: {exc}") from exc
        if check.returncode!=0:
            detail=(check.stderr or check.stdout or "unknown error").strip()
            raise RuntimeError(f"Worker executable ใช้งานไม่ได้: {detail}")
        return worker.resolve()

    python_exe=Path(python_exe or paths.source_python or "")
    if python_exe.name.lower()!="python.exe":
        raise RuntimeError("Bot subprocess ต้องใช้ python.exe เท่านั้น ห้ามใช้ pythonw.exe")
    if not python_exe.is_file():
        raise RuntimeError(f"ไม่พบ Python สำหรับ Bot: {python_exe}")
    if python_exe.stat().st_size<=0:
        raise RuntimeError(f"ไฟล์ Python สำหรับ Bot เสียหรือว่างเปล่า: {python_exe}")

    common_options={
        "cwd":str(paths.app_dir),
        "capture_output":True,
        "text":True,
        "encoding":"utf-8",
        "errors":"replace",
        "timeout":RUNTIME_CHECK_TIMEOUT_SECONDS,
        "creationflags":getattr(subprocess,"CREATE_NO_WINDOW",0),
    }
    try:
        version_check=subprocess.run([str(python_exe),"--version"],**common_options)
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            f"Runtime check timeout ขณะเรียก python.exe --version ({RUNTIME_CHECK_TIMEOUT_SECONDS} วินาที)"
        ) from exc
    except (OSError,subprocess.SubprocessError) as exc:
        raise RuntimeError(f"เรียก Python สำหรับ Bot ไม่สำเร็จ: {exc}") from exc
    if version_check.returncode!=0:
        detail=(version_check.stderr or version_check.stdout or "unknown error").strip()
        raise RuntimeError(f"Python สำหรับ Bot ใช้งานไม่ได้: {detail}")

    import_statement="import "+", ".join(REQUIRED_BOT_MODULES)
    try:
        dependency_check=subprocess.run(
            [str(python_exe),"-c",import_statement],
            **common_options,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            f"Runtime check timeout ขณะตรวจ dependencies ({RUNTIME_CHECK_TIMEOUT_SECONDS} วินาที)"
        ) from exc
    except (OSError,subprocess.SubprocessError) as exc:
        raise RuntimeError(f"ตรวจ dependencies ของ Bot ไม่สำเร็จ: {exc}") from exc
    if dependency_check.returncode!=0:
        detail=(dependency_check.stderr or dependency_check.stdout or "unknown error").strip()
        raise RuntimeError(
            "Python environment ขาด dependencies ที่จำเป็น "
            f"({', '.join(REQUIRED_BOT_MODULES)}): {detail}"
        )
    return python_exe.resolve()


def _probe_error_result(status,message,duration_ms=0):
    return {
        "status":status,
        "connected":False,
        "login_verified":False,
        "active_job_form":False,
        "active_job_forms":[],
        "matches":[],
        "login_signals":[],
        "errors":[str(message)],
        "timed_out":status=="TIMEOUT",
        "duration_ms":int(duration_ms or 0),
    }


def build_tvc_probe_command(runtime_executable: Path,runtime_paths=None):
    """Build an isolated probe command for source and frozen deployments."""
    paths=runtime_paths or RUNTIME_PATHS
    if paths is None:
        raise RuntimeError(RUNTIME_PATH_ERROR or "ไม่พบ runtime paths")
    executable=Path(runtime_executable)
    if paths.frozen:
        return [str(executable),"--probe-tvc"]
    probe_script=paths.app_dir/"src"/"tvc_probe.py"
    if not probe_script.is_file():
        raise FileNotFoundError(f"ไม่พบ T.V.C probe script: {probe_script}")
    return [str(executable),"-u",str(probe_script)]


def check_tvc_client(
    runtime_executable: Path,
    *,
    runtime_paths=None,
    timeout_seconds=TVC_PROBE_TIMEOUT_SECONDS,
    runner=None,
):
    """Run the read-only probe in a killable subprocess with a hard timeout."""
    runner=runner or subprocess.run
    paths=runtime_paths or RUNTIME_PATHS
    command=build_tvc_probe_command(runtime_executable,paths)
    started=time.monotonic()
    try:
        completed=runner(
            command,
            cwd=str(paths.app_dir),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
            creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0),
        )
    except subprocess.TimeoutExpired:
        duration_ms=round((time.monotonic()-started)*1000)
        return _probe_error_result(
            "TIMEOUT",
            f"T.V.C probe timeout หลัง {timeout_seconds} วินาที",
            duration_ms,
        )
    except (OSError,subprocess.SubprocessError) as exc:
        duration_ms=round((time.monotonic()-started)*1000)
        return _probe_error_result(
            "ERROR",f"เรียก T.V.C probe ไม่สำเร็จ: {exc}",duration_ms
        )

    payload=None
    for line in reversed((completed.stdout or "").splitlines()):
        if line.startswith(TVC_PROBE_RESULT_PREFIX):
            payload=line[len(TVC_PROBE_RESULT_PREFIX):]
            break
    if payload is None:
        detail=(completed.stderr or completed.stdout or "ไม่มีผลลัพธ์จาก probe").strip()
        return _probe_error_result("ERROR",detail)
    try:
        result=json.loads(payload)
    except (TypeError,json.JSONDecodeError) as exc:
        return _probe_error_result("ERROR",f"อ่านผล T.V.C probe ไม่สำเร็จ: {exc}")
    connected=bool(result.get("connected"))
    login_verified=bool(result.get("login_verified"))
    result.setdefault(
        "status",
        "READY" if connected and login_verified else "FOUND" if connected else "NOT_FOUND",
    )
    result.setdefault("timed_out",result.get("status")=="TIMEOUT")
    result.setdefault("errors",[])
    result.setdefault("duration_ms",round((time.monotonic()-started)*1000))
    return result


def check_excel_access(path: Path):
    """Best-effort, non-mutating check that the workbook can be opened for Bot save."""
    path=Path(path)
    with path.open("r+b"):
        pass
    return True


def load_photo_asset(path,master=None,image_factory=None):
    """Load an optional Tk image and return None for missing/invalid assets."""
    path=Path(path)
    if not path.is_file():
        return None
    factory=image_factory or tk.PhotoImage
    try:
        return factory(file=str(path),master=master)
    except (OSError,tk.TclError,ValueError):
        return None


def apply_optional_window_icon(root,icon_path=ASSET_DIR/"app_icon.ico"):
    try:
        if Path(icon_path).is_file():
            root.iconbitmap(default=str(icon_path))
            return True
    except (OSError,tk.TclError):
        pass
    return False


def _uncertain_issues_from_errors(errors):
    issues=[]
    for error in errors:
        result=str(error.get("bot_result") or "")
        if "UNCERTAIN_TVC_SAVE" not in result.upper():
            continue
        issues.append({
            "reason":"UNCERTAIN_TVC_SAVE",
            "job_ref":str(error.get("job_ref") or ""),
            "bot_status":"ERROR",
            "bot_result":result,
            "message":(
                "พบงานที่ไม่ยืนยันผลการบันทึก T.V.C "
                "กรุณาตรวจสอบก่อนเริ่มใหม่"
            ),
        })
    return issues


def perform_precheck(
    paths,
    runtime_validator=None,
    tvc_checker=None,
    workbook_validator=None,
    access_checker=None,
    stats_reader=None,
    errors_reader=None,
    safety_reader=None,
):
    """Run every blocking pre-check without touching Tk or automating T.V.C."""
    runtime_validator=runtime_validator or validate_bot_runtime
    tvc_checker=tvc_checker or check_tvc_client
    workbook_validator=workbook_validator or validate_workbook
    access_checker=access_checker or check_excel_access
    stats_reader=stats_reader or get_job_stats
    custom_errors_reader=errors_reader is not None
    errors_reader=errors_reader or get_job_errors
    if safety_reader is None and not custom_errors_reader:
        safety_reader=get_safety_issues
    result={
        "runtime":{"ready":False,"python":None,"message":""},
        "tvc":{
            "status":"PENDING",
            "ready":False,
            "connected":False,
            "login_verified":False,
            "active_job_form":False,
            "active_job_forms":[],
            "login_signals":[],
            "message":"",
        },
        "items":[],
        "queue_ready":False,
        "total_wait":0,
        "ready":False,
    }

    python_exe=None
    try:
        python_exe=runtime_validator()
        result["runtime"]={
            "ready":True,
            "python":str(python_exe),
            "message":f"Ready: {python_exe}",
        }
    except Exception as exc:
        result["runtime"]["message"]=str(exc)

    if python_exe is not None:
        try:
            probe=tvc_checker(Path(python_exe))
            probe_status=str(probe.get("status") or "").upper()
            connected=bool(probe.get("connected"))
            login_verified=bool(probe.get("login_verified"))
            if not probe_status:
                probe_status=(
                    "READY" if connected and login_verified
                    else "FOUND" if connected
                    else "NOT_FOUND"
                )
            tvc_ready=connected and login_verified and probe_status=="READY"
            if probe_status=="TIMEOUT":
                message=(
                    f"T.V.C probe หมดเวลา ({probe.get('duration_ms',0)} ms) "
                    "กรุณาลองตรวจสอบใหม่"
                )
            elif probe_status=="ERROR":
                errors=list(probe.get("errors") or [])
                message=errors[0] if errors else "T.V.C probe เกิดข้อผิดพลาด"
            elif not connected:
                message="ไม่พบ T.V.C Client กรุณาเปิดและเข้าสู่ระบบก่อน"
            elif not login_verified:
                message="พบ T.V.C Client แต่ยังยืนยันการเข้าสู่ระบบไม่ได้"
            else:
                message="Connected"
            result["tvc"]={
                "ready":tvc_ready,
                "connected":connected,
                "login_verified":login_verified,
                "active_job_form":bool(probe.get("active_job_form")),
                "active_job_forms":list(probe.get("active_job_forms") or []),
                "message":message,
                "matches":list(probe.get("matches") or []),
                "login_signals":list(probe.get("login_signals") or []),
                "status":probe_status,
                "timed_out":bool(probe.get("timed_out")),
                "errors":list(probe.get("errors") or []),
                "duration_ms":int(probe.get("duration_ms",0) or 0),
            }
        except Exception as exc:
            result["tvc"]["status"]="ERROR"
            result["tvc"]["message"]=str(exc)
    else:
        result["tvc"]["status"]="SKIPPED"
        result["tvc"]["message"]="ยังไม่ได้ตรวจ T.V.C เพราะ Runtime ไม่พร้อม"

    for raw_path in paths:
        path=normalize_excel_path(raw_path)
        item_result={
            "path":str(path),
            "status":"INVALID",
            "stats":{},
            "errors":[],
            "safety_issues":[],
            "message":"",
        }
        try:
            workbook_validator(path)
            access_checker(path)
            stats=stats_reader(path)
            errors=errors_reader(path)
            issues=(
                safety_reader(path)
                if safety_reader is not None
                else _uncertain_issues_from_errors(errors)
            )
            if int(stats.get("RUNNING",0) or 0)>0:
                raise RuntimeError(
                    f"พบ JOB RUNNING {stats['RUNNING']} รายการ ต้องตรวจ Excel ก่อนเริ่มใหม่"
                )
            result["total_wait"]+=int(stats.get("WAIT",0) or 0)
            item_result.update(
                status="REVIEW_REQUIRED" if issues else "READY",
                stats=dict(stats),
                errors=list(errors),
                safety_issues=list(issues),
                message=(issues[0].get("message") if issues else "READY"),
            )
        except Exception as exc:
            item_result["message"]=str(exc)
        result["items"].append(item_result)

    result["queue_ready"]=(
        bool(result["items"])
        and all(item["status"]=="READY" for item in result["items"])
        and result["total_wait"]>0
    )
    if not result["items"]:
        result["queue_message"]="ยังไม่ได้เลือก Excel"
    elif any(item["status"]=="REVIEW_REQUIRED" for item in result["items"]):
        result["queue_message"]="พบ Safety Lock กรุณาตรวจสอบงานก่อนเริ่มใหม่"
    elif not all(item["status"]=="READY" for item in result["items"]):
        result["queue_message"]="มีไฟล์ INVALID กรุณาแก้ไขหรือลบออก"
    elif result["total_wait"]<=0:
        result["queue_message"]="Queue ต้องมี JOB สถานะ WAIT อย่างน้อย 1 รายการ"
    else:
        result["queue_message"]=f"Ready | WAIT รวม {result['total_wait']}"
    result["ready"]=(
        result["runtime"]["ready"]
        and result["tvc"]["ready"]
        and result["queue_ready"]
    )
    return result


def format_queue_summary(items,outcome):
    summary=build_queue_summary(items)
    title={
        "COMPLETE_SUCCESS":"Run Complete",
        "COMPLETE_WITH_ERRORS":"Completed with Errors",
        "STOPPED":"Queue Stopped",
        "INCOMPLETE":"Queue Incomplete",
        "FAILED":"Queue Failed",
        "COMPLETE":"Run Complete",
        "ERROR":"Queue Error",
    }.get(outcome,f"Queue {outcome.title()}")
    lines=[
        title,
        "",
        f"Files: {summary['files_completed']}/{summary['files_total']} completed",
        f"Files Error: {summary['files_error']}",
        "Jobs:",
        f"DONE: {summary['DONE']}",
        f"ERROR: {summary['ERROR']}",
        f"WAIT: {summary['WAIT']}",
    ]
    if summary["error_files"] or summary["problems"]:
        lines.extend(["","Problem:"])
        lines.extend(summary["error_files"])
        lines.extend(summary["problems"])
    return "\n".join(lines),summary


def determine_final_outcome(requested_outcome,summary,refresh_failed=False):
    """Map an execution stop reason plus fresh Excel totals to a user-facing outcome."""
    if refresh_failed:
        return "FAILED"
    if requested_outcome=="STOPPED":
        return "STOPPED"
    if requested_outcome in {"ERROR","FAILED"}:
        return "FAILED"
    if int(summary.get("ERROR",0) or 0)>0 or int(summary.get("files_error",0) or 0)>0:
        return "COMPLETE_WITH_ERRORS"
    unfinished=sum(
        int(summary.get(key,0) or 0)
        for key in ("WAIT","RUNNING","OTHER")
    )
    if unfinished>0:
        return "INCOMPLETE"
    return "COMPLETE_SUCCESS"


def evaluate_safety_revalidation(
    precheck_result,locks,inspections,current_keys,include_resolved_keys=False
):
    """Evaluate each persisted lock independently from current Queue membership."""
    runtime_ready=bool(precheck_result.get("runtime",{}).get("ready"))
    tvc_ready=bool(precheck_result.get("tvc",{}).get("ready"))
    active_job_form=bool(precheck_result.get("tvc",{}).get("active_job_form"))
    if not (runtime_ready and tvc_ready):
        result=(False,["read-only Runtime/T.V.C validation ยังไม่ผ่าน"])
        return (*result,tuple()) if include_resolved_keys else result
    unresolved=[]
    resolved_keys=[]
    for key,lock in locks.items():
        reason=str(lock.get("outcome") or lock.get("reason") or "")
        dirty_form_lock=reason=="DIRTY_TVC_FORM_POSSIBLE"
        if dirty_form_lock and active_job_form:
            unresolved.append(
                "ยังพบหน้าเพิ่มใบงาน/active JOB form ใน T.V.C กรุณาปิดหรือยกเลิกก่อน"
            )
            continue
        inspection=inspections.get(key)
        if not inspection:
            unresolved.append(f"{Path(key).name}: ไม่มีผลตรวจ recovery")
            continue
        safety_issues=list(inspection.get("safety_issues") or [])
        if safety_issues:
            unresolved.append(
                f"{Path(key).name}: ยังมี UNCERTAIN_TVC_SAVE ใน workbook"
            )
            continue
        running_count=int(inspection.get("running_count",0) or 0)
        outcome=str(inspection.get("outcome") or "")
        clean=(
            outcome=="already_clean"
            and bool(inspection.get("verified"))
            and running_count==0
        )
        if not clean:
            unresolved.append(
                f"{Path(key).name}: {inspection.get('message') or lock.get('message') or outcome}"
            )
            continue
        resolved_keys.append(key)
    result=(not unresolved,unresolved)
    return (*result,tuple(resolved_keys)) if include_resolved_keys else result


def load_persisted_safety_locks(path):
    """Load metadata with an explicit health state; existing failures fail closed."""
    path=Path(path)
    try:
        metadata_stat=path.stat()
        if not stat.S_ISREG(metadata_stat.st_mode):
            return SafetyMetadataLoadResult(
                {},SAFETY_METADATA_CORRUPT,f"Safety metadata ไม่ใช่ไฟล์: {path}"
            )
        raw=path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return SafetyMetadataLoadResult({},SAFETY_METADATA_MISSING)
    except UnicodeError as exc:
        return SafetyMetadataLoadResult(
            {},SAFETY_METADATA_CORRUPT,
            f"Safety metadata encoding ไม่ถูกต้องหรือไฟล์เสีย: {exc}",
        )
    except OSError as exc:
        return SafetyMetadataLoadResult(
            {},SAFETY_METADATA_UNREADABLE,f"อ่าน Safety metadata ไม่สำเร็จ: {exc}"
        )
    try:
        payload=json.loads(raw)
    except (ValueError,TypeError) as exc:
        return SafetyMetadataLoadResult(
            {},SAFETY_METADATA_CORRUPT,f"Safety metadata เสียหรือ JSON ไม่ถูกต้อง: {exc}"
        )
    if not isinstance(payload,dict):
        return SafetyMetadataLoadResult(
            {},SAFETY_METADATA_CORRUPT,"Safety metadata ต้องเป็น JSON object"
        )
    locks={}
    for key,value in payload.items():
        if not isinstance(value,dict) or not value.get("path"):
            return SafetyMetadataLoadResult(
                {},SAFETY_METADATA_CORRUPT,f"Safety metadata entry ไม่ถูกต้อง: {key}"
            )
        try:
            normalized_path=normalize_excel_path(value["path"])
        except (OSError,TypeError,ValueError) as exc:
            return SafetyMetadataLoadResult(
                {},SAFETY_METADATA_CORRUPT,
                f"Safety metadata path ไม่ถูกต้อง ({key}): {exc}",
            )
        normalized=os.path.normcase(str(normalized_path))
        locks[normalized]={
            "path":str(normalized_path),
            "outcome":str(value.get("outcome") or "failed"),
            "message":str(value.get("message") or "ต้องตรวจสอบงาน"),
            "job_ref":str(value.get("job_ref") or ""),
        }
    return SafetyMetadataLoadResult(locks,SAFETY_METADATA_HEALTHY)


def save_persisted_safety_locks(path,locks):
    """Persist lock metadata with flush/fsync and atomic replacement."""
    path=Path(path)
    path.parent.mkdir(parents=True,exist_ok=True)
    if not locks:
        path.unlink(missing_ok=True)
        return
    temporary=path.with_name(
        f".{path.name}.{os.getpid()}.{uuid.uuid4().hex}.tmp"
    )
    try:
        with temporary.open("w",encoding="utf-8",newline="\n") as stream:
            json.dump(locks,stream,ensure_ascii=False,indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(path)
    except Exception:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def new_stop_file_path(runtime_paths):
    """Allocate a collision-resistant cooperative-stop path under writable data."""
    runtime_dir=Path(runtime_paths.runtime_temp_dir)
    runtime_dir.mkdir(parents=True,exist_ok=True)
    return runtime_dir/f"tvc_bot_stop_{os.getpid()}_{uuid.uuid4().hex}.flag"


def prepare_worker_launch(runtime_paths,excel_path,executable):
    """Prepare all filesystem/argv prerequisites before entering RUNNING state."""
    stop_file=new_stop_file_path(runtime_paths)
    stop_file.unlink(missing_ok=True)
    command=build_worker_command(
        runtime_paths,
        excel_path,
        stop_file,
        executable=executable,
    )
    return stop_file,command


def write_startup_error_log(exc):
    """Best-effort traceback logging that never hides the original startup error."""
    if RUNTIME_PATHS is None:
        return
    try:
        RUNTIME_PATHS.logs_dir.mkdir(parents=True,exist_ok=True)
        log_file=RUNTIME_PATHS.logs_dir/"gui_startup_error.log"
        detail="".join(traceback.format_exception(type(exc),exc,exc.__traceback__))
        with log_file.open("a",encoding="utf-8",newline="\n") as stream:
            stream.write(detail)
            if not detail.endswith("\n"):
                stream.write("\n")
    except Exception:
        pass


def show_startup_error(root,exc):
    """Show a visible startup failure, destroy any Tk root, and never re-raise."""
    write_startup_error_log(exc)
    dialog_root=root
    if dialog_root is None:
        try:
            dialog_root=tk.Tk()
            dialog_root.withdraw()
        except Exception:
            dialog_root=None
    try:
        if dialog_root is not None:
            dialog_root.deiconify()
            dialog_root.update_idletasks()
        messagebox.showerror(
            STARTUP_ERROR_TITLE,
            f"ไม่สามารถเปิด T.V.C JOB BOT ได้\n\nสาเหตุ: {exc}",
            parent=dialog_root,
        )
    except Exception:
        pass
    finally:
        if dialog_root is not None:
            try:
                dialog_root.destroy()
            except Exception:
                pass


class TVCControlApp:
    def __init__(self,root: tk.Tk):
        self.root=root
        self.runtime_paths=RUNTIME_PATHS
        self.safety_state_file=(
            self.runtime_paths.writable_data_dir/"safety_locks.json"
        )
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        self.root.geometry("1120x840")
        self.root.minsize(960,720)
        apply_optional_window_icon(self.root)

        self.process: subprocess.Popen | None=None
        self.stop_file: Path | None=None
        self.events: queue.Queue=queue.Queue()
        self.excel_queue=ExcelQueue()
        self.run_controller: QueueRunController | None=None
        self.current_file_index=-1
        self.queue_revision=0
        self.queue_running=False
        self.precheck_valid=False
        self.precheck_generation=0
        self.precheck_in_progress=False
        self.precheck_purpose=""
        self.start_pending=False
        self.finalization_in_progress=False
        metadata=load_persisted_safety_locks(self.safety_state_file)
        self.safety_locks=metadata.locks
        self.safety_metadata_health=metadata.health
        self.safety_metadata_error=metadata.message
        self.tvc_connected=False
        self.tvc_login_verified=False
        self.tvc_error=""
        self.asset_images={}
        self.stats_loading=False
        self.valid_excel=False
        self.stop_request_sent=False
        self.stop_event_seen=False
        self.last_batch_success: bool | None=None
        self.current_job_ref=""
        self.force_stop_used=False
        self.force_retry_available=False
        self.recovery_in_progress=False
        self.recovery_failed=(
            bool(self.safety_locks)
            or self.safety_metadata_health in SAFETY_METADATA_FAILURE_STATES
        )
        self.stop_event_phase=""
        self.last_return_code: int | None=None
        self.closing=False
        self.runtime_check_in_progress=False
        self.runtime_valid=False
        self.runtime_error=""
        self.bot_python: Path | None=None

        self.excel_var=tk.StringVar()
        self.status_var=tk.StringVar(value="กำลังตรวจสอบระบบ...")
        self.runtime_status_var=tk.StringVar(value="กำลังตรวจสอบ")
        self.tvc_status_var=tk.StringVar(value="กำลังตรวจสอบ")
        self.queue_status_var=tk.StringVar(value="ยังไม่มีไฟล์")
        self.safety_status_var=tk.StringVar(
            value=(
                f"FAIL_CLOSED ({self.safety_metadata_health})"
                if self.safety_metadata_health in SAFETY_METADATA_FAILURE_STATES
                else f"LOCKED ({len(self.safety_locks)})"
                if self.safety_locks
                else "ไม่ล็อก"
            )
        )
        self.safety_detail_var=tk.StringVar(value="")
        self.wait_var=tk.StringVar(value="0")
        self.done_var=tk.StringVar(value="0")
        self.error_var=tk.StringVar(value="0")
        self.current_file_var=tk.StringVar(value="- / 0")
        self.current_job_var=tk.StringVar(value="- / 0")
        self.progress_text_var=tk.StringVar(value="0 / 0")
        self.overall_progress_text_var=tk.StringVar(value="0 / 0")

        self._configure_styles()
        self._build_ui()
        self._refresh_safety_status()
        LOG_DIR.mkdir(parents=True,exist_ok=True)
        self._load_default_excel()

        self.root.protocol("WM_DELETE_WINDOW",self.on_close)
        self.root.after(100,self._drain_events)
        self.root.after(STATS_INTERVAL_MS,self._stats_tick)
        self.root.after(0,self.retry_precheck)

    def _configure_styles(self):
        style=ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Header.TLabel",font=("Segoe UI",22,"bold"))
        style.configure("Subheader.TLabel",font=("Segoe UI",10))
        style.configure("Status.TLabel",font=("Segoe UI",13,"bold"))
        style.configure("MetricTitle.TLabel",font=("Segoe UI",10))
        style.configure("MetricValue.TLabel",font=("Segoe UI",20,"bold"))
        style.configure("Action.TButton",font=("Segoe UI",10,"bold"),padding=(14,8))

    def _build_ui(self):
        container=ttk.Frame(self.root,padding=16)
        container.pack(fill="both",expand=True)
        container.columnconfigure(0,weight=1)
        container.rowconfigure(2,weight=1)
        container.rowconfigure(6,weight=2)

        header=ttk.Frame(container)
        header.grid(row=0,column=0,sticky="ew",pady=(0,10))
        header.columnconfigure(1,weight=1)
        self.mascot_label=ttk.Label(
            header,text="BOT",anchor="center",font=("Segoe UI",15,"bold"),width=8
        )
        self.mascot_label.grid(row=0,column=0,rowspan=2,sticky="nsw",padx=(0,12))
        ttk.Label(header,text=APP_NAME,style="Header.TLabel").grid(row=0,column=1,sticky="sw")
        ttk.Label(header,text="Automation Control",style="Subheader.TLabel").grid(
            row=1,column=1,sticky="nw"
        )
        ttk.Label(header,text=f"v{APP_VERSION}",font=("Segoe UI",12,"bold")).grid(
            row=0,column=2,rowspan=2,sticky="ne"
        )
        self._set_mascot("READY")

        precheck=ttk.LabelFrame(container,text="Connection / Pre-check",padding=9)
        precheck.grid(row=1,column=0,sticky="ew",pady=(0,9))
        for column in range(4):
            precheck.columnconfigure(column,weight=1)
        ttk.Label(precheck,text="T.V.C Status:").grid(row=0,column=0,sticky="w")
        ttk.Label(precheck,textvariable=self.tvc_status_var,font=("Segoe UI",10,"bold")).grid(
            row=1,column=0,sticky="w"
        )
        ttk.Label(precheck,text="Runtime:").grid(row=0,column=1,sticky="w")
        ttk.Label(precheck,textvariable=self.runtime_status_var,font=("Segoe UI",10,"bold")).grid(
            row=1,column=1,sticky="w"
        )
        ttk.Label(precheck,text="Queue:").grid(row=0,column=2,sticky="w")
        ttk.Label(precheck,textvariable=self.queue_status_var,font=("Segoe UI",10,"bold")).grid(
            row=1,column=2,sticky="w"
        )
        ttk.Label(precheck,text="Safety Lock:").grid(row=0,column=3,sticky="w")
        ttk.Label(precheck,textvariable=self.safety_status_var,font=("Segoe UI",10,"bold")).grid(
            row=1,column=3,sticky="w"
        )
        ttk.Label(
            precheck,
            textvariable=self.safety_detail_var,
            justify="left",
            wraplength=1040,
        ).grid(row=2,column=0,columnspan=4,sticky="ew",pady=(6,0))

        queue_frame=ttk.LabelFrame(container,text="Excel Queue",padding=8)
        queue_frame.grid(row=2,column=0,sticky="nsew",pady=(0,8))
        queue_frame.columnconfigure(0,weight=1)
        queue_frame.rowconfigure(0,weight=1)
        columns=("path","status","wait","done","error")
        self.queue_tree=ttk.Treeview(
            queue_frame,columns=columns,show="headings",height=6,selectmode="extended"
        )
        self.queue_tree.heading("path",text="Excel file")
        self.queue_tree.heading("status",text="Status")
        self.queue_tree.heading("wait",text="WAIT")
        self.queue_tree.heading("done",text="DONE")
        self.queue_tree.heading("error",text="ERROR")
        self.queue_tree.column("path",width=580,anchor="w")
        self.queue_tree.column("status",width=100,anchor="center")
        for name in ("wait","done","error"):
            self.queue_tree.column(name,width=75,anchor="center",stretch=False)
        self.queue_tree.grid(row=0,column=0,sticky="nsew")
        queue_scroll=ttk.Scrollbar(queue_frame,orient="vertical",command=self.queue_tree.yview)
        queue_scroll.grid(row=0,column=1,sticky="ns")
        self.queue_tree.configure(yscrollcommand=queue_scroll.set)
        self.queue_tree.bind("<<TreeviewSelect>>",self._on_queue_selection)

        controls=ttk.Frame(container)
        controls.grid(row=3,column=0,sticky="ew",pady=(0,8))
        self.browse_button=ttk.Button(controls,text="+ เพิ่ม Excel",command=self.choose_excel)
        self.browse_button.pack(side="left",padx=(0,6))
        self.remove_button=ttk.Button(controls,text="ลบไฟล์",command=self.remove_selected_files)
        self.remove_button.pack(side="left",padx=(0,6))
        self.clear_button=ttk.Button(controls,text="ล้างรายการ",command=self.clear_files)
        self.clear_button.pack(side="left",padx=(0,6))
        self.up_button=ttk.Button(controls,text="ขึ้น",command=lambda:self.move_selected_file(-1))
        self.up_button.pack(side="left",padx=(0,4))
        self.down_button=ttk.Button(controls,text="ลง",command=lambda:self.move_selected_file(1))
        self.down_button.pack(side="left",padx=(0,10))
        self.runtime_button=ttk.Button(
            controls,text="ตรวจสอบระบบ",command=self.retry_precheck
        )
        self.runtime_button.pack(side="left",padx=(0,6))
        self.safety_button=ttk.Button(
            controls,
            text="ตรวจสอบหลังแก้ไข",
            command=self.retry_safety_validation,
            state="disabled",
        )
        self.safety_button.pack(side="left")

        progress_frame=ttk.LabelFrame(container,text="Progress",padding=9)
        progress_frame.grid(row=4,column=0,sticky="ew",pady=(0,8))
        progress_frame.columnconfigure(1,weight=1)
        progress_frame.columnconfigure(3,weight=1)
        ttk.Label(progress_frame,text="สถานะ Bot:").grid(row=0,column=0,sticky="w")
        self.status_label=ttk.Label(progress_frame,textvariable=self.status_var,style="Status.TLabel")
        self.status_label.grid(row=0,column=1,sticky="w")
        ttk.Label(progress_frame,text="Current File:").grid(row=0,column=2,sticky="e",padx=(12,5))
        ttk.Label(progress_frame,textvariable=self.current_file_var,font=("Segoe UI",11,"bold")).grid(
            row=0,column=3,sticky="w"
        )
        ttk.Label(progress_frame,text="Current JOB:").grid(row=1,column=0,sticky="w",pady=(6,0))
        ttk.Label(progress_frame,textvariable=self.current_job_var,font=("Segoe UI",11,"bold")).grid(
            row=1,column=1,sticky="w",pady=(6,0)
        )
        ttk.Label(progress_frame,text="File Progress").grid(row=2,column=0,sticky="w",pady=(7,0))
        self.progress=ttk.Progressbar(progress_frame,mode="determinate",maximum=1,value=0)
        self.progress.grid(row=2,column=1,columnspan=2,sticky="ew",padx=(8,8),pady=(7,0))
        ttk.Label(progress_frame,textvariable=self.progress_text_var,font=("Segoe UI",10,"bold")).grid(
            row=2,column=3,sticky="w",pady=(7,0)
        )
        ttk.Label(progress_frame,text="Overall Queue").grid(row=3,column=0,sticky="w",pady=(6,0))
        self.overall_progress=ttk.Progressbar(
            progress_frame,mode="determinate",maximum=1,value=0
        )
        self.overall_progress.grid(row=3,column=1,columnspan=2,sticky="ew",padx=(8,8),pady=(6,0))
        ttk.Label(
            progress_frame,textvariable=self.overall_progress_text_var,font=("Segoe UI",10,"bold")
        ).grid(row=3,column=3,sticky="w",pady=(6,0))

        metrics_actions=ttk.Frame(container)
        metrics_actions.grid(row=5,column=0,sticky="ew",pady=(0,8))
        metrics=ttk.Frame(metrics_actions)
        metrics.pack(side="left",fill="x",expand=True)
        for column in range(3):
            metrics.columnconfigure(column,weight=1)
        self._metric_card(metrics,0,"WAIT",self.wait_var)
        self._metric_card(metrics,1,"DONE",self.done_var)
        self._metric_card(metrics,2,"ERROR",self.error_var)

        actions=ttk.Frame(metrics_actions)
        actions.pack(side="right",padx=(12,0))
        self.start_button=ttk.Button(
            actions,text="เริ่ม Bot",style="Action.TButton",command=self.start_bot,state="disabled"
        )
        self.start_button.pack(side="left",padx=(0,8))
        self.stop_button=ttk.Button(
            actions,text="หยุดหลังจบ JOB",style="Action.TButton",command=self.stop_bot,state="disabled"
        )
        self.stop_button.pack(side="left",padx=(0,8))
        ttk.Button(
            actions,text="เปิด Logs",style="Action.TButton",command=self.open_logs
        ).pack(side="left")

        log_frame=ttk.LabelFrame(container,text="Bot Log (real-time)",padding=8)
        log_frame.grid(row=6,column=0,sticky="nsew")
        log_frame.columnconfigure(0,weight=1)
        log_frame.rowconfigure(0,weight=1)
        self.log_text=scrolledtext.ScrolledText(
            log_frame,
            wrap="word",
            state="disabled",
            font=("Consolas",10),
            background="#111827",
            foreground="#e5e7eb",
            insertbackground="#e5e7eb",
        )
        self.log_text.grid(row=0,column=0,sticky="nsew")

    def _set_mascot(self,state):
        state=state if state in MASCOT_FILES else "READY"
        image=self.asset_images.get(state)
        if state not in self.asset_images:
            image=load_photo_asset(ASSET_DIR/MASCOT_FILES[state],master=self.root)
            self.asset_images[state]=image
        if image is None:
            self.mascot_label.configure(image="",text="BOT")
        else:
            self.mascot_label.configure(image=image,text="")

    def _metric_card(self,parent,column,title,variable):
        box=ttk.LabelFrame(parent,text=title,padding=10)
        left=0 if column==0 else 6
        right=6 if column<2 else 0
        box.grid(row=0,column=column,sticky="nsew",padx=(left,right))
        ttk.Label(box,text=title,style="MetricTitle.TLabel").pack()
        ttk.Label(box,textvariable=variable,style="MetricValue.TLabel").pack()

    def _load_default_excel(self):
        try:
            cfg=configparser.ConfigParser()
            cfg.read(self.runtime_paths.config_file,encoding="utf-8")
            configured=Path(cfg.get("excel","file"))
            path=configured if configured.is_absolute() else BASE/configured
            if path.is_file():
                self.excel_queue.add_paths([path])
                self.queue_revision+=1
                self.excel_var.set(str(normalize_excel_path(path)))
                self._refresh_queue_tree()
        except Exception as exc:
            self.append_log(f"ไม่สามารถโหลดไฟล์เริ่มต้นจาก config.ini: {exc}")

    def choose_excel(self):
        if not self._queue_edit_allowed():
            return
        initial=self.excel_var.get().strip()
        initial_dir=str(Path(initial).parent) if initial else str(BASE/"data")
        selected=filedialog.askopenfilenames(
            title="เพิ่มไฟล์ Excel เข้า Queue",
            initialdir=initial_dir,
            filetypes=[("Excel Workbook","*.xlsx")],
        )
        if not selected:
            return
        accepted=[]
        rejected=[]
        for raw_path in selected:
            path=normalize_excel_path(raw_path)
            if path.suffix.lower()==".xlsx":
                accepted.append(path)
            else:
                rejected.append(path)
        added,duplicates=self.excel_queue.add_paths(accepted)
        if added:
            self.excel_var.set(str(added[-1].path))
            self.append_log(f"เพิ่ม Excel {len(added)} ไฟล์เข้า Queue")
        if duplicates:
            self.append_log(
                "ข้าม path ซ้ำ: "+", ".join(str(path) for path in duplicates)
            )
        if rejected:
            messagebox.showerror(
                "ไฟล์ไม่ถูกต้อง",
                "รองรับเฉพาะไฟล์ .xlsx\n\n" + "\n".join(str(path) for path in rejected),
                parent=self.root,
            )
        self._invalidate_precheck()
        self._refresh_queue_tree()
        self.retry_precheck()

    def _queue_edit_allowed(self):
        return not (
            self.queue_running
            or self._is_running()
            or self.recovery_in_progress
            or self.start_pending
            or self.finalization_in_progress
            or self.excel_queue.locked
            or (
                self.precheck_in_progress
                and self.precheck_purpose in {"start","safety"}
            )
            or self.closing
        )

    def _safety_lock_active(self):
        return (
            bool(getattr(self,"safety_locks",{}))
            or self._metadata_fail_closed()
        )

    def _metadata_fail_closed(self):
        return getattr(
            self,"safety_metadata_health",SAFETY_METADATA_MISSING
        ) in SAFETY_METADATA_FAILURE_STATES

    def _set_safety_lock(self,path,outcome,message="",job_ref=""):
        if not hasattr(self,"safety_locks"):
            return
        path=normalize_excel_path(path)
        key=os.path.normcase(str(path))
        self.safety_locks[key]={
            "path":str(path),
            "outcome":str(outcome or "failed"),
            "message":str(message or outcome or "ต้องตรวจสอบงาน"),
            "job_ref":str(job_ref or ""),
        }
        self.recovery_failed=True
        self._persist_safety_locks()
        self._refresh_safety_status()

    def _persist_safety_locks(self,allow_repair=False):
        state_file=getattr(self,"safety_state_file",None)
        if state_file is None:
            return True
        health=getattr(self,"safety_metadata_health",SAFETY_METADATA_MISSING)
        if health in {SAFETY_METADATA_CORRUPT,SAFETY_METADATA_UNREADABLE} and not allow_repair:
            self.recovery_failed=True
            self._refresh_safety_status()
            if hasattr(self,"append_log"):
                self.append_log(
                    "ไม่ overwrite Safety metadata ที่เสีย/อ่านไม่ได้โดยอัตโนมัติ; "
                    "ให้แก้ไฟล์แล้วกดตรวจสอบหลังแก้ไข"
                )
            return False
        try:
            save_persisted_safety_locks(state_file,self.safety_locks)
        except Exception as exc:
            self.safety_metadata_health=SAFETY_METADATA_WRITE_FAILED
            self.safety_metadata_error=f"{SAFETY_PERSISTENCE_ERROR}: {exc}"
            self.recovery_failed=True
            if hasattr(self,"_set_status"):
                self._set_status(SAFETY_PERSISTENCE_ERROR)
            if hasattr(self,"append_log"):
                self.append_log(self.safety_metadata_error)
            self._refresh_safety_status()
            return False
        self.safety_metadata_health=(
            SAFETY_METADATA_HEALTHY
            if self.safety_locks else SAFETY_METADATA_MISSING
        )
        self.safety_metadata_error=""
        self._refresh_safety_status()
        return True

    def _retry_safety_metadata(self):
        """Explicit user action to retry a failed write or reload repaired metadata."""
        health=getattr(self,"safety_metadata_health",SAFETY_METADATA_MISSING)
        if health==SAFETY_METADATA_WRITE_FAILED:
            return self._persist_safety_locks(allow_repair=True)
        if health not in {SAFETY_METADATA_CORRUPT,SAFETY_METADATA_UNREADABLE}:
            return True

        loaded=load_persisted_safety_locks(self.safety_state_file)
        if loaded.health in SAFETY_METADATA_FAILURE_STATES:
            self.safety_metadata_health=loaded.health
            self.safety_metadata_error=loaded.message
            self.recovery_failed=True
            self._set_status(SAFETY_PERSISTENCE_ERROR)
            self.append_log(loaded.message or SAFETY_PERSISTENCE_ERROR)
            self._refresh_safety_status()
            self._update_buttons()
            return False

        in_memory=dict(self.safety_locks)
        self.safety_locks=dict(loaded.locks)
        self.safety_locks.update(in_memory)
        self.safety_metadata_health=loaded.health
        self.safety_metadata_error=""
        if self.safety_locks!=loaded.locks:
            if not self._persist_safety_locks(allow_repair=True):
                return False
        else:
            self.recovery_failed=bool(self.safety_locks)
            self._refresh_safety_status()
        self.append_log("โหลด Safety metadata หลังผู้ใช้แก้ไขสำเร็จ")
        return True

    def _refresh_safety_status(self):
        if not hasattr(self,"safety_status_var"):
            return
        health=getattr(self,"safety_metadata_health",SAFETY_METADATA_MISSING)
        if health in SAFETY_METADATA_FAILURE_STATES:
            self.safety_status_var.set(f"FAIL_CLOSED ({health})")
        elif getattr(self,"safety_locks",{}):
            self.safety_status_var.set(f"LOCKED ({len(self.safety_locks)})")
        else:
            self.safety_status_var.set("ไม่ล็อก")
        if not hasattr(self,"safety_detail_var"):
            return
        lines=[]
        if health in SAFETY_METADATA_FAILURE_STATES:
            detail=self.safety_metadata_error or SAFETY_PERSISTENCE_ERROR
            lines.append(
                f"METADATA | health={health} | path={self.safety_state_file} | error={detail}"
            )
        for lock in getattr(self,"safety_locks",{}).values():
            path=str(lock.get("path") or "")
            lock_type=str(lock.get("outcome") or lock.get("type") or "UNKNOWN")
            reason=str(lock.get("message") or lock.get("reason") or lock_type)
            lines.append(
                "LOCK | "
                f"filename={Path(path).name} | path={path} | reason={reason} | "
                f"type={lock_type} | status=UNRESOLVED"
            )
        self.safety_detail_var.set("\n".join(lines))

    def _invalidate_precheck(self):
        self.queue_revision+=1
        self.precheck_generation+=1
        # Detach a bounded manual/startup check immediately. Its eventual event
        # carries the old generation and cannot overwrite the replacement.
        self.precheck_in_progress=False
        self.precheck_purpose=""
        self.runtime_check_in_progress=False
        self.precheck_valid=False
        self.valid_excel=False
        self.runtime_valid=False
        self.tvc_connected=False
        self.tvc_login_verified=False
        self.bot_python=None
        for item in self.excel_queue.items:
            item.status="PENDING"
            item.message=""
        self.queue_status_var.set("รอตรวจสอบ")
        self._update_buttons()

    def _refresh_queue_tree(self,selection_index=None):
        if not hasattr(self,"queue_tree"):
            return
        selected_paths={
            self.queue_tree.item(item_id,"values")[0]
            for item_id in self.queue_tree.selection()
            if self.queue_tree.item(item_id,"values")
        }
        for item_id in self.queue_tree.get_children():
            self.queue_tree.delete(item_id)
        for index,item in enumerate(self.excel_queue.items):
            stats=item.stats
            item_id=self.queue_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    str(item.path),
                    QUEUE_STATUS_LABELS.get(item.status,item.status),
                    stats.get("WAIT",0),
                    stats.get("DONE",0),
                    stats.get("ERROR",0),
                ),
            )
            if str(item.path) in selected_paths:
                self.queue_tree.selection_add(item_id)
        if selection_index is not None and 0<=selection_index<len(self.excel_queue.items):
            item_id=str(selection_index)
            self.queue_tree.selection_set(item_id)
            self.queue_tree.focus(item_id)
            self.queue_tree.see(item_id)

    def _selected_queue_indices(self):
        if not hasattr(self,"queue_tree"):
            return []
        return sorted(int(item_id) for item_id in self.queue_tree.selection())

    def _on_queue_selection(self,_event=None):
        indices=self._selected_queue_indices()
        if indices:
            self.excel_var.set(str(self.excel_queue.items[indices[0]].path))

    def remove_selected_files(self):
        if not self._queue_edit_allowed():
            return
        indices=self._selected_queue_indices()
        if not indices:
            return
        removed_paths=[self.excel_queue.items[index].path for index in indices]
        self.excel_queue.remove_indices(indices)
        for path in removed_paths:
            key=os.path.normcase(str(normalize_excel_path(path)))
            if key in self.safety_locks:
                self.append_log(
                    f"นำไฟล์ที่มี Safety Lock ออกจาก Queue: {path} "
                    "(workbook เดิมไม่ถูกแก้ไข; ต้องกดตรวจสอบหลังแก้ไข)"
                )
        self.excel_var.set("")
        self._invalidate_precheck()
        self._refresh_queue_tree()
        self._refresh_progress()
        self.retry_precheck()

    def clear_files(self):
        if not self._queue_edit_allowed() or not self.excel_queue.items:
            return
        removed=self.excel_queue.clear()
        for item in removed:
            if item.key in self.safety_locks:
                self.append_log(
                    f"นำไฟล์ที่มี Safety Lock ออกจาก Queue: {item.path} "
                    "(workbook เดิมไม่ถูกแก้ไข; ต้องกดตรวจสอบหลังแก้ไข)"
                )
        self.excel_var.set("")
        self._invalidate_precheck()
        self._refresh_queue_tree()
        self._refresh_progress()
        self.retry_precheck()

    def move_selected_file(self,offset):
        if not self._queue_edit_allowed():
            return
        indices=self._selected_queue_indices()
        if len(indices)!=1:
            return
        new_index=self.excel_queue.move(indices[0],offset)
        self._invalidate_precheck()
        self._refresh_queue_tree(selection_index=new_index)
        self.retry_precheck()

    def retry_precheck(self):
        return self._begin_precheck("manual")

    def _begin_precheck(self,purpose):
        if (
            self.queue_running
            or self._is_running()
            or self.recovery_in_progress
            or self.precheck_in_progress
            or self.closing
        ):
            return
        self.precheck_generation+=1
        generation=self.precheck_generation
        paths=[item.path for item in self.excel_queue.items]
        revision=getattr(self,"queue_revision",0)
        path_keys=tuple(item.key for item in self.excel_queue.items)
        self.precheck_in_progress=True
        self.precheck_purpose=purpose
        self.runtime_check_in_progress=True
        self.precheck_valid=False
        self.valid_excel=False
        self.runtime_valid=False
        self.runtime_error=""
        self.bot_python=None
        self.tvc_connected=False
        self.tvc_login_verified=False
        self.runtime_status_var.set("กำลังตรวจสอบ")
        self.tvc_status_var.set("กำลังตรวจสอบ")
        self.queue_status_var.set("กำลังตรวจสอบ")
        for item in self.excel_queue.items:
            item.status="CHECKING"
        self._refresh_queue_tree()
        if purpose=="start":
            self._set_status("กำลังตรวจสอบก่อนเริ่ม...")
            self.append_log("Start-time revalidation: Runtime / T.V.C / Excel / Queue")
        else:
            self._set_status("กำลังตรวจสอบระบบ...")
            self.append_log("เริ่ม Pre-check: Runtime / T.V.C / Excel / Queue")
        self._update_buttons()
        threading.Thread(
            target=self._precheck_worker,
            args=(generation,paths,purpose,revision,path_keys),
            daemon=True,
        ).start()

    def _precheck_worker(self,generation,paths,purpose="manual",revision=0,path_keys=()):
        try:
            result=perform_precheck(paths)
        except Exception as exc:
            result={
                "runtime":{"ready":False,"python":None,"message":str(exc)},
                "tvc":{"ready":False,"connected":False,"login_verified":False,"message":str(exc)},
                "items":[],
                "queue_ready":False,
                "queue_message":str(exc),
                "total_wait":0,
                "ready":False,
            }
        self.events.put(
            ("precheck_result",generation,result,purpose,revision,tuple(path_keys))
        )

    def _on_precheck_result(
        self,generation,result,purpose="manual",revision=None,path_keys=None
    ):
        if generation!=self.precheck_generation:
            return
        self.precheck_in_progress=False
        self.runtime_check_in_progress=False
        current_keys=tuple(item.key for item in self.excel_queue.items)
        snapshot_changed=(
            revision is not None
            and (
                revision!=getattr(self,"queue_revision",0)
                or tuple(path_keys or ())!=current_keys
            )
        )
        if snapshot_changed:
            self.precheck_valid=False
            self.valid_excel=False
            self.start_pending=False
            self.excel_queue.locked=False
            for item in self.excel_queue.items:
                item.status="PENDING"
            self.queue_status_var.set("Queue เปลี่ยนระหว่างตรวจ")
            self._set_status("Error")
            self.append_log(
                "ยกเลิก Start: Queue revision/path set เปลี่ยนระหว่าง Start-time revalidation"
            )
            self._refresh_queue_tree()
            self._update_buttons()
            return
        runtime=result["runtime"]
        tvc=result["tvc"]
        self.runtime_valid=bool(runtime["ready"])
        self.runtime_error="" if self.runtime_valid else runtime.get("message","")
        self.bot_python=Path(runtime["python"]) if runtime.get("python") else None
        self.tvc_connected=bool(tvc.get("connected"))
        self.tvc_login_verified=bool(tvc.get("login_verified"))
        self.tvc_error="" if tvc.get("ready") else tvc.get("message","")
        self.runtime_status_var.set("พร้อม" if self.runtime_valid else "Error")
        tvc_probe_status=str(tvc.get("status") or "").upper()
        if tvc.get("ready"):
            tvc_label="พร้อม"
        elif tvc_probe_status=="TIMEOUT":
            tvc_label="หมดเวลาตรวจสอบ"
        elif tvc_probe_status=="ERROR":
            tvc_label="Error"
        elif tvc_probe_status in {"PENDING","SKIPPED"}:
            tvc_label="ยังไม่ได้ตรวจ"
        elif tvc_probe_status in {"LOGIN_REQUIRED","FOUND_NOT_READY","FOUND"}:
            tvc_label="พบโปรแกรม - ยังไม่ยืนยัน Login"
        elif tvc.get("connected"):
            tvc_label="ยังไม่พร้อม"
        else:
            tvc_label="ยังไม่พบโปรแกรม"
        self.tvc_status_var.set(tvc_label)

        by_path={str(item.path):item for item in self.excel_queue.items}
        for item_result in result.get("items",[]):
            item=by_path.get(item_result["path"])
            if item is None:
                continue
            item.status=item_result["status"]
            item.stats=dict(item_result.get("stats") or item.stats)
            item.error_jobs=list(item_result.get("errors") or [])
            item.message=item_result.get("message","")
            safety_issues=list(item_result.get("safety_issues") or [])
            if safety_issues:
                issue=next(
                    (
                        candidate for candidate in safety_issues
                        if candidate.get("reason")=="UNCERTAIN_TVC_SAVE"
                    ),
                    safety_issues[0],
                )
                self._set_safety_lock(
                    item.path,
                    issue.get("reason") or "failed",
                    issue.get("message") or item.message,
                    job_ref=issue.get("job_ref") or "",
                )
                self.append_log(
                    f"SAFETY LOCK {item.path}: "
                    f"{issue.get('message') or item.message}"
                )
            if item.status=="INVALID":
                self.append_log(f"INVALID {item.path}: {item.message}")
        self.queue_status_var.set(result.get("queue_message","Invalid"))
        self.precheck_valid=bool(result.get("ready"))
        self.valid_excel=bool(result.get("queue_ready"))
        safety_locked=self._safety_lock_active()
        if self._metadata_fail_closed():
            self._set_status(SAFETY_PERSISTENCE_ERROR)
            self._set_mascot("ERROR")
            self.append_log(
                self.safety_metadata_error or SAFETY_PERSISTENCE_ERROR
            )
        elif safety_locked:
            self._set_status("ต้องตรวจสอบงานก่อนเริ่มใหม่")
            self._set_mascot("ERROR")
            self.append_log(
                "Pre-check เสร็จ แต่ Safety Lock ยัง active; "
                "ให้ตรวจ T.V.C/Excel แล้วกด 'ตรวจสอบหลังแก้ไข'"
            )
        elif self.precheck_valid:
            self._set_status("พร้อมใช้งาน")
            self._set_mascot("READY")
            self.append_log("Pre-check ผ่านครบทุกหัวข้อ")
        else:
            item_statuses={item.status for item in self.excel_queue.items}
            waiting_for_setup=(
                self.runtime_valid
                and tvc_probe_status not in {"TIMEOUT","ERROR"}
                and "INVALID" not in item_statuses
                and "REVIEW_REQUIRED" not in item_statuses
            )
            self._set_status("รอความพร้อม" if waiting_for_setup else "Error")
            self._set_mascot("READY" if waiting_for_setup else "ERROR")
            if not self.runtime_valid:
                self.append_log(f"Runtime Error: {self.runtime_error}")
            if not tvc.get("ready"):
                self.append_log(f"T.V.C: {tvc.get('message','Not Found')}")
            if not result.get("queue_ready"):
                self.append_log(f"Queue: {result.get('queue_message','Invalid')}")
        self._refresh_queue_tree()
        self._refresh_progress()
        if purpose=="start":
            if self.precheck_valid and not safety_locked:
                self.append_log("Start-time revalidation ผ่าน กำลังเริ่ม Queue...")
                self._start_validated_queue()
                return
            self.start_pending=False
            self.excel_queue.locked=False
            self.append_log("Start-time revalidation ไม่ผ่าน จึงไม่เริ่ม subprocess")
        self._update_buttons()

    def retry_safety_validation(self):
        if (
            not self._safety_lock_active()
            or self.queue_running
            or self._is_running()
            or self.recovery_in_progress
            or self.precheck_in_progress
            or self.closing
        ):
            return
        if self._metadata_fail_closed():
            if not self._retry_safety_metadata():
                return
            if not self.safety_locks:
                self.recovery_failed=False
                self._refresh_safety_status()
                self.retry_precheck()
                return
        self.precheck_generation+=1
        generation=self.precheck_generation
        revision=self.queue_revision
        paths=[item.path for item in self.excel_queue.items]
        path_keys=tuple(item.key for item in self.excel_queue.items)
        locks={key:dict(value) for key,value in self.safety_locks.items()}
        self.precheck_in_progress=True
        self.precheck_purpose="safety"
        self.runtime_check_in_progress=True
        self.precheck_valid=False
        self.valid_excel=False
        self._set_status("กำลังตรวจสอบก่อนเริ่ม...")
        self.safety_status_var.set("กำลังตรวจสอบ")
        for item in self.excel_queue.items:
            item.status="CHECKING"
        self._refresh_queue_tree()
        self.append_log(
            "ผู้ใช้ยืนยันตรวจ T.V.C/Excel แล้ว; เริ่ม read-only Safety revalidation "
            "(ไม่แก้ UNCERTAIN_TVC_SAVE และไม่ retry JOB)"
        )
        self._update_buttons()
        threading.Thread(
            target=self._safety_validation_worker,
            args=(generation,revision,paths,path_keys,locks),
            daemon=True,
        ).start()

    def _safety_validation_worker(
        self,generation,revision,paths,path_keys,locks
    ):
        precheck_result=perform_precheck(paths)
        inspections={}
        for key,lock in locks.items():
            try:
                lock_path=Path(lock["path"])
                inspection=inspect_recovery_state(
                    lock_path,
                    current_job_ref=lock.get("job_ref",""),
                )
                inspection=dict(inspection)
                inspection["safety_issues"]=get_safety_issues(lock_path)
                inspections[key]=inspection
            except Exception as exc:
                inspections[key]={
                    "outcome":"failed",
                    "verified":False,
                    "running_count":-1,
                    "message":str(exc),
                }
        resolved,unresolved,resolved_keys=evaluate_safety_revalidation(
            precheck_result,
            locks,
            inspections,
            path_keys,
            include_resolved_keys=True,
        )
        self.events.put(
            (
                "safety_revalidation_result",
                generation,
                revision,
                tuple(path_keys),
                precheck_result,
                resolved,
                unresolved,
                resolved_keys,
            )
        )

    def _on_safety_revalidation_result(
        self,generation,revision,path_keys,precheck_result,resolved,unresolved,
        resolved_keys=None,
    ):
        current_keys=tuple(item.key for item in self.excel_queue.items)
        snapshot_current=(
            generation==self.precheck_generation
            and revision==self.queue_revision
            and tuple(path_keys)==current_keys
        )
        if not snapshot_current:
            self.precheck_in_progress=False
            self.runtime_check_in_progress=False
            self.safety_status_var.set("LOCKED")
            self._set_status("ต้องตรวจสอบงานก่อนเริ่มใหม่")
            self.append_log("Safety revalidation ถูกยกเลิกเพราะ Queue เปลี่ยนระหว่างตรวจ")
            self._update_buttons()
            return
        self._on_precheck_result(
            generation,
            precheck_result,
            purpose="safety",
            revision=revision,
            path_keys=path_keys,
        )
        if resolved_keys is None:
            resolved_keys=tuple(self.safety_locks) if resolved else tuple()
        released=[]
        for key in resolved_keys:
            lock=self.safety_locks.pop(key,None)
            if lock is not None:
                released.append(str(lock.get("path") or key))
        persisted=True
        if released:
            persisted=self._persist_safety_locks()

        if resolved and not self.safety_locks and persisted and not self._metadata_fail_closed():
            self.recovery_failed=False
            self._refresh_safety_status()
            if self.precheck_valid:
                self._set_status("พร้อมใช้งาน")
                self._set_mascot("READY")
            else:
                self._set_status("Error")
                self._set_mascot("ERROR")
            self.append_log(
                "Safety Lock ถูกปลดหลัง read-only revalidation ผ่าน"
                + (f": {', '.join(released)}" if released else "")
            )
            if not self.precheck_valid:
                self.append_log(
                    "Safety Lock ถูกปลดแล้ว แต่ Queue ยังเริ่มไม่ได้ "
                    "(เช่น ไม่มี WAIT หรือ Queue ว่าง)"
                )
        else:
            self.recovery_failed=True
            self._refresh_safety_status()
            if self._metadata_fail_closed():
                self._set_status(SAFETY_PERSISTENCE_ERROR)
            else:
                self._set_status("ต้องตรวจสอบงานก่อนเริ่มใหม่")
            for path in released:
                self.append_log(f"Safety Lock ถูกปลดเฉพาะไฟล์ที่ตรวจผ่าน: {path}")
            for message in unresolved:
                self.append_log(f"Safety Lock: {message}")
        self._update_buttons()

    def _select_excel(self,path: Path,show_error: bool):
        path=path.expanduser().resolve()
        try:
            validate_workbook(path)
            stats=get_job_stats(path)
        except Exception as exc:
            self.valid_excel=False
            self.excel_var.set(str(path))
            self._update_buttons()
            self._set_status("Error")
            self.append_log(f"Excel ไม่ถูกต้อง: {exc}")
            if show_error:
                messagebox.showerror("ไฟล์ Excel ไม่ถูกต้อง",str(exc),parent=self.root)
            return False

        self.valid_excel=True
        self.excel_var.set(str(path))
        self.current_job_var.set("- / 0")
        self._apply_stats(str(path),stats)

        if stats["RUNNING"]>0:
            self.recovery_failed=True
            self._set_status("Error")
            message=(
                f"พบ JOB RUNNING {stats['RUNNING']} รายการ "
                "ต้องตรวจ Excel ก่อนเริ่มใหม่"
            )
            self.append_log(message)
            self._update_buttons()
            if show_error:
                messagebox.showerror("พบ JOB RUNNING",message,parent=self.root)
            return False

        self.recovery_failed=False
        if self.runtime_check_in_progress:
            self._set_status("กำลังตรวจสอบระบบ...")
        elif self.runtime_valid:
            self._set_status("พร้อมใช้งาน")
        elif self.runtime_error:
            self._set_status("Error")
        self.append_log(f"เลือกไฟล์ Excel: {path}")
        self._update_buttons()
        self._request_stats()
        return True

    def retry_runtime_check(self):
        if (
            self._is_running()
            or self.recovery_in_progress
            or self.runtime_check_in_progress
            or self.closing
        ):
            return
        self.runtime_check_in_progress=True
        self.runtime_valid=False
        self.runtime_error=""
        self.bot_python=None
        self._set_status("กำลังตรวจสอบระบบ...")
        runtime_name=(
            "Worker executable"
            if getattr(self,"runtime_paths",RUNTIME_PATHS).frozen
            else "Python environment และ dependencies"
        )
        self.append_log(f"กำลังตรวจสอบ {runtime_name}...")
        self._update_buttons()
        threading.Thread(target=self._runtime_check_worker,daemon=True).start()

    def _runtime_check_worker(self):
        try:
            python_exe=validate_bot_runtime()
            self.events.put(("runtime_check_result",True,python_exe,""))
        except Exception as exc:
            self.events.put(("runtime_check_result",False,None,str(exc)))

    def _on_runtime_check_result(self,success,python_exe,error):
        self.runtime_check_in_progress=False
        if success:
            self.runtime_valid=True
            self.runtime_error=""
            self.bot_python=Path(python_exe)
            if self.recovery_failed:
                self._set_status("Error")
            else:
                self._set_status("พร้อมใช้งาน")
            self.append_log(f"ตรวจระบบผ่าน: {self.bot_python}")
        else:
            self.runtime_valid=False
            self.runtime_error=error or "ไม่ทราบสาเหตุ"
            self.bot_python=None
            self._set_status("Error")
            frozen=getattr(self,"runtime_paths",RUNTIME_PATHS).frozen
            message=(
                (
                    "Worker executable ใช้งานไม่ได้ กรุณาติดตั้งไฟล์ "
                    "TVC Bot Worker.exe ไว้ข้าง GUI\n\n"
                    if frozen else
                    "Python environment ใช้งานไม่ได้ กรุณาติดตั้ง/ซ่อม virtual environment\n\n"
                )
                +self.runtime_error
            )
            self.append_log(message.replace("\n\n"," | "))
            messagebox.showerror(
                "Worker ใช้งานไม่ได้" if frozen else "Python environment ใช้งานไม่ได้",
                message,
                parent=self.root,
            )
        self._update_buttons()

    def start_bot(self):
        if not hasattr(self,"excel_queue"):
            return self._start_single_legacy()
        if self.queue_running or self._is_running() or self.precheck_in_progress:
            return
        if self._safety_lock_active():
            self._set_status("ต้องตรวจสอบงานก่อนเริ่มใหม่")
            self.append_log(
                "Start ถูกบล็อกโดย Safety Lock; กรุณากด 'ตรวจสอบหลังแก้ไข'"
            )
            self._update_buttons()
            return
        if not self.excel_queue.items:
            self._set_status("Error")
            self.append_log("ยังไม่มีไฟล์ Excel ใน Queue")
            self._update_buttons()
            return
        self.start_pending=True
        self.excel_queue.locked=True
        self._set_status("กำลังตรวจสอบก่อนเริ่ม...")
        self._update_buttons()
        self._begin_precheck("start")

    def _start_validated_queue(self):
        totals=self.excel_queue.totals()
        if (
            not self.start_pending
            or not self.precheck_valid
            or self._safety_lock_active()
            or self.bot_python is None
            or not self.tvc_connected
            or not self.tvc_login_verified
            or not self.excel_queue.all_ready()
            or totals["WAIT"]<=0
        ):
            self.start_pending=False
            self.excel_queue.locked=False
            self._set_status(
                "ต้องตรวจสอบงานก่อนเริ่มใหม่"
                if self._safety_lock_active()
                else "Error"
            )
            self.append_log("ยกเลิก Start: ผล Start-time revalidation ไม่อนุญาตให้เริ่ม")
            self._update_buttons()
            return
        self.start_pending=False
        self.queue_running=True
        self.run_controller=QueueRunController(len(self.excel_queue.items))
        self.stop_request_sent=False
        self.stop_event_seen=False
        self.force_stop_used=False
        self.recovery_failed=False
        self.current_file_index=self.run_controller.start()
        self._set_status("กำลังทำงาน")
        self._set_mascot("RUNNING")
        self.append_log(
            f"เริ่ม Queue {len(self.excel_queue.items)} ไฟล์ | WAIT รวม {totals['WAIT']}"
        )
        self._update_buttons()
        self._start_queue_item(self.current_file_index)

    def _rollback_queue_launch_failure(self,item,exc):
        message=f"{STOP_FILE_PREPARATION_ERROR}: {exc}"
        self.process=None
        self.stop_file=None
        item.status="ERROR"
        item.message=message
        if self.run_controller is not None:
            self.run_controller.complete_current(1)
        self.queue_running=False
        self.finalization_in_progress=False
        self.excel_queue.locked=False
        self.start_pending=False
        self.precheck_valid=False
        self.valid_excel=False
        self.queue_status_var.set("Error")
        self._set_status(STOP_FILE_PREPARATION_ERROR)
        self._set_mascot("ERROR")
        self.append_log(message)
        self._refresh_queue_tree(selection_index=self.current_file_index)
        self._refresh_progress()
        self._update_buttons()
        messagebox.showerror(
            STOP_FILE_PREPARATION_ERROR,
            message,
            parent=self.root,
        )

    def _start_queue_item(self,index):
        if (
            not self.queue_running
            or self.stop_request_sent
            or index is None
            or not (0<=index<len(self.excel_queue.items))
        ):
            self._finish_queue("STOPPED" if self.stop_request_sent else "ERROR")
            return
        self.current_file_index=index
        item=self.excel_queue.items[index]
        self.excel_var.set(str(item.path))
        self.current_file_var.set(f"{index+1} / {len(self.excel_queue.items)}")
        self.current_job_var.set("- / 0")
        self.current_job_ref=""
        self.stop_event_seen=False
        self.stop_event_phase=""
        self.last_batch_success=None
        self.force_stop_used=False
        self.force_retry_available=False
        self.recovery_in_progress=False
        self.last_return_code=None
        runtime_paths=getattr(self,"runtime_paths",RUNTIME_PATHS)
        try:
            stop_file,command=prepare_worker_launch(
                runtime_paths,
                item.path,
                self.bot_python,
            )
        except Exception as exc:
            self._rollback_queue_launch_failure(item,exc)
            return
        self.stop_file=stop_file
        item.status="RUNNING"
        item.message=""
        environment=os.environ.copy()
        environment["PYTHONIOENCODING"]="utf-8"
        try:
            self.process=subprocess.Popen(
                command,
                cwd=str(runtime_paths.app_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0),
                env=environment,
            )
        except Exception as exc:
            self.process=None
            self.stop_file=None
            item.status="ERROR"
            item.message=str(exc)
            if self.run_controller is not None:
                self.run_controller.complete_current(1)
            self.append_log(f"เริ่ม Bot สำหรับ {item.path.name} ไม่สำเร็จ: {exc}")
            self._finish_queue("ERROR")
            return
        self.append_log(
            f"เริ่มไฟล์ {index+1}/{len(self.excel_queue.items)}: {item.path}"
        )
        self._refresh_queue_tree(selection_index=index)
        self._refresh_progress()
        self._update_buttons()
        threading.Thread(target=self._read_process_output,daemon=True).start()

    def _start_single_legacy(self):
        if self._is_running():
            return
        if self.runtime_check_in_progress or not self.runtime_valid or self.bot_python is None:
            self._set_status("Error")
            self.append_log("ยังไม่สามารถเริ่ม Bot ได้: Python environment ยังไม่ผ่านการตรวจสอบ")
            self._update_buttons()
            return
        excel_text=self.excel_var.get().strip()
        if not excel_text or not self._select_excel(Path(excel_text),show_error=True):
            return
        python_exe=self.bot_python

        runtime_paths=getattr(self,"runtime_paths",RUNTIME_PATHS)
        try:
            stop_file,command=prepare_worker_launch(
                runtime_paths,
                self.excel_var.get(),
                python_exe,
            )
        except Exception as exc:
            self.process=None
            self.stop_file=None
            message=f"{STOP_FILE_PREPARATION_ERROR}: {exc}"
            self._set_status(STOP_FILE_PREPARATION_ERROR)
            self.append_log(message)
            self._update_buttons()
            messagebox.showerror(
                STOP_FILE_PREPARATION_ERROR,
                message,
                parent=self.root,
            )
            return
        self.stop_file=stop_file
        environment=os.environ.copy()
        environment["PYTHONIOENCODING"]="utf-8"

        try:
            self.process=subprocess.Popen(
                command,
                cwd=str(runtime_paths.app_dir),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=getattr(subprocess,"CREATE_NO_WINDOW",0),
                env=environment,
            )
        except Exception as exc:
            self.process=None
            self.stop_file=None
            self._set_status("Error")
            self.append_log(f"เริ่ม Bot ไม่สำเร็จ: {exc}")
            messagebox.showerror("เริ่ม Bot ไม่สำเร็จ",str(exc),parent=self.root)
            self._update_buttons()
            return

        self.stop_request_sent=False
        self.stop_event_seen=False
        self.last_batch_success=None
        self.current_job_ref=""
        self.force_stop_used=False
        self.force_retry_available=False
        self.recovery_in_progress=False
        self.stop_event_phase=""
        self.last_return_code=None
        self._set_status("กำลังทำงาน")
        self.append_log("เริ่ม Bot...")
        self._update_buttons()
        threading.Thread(target=self._read_process_output,daemon=True).start()

    def _read_process_output(self):
        process=self.process
        if process is None:
            return
        try:
            assert process.stdout is not None
            for line in process.stdout:
                self.events.put(("line",line.rstrip("\r\n")))
        except Exception as exc:
            self.events.put(("line",f"อ่าน log จาก Bot ไม่สำเร็จ: {exc}"))
        finally:
            return_code=process.wait()
            if process.stdout is not None:
                process.stdout.close()
            self.events.put(("process_end",process,return_code))

    def stop_bot(self):
        if self._is_running() or getattr(self,"queue_running",False):
            self._request_stop()

    def _request_stop(self):
        # Intentional contract: finish the current JOB's complete safe flow.
        # The stop flag is consumed only at JOB boundaries by bot.py.
        process=self.process
        if self.stop_request_sent:
            return
        if process is None or process.poll() is not None:
            if getattr(self,"queue_running",False):
                self.stop_request_sent=True
                self.append_log("รับคำขอหยุด Queue แล้ว จะไม่เริ่มไฟล์ถัดไป")
                self._finish_queue("STOPPED")
            return

        if self.force_retry_available:
            self.force_retry_available=False
            self.stop_request_sent=True
            self._set_status("กำลังหยุดหลังจบ JOB ปัจจุบัน")
            self.append_log("กำลังลองบังคับหยุด process อีกครั้ง...")
            self._update_buttons()
            self._force_stop_if_running(process)
            return

        try:
            assert self.stop_file is not None
            self.stop_file.write_text("stop",encoding="utf-8")
        except Exception as exc:
            self.stop_request_sent=False
            self._set_status("กำลังทำงาน")
            self.append_log(f"ส่งคำขอหยุดไม่สำเร็จ: {exc}")
            self._update_buttons()
            messagebox.showerror(
                "ส่งคำขอหยุดไม่สำเร็จ",
                f"{exc}\n\nBot ยังทำงานอยู่ และสามารถกดหยุดเพื่อลองใหม่ได้",
                parent=self.root,
            )
            return

        self.stop_request_sent=True
        self._set_status("กำลังหยุดหลังจบ JOB ปัจจุบัน")
        self.append_log("ส่งคำขอหยุดแล้ว: Bot จะทำ JOB ปัจจุบันให้จบก่อนหยุด")
        self._update_buttons()
        self.root.after(STOP_TIMEOUT_MS,lambda: self._force_stop_if_running(process))

    def _force_stop_if_running(self,process):
        if self.process is not process or process.poll() is not None:
            return
        self.force_stop_used=True
        current=self.current_job_ref or "ไม่ทราบ JOB"
        self.append_log(
            f"Bot ไม่จบภายในเวลาที่กำหนด กำลังบังคับหยุด process tree (JOB: {current})..."
        )
        threading.Thread(target=self._terminate_process_tree,args=(process,),daemon=True).start()

    def _terminate_process_tree(self,process):
        errors=[]
        try:
            parent=psutil.Process(process.pid)
            children=parent.children(recursive=True)
            for child in children:
                child.terminate()
            parent.terminate()
            _,alive=psutil.wait_procs(children+[parent],timeout=3)
            for item in alive:
                item.kill()
        except psutil.NoSuchProcess:
            pass
        except Exception as exc:
            errors.append(f"psutil: {exc}")

        if process.poll() is None:
            try:
                process.kill()
            except Exception as exc:
                errors.append(f"process.kill: {exc}")

        if process.poll() is None:
            try:
                process.wait(timeout=3)
            except subprocess.TimeoutExpired:
                pass
            except Exception as exc:
                errors.append(f"process.wait: {exc}")

        if process.poll() is None:
            message="; ".join(errors) or "process ยังไม่จบหลัง terminate/kill"
            self.events.put(("force_termination_failed",process,message))

    def open_logs(self):
        try:
            LOG_DIR.mkdir(parents=True,exist_ok=True)
            os.startfile(str(LOG_DIR))
        except Exception as exc:
            messagebox.showerror(
                "ไม่สามารถเปิดโฟลเดอร์ Logs ได้",
                str(exc),
                parent=self.root,
            )

    def _is_running(self):
        return self.process is not None and self.process.poll() is None

    def _update_buttons(self):
        running=self._is_running()
        if not hasattr(self,"excel_queue"):
            blocked=(
                running
                or self.recovery_in_progress
                or self.recovery_failed
                or self.runtime_check_in_progress
                or not self.runtime_valid
            )
            self.start_button.configure(
                state="disabled" if blocked or not self.valid_excel or self.closing else "normal"
            )
            self.stop_button.configure(
                state="normal" if running and not self.stop_request_sent else "disabled"
            )
            self.browse_button.configure(
                state="disabled" if running or self.recovery_in_progress else "normal"
            )
            self.runtime_button.configure(
                state=(
                    "disabled"
                    if running or self.recovery_in_progress or self.runtime_check_in_progress or self.closing
                    else "normal"
                )
            )
            return
        queue_running=self.queue_running
        blocked=(
            running
            or queue_running
            or self.start_pending
            or self.finalization_in_progress
            or self.recovery_in_progress
            or self.recovery_failed
            or self._safety_lock_active()
            or self.precheck_in_progress
            or not self.precheck_valid
            or not self.runtime_valid
        )
        self.start_button.configure(
            state=(
                "disabled"
                if blocked
                or self.closing
                or not self.excel_queue.all_ready()
                or self.excel_queue.totals()["WAIT"]<=0
                else "normal"
            )
        )
        self.stop_button.configure(
            state="normal" if queue_running and not self.stop_request_sent else "disabled"
        )
        edit_state="normal" if self._queue_edit_allowed() else "disabled"
        for button in (
            self.browse_button,
            self.remove_button,
            self.clear_button,
            self.up_button,
            self.down_button,
        ):
            button.configure(state=edit_state)
        self.runtime_button.configure(
            state=(
                "disabled"
                if queue_running
                or self.start_pending
                or self.recovery_in_progress
                or self.precheck_in_progress
                or self.closing
                else "normal"
            )
        )
        self.safety_button.configure(
            state=(
                "normal"
                if self._safety_lock_active()
                and not queue_running
                and not self.start_pending
                and not self.recovery_in_progress
                and not self.precheck_in_progress
                and not self.closing
                else "disabled"
            )
        )

    def _set_status(self,text):
        self.status_var.set(text)
        colors={
            "พร้อมใช้งาน":"#2563eb",
            "กำลังตรวจสอบระบบ...":"#d97706",
            "กำลังตรวจสอบก่อนเริ่ม...":"#d97706",
            "รอความพร้อม":"#2563eb",
            "กำลังทำงาน":"#d97706",
            "กำลังหยุดหลังจบ JOB ปัจจุบัน":"#d97706",
            "หยุดแล้ว":"#2563eb",
            "หยุดไม่สำเร็จ - process ยังทำงาน":"#b91c1c",
            "สำเร็จ":"#15803d",
            "เสร็จสิ้นพร้อมข้อผิดพลาด":"#b91c1c",
            "ต้องตรวจสอบงานก่อนเริ่มใหม่":"#b91c1c",
            "ผลการทำงานไม่ครบ":"#b91c1c",
            "Error":"#b91c1c",
        }
        self.status_label.configure(foreground=colors.get(text,"#111827"))
        if hasattr(self,"mascot_label"):
            if text=="กำลังทำงาน" or text=="กำลังหยุดหลังจบ JOB ปัจจุบัน":
                self._set_mascot("RUNNING")
            elif text=="สำเร็จ":
                self._set_mascot("SUCCESS")
            elif text in {
                "Error",
                "หยุดไม่สำเร็จ - process ยังทำงาน",
                "เสร็จสิ้นพร้อมข้อผิดพลาด",
                "ต้องตรวจสอบงานก่อนเริ่มใหม่",
                "ผลการทำงานไม่ครบ",
            }:
                self._set_mascot("ERROR")
            else:
                self._set_mascot("READY")

    def append_log(self,message):
        if message is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end",str(message)+"\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _handle_line(self,line):
        if line.startswith(EVENT_PREFIX):
            try:
                event=json.loads(line[len(EVENT_PREFIX):])
            except json.JSONDecodeError:
                self.append_log(line)
                return
            self._handle_bot_event(event)
        else:
            self.append_log(line)

    def _handle_bot_event(self,event):
        event_name=event.get("event")
        if event_name=="batch_start":
            self._set_status("กำลังทำงาน")
        elif event_name=="job_start":
            self.current_job_ref=str(event.get("ref") or "")
            self.current_job_var.set(f"{event.get('ref','-')} / {event.get('total',0)}")
            self._request_stats()
        elif event_name=="job_done":
            self._request_stats()
        elif event_name=="job_error":
            self._set_status("Error")
            self.last_batch_success=False
            self._request_stats()
        elif event_name=="commit_state":
            self.append_log(f"Commit state: {event.get('state','')}")
        elif event_name=="stopped":
            self.stop_event_seen=True
            self.stop_event_phase=str(event.get("phase") or "")
            self._request_stats()
        elif event_name=="batch_complete":
            self.last_batch_success=bool(event.get("success"))
            self._request_stats()

    def _stats_tick(self):
        if self.root.winfo_exists():
            self._request_stats()
            self.root.after(STATS_INTERVAL_MS,self._stats_tick)

    def _request_stats(self):
        if self.stats_loading:
            return
        path_text=self.excel_var.get().strip()
        if not path_text:
            return
        path=Path(path_text)
        self.stats_loading=True
        threading.Thread(target=self._load_stats,args=(path,),daemon=True).start()

    def _load_stats(self,path):
        try:
            stats=get_job_stats(path)
            self.events.put(("stats",str(path),stats))
        except Exception as exc:
            self.events.put(("stats_error",str(path),str(exc)))

    def _apply_stats(self,path_text,stats):
        if hasattr(self,"excel_queue"):
            normalized=str(normalize_excel_path(path_text))
            for item in self.excel_queue.items:
                if str(item.path)==normalized:
                    item.stats=dict(stats)
                    break
            self._refresh_queue_tree(selection_index=self.current_file_index if self.queue_running else None)
            self._refresh_progress()
            return
        if path_text==self.excel_var.get():
            self.wait_var.set(str(stats["WAIT"]))
            self.done_var.set(str(stats["DONE"]))
            self.error_var.set(str(stats["ERROR"]))
            total=stats["TOTAL"]
            completed=stats["COMPLETED"]
            self.progress.configure(maximum=max(total,1),value=min(completed,total))
            self.progress_text_var.set(f"{completed} / {total}")

    def _refresh_progress(self):
        if not hasattr(self,"excel_queue"):
            return
        progress=calculate_queue_progress(
            self.excel_queue.items,self.current_file_index
        )
        current=progress["current"]
        overall=progress["overall"]
        self.wait_var.set(str(overall["WAIT"]))
        self.done_var.set(str(overall["DONE"]))
        self.error_var.set(str(overall["ERROR"]))
        current_total=current["TOTAL"]
        current_processed=min(progress["current_processed"],current_total)
        overall_total=overall["TOTAL"]
        overall_processed=min(progress["overall_processed"],overall_total)
        self.progress.configure(
            maximum=max(current_total,1),value=current_processed
        )
        self.progress_text_var.set(f"{current_processed} / {current_total} JOB")
        self.overall_progress.configure(
            maximum=max(overall_total,1),value=overall_processed
        )
        self.overall_progress_text_var.set(f"{overall_processed} / {overall_total} JOB")
        if self.current_file_index<0:
            self.current_file_var.set(f"- / {len(self.excel_queue.items)}")

    def _drain_events(self):
        try:
            while True:
                event=self.events.get_nowait()
                kind=event[0]
                if kind=="line":
                    self._handle_line(event[1])
                elif kind=="stats":
                    self.stats_loading=False
                    self._apply_stats(event[1],event[2])
                elif kind=="stats_error":
                    self.stats_loading=False
                    if event[1]==self.excel_var.get():
                        self.append_log(f"อัปเดตสถิติไม่ได้ชั่วคราว: {event[2]}")
                elif kind=="process_end":
                    self._on_process_end(event[1],event[2])
                elif kind=="reconciliation":
                    self._on_reconciliation(event[1],event[2])
                elif kind=="force_termination_failed":
                    self._on_force_termination_failed(event[1],event[2])
                elif kind=="runtime_check_result":
                    self._on_runtime_check_result(event[1],event[2],event[3])
                elif kind=="precheck_result":
                    self._on_precheck_result(
                        event[1],event[2],event[3],event[4],event[5]
                    )
                elif kind=="queue_file_result":
                    self._on_queue_file_result(
                        event[1],event[2],event[3],event[4],event[5]
                    )
                elif kind=="queue_recovery_result":
                    self._on_queue_recovery_result(
                        event[1],event[2],event[3],event[4],event[5]
                    )
                elif kind=="safety_revalidation_result":
                    self._on_safety_revalidation_result(
                        event[1],event[2],event[3],event[4],event[5],event[6],event[7]
                    )
                elif kind=="queue_finalized":
                    self._on_queue_finalized(event[1],event[2])
        except queue.Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(100,self._drain_events)

    def _on_process_end(self,process,return_code):
        if self.process is not process:
            return
        self.process=None
        self.last_return_code=return_code
        if self.stop_file is not None:
            try:
                self.stop_file.unlink(missing_ok=True)
            except OSError:
                pass
            self.stop_file=None

        if getattr(self,"queue_running",False) and return_code==0:
            index=self.current_file_index
            path=self.excel_queue.items[index].path
            worker_success=self.last_batch_success is not False
            self.append_log(f"subprocess จบสำหรับ {path.name}; กำลังตรวจผลใน Excel...")
            threading.Thread(
                target=self._verify_queue_file,
                args=(index,path,return_code,worker_success),
                daemon=True,
            ).start()
            self._update_buttons()
            return

        if self.force_stop_used or return_code!=0:
            self.recovery_in_progress=True
            self._update_buttons()
            self.append_log("subprocess จบแล้ว กำลัง reconcile และ verify สถานะ Excel...")
            excel_path=Path(self.excel_var.get())
            current_ref=self.current_job_ref
            expected_no_active=(
                self.stop_event_seen
                and self.stop_event_phase=="before_job"
                and not current_ref
            )
            detail=(
                "ถูกบังคับหยุดก่อน save"
                if self.force_stop_used
                else "process จบผิดปกติก่อน save"
            )
            reason=f"{DIRTY_TVC_FORM_POSSIBLE} | {detail}"
            threading.Thread(
                target=self._reconcile_after_exit,
                args=(excel_path,current_ref,reason,expected_no_active),
                daemon=True,
            ).start()
            return

        if return_code==0 and self.last_batch_success is not False:
            self._set_status("สำเร็จ")
            self.append_log("Bot ทำงานเสร็จสมบูรณ์")
        else:
            self._set_status("Error")
            self.append_log(f"Bot จบการทำงานด้วยรหัส {return_code}")

        self._update_buttons()
        self._request_stats()
        if self.closing:
            self.root.after(100,self._destroy)

    def _verify_queue_file(self,index,path,return_code,worker_success):
        try:
            stats=get_job_stats(path)
            errors=get_job_errors(path)
            error=""
        except Exception as exc:
            stats={}
            errors=[]
            error=str(exc)
        self.events.put(
            ("queue_file_result",index,return_code,worker_success,stats,(errors,error))
        )

    def _on_queue_file_result(self,index,return_code,worker_success,stats,error_payload):
        if not self.queue_running or not (0<=index<len(self.excel_queue.items)):
            return
        errors,error=error_payload
        item=self.excel_queue.items[index]
        if stats:
            item.stats=dict(stats)
        item.error_jobs=list(errors)
        if error:
            item.status="ERROR"
            item.message=f"ตรวจผล Excel ไม่สำเร็จ: {error}"
            if self.run_controller is not None:
                self.run_controller.complete_current(1)
            self.append_log(f"{item.path.name}: {item.message}")
            self._finish_queue("ERROR")
            return
        if return_code!=0 or not worker_success:
            item.status="ERROR"
            item.message=f"Bot จบด้วยรหัส {return_code}"
            if self.run_controller is not None:
                self.run_controller.complete_current(1)
            self._finish_queue("ERROR")
            return
        if int(item.stats.get("RUNNING",0) or 0)>0:
            item.status="ERROR"
            item.message="ยังพบ JOB RUNNING หลัง subprocess จบ"
            if self.run_controller is not None:
                self.run_controller.complete_current(1)
            self._finish_queue("ERROR")
            return
        if int(item.stats.get("WAIT",0) or 0)>0:
            item.status="ERROR"
            item.message="subprocess จบด้วย exit 0 แต่ยังมี JOB WAIT เหลืออยู่"
            if self.run_controller is not None:
                self.run_controller.complete_current(1)
            self._finish_queue("ERROR")
            return

        item.status="DONE"
        item.message="ประมวลผล WAIT ครบแล้ว"
        self.append_log(f"DONE file {index+1}/{len(self.excel_queue.items)}: {item.path.name}")
        next_index=(
            self.run_controller.complete_current(
                0,stop_requested=self.stop_request_sent
            )
            if self.run_controller is not None
            else None
        )
        self._refresh_queue_tree(selection_index=index)
        self._refresh_progress()
        if next_index is None:
            outcome=self.run_controller.outcome if self.run_controller is not None else "COMPLETE"
            self._finish_queue(outcome)
            return
        if self.stop_request_sent:
            self._finish_queue("STOPPED")
            return
        self._start_queue_item(next_index)

    def _finish_queue(self,outcome):
        if (
            not getattr(self,"queue_running",False)
            or getattr(self,"finalization_in_progress",False)
        ):
            return
        if (
            outcome=="STOPPED"
            and 0<=self.current_file_index<len(self.excel_queue.items)
            and self.excel_queue.items[self.current_file_index].status=="RUNNING"
        ):
            self.excel_queue.items[self.current_file_index].status="STOPPED"
            self.excel_queue.items[self.current_file_index].message="ผู้ใช้สั่งหยุด Queue"
        self.finalization_in_progress=True
        self.queue_status_var.set("กำลังสรุปผล")
        self.append_log("กำลัง refresh สถิติทุก Excel เพื่อสรุปผล Queue...")
        self._update_buttons()
        snapshot=[(item.key,item.path) for item in self.excel_queue.items]
        threading.Thread(
            target=self._finalize_queue_worker,
            args=(outcome,snapshot),
            daemon=True,
        ).start()

    def _finalize_queue_worker(self,requested_outcome,snapshot):
        results=[]
        for key,path in snapshot:
            try:
                stats=get_job_stats(path)
                errors=get_job_errors(path)
                error=""
            except Exception as exc:
                stats={}
                errors=[]
                error=str(exc)
            results.append((key,stats,errors,error))
        self.events.put(("queue_finalized",requested_outcome,results))

    def _on_queue_finalized(self,requested_outcome,results):
        if not getattr(self,"queue_running",False):
            return
        by_key={item.key:item for item in self.excel_queue.items}
        refresh_failed=False
        for key,stats,errors,error in results:
            item=by_key.get(key)
            if item is None:
                refresh_failed=True
                continue
            if stats:
                item.stats=dict(stats)
            item.error_jobs=list(errors)
            if error:
                refresh_failed=True
                item.status="ERROR"
                item.message=(
                    f"{item.message} | " if item.message else ""
                )+f"refresh final stats ไม่สำเร็จ: {error}"
        summary=build_queue_summary(self.excel_queue.items)
        final_outcome=determine_final_outcome(
            requested_outcome,summary,refresh_failed=refresh_failed
        )
        self.queue_running=False
        self.finalization_in_progress=False
        self.excel_queue.locked=False
        self.start_pending=False
        self.precheck_valid=False
        self.valid_excel=False
        self.queue_status_var.set(
            {
                "COMPLETE_SUCCESS":"Completed",
                "COMPLETE_WITH_ERRORS":"Completed with errors",
                "STOPPED":"Stopped",
                "INCOMPLETE":"Incomplete",
                "FAILED":"Error",
            }.get(final_outcome,final_outcome)
        )
        if final_outcome=="COMPLETE_SUCCESS":
            self._set_status("สำเร็จ")
            self.append_log("Queue ทำงานครบทุกไฟล์แล้ว")
        elif final_outcome=="COMPLETE_WITH_ERRORS":
            self._set_status("เสร็จสิ้นพร้อมข้อผิดพลาด")
            self.append_log("Queue ทำงานครบ แต่ Summary ยังพบ JOB ERROR")
        elif final_outcome=="STOPPED":
            self._set_status("หยุดแล้ว")
            self.append_log("Queue หยุดแล้ว และจะไม่เริ่มไฟล์ถัดไป")
        elif final_outcome=="INCOMPLETE":
            self._set_status("ผลการทำงานไม่ครบ")
            self.append_log("Queue จบโดยไม่ได้ Stop แต่ยังมี WAIT/RUNNING/OTHER เหลือ")
        else:
            self._set_status("Error")
            self.append_log("Queue หยุดทันทีเนื่องจาก Error/Recovery/Uncertain state")
        if self._safety_lock_active():
            if self._metadata_fail_closed():
                self._set_status(SAFETY_PERSISTENCE_ERROR)
            else:
                self._set_status("ต้องตรวจสอบงานก่อนเริ่มใหม่")
            self._refresh_safety_status()
            self.append_log(
                "Safety Lock ยัง active; Start จะถูกปิดจนกว่าจะกดตรวจสอบหลังแก้ไขและผ่าน"
            )
        self._refresh_queue_tree(selection_index=self.current_file_index)
        self._refresh_progress()
        self._update_buttons()
        text,summary=format_queue_summary(self.excel_queue.items,final_outcome)
        if not self.closing:
            if final_outcome=="COMPLETE_SUCCESS":
                messagebox.showinfo("Run Complete",text,parent=self.root)
            else:
                messagebox.showwarning("Queue Summary",text,parent=self.root)
        elif self.process is None and not self.recovery_in_progress:
            self.root.after(100,self._destroy)

    def _reconcile_after_exit(self,excel_path,current_ref,reason,expected_no_active):
        try:
            result=reconcile_process_exit(
                excel_path,
                current_ref,
                precommit_result=reason,
                expected_no_active=expected_no_active,
            )
        except Exception as exc:
            result={
                "outcome":"failed",
                "message":str(exc),
                "verified":False,
                "job_ref":current_ref,
                "job_reset":False,
                "services_reset":0,
            }
        self.events.put(("reconciliation",result,self.last_return_code))

    def _on_reconciliation(self,payload,return_code):
        self.recovery_in_progress=False
        outcome=payload.get("outcome","failed")
        verified=bool(payload.get("verified"))
        running_count=payload.get("running_count")
        ref=payload.get("job_ref") or self.current_job_ref or "ไม่พบ JOB เป้าหมาย"
        self.append_log(
            "Reconciliation: "
            f"outcome={outcome}, JOB={ref}, verified={verified}, "
            f"RUNNING={running_count if running_count is not None else 'unknown'}, "
            f"JOB reset={payload.get('job_reset',False)}, "
            f"service reset={payload.get('services_reset',0)}"
        )
        warning=payload.get("warning")
        if warning:
            self.append_log(f"คำเตือน recovery: {warning}")

        no_running=running_count==0
        safe=verified and no_running and outcome in {"recovered","already_clean"}
        uncertain=outcome=="uncertain_commit" and verified and no_running
        # A timeout fallback is inherently uncertain even when Excel happens to
        # look clean: the worker may have been killed before its next workbook
        # write while an editable JOB form is still open in T.V.C.
        dirty_form_possible=safe and (
            outcome=="recovered" or self.force_stop_used
        )
        if dirty_form_possible:
            self.recovery_failed=True
            self._set_status("ต้องตรวจสอบงานก่อนเริ่มใหม่")
            message=(
                "Bot ถูกหยุดผิดปกติ อาจมีข้อมูลค้างในหน้า T.V.C\n"
                "กรุณาตรวจสอบ/ปิดหน้าใบงานก่อนเริ่มใหม่"
            )
            self.append_log(message.replace("\n"," | "))
            self.closing=False
        elif uncertain:
            self.recovery_failed=True
            self._set_status("Error")
            message=payload.get("message") or "UNCERTAIN_TVC_SAVE"
            self.append_log(message)
            messagebox.showerror(
                "ต้องตรวจสอบ T.V.C ก่อน retry",
                message,
                parent=self.root,
            )
            self.closing=False
        elif not safe:
            self.recovery_failed=True
            self._set_status("Error")
            message=payload.get("message") or "reconciliation ไม่สำเร็จ"
            self.append_log(message)
            messagebox.showerror(
                "ตรวจสถานะไม่สำเร็จ",
                message,
                parent=self.root,
            )
            self.closing=False
        else:
            self.recovery_failed=False
            user_stop=(
                self.stop_event_seen
                or self.force_stop_used
                or return_code==2
            )
            if user_stop and self.last_batch_success is not False:
                self._set_status("หยุดแล้ว")
                self.append_log("Bot หยุดและตรวจสถานะ Excel เรียบร้อยแล้ว")
            else:
                self._set_status("Error")
                self.append_log(
                    f"Bot จบผิดปกติด้วยรหัส {return_code}; "
                    "Excel ถูก reconcile และ verify แล้ว"
                )

        if (dirty_form_possible or uncertain or not safe) and hasattr(self,"safety_locks"):
            lock_path=(
                self.excel_queue.items[self.current_file_index].path
                if getattr(self,"queue_running",False)
                and 0<=self.current_file_index<len(self.excel_queue.items)
                else Path(self.excel_var.get())
            )
            self._set_safety_lock(
                lock_path,
                "DIRTY_TVC_FORM_POSSIBLE" if dirty_form_possible else outcome,
                (
                    "Bot ถูกหยุดผิดปกติ อาจมีข้อมูลค้างในหน้า T.V.C "
                    "กรุณาตรวจสอบ/ปิดหน้าใบงานก่อนเริ่มใหม่"
                    if dirty_form_possible
                    else payload.get("message") or "ต้องตรวจสอบงานก่อนเริ่มใหม่"
                ),
                job_ref=payload.get("job_ref") or self.current_job_ref,
            )
            if self._metadata_fail_closed():
                self._set_status(SAFETY_PERSISTENCE_ERROR)
                self.append_log(
                    "Recovery ยังไม่ถือว่าปลอดภัย: Safety Lock persistence ไม่สำเร็จ; "
                    "ห้าม restart เพื่อข้ามการตรวจสอบ"
                )
            else:
                self._set_status("ต้องตรวจสอบงานก่อนเริ่มใหม่")

        if getattr(self,"queue_running",False):
            index=self.current_file_index
            item=self.excel_queue.items[index]
            user_stop=(
                self.stop_event_seen
                or self.force_stop_used
                or return_code==2
            )
            queue_outcome=(
                "STOPPED"
                if safe and user_stop and not self._metadata_fail_closed()
                else "ERROR"
            )
            item.status="STOPPED" if queue_outcome=="STOPPED" else "ERROR"
            item.message=payload.get("message") or (
                "หยุดโดยผู้ใช้" if queue_outcome=="STOPPED" else f"Recovery: {outcome}"
            )
            if self.run_controller is not None:
                self.run_controller.complete_current(
                    return_code,stop_requested=(queue_outcome=="STOPPED")
                )
            threading.Thread(
                target=self._refresh_queue_recovery_result,
                args=(index,item.path,queue_outcome),
                daemon=True,
            ).start()
            self._refresh_queue_tree(selection_index=index)
            self._update_buttons()
            return

        self._update_buttons()
        self._request_stats()
        if self.closing and safe:
            self.root.after(100,self._destroy)

    def _refresh_queue_recovery_result(self,index,path,queue_outcome):
        try:
            stats=get_job_stats(path)
            errors=get_job_errors(path)
            error=""
        except Exception as exc:
            stats={}
            errors=[]
            error=str(exc)
        self.events.put(
            ("queue_recovery_result",index,queue_outcome,stats,errors,error)
        )

    def _on_queue_recovery_result(self,index,queue_outcome,stats,errors,error):
        if not self.queue_running or not (0<=index<len(self.excel_queue.items)):
            return
        item=self.excel_queue.items[index]
        if stats:
            item.stats=dict(stats)
        item.error_jobs=list(errors)
        if error:
            item.message=f"{item.message} | อ่านผลหลัง recovery ไม่สำเร็จ: {error}"
            queue_outcome="ERROR"
            item.status="ERROR"
        self._finish_queue(queue_outcome)

    def _on_force_termination_failed(self,process,message):
        if self.process is not process or process.poll() is not None:
            return
        self.force_stop_used=False
        self.force_retry_available=True
        self.stop_request_sent=False
        self._set_status("หยุดไม่สำเร็จ - process ยังทำงาน")
        self.append_log(f"บังคับหยุดไม่สำเร็จ: {message}")
        self._update_buttons()

    def on_close(self):
        if self.recovery_in_progress:
            self.closing=True
            self.append_log("กำลังตรวจสถานะ Excel หลังบังคับหยุด กรุณารอสักครู่...")
            return
        if self._is_running() or getattr(self,"queue_running",False):
            confirmed=messagebox.askyesno(
                "Bot กำลังทำงาน",
                "ต้องการหยุด Queue หลัง safe checkpoint และปิดโปรแกรมหรือไม่?",
                parent=self.root,
            )
            if not confirmed:
                return
            self.closing=True
            self._update_buttons()
            self._request_stop()
            return
        self._destroy()

    def _destroy(self):
        if self.stop_file is not None:
            try:
                self.stop_file.unlink(missing_ok=True)
            except OSError:
                pass
        self.root.destroy()


def main(mutex_factory=WindowsSingleInstance):
    guard=None
    root=None
    try:
        try:
            guard=mutex_factory()
            acquired=guard.acquire()
        except Exception as exc:
            show_startup_error(root,exc)
            return 1
        if not acquired:
            try:
                root=tk.Tk()
                root.withdraw()
                messagebox.showwarning(
                    APP_NAME,
                    "T.V.C JOB BOT เปิดใช้งานอยู่แล้ว",
                    parent=root,
                )
            except Exception as exc:
                show_startup_error(root,exc)
                root=None
                return 1
            finally:
                if root is not None:
                    try:
                        root.destroy()
                    except Exception:
                        pass
                    root=None
            return 0

        try:
            root=tk.Tk()
            root.withdraw()
            root.title(f"{APP_NAME} v{APP_VERSION}")
            if RUNTIME_PATH_ERROR:
                raise RuntimeError(RUNTIME_PATH_ERROR)
            initialize_gui_dependencies()
            TVCControlApp(root)
        except Exception as exc:
            show_startup_error(root,exc)
            root=None
            return 1
        root.deiconify()
        root.mainloop()
        return 0
    finally:
        if guard is not None:
            try:
                guard.release()
            except Exception:
                pass


if __name__=="__main__":
    raise SystemExit(main())
