"""Resolve source and frozen runtime paths without embedding machine-specific paths."""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import sys
import uuid


GUI_EXECUTABLE_NAME = "TVC Bot Control.exe"
WORKER_EXECUTABLE_NAME = "TVC Bot Worker.exe"
LOCAL_DATA_DIRECTORY_NAME = "TVC_JOB_BOT"


@dataclass(frozen=True)
class RuntimePaths:
    frozen: bool
    app_dir: Path
    resource_dir: Path
    config_file: Path
    field_map_file: Path
    assets_dir: Path
    writable_data_dir: Path
    logs_dir: Path
    screenshots_dir: Path
    runtime_temp_dir: Path
    source_python: Path | None
    bot_script: Path | None
    worker_executable: Path | None


def _external_or_bundled(app_dir: Path,resource_dir: Path,name: str) -> Path:
    external=app_dir/name
    if external.exists() or app_dir==resource_dir:
        return external
    return resource_dir/name


def probe_writable_directory(
    directory: str | Path,
    *,
    fsync_func=None,
    replace_func=None,
    unlink_func=None,
) -> bool:
    """Exercise the writes runtime state needs, including atomic replacement."""
    directory=Path(directory)
    fsync_func=fsync_func or os.fsync
    replace_func=replace_func or os.replace
    unlink_func=unlink_func or (lambda path: path.unlink())
    token=f"{os.getpid()}_{uuid.uuid4().hex}"
    initial=directory/f".tvc_write_probe_{token}.initial"
    replacement=directory/f".tvc_write_probe_{token}.replacement"
    target=directory/f".tvc_write_probe_{token}.target"
    success=False
    try:
        directory.mkdir(parents=True,exist_ok=True)
        with initial.open("wb") as stream:
            stream.write(b"initial")
            stream.flush()
            fsync_func(stream.fileno())
        with replacement.open("wb") as stream:
            stream.write(b"replacement")
            stream.flush()
            fsync_func(stream.fileno())
        replace_func(initial,target)
        replace_func(replacement,target)
        if target.read_bytes()!=b"replacement":
            raise OSError("writable-directory probe readback mismatch")
        unlink_func(target)
        success=True
    except Exception:
        success=False
    finally:
        for candidate in (initial,replacement,target):
            try:
                if candidate.exists():
                    unlink_func(candidate)
            except Exception:
                success=False
    return success


def _writable_data_root(app_dir: Path,writable_probe=None) -> Path:
    probe=writable_probe or probe_writable_directory
    if probe(app_dir):
        return app_dir
    local_app_data=os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise RuntimeError(
            "ไม่พบ writable data directory: app directory ใช้งานไม่ได้และไม่พบ LOCALAPPDATA"
        )
    fallback=Path(local_app_data)/LOCAL_DATA_DIRECTORY_NAME
    if probe(fallback):
        return fallback
    raise RuntimeError(
        "ไม่พบ writable data directory: ทั้ง app directory และ "
        f"LOCALAPPDATA fallback ใช้งานไม่ได้ ({fallback})"
    )


def resolve_runtime_paths(
    *,
    frozen: bool | None=None,
    executable: str | Path | None=None,
    module_file: str | Path | None=None,
    bundle_dir: str | Path | None=None,
    writable_probe=None,
) -> RuntimePaths:
    """Return paths for source/dev mode or a simulated/real PyInstaller runtime."""
    is_frozen=bool(getattr(sys,"frozen",False)) if frozen is None else bool(frozen)
    module_path=Path(module_file or __file__).resolve()

    if is_frozen:
        app_dir=Path(executable or sys.executable).resolve().parent
        detected_bundle=bundle_dir or getattr(sys,"_MEIPASS",app_dir)
        resource_dir=Path(detected_bundle).resolve()
        source_python=None
        bot_script=None
        worker_executable=app_dir/WORKER_EXECUTABLE_NAME
    else:
        app_dir=module_path.parents[1]
        resource_dir=app_dir
        source_python=app_dir/".venv"/"Scripts"/"python.exe"
        bot_script=app_dir/"src"/"bot.py"
        worker_executable=None

    writable_data_dir=_writable_data_root(app_dir,writable_probe)
    return RuntimePaths(
        frozen=is_frozen,
        app_dir=app_dir,
        resource_dir=resource_dir,
        config_file=_external_or_bundled(app_dir,resource_dir,"config.ini"),
        field_map_file=_external_or_bundled(app_dir,resource_dir,"field_map.json"),
        assets_dir=_external_or_bundled(app_dir,resource_dir,"assets"),
        writable_data_dir=writable_data_dir,
        logs_dir=writable_data_dir/"logs",
        screenshots_dir=writable_data_dir/"screenshots",
        runtime_temp_dir=writable_data_dir/"runtime",
        source_python=source_python,
        bot_script=bot_script,
        worker_executable=worker_executable,
    )


def build_worker_command(
    paths: RuntimePaths,
    excel_path: str | Path,
    stop_file: str | Path,
    *,
    executable: str | Path | None=None,
) -> list[str]:
    """Build an argv list; keeping it as a list safely supports spaces in paths."""
    if paths.frozen:
        target=Path(executable or paths.worker_executable or "")
        command=[str(target)]
    else:
        target=Path(executable or paths.source_python or "")
        if paths.bot_script is None:
            raise RuntimeError("ไม่พบ source bot.py")
        command=[str(target),"-u",str(paths.bot_script)]
    return command+["--excel",str(Path(excel_path)),"--stop-file",str(Path(stop_file))]
