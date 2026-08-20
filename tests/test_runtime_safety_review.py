from pathlib import Path
import os
import subprocess
import sys
import tempfile
import unittest
from unittest import mock


SRC=Path(__file__).resolve().parents[1]/"src"
sys.path.insert(0,str(SRC))

import runtime_paths
from runtime_paths import probe_writable_directory,resolve_runtime_paths
import tvc_control
from single_instance import GUI_MUTEX_NAME,WindowsSingleInstance


class Value:
    def __init__(self,value=""):
        self.value=value

    def set(self,value):
        self.value=value

    def get(self):
        return self.value


class FakeMutexApi:
    def __init__(self):
        self.references={}
        self.handles={}
        self.next_handle=1

    def create(self,name):
        already_exists=self.references.get(name,0)>0
        handle=self.next_handle
        self.next_handle+=1
        self.references[name]=self.references.get(name,0)+1
        self.handles[handle]=name
        return handle,already_exists

    def close(self,handle):
        name=self.handles.pop(handle)
        remaining=self.references[name]-1
        if remaining:
            self.references[name]=remaining
        else:
            del self.references[name]

    def crash(self,handle):
        self.close(handle)


class SingleInstanceTests(unittest.TestCase):
    @unittest.skipUnless(os.name=="nt","Windows named mutex integration")
    def test_real_windows_mutex_blocks_another_process_and_clears_after_crash(self):
        code=(
            f"import sys;sys.path.insert(0,{str(SRC)!r});"
            "from single_instance import WindowsSingleInstance;"
            "m=WindowsSingleInstance();"
            "print('ACQUIRED' if m.acquire() else 'BLOCKED',flush=True);"
            "sys.stdin.readline();m.release()"
        )
        first=subprocess.Popen(
            [sys.executable,"-c",code],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        try:
            self.assertEqual(first.stdout.readline().strip(),"ACQUIRED")
            blocked=subprocess.run(
                [sys.executable,"-c",code],
                input="\n",
                capture_output=True,
                text=True,
                timeout=10,
            )
            self.assertEqual(blocked.returncode,0,blocked.stderr)
            self.assertEqual(blocked.stdout.strip(),"BLOCKED")
            first.kill()
            first.communicate(timeout=10)
            after_crash=subprocess.run(
                [sys.executable,"-c",code],
                input="\n",
                capture_output=True,
                text=True,
                timeout=10,
            )
            self.assertEqual(after_crash.returncode,0,after_crash.stderr)
            self.assertEqual(after_crash.stdout.strip(),"ACQUIRED")
        finally:
            if first.poll() is None:
                first.kill()
                first.communicate(timeout=10)
            for stream in (first.stdin,first.stdout,first.stderr):
                if stream is not None:
                    stream.close()

    def test_second_instance_is_blocked_then_opens_after_release(self):
        api=FakeMutexApi()
        first=WindowsSingleInstance(api=api)
        second=WindowsSingleInstance(api=api)
        self.assertTrue(first.acquire())
        self.assertFalse(second.acquire())
        first.release()
        self.assertTrue(second.acquire())
        second.release()

    def test_crashed_owner_does_not_leave_permanent_stale_lock(self):
        api=FakeMutexApi()
        first=WindowsSingleInstance(api=api)
        self.assertTrue(first.acquire())
        handle=first._handle
        api.crash(handle)
        first._handle=None
        replacement=WindowsSingleInstance(api=api)
        self.assertTrue(replacement.acquire())
        replacement.release()

    def test_mutex_identity_is_same_for_source_and_frozen_launches(self):
        source=WindowsSingleInstance()
        frozen=WindowsSingleInstance()
        self.assertEqual(source.name,GUI_MUTEX_NAME)
        self.assertEqual(frozen.name,GUI_MUTEX_NAME)

    def test_second_writer_cannot_overwrite_registry(self):
        api=FakeMutexApi()
        with tempfile.TemporaryDirectory(prefix="tvc two writer registry ") as tmp:
            state_file=Path(tmp)/"safety_locks.json"
            first=WindowsSingleInstance(api=api)
            second=WindowsSingleInstance(api=api)
            first_path=Path(tmp)/"first.xlsx"
            second_path=Path(tmp)/"second.xlsx"
            first_key=os.path.normcase(str(first_path.resolve()))
            second_key=os.path.normcase(str(second_path.resolve()))
            self.assertTrue(first.acquire())
            tvc_control.save_persisted_safety_locks(
                state_file,
                {first_key:{"path":str(first_path),"outcome":"dirty"}},
            )
            self.assertFalse(second.acquire())
            self.assertEqual(
                set(tvc_control.load_persisted_safety_locks(state_file).locks),
                {first_key},
            )
            first.release()
            self.assertTrue(second.acquire())
            current=tvc_control.load_persisted_safety_locks(state_file).locks
            current[second_key]={
                "path":str(second_path),
                "outcome":"dirty",
            }
            tvc_control.save_persisted_safety_locks(state_file,current)
            second.release()
            self.assertEqual(
                set(tvc_control.load_persisted_safety_locks(state_file).locks),
                {first_key,second_key},
            )

    def test_registry_atomic_temp_name_contains_pid_and_uuid(self):
        with tempfile.TemporaryDirectory(prefix="tvc unique metadata temp ") as tmp:
            state_file=Path(tmp)/"safety_locks.json"
            opened=[]

            def fail_open(path,*_args,**_kwargs):
                opened.append(path)
                raise OSError("stop after temp allocation")

            token=type("Token",(),{"hex":"unique_token"})()
            with (
                mock.patch.object(tvc_control.uuid,"uuid4",return_value=token),
                mock.patch.object(Path,"open",new=fail_open),
            ):
                with self.assertRaisesRegex(OSError,"temp allocation"):
                    tvc_control.save_persisted_safety_locks(
                        state_file,
                        {"x":{"path":str(Path(tmp)/"x.xlsx"),"outcome":"dirty"}},
                    )
            self.assertEqual(len(opened),1)
            self.assertEqual(
                opened[0].name,
                f".safety_locks.json.{os.getpid()}.unique_token.tmp",
            )


class WritableDirectoryTests(unittest.TestCase):
    def test_frozen_worker_help_uses_fallback_without_writing_app_directory(self):
        with tempfile.TemporaryDirectory(prefix="tvc read only worker ") as tmp:
            root=Path(tmp)
            app_dir=root/"Program Files"/"TVC JOB BOT"
            resource_dir=root/"Bundled Resources"
            writable=root/"Local App Data"/"TVC_JOB_BOT"
            app_dir.mkdir(parents=True)
            resource_dir.mkdir(parents=True)
            code=f"""
from pathlib import Path
import runpy
import sys
sys.path.insert(0,{str(SRC)!r})
import runtime_paths
app_dir=Path({str(app_dir)!r})
resource_dir=Path({str(resource_dir)!r})
writable=Path({str(writable)!r})
layout=runtime_paths.RuntimePaths(
    frozen=True,
    app_dir=app_dir,
    resource_dir=resource_dir,
    config_file=resource_dir/'config.ini',
    field_map_file=resource_dir/'field_map.json',
    assets_dir=resource_dir/'assets',
    writable_data_dir=writable,
    logs_dir=writable/'logs',
    screenshots_dir=writable/'screenshots',
    runtime_temp_dir=writable/'runtime',
    source_python=None,
    bot_script=None,
    worker_executable=app_dir/'TVC Bot Worker.exe',
)
runtime_paths.resolve_runtime_paths=lambda:layout
original_mkdir=Path.mkdir
def guarded_mkdir(path,*args,**kwargs):
    try:
        path.resolve().relative_to(app_dir.resolve())
    except ValueError:
        return original_mkdir(path,*args,**kwargs)
    raise PermissionError('simulated read-only application directory')
Path.mkdir=guarded_mkdir
try:
    runpy.run_path({str(SRC/'bot_worker.py')!r},run_name='__main__')
finally:
    Path.mkdir=original_mkdir
"""
            result=subprocess.run(
                [sys.executable,"-c",code,"--help"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=15,
            )
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertIn("--excel",result.stdout)
            self.assertTrue((writable/"logs").is_dir())
            self.assertTrue((writable/"screenshots").is_dir())
            self.assertFalse((app_dir/"logs").exists())
            self.assertFalse((app_dir/"screenshots").exists())
            self.assertFalse((app_dir/"runtime").exists())

    def test_real_probe_supports_spaces_and_cleans_up(self):
        with tempfile.TemporaryDirectory(prefix="tvc writable root ") as tmp:
            directory=Path(tmp)/"Writable Data With Spaces"
            self.assertTrue(probe_writable_directory(directory))
            self.assertEqual(list(directory.iterdir()),[])

    def test_probe_detects_create_fsync_replace_and_delete_failures(self):
        with tempfile.TemporaryDirectory(prefix="tvc probe failures ") as tmp:
            directory=Path(tmp)/"data"
            with mock.patch.object(Path,"mkdir",side_effect=OSError("create failed")):
                self.assertFalse(probe_writable_directory(directory))
            self.assertFalse(
                probe_writable_directory(
                    directory,fsync_func=lambda _fd: (_ for _ in ()).throw(OSError("fsync"))
                )
            )
            self.assertFalse(
                probe_writable_directory(
                    directory,replace_func=lambda _src,_dst: (_ for _ in ()).throw(
                        OSError("replace")
                    )
                )
            )
            self.assertFalse(
                probe_writable_directory(
                    directory,unlink_func=lambda _path: (_ for _ in ()).throw(
                        OSError("delete")
                    )
                )
            )

    def test_app_directory_is_used_only_when_real_probe_succeeds(self):
        with tempfile.TemporaryDirectory(prefix="tvc resolver source ") as tmp:
            root=Path(tmp)/"Source App With Spaces"
            (root/"src").mkdir(parents=True)
            paths=resolve_runtime_paths(
                frozen=False,module_file=root/"src"/"runtime_paths.py"
            )
            self.assertEqual(paths.writable_data_dir,root)
            self.assertEqual(paths.runtime_temp_dir,root/"runtime")

    def test_os_access_true_does_not_override_real_probe_failure(self):
        with tempfile.TemporaryDirectory(prefix="tvc local fallback ") as tmp:
            root=Path(tmp)/"Read Only App"
            local=Path(tmp)/"Local App Data"
            calls=[]

            def probe(path):
                calls.append(Path(path))
                return Path(path)==local/"TVC_JOB_BOT"

            with (
                mock.patch.dict(os.environ,{"LOCALAPPDATA":str(local)}),
                mock.patch.object(runtime_paths.os,"access",return_value=True) as access,
            ):
                paths=resolve_runtime_paths(
                    frozen=True,
                    executable=root/"TVC Bot Control.exe",
                    bundle_dir=root/"bundle",
                    writable_probe=probe,
                )
            access.assert_not_called()
            self.assertEqual(calls,[root,local/"TVC_JOB_BOT"])
            self.assertEqual(paths.writable_data_dir,local/"TVC_JOB_BOT")
            self.assertEqual(paths.logs_dir,local/"TVC_JOB_BOT"/"logs")
            self.assertEqual(paths.runtime_temp_dir,local/"TVC_JOB_BOT"/"runtime")
            self.assertEqual(paths.config_file,root/"bundle"/"config.ini")

            stop_path=tvc_control.new_stop_file_path(paths)
            self.assertEqual(stop_path.parent,local/"TVC_JOB_BOT"/"runtime")
            self.assertTrue(stop_path.parent.is_dir())
            self.assertFalse((root/"runtime").exists())

    def test_both_app_and_local_fallback_failure_is_fatal(self):
        with tempfile.TemporaryDirectory(prefix="tvc fatal resolver ") as tmp:
            root=Path(tmp)/"App"
            local=Path(tmp)/"Local"
            with mock.patch.dict(os.environ,{"LOCALAPPDATA":str(local)}):
                with self.assertRaisesRegex(RuntimeError,"writable data directory"):
                    resolve_runtime_paths(
                        frozen=True,
                        executable=root/"TVC Bot Control.exe",
                        bundle_dir=root/"bundle",
                        writable_probe=lambda _path:False,
                    )

    def test_missing_localappdata_after_app_failure_is_fatal(self):
        with tempfile.TemporaryDirectory(prefix="tvc no local resolver ") as tmp:
            root=Path(tmp)/"App"
            with mock.patch.dict(os.environ,{},clear=True):
                with self.assertRaisesRegex(RuntimeError,"LOCALAPPDATA"):
                    resolve_runtime_paths(
                        frozen=False,
                        module_file=root/"src"/"runtime_paths.py",
                        writable_probe=lambda _path:False,
                    )

    def test_source_mode_uses_same_local_fallback_policy(self):
        with tempfile.TemporaryDirectory(prefix="tvc source fallback ") as tmp:
            root=Path(tmp)/"Source App"
            local=Path(tmp)/"Local Data"
            expected=local/"TVC_JOB_BOT"
            with mock.patch.dict(os.environ,{"LOCALAPPDATA":str(local)}):
                paths=resolve_runtime_paths(
                    frozen=False,
                    module_file=root/"src"/"runtime_paths.py",
                    writable_probe=lambda path:Path(path)==expected,
                )
            self.assertEqual(paths.writable_data_dir,expected)
            self.assertEqual(paths.config_file,root/"config.ini")


class SafetyStatusDetailTests(unittest.TestCase):
    @staticmethod
    def app(locks,health=tvc_control.SAFETY_METADATA_HEALTHY,error=""):
        app=tvc_control.TVCControlApp.__new__(tvc_control.TVCControlApp)
        app.safety_locks=dict(locks)
        app.safety_metadata_health=health
        app.safety_metadata_error=error
        app.safety_state_file=Path("C:/Users/Test/AppData/Local/TVC_JOB_BOT/safety_locks.json")
        app.safety_status_var=Value()
        app.safety_detail_var=Value()
        app.recovery_failed=bool(locks)
        app._set_status=lambda _text:None
        app.append_log=lambda _text:None
        return app

    @staticmethod
    def lock(path,reason,lock_type):
        return {
            "path":str(path),
            "message":reason,
            "outcome":lock_type,
            "job_ref":"TEST-001",
        }

    def test_write_failed_shows_metadata_error_and_one_lock_fields_together(self):
        path=Path("C:/Input/jobs one.xlsx")
        app=self.app(
            {"one":self.lock(path,"user stopped","DIRTY_TVC_FORM_POSSIBLE")},
            tvc_control.SAFETY_METADATA_WRITE_FAILED,
            "disk full",
        )
        app._refresh_safety_status()
        detail=app.safety_detail_var.get()
        self.assertIn("health=WRITE_FAILED",detail)
        self.assertIn("error=disk full",detail)
        self.assertIn("filename=jobs one.xlsx",detail)
        self.assertIn(f"path={path}",detail)
        self.assertIn("reason=user stopped",detail)
        self.assertIn("type=DIRTY_TVC_FORM_POSSIBLE",detail)
        self.assertIn("status=UNRESOLVED",detail)

    def test_multiple_locks_are_all_visible(self):
        locks={
            "one":self.lock(Path("C:/Input/one.xlsx"),"first reason","TYPE_ONE"),
            "two":self.lock(Path("C:/Input/two.xlsx"),"second reason","TYPE_TWO"),
        }
        app=self.app(
            locks,
            tvc_control.SAFETY_METADATA_WRITE_FAILED,
            "replace denied",
        )
        app._refresh_safety_status()
        detail=app.safety_detail_var.get()
        self.assertIn("health=WRITE_FAILED",detail)
        self.assertIn("filename=one.xlsx",detail)
        self.assertIn("filename=two.xlsx",detail)
        self.assertEqual(detail.count("status=UNRESOLVED"),2)

    def test_partial_unlock_write_failure_displays_true_memory_state(self):
        first=Path("C:/Input/first.xlsx")
        second=Path("C:/Input/second.xlsx")
        app=self.app({
            "first":self.lock(first,"resolved","TYPE_FIRST"),
            "second":self.lock(second,"still dirty","TYPE_SECOND"),
        })
        app.safety_state_file=Path("C:/No Write/safety_locks.json")
        del app.safety_locks["first"]
        with mock.patch.object(
            tvc_control,
            "save_persisted_safety_locks",
            side_effect=PermissionError("denied"),
        ):
            self.assertFalse(app._persist_safety_locks())
        detail=app.safety_detail_var.get()
        self.assertIn("health=WRITE_FAILED",detail)
        self.assertNotIn("filename=first.xlsx",detail)
        self.assertIn("filename=second.xlsx",detail)
        self.assertIn("reason=still dirty",detail)
        self.assertTrue(app._safety_lock_active())


class GuiFilesystemCallbackTests(unittest.TestCase):
    def test_open_logs_mkdir_failure_is_handled_by_dialog(self):
        app=tvc_control.TVCControlApp.__new__(tvc_control.TVCControlApp)
        app.root=mock.Mock()
        with (
            mock.patch.object(Path,"mkdir",side_effect=PermissionError("denied")),
            mock.patch.object(tvc_control.os,"startfile") as startfile,
            mock.patch.object(tvc_control.messagebox,"showerror") as dialog,
        ):
            app.open_logs()
        startfile.assert_not_called()
        dialog.assert_called_once()
        self.assertEqual(
            dialog.call_args.args[0],
            "ไม่สามารถเปิดโฟลเดอร์ Logs ได้",
        )
        self.assertIn("denied",dialog.call_args.args[1])

    def test_open_logs_startfile_failure_is_handled_by_dialog(self):
        app=tvc_control.TVCControlApp.__new__(tvc_control.TVCControlApp)
        app.root=mock.Mock()
        with (
            mock.patch.object(Path,"mkdir"),
            mock.patch.object(
                tvc_control.os,
                "startfile",
                side_effect=OSError("shell open failed"),
            ),
            mock.patch.object(tvc_control.messagebox,"showerror") as dialog,
        ):
            app.open_logs()
        dialog.assert_called_once()
        self.assertEqual(
            dialog.call_args.args[0],
            "ไม่สามารถเปิดโฟลเดอร์ Logs ได้",
        )
        self.assertIn("shell open failed",dialog.call_args.args[1])


if __name__=="__main__":
    unittest.main(verbosity=2)
