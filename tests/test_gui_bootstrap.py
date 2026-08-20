from pathlib import Path
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock


SRC=Path(__file__).resolve().parents[1]/"src"
sys.path.insert(0,str(SRC))

import tvc_control


class FakeRoot:
    def __init__(self,events):
        self.events=events
        self.events.append("root_created")

    def withdraw(self):
        self.events.append("withdraw")

    def title(self,_text):
        self.events.append("title")

    def deiconify(self):
        self.events.append("deiconify")

    def update_idletasks(self):
        self.events.append("update_idletasks")

    def destroy(self):
        self.events.append("destroy")

    def mainloop(self):
        self.events.append("mainloop")


class FakeMutex:
    def __init__(self,acquired=True):
        self.should_acquire=acquired
        self.released=False

    def acquire(self):
        return self.should_acquire

    def release(self):
        self.released=True


class BootstrapTests(unittest.TestCase):
    @staticmethod
    def _importer_missing(missing):
        def importer(name):
            if name==missing:
                if missing=="excel_io":
                    raise ImportError("mock excel_io import failure")
                raise ModuleNotFoundError(f"No module named '{missing}'",name=missing)
            if name=="excel_io":
                return type(
                    "ExcelModule",
                    (),
                    {
                        "get_job_stats":staticmethod(lambda *_args,**_kwargs:{}),
                        "get_job_errors":staticmethod(lambda *_args,**_kwargs:[]),
                        "get_safety_issues":staticmethod(lambda *_args,**_kwargs:[]),
                        "inspect_recovery_state":staticmethod(lambda *_args,**_kwargs:{}),
                        "reconcile_process_exit":staticmethod(lambda *_args,**_kwargs:{}),
                        "validate_workbook":staticmethod(lambda *_args,**_kwargs:True),
                        "DIRTY_TVC_FORM_POSSIBLE":"DIRTY_TVC_FORM_POSSIBLE",
                    },
                )
            if name=="tvc_probe":
                return type(
                    "ProbeModule",
                    (),
                    {"probe":staticmethod(lambda:{
                        "connected":True,
                        "login_verified":True,
                        "active_job_form":False,
                    })},
                )
            return object()
        return importer

    def test_module_imports_without_third_party_site_packages(self):
        console_python=Path(sys.base_prefix)/"python.exe"
        if not console_python.is_file() or console_python.stat().st_size<=0:
            self.skipTest("ไม่มี console Python สำหรับ startup subprocess")
        code=(
            f"import sys;sys.path.insert(0,{str(SRC)!r});"
            "import tvc_control;print('BOOTSTRAP_OK')"
        )
        result=subprocess.run(
            [str(console_python),"-S","-c",code],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertIn("BOOTSTRAP_OK",result.stdout)

    def test_startup_missing_psutil_has_specific_error(self):
        with self.assertRaises(tvc_control.StartupDependencyError) as captured:
            tvc_control.initialize_gui_dependencies(self._importer_missing("psutil"))
        self.assertEqual(captured.exception.dependency,"psutil")
        self.assertIn("ไม่พบ psutil",str(captured.exception))

    def test_startup_missing_openpyxl_has_specific_error(self):
        with self.assertRaises(tvc_control.StartupDependencyError) as captured:
            tvc_control.initialize_gui_dependencies(self._importer_missing("openpyxl"))
        self.assertEqual(captured.exception.dependency,"openpyxl")
        self.assertIn("ไม่พบ openpyxl",str(captured.exception))

    def test_startup_excel_io_import_failure_has_specific_error(self):
        with self.assertRaises(tvc_control.StartupDependencyError) as captured:
            tvc_control.initialize_gui_dependencies(self._importer_missing("excel_io"))
        self.assertEqual(captured.exception.dependency,"excel_io")
        self.assertIn("excel_io",str(captured.exception))

    def test_root_and_error_dialog_exist_before_startup_exit(self):
        events=[]
        root=FakeRoot(events)

        def fail_after_root():
            events.append("dependency_check")
            raise tvc_control.StartupDependencyError("psutil","mock missing")

        def show_error(*_args,**_kwargs):
            events.append("dialog")

        with (
            mock.patch.object(tvc_control.tk,"Tk",return_value=root),
            mock.patch.object(
                tvc_control,
                "initialize_gui_dependencies",
                side_effect=fail_after_root,
            ),
            mock.patch.object(tvc_control.messagebox,"showerror",side_effect=show_error),
            mock.patch.object(tvc_control,"write_startup_error_log"),
            mock.patch.object(tvc_control,"TVCControlApp") as app,
        ):
            result=tvc_control.main()

        self.assertEqual(result,1)
        app.assert_not_called()
        self.assertLess(events.index("root_created"),events.index("dependency_check"))
        self.assertLess(events.index("dependency_check"),events.index("dialog"))
        self.assertLess(events.index("dialog"),events.index("destroy"))

    def test_successful_bootstrap_enters_gui_mainloop(self):
        events=[]
        root=FakeRoot(events)
        with (
            mock.patch.object(tvc_control.tk,"Tk",return_value=root),
            mock.patch.object(tvc_control,"initialize_gui_dependencies",return_value=True),
            mock.patch.object(tvc_control,"TVCControlApp") as app,
        ):
            result=tvc_control.main()
        self.assertEqual(result,0)
        app.assert_called_once_with(root)
        self.assertIn("mainloop",events)

    def test_second_instance_shows_exact_message_and_exits_before_app(self):
        events=[]
        root=FakeRoot(events)
        mutex=FakeMutex(acquired=False)
        with (
            mock.patch.object(tvc_control.tk,"Tk",return_value=root),
            mock.patch.object(tvc_control.messagebox,"showwarning") as warning,
            mock.patch.object(tvc_control,"TVCControlApp") as app,
        ):
            result=tvc_control.main(mutex_factory=lambda:mutex)
        self.assertEqual(result,0)
        warning.assert_called_once()
        self.assertEqual(
            warning.call_args.args[1],
            "T.V.C JOB BOT เปิดใช้งานอยู่แล้ว",
        )
        app.assert_not_called()
        self.assertTrue(mutex.released)

    def test_mutex_is_released_after_startup_failure(self):
        events=[]
        root=FakeRoot(events)
        mutex=FakeMutex(acquired=True)
        with (
            mock.patch.object(tvc_control.tk,"Tk",return_value=root),
            mock.patch.object(
                tvc_control,
                "initialize_gui_dependencies",
                side_effect=tvc_control.StartupDependencyError("psutil","missing"),
            ),
            mock.patch.object(tvc_control.messagebox,"showerror"),
            mock.patch.object(tvc_control,"write_startup_error_log"),
        ):
            result=tvc_control.main(mutex_factory=lambda:mutex)
        self.assertEqual(result,1)
        self.assertTrue(mutex.released)

    def test_unusable_app_and_fallback_is_fatal_before_dependencies(self):
        events=[]
        root=FakeRoot(events)
        mutex=FakeMutex(acquired=True)
        with (
            mock.patch.object(tvc_control.tk,"Tk",return_value=root),
            mock.patch.object(tvc_control,"RUNTIME_PATH_ERROR","no writable root"),
            mock.patch.object(tvc_control,"initialize_gui_dependencies") as dependencies,
            mock.patch.object(tvc_control.messagebox,"showerror") as dialog,
            mock.patch.object(tvc_control,"write_startup_error_log"),
        ):
            result=tvc_control.main(mutex_factory=lambda:mutex)
        self.assertEqual(result,1)
        dependencies.assert_not_called()
        self.assertIn("no writable root",dialog.call_args.args[1])
        self.assertTrue(mutex.released)

    def test_constructor_failure_shows_startup_dialog_destroys_root_and_releases_mutex(self):
        events=[]
        root=FakeRoot(events)
        mutex=FakeMutex(acquired=True)

        def show_error(*_args,**_kwargs):
            events.append("dialog")

        with (
            mock.patch.object(tvc_control.tk,"Tk",return_value=root),
            mock.patch.object(tvc_control,"initialize_gui_dependencies",return_value=True),
            mock.patch.object(
                tvc_control,
                "TVCControlApp",
                side_effect=PermissionError("logs directory denied"),
            ),
            mock.patch.object(tvc_control.messagebox,"showerror",side_effect=show_error) as dialog,
            mock.patch.object(tvc_control,"write_startup_error_log") as error_log,
        ):
            result=tvc_control.main(mutex_factory=lambda:mutex)

        self.assertEqual(result,1)
        error_log.assert_called_once()
        self.assertIsInstance(error_log.call_args.args[0],PermissionError)
        self.assertEqual(dialog.call_args.args[0],tvc_control.STARTUP_ERROR_TITLE)
        self.assertIn("logs directory denied",dialog.call_args.args[1])
        self.assertIn("destroy",events)
        self.assertNotIn("mainloop",events)
        self.assertTrue(mutex.released)

    def test_tk_root_creation_failure_is_clean_and_releases_mutex(self):
        mutex=FakeMutex(acquired=True)
        with (
            mock.patch.object(
                tvc_control.tk,
                "Tk",
                side_effect=PermissionError("Tk unavailable"),
            ),
            mock.patch.object(tvc_control.messagebox,"showerror") as dialog,
            mock.patch.object(tvc_control,"write_startup_error_log"),
        ):
            result=tvc_control.main(mutex_factory=lambda:mutex)
        self.assertEqual(result,1)
        self.assertEqual(dialog.call_args.args[0],tvc_control.STARTUP_ERROR_TITLE)
        self.assertIn("Tk unavailable",dialog.call_args.args[1])
        self.assertTrue(mutex.released)

    def test_mutex_acquire_failure_is_reported_and_released(self):
        events=[]
        root=FakeRoot(events)
        mutex=FakeMutex(acquired=True)
        mutex.acquire=mock.Mock(side_effect=OSError("mutex denied"))
        with (
            mock.patch.object(tvc_control.tk,"Tk",return_value=root),
            mock.patch.object(tvc_control.messagebox,"showerror") as dialog,
            mock.patch.object(tvc_control,"write_startup_error_log"),
        ):
            result=tvc_control.main(mutex_factory=lambda:mutex)
        self.assertEqual(result,1)
        self.assertIn("mutex denied",dialog.call_args.args[1])
        self.assertIn("destroy",events)
        self.assertTrue(mutex.released)

    def test_startup_error_logger_writes_traceback_when_log_path_works(self):
        with tempfile.TemporaryDirectory(prefix="tvc startup error log ") as tmp:
            runtime=SimpleNamespace(logs_dir=Path(tmp)/"logs")
            with mock.patch.object(tvc_control,"RUNTIME_PATHS",runtime):
                try:
                    raise PermissionError("constructor denied")
                except PermissionError as exc:
                    tvc_control.write_startup_error_log(exc)
            log_file=runtime.logs_dir/"gui_startup_error.log"
            self.assertTrue(log_file.is_file())
            detail=log_file.read_text(encoding="utf-8")
            self.assertIn("PermissionError",detail)
            self.assertIn("constructor denied",detail)


if __name__=="__main__":
    unittest.main(verbosity=2)
