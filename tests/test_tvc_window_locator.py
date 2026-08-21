from pathlib import Path
from types import SimpleNamespace
import sys
import tempfile
import unittest
from unittest import mock


SRC=Path(__file__).resolve().parents[1]/"src"
sys.path.insert(0,str(SRC))

import tvc_driver
import tvc_probe
import tvc_window_locator


LOGGED_IN_TITLE=(
    "T.V.C Client [Version 1.8.2] "
    "\u0e40\u0e02\u0e49\u0e32\u0e2a\u0e39\u0e48\u0e23\u0e30\u0e1a\u0e1a"
    "\u0e42\u0e14\u0e22 ADMINPREMIUM []"
)


class Info:
    def __init__(self,handle=1234,process_id=4321,automation_id=""):
        self.handle=handle
        self.process_id=process_id
        self.automation_id=automation_id


class Window:
    def __init__(self,title,class_name="WindowsForms10.Window",handle=1234,pid=4321):
        self.title=title
        self._class_name=class_name
        self.element_info=Info(handle,process_id=pid)

    def window_text(self):
        return self.title

    def class_name(self):
        return self._class_name

    def children(self):
        raise AssertionError("top-level locator must not inspect controls")


class DesktopFixture:
    def __init__(self,windows):
        self._windows=list(windows)

    def windows(self,**_kwargs):
        return list(self._windows)


class SharedWindowLocatorTests(unittest.TestCase):
    def test_verified_main_title_family_matches_actual_caption(self):
        self.assertTrue(tvc_window_locator.is_tvc_main_title(LOGGED_IN_TITLE))
        self.assertTrue(
            tvc_window_locator.is_tvc_main_title("T.V.C Client [Version 1.8.2]")
        )
        self.assertFalse(
            tvc_window_locator.is_tvc_main_title("Report for T.V.C Client")
        )

    def test_unrelated_class_containing_tvc_is_ignored(self):
        unrelated=Window("Unrelated application",class_name="NotVCWindow")
        result=tvc_window_locator.locate_tvc_main_window(
            "win32",
            desktop_factory=lambda **_kwargs:DesktopFixture([unrelated]),
        )
        self.assertEqual(result.status,"NOT_FOUND")
        self.assertIsNone(result.selected)

    def test_probe_and_driver_resolve_same_win32_window(self):
        main=Window(LOGGED_IN_TITLE)
        unrelated=Window("Visual Studio Code",handle=77,pid=88)
        calls=[]

        def desktop_factory(*,backend):
            calls.append(backend)
            if backend!="win32":
                raise AssertionError("positive main title must not require UIA")
            return DesktopFixture([unrelated,main])

        located=tvc_window_locator.locate_tvc_main_window(
            "win32",desktop_factory=desktop_factory
        )
        self.assertIs(located.selected.window,main)
        probe_result=tvc_probe.probe(desktop_factory=desktop_factory)
        self.assertEqual(probe_result["status"],"READY")
        self.assertIn("logged_in_user_title",probe_result["login_signals"])
        self.assertEqual(probe_result["main_window"]["handle"],main.element_info.handle)
        self.assertEqual(probe_result["main_window"]["pid"],main.element_info.process_id)
        self.assertEqual(calls,["win32","win32"])

        driver=tvc_driver.TVCDriver(r"^เพิ่มใบงาน \(JOB\)$",backend="win32")
        with mock.patch.object(
            tvc_driver.tvc_window_locator,
            "locate_tvc_main_window",
            return_value=located,
        ) as locator:
            self.assertIs(driver._find_tvc_main_window(),main)
        locator.assert_called_once_with("win32",timeout_seconds=1.5)

    def test_diagnostic_reports_attach_metadata_without_control_walk(self):
        main=Window(LOGGED_IN_TITLE,handle=9012,pid=3456)
        with tempfile.TemporaryDirectory(prefix="tvc_locator_") as temp_dir:
            config_file=Path(temp_dir)/"config.ini"
            config_file.write_text("[tvc]\nbackend = win32\n",encoding="utf-8")
            result=tvc_window_locator.diagnose_tvc_main_window(
                SimpleNamespace(config_file=config_file),
                desktop_factory=lambda **_kwargs:DesktopFixture([main]),
            )
        self.assertEqual(result["status"],"FOUND")
        self.assertTrue(result["driver_can_attach"])
        self.assertEqual(result["backend"],"win32")
        self.assertEqual(result["selected"]["handle"],9012)
        self.assertEqual(result["selected"]["pid"],3456)

    def test_deadline_is_explicit(self):
        class Clock:
            def __init__(self):
                self.value=0.0
            def __call__(self):
                self.value+=0.2
                return self.value

        result=tvc_window_locator.locate_tvc_main_window(
            "win32",
            timeout_seconds=0.1,
            desktop_factory=lambda **_kwargs:DesktopFixture([]),
            clock=Clock(),
        )
        self.assertEqual(result.status,"TIMEOUT")
        self.assertTrue(result.timed_out)


if __name__=="__main__":
    unittest.main()
