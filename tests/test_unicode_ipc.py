import io
import json
import os
from pathlib import Path
import subprocess
import sys
import unittest
from unittest import mock


SRC=Path(__file__).resolve().parents[1]/"src"
sys.path.insert(0,str(SRC))

import bot
import bot_worker
import tvc_control
import tvc_probe
import tvc_window_locator
from utils import configure_utf8_stdio


THAI_TITLE="T.V.C Client - รายการใบงาน (JOB)"
THAI_CONTROL="ข้อมูลรถ - งานปัจจุบัน"


def ready_probe_result():
    return {
        "status":"READY",
        "connected":True,
        "login_verified":True,
        "active_job_form":True,
        "active_job_forms":[THAI_CONTROL],
        "matches":[THAI_TITLE],
        "login_signals":["job_list_title"],
        "errors":[],
        "timed_out":False,
        "duration_ms":12,
    }


class UnicodeIpcTests(unittest.TestCase):
    def test_stdio_helper_reconfigures_both_streams(self):
        stdout=mock.Mock()
        stderr=mock.Mock()
        with mock.patch.object(sys,"stdout",stdout),mock.patch.object(sys,"stderr",stderr):
            configure_utf8_stdio()
        stdout.reconfigure.assert_called_once_with(encoding="utf-8",errors="replace")
        stderr.reconfigure.assert_called_once_with(encoding="utf-8",errors="replace")

    def test_stdio_helper_ignores_stream_reconfigure_failure(self):
        broken_stream=mock.Mock()
        broken_stream.reconfigure.side_effect=RuntimeError("unsupported stream")
        with mock.patch.object(sys,"stdout",broken_stream), \
             mock.patch.object(sys,"stderr",broken_stream):
            configure_utf8_stdio()

    def test_probe_main_is_ascii_safe_on_cp1252_and_roundtrips_thai(self):
        raw=io.BytesIO()
        cp1252_stream=io.TextIOWrapper(raw,encoding="cp1252",errors="strict")
        original_stdout=sys.stdout
        try:
            sys.stdout=cp1252_stream
            with mock.patch.object(tvc_probe,"probe",return_value=ready_probe_result()):
                return_code=tvc_probe.main()
            cp1252_stream.flush()
        finally:
            sys.stdout=original_stdout

        output=raw.getvalue().decode("ascii").strip()
        self.assertEqual(return_code,0)
        self.assertTrue(output.startswith(tvc_probe.RESULT_PREFIX))
        self.assertTrue(output.isascii())
        payload=json.loads(output[len(tvc_probe.RESULT_PREFIX):])
        self.assertEqual(payload["matches"],[THAI_TITLE])
        self.assertEqual(payload["active_job_forms"],[THAI_CONTROL])

    def test_bot_event_json_is_ascii_safe_and_roundtrips_thai(self):
        output=io.StringIO()
        with mock.patch.object(sys,"stdout",output):
            bot.emit_event("probe_fixture",title=THAI_TITLE,control=THAI_CONTROL)
        line=output.getvalue().strip()
        self.assertTrue(line.isascii())
        payload=json.loads(line[len(bot.EVENT_PREFIX):])
        self.assertEqual(payload["title"],THAI_TITLE)
        self.assertEqual(payload["control"],THAI_CONTROL)

    def test_gui_decodes_ascii_safe_probe_json_as_utf8(self):
        probe_payload=json.dumps(ready_probe_result(),ensure_ascii=True)
        completed=subprocess.CompletedProcess(
            args=["worker","--probe-tvc"],
            returncode=0,
            stdout=tvc_control.TVC_PROBE_RESULT_PREFIX+probe_payload+"\n",
            stderr="",
        )
        runner=mock.Mock(return_value=completed)
        paths=mock.Mock(app_dir=Path.cwd(),frozen=True)

        result=tvc_control.check_tvc_client(
            Path("TVC Bot Worker.exe"),
            runtime_paths=paths,
            runner=runner,
        )

        self.assertEqual(result["status"],"READY")
        self.assertEqual(result["matches"],[THAI_TITLE])
        self.assertEqual(result["active_job_forms"],[THAI_CONTROL])
        self.assertEqual(runner.call_args.kwargs["encoding"],"utf-8")
        self.assertEqual(runner.call_args.kwargs["errors"],"replace")

    def test_worker_configures_stdio_before_dispatching_probe(self):
        order=[]
        fake_probe=mock.Mock(side_effect=lambda:order.append("probe") or 0)
        fake_module=mock.Mock(main=fake_probe)
        with mock.patch.object(bot_worker,"configure_utf8_stdio",side_effect=lambda:order.append("stdio")), \
             mock.patch.object(sys,"argv",["bot_worker.py","--probe-tvc"]), \
             mock.patch.dict(sys.modules,{"tvc_probe":fake_module}):
            return_code=bot_worker.main()
        self.assertEqual(return_code,0)
        self.assertEqual(order,["stdio","probe"])

    def test_worker_dispatches_read_only_window_diagnostic(self):
        order=[]
        fake_diagnostic=mock.Mock(
            side_effect=lambda:order.append("diagnostic") or 0
        )
        fake_module=mock.Mock(diagnostic_main=fake_diagnostic)
        with mock.patch.object(
            bot_worker,
            "configure_utf8_stdio",
            side_effect=lambda:order.append("stdio"),
        ), mock.patch.object(
            sys,"argv",["bot_worker.py","--diagnose-tvc-window"]
        ), mock.patch.dict(sys.modules,{"tvc_window_locator":fake_module}):
            return_code=bot_worker.main()
        self.assertEqual(return_code,0)
        self.assertEqual(order,["stdio","diagnostic"])

    def test_standalone_probe_under_cp1252_never_raises_unicode_encode_error(self):
        environment=os.environ.copy()
        environment["PYTHONIOENCODING"]="cp1252"
        result=subprocess.run(
            [sys.executable,str(SRC/"tvc_probe.py")],
            cwd=str(SRC.parent),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=8,
            env=environment,
        )
        self.assertIn(result.returncode,{0,2})
        self.assertNotIn("UnicodeEncodeError",result.stderr)
        lines=[line for line in result.stdout.splitlines() if line.startswith(tvc_probe.RESULT_PREFIX)]
        self.assertTrue(lines,result.stdout+result.stderr)
        self.assertTrue(lines[-1].isascii())
        json.loads(lines[-1][len(tvc_probe.RESULT_PREFIX):])


if __name__=="__main__":
    unittest.main(verbosity=2)
