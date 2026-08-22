from __future__ import annotations

from datetime import datetime, date
from pathlib import Path
import re
import time

from pywinauto import Desktop
from pywinauto.keyboard import send_keys
from pywinauto.mouse import click as mouse_click
import tvc_window_locator


class TVCDriver:
    def __init__(self, title_regex, backend="win32", delay=0.45, typing_interval=0.03):
        self.title_regex = title_regex
        self.backend = backend
        self.delay = delay
        self.typing_interval = typing_interval
        self.window = None

    # =========================================================
    # WINDOW / CONTROL
    # =========================================================
    def _safe_windows(self):
        last_error = None
        for _ in range(5):
            try:
                return Desktop(backend=self.backend).windows()
            except Exception as e:
                last_error = e
                time.sleep(0.15)

        if last_error:
            raise last_error

        return []

    def _has_auto_id(self, window, auto_id):
        try:
            controls = window.descendants()
        except Exception:
            return False

        for ctrl in controls:
            try:
                aid = str(getattr(ctrl.element_info, "automation_id", "") or "")
                if aid == auto_id:
                    return True
            except Exception:
                pass

        return False

    def _find_job_window(self):
        # 1) Exact/regex title such as "เพิ่มใบงาน (JOB)"
        for w in self._safe_windows():
            try:
                title = w.window_text()
                if title and re.search(self.title_regex, title, re.I):
                    return w
            except Exception:
                pass

        # 2) Fallback: title may change after entering vehicle data.
        # Identify the JOB form from stable controls.
        for w in self._safe_windows():
            try:
                if (
                    self._has_auto_id(w, "ButtonX3")
                    and self._has_auto_id(w, "ListView1")
                    and self._has_auto_id(w, "Tno")
                ):
                    return w
            except Exception:
                pass

        return None

    def _find_tvc_main_window(self):
        """Resolve the same verified top-level window accepted by pre-check."""
        result=tvc_window_locator.locate_tvc_main_window(
            self.backend,
            timeout_seconds=1.5,
        )
        if result.selected is None:
            return None
        return result.selected.window

    def _wait_job_window(self, timeout=6.0):
        end = time.time() + timeout

        while time.time() < end:
            w = self._find_job_window()
            if w is not None:
                return w
            time.sleep(0.15)

        return None

    def open_new_job_form(self):
        """
        v0.6.3
        เปิดหน้าเพิ่มใบงานเองจากหน้ารายการใบงาน

        จากการ Inspect เครื่องจริง:
          ปุ่ม "ใบงาน"
            RECT = (361,61)-(404,117)

          popup dropdown
            TYPE = Menu
            RECT ตัวอย่าง = (361,117)-(619,210)

          "ออกใบงาน (JOB)" = แถวแรกของ popup

        ใช้ตำแหน่งแบบ relative กับหน้าต่าง T.V.C
        เพื่อไม่ผูกกับตำแหน่ง absolute ของจอ
        """

        # If JOB form is already open, do nothing.
        existing = self._find_job_window()
        if existing is not None:
            self.window = existing
            try:
                self.window.set_focus()
            except Exception:
                pass
            print("OPEN JOB: หน้าเพิ่มใบงานเปิดอยู่แล้ว")
            return existing

        main = self._find_tvc_main_window()
        if main is None:
            raise RuntimeError(
                "ไม่พบหน้าต่างหลัก T.V.C Client สำหรับเปิดใบงานใหม่"
            )

        try:
            main.set_focus()
        except Exception:
            pass

        time.sleep(0.35)

        main_rect = main.rectangle()

        # Coordinates found from the actual TVC machine.
        # Keep them relative to the TVC main window so moving the app still works.
        job_menu_x = main_rect.left + 383
        job_menu_y = main_rect.top + 89

        print(
            "OPEN JOB: click 'ใบงาน' "
            f"at ({job_menu_x}, {job_menu_y})"
        )
        mouse_click(coords=(job_menu_x, job_menu_y))

        # Wait for custom dropdown Menu to appear.
        d = Desktop(backend="uia")
        menu = None
        end = time.time() + 2.5

        # Probe inside the first menu row based on the inspected layout.
        probe_x = main_rect.left + 440
        probe_y = main_rect.top + 142

        while time.time() < end:
            try:
                candidate = d.from_point(probe_x, probe_y)
                control_type = str(
                    getattr(candidate.element_info, "control_type", "") or ""
                )

                if control_type == "Menu":
                    menu = candidate
                    break
            except Exception:
                pass

            time.sleep(0.10)

        if menu is None:
            raise RuntimeError(
                "กดเมนู 'ใบงาน' แล้ว แต่ไม่พบ dropdown Menu "
                "Bot หยุดเพื่อป้องกันการคลิกผิดตำแหน่ง"
            )

        r = menu.rectangle()

        # First row = "ออกใบงาน (JOB)"
        click_x = (r.left + r.right) // 2
        click_y = r.top + 35

        print(
            "OPEN JOB: พบ dropdown Menu "
            f"RECT={r}; click first row 'ออกใบงาน (JOB)' (offset +35) "
            f"at ({click_x}, {click_y})"
        )

        mouse_click(coords=(click_x, click_y))

        # Verify that the new JOB form really opened.
        job_window = self._wait_job_window(timeout=6.0)
        if job_window is None:
            raise RuntimeError(
                "คลิก 'ออกใบงาน (JOB)' แล้ว "
                "แต่ไม่พบหน้า 'เพิ่มใบงาน (JOB)' ภายในเวลาที่กำหนด"
            )

        self.window = job_window

        try:
            self.window.set_focus()
        except Exception:
            pass

        print("OPEN JOB: เปิดหน้าเพิ่มใบงาน (JOB) สำเร็จ")
        time.sleep(max(self.delay, 0.4))
        return job_window

    def connect(self):
        self.window = None

        # First try an already-open JOB form.
        w = self._find_job_window()
        if w is not None:
            self.window = w
            try:
                self.window.set_focus()
            except Exception:
                pass
            return w

        # v0.6.3: if the JOB form isn't open, create it automatically.
        print("OPEN JOB: ยังไม่พบหน้าเพิ่มใบงาน -> เปิดใบงานใหม่อัตโนมัติ")
        return self.open_new_job_form()

    def refresh_window(self):
        if self.window is not None:
            try:
                if self.window.exists() and self._has_auto_id(self.window, "ButtonX3"):
                    return self.window
            except Exception:
                pass

        return self.connect()

    def _controls(self):
        self.refresh_window()

        try:
            return self.window.descendants()
        except Exception:
            self.connect()
            return self.window.descendants()

    def find(self, spec):
        auto_id = spec["auto_id"]
        class_contains = spec.get("class_contains")

        for ctrl in self._controls():
            try:
                info = ctrl.element_info
                aid = str(getattr(info, "automation_id", "") or "")
                class_name = str(getattr(info, "class_name", "") or "")

                if aid != auto_id:
                    continue

                if class_contains and class_contains.lower() not in class_name.lower():
                    continue

                return ctrl

            except Exception:
                pass

        raise RuntimeError(f"ไม่พบ control {auto_id}")

    # =========================================================
    # INPUT
    # =========================================================
    def set_text(self, spec, value, press_enter=False):
        if value is None or str(value).strip() == "":
            return

        ctrl = self.find(spec)
        ctrl.set_focus()

        try:
            ctrl.set_edit_text(str(value))
        except Exception:
            ctrl.click_input()
            send_keys("^a{BACKSPACE}")
            send_keys(str(value), with_spaces=True)

        if press_enter:
            send_keys("{ENTER}")

        time.sleep(self.delay)

    def set_date(self, spec, value):
        if value is None or str(value).strip() == "":
            return

        if isinstance(value, datetime):
            dt = value
        elif isinstance(value, date):
            dt = datetime(value.year, value.month, value.day)
        else:
            raw = str(value).strip()
            dt = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
                try:
                    dt = datetime.strptime(raw, fmt)
                    break
                except ValueError:
                    pass

            if dt is None:
                raise RuntimeError(f"รูปแบบวันที่ไม่รองรับ: {value}")

        ctrl = self.find(spec)
        ctrl.set_focus()

        try:
            ctrl.set_time(
                dt.year, dt.month, dt.day,
                dt.hour, dt.minute, dt.second, 0
            )
        except Exception:
            ctrl.click_input()
            send_keys("^a")
            send_keys(dt.strftime("%d/%m/%Y"), with_spaces=True)

        time.sleep(self.delay)

    def click(self, spec):
        ctrl = self.find(spec)

        try:
            ctrl.click()
        except Exception:
            ctrl.click_input()

        time.sleep(self.delay)

    # =========================================================
    # LIST / TEXT
    # =========================================================
    def list_count(self, spec):
        ctrl = self.find(spec)

        try:
            return ctrl.item_count()
        except Exception:
            try:
                return len(ctrl.items())
            except Exception:
                return None

    def get_text(self, spec):
        ctrl = self.find(spec)

        try:
            return ctrl.window_text()
        except Exception:
            return ""

    # =========================================================
    # SCREENSHOT
    # =========================================================
    def screenshot(self, path: Path):
        self.refresh_window()

        try:
            self.window.capture_as_image().save(path)
        except Exception:
            self.connect()
            self.window.capture_as_image().save(path)

    # =========================================================
    # DUPLICATE CUSTOMER
    # =========================================================
    def _find_duplicate_customer_window(self):
        for w in self._safe_windows():
            try:
                title = (w.window_text() or "").strip()
                if title == "ตรวจสอบรายชื่อ":
                    return w
            except Exception:
                pass

        return None

    def _wait_duplicate_customer_window(self, timeout=2.5):
        end = time.time() + timeout

        while time.time() < end:
            w = self._find_duplicate_customer_window()
            if w is not None:
                return w

            time.sleep(0.15)

        return None

    def _find_control_by_auto_id_in_window(self, window, auto_id):
        try:
            controls = window.descendants()
        except Exception:
            controls = []

        for ctrl in controls:
            try:
                aid = str(getattr(ctrl.element_info, "automation_id", "") or "")

                if aid == auto_id:
                    return ctrl
            except Exception:
                pass

        return None

    def _click_update_existing_customer(self, duplicate_window):
        """
        ตรวจจาก Inspect จริง:
            ButtonX1 = อับเดทข้อมูลลงในรหัสเก่า  <-- ต้องกด
            ButtonX2 = สร้างเป็นรหัสลูกค้าใหม่   <-- ห้ามกด

        v0.6.2:
            ใช้ click_input() เป็นหลัก เพื่อให้เป็น physical mouse click
            สำหรับ custom WinForms control ของ T.V.C
        """

        update_button = self._find_control_by_auto_id_in_window(
            duplicate_window,
            "ButtonX1",
        )

        if update_button is None:
            raise RuntimeError(
                "พบหน้าต่าง 'ตรวจสอบรายชื่อ' "
                "แต่ไม่พบ ButtonX1 (อับเดทข้อมูลลงในรหัสเก่า) "
                "Bot หยุดเพื่อป้องกันการสร้างรหัสลูกค้าใหม่"
            )

        # Safety check: ถ้า ButtonX2 ไม่ต้องแตะ
        _create_new_button = self._find_control_by_auto_id_in_window(
            duplicate_window,
            "ButtonX2",
        )

        try:
            button_text = (update_button.window_text() or "").strip()
        except Exception:
            button_text = ""

        print(
            "DUPLICATE CUSTOMER: "
            f"physical click ButtonX1 = {button_text!r}"
        )

        try:
            duplicate_window.set_focus()
        except Exception:
            pass

        time.sleep(0.2)

        try:
            update_button.set_focus()
        except Exception:
            pass

        time.sleep(0.1)

        # IMPORTANT: use physical mouse click first
        try:
            update_button.click_input()
        except Exception as e:
            raise RuntimeError(
                "พบ ButtonX1 แต่ physical click_input() ไม่สำเร็จ: "
                f"{e}"
            )

        time.sleep(max(self.delay, 0.6))

    # =========================================================
    # SAVE FLOW
    # =========================================================
    def save_flow_yes_then_no(
        self,
        save_spec=None,
        dialog_wait=1.0,
    ):
        """
        NORMAL:
            F10 -> Y -> N

        DUPLICATE:
            F10
            -> Window 'ตรวจสอบรายชื่อ'
            -> physical click ButtonX1
               'อับเดทข้อมูลลงในรหัสเก่า'
            -> Y
            -> N

        ButtonX2 ห้ามกด
        """

        self.refresh_window()

        try:
            self.window.set_focus()
        except Exception:
            pass

        time.sleep(0.25)

        # 1) Save
        send_keys("{F10}")

        # 2) Duplicate branch
        duplicate_window = self._wait_duplicate_customer_window(
            timeout=2.5
        )

        duplicate_found = duplicate_window is not None

        if duplicate_found:
            print(
                "DUPLICATE CUSTOMER: "
                "พบหน้าต่างตรวจสอบรายชื่อ"
            )

            self._click_update_existing_customer(
                duplicate_window
            )

            time.sleep(dialog_wait)

        else:
            print(
                "DUPLICATE CUSTOMER: "
                "ไม่พบ -> flow ปกติ"
            )

        # 3) Save confirmation
        send_keys("y")
        time.sleep(dialog_wait)

        # 4) Print confirmation
        send_keys("n")
        time.sleep(dialog_wait)

        return {
            "duplicate_customer": duplicate_found,
            "duplicate_action": (
                "UPDATE_EXISTING_BUTTONX1_PHYSICAL_CLICK"
                if duplicate_found
                else "NONE"
            ),
            "first_clicked": "YES",
            "second_clicked": "NO",
        }
