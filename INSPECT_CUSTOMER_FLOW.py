from __future__ import annotations

from pathlib import Path
from pywinauto import Desktop

OUT = Path("customer_flow_inspect.txt")

TARGET_TITLES = (
    "เพิ่มใบงาน (JOB)",
    "ค้นหาลูกค้า",
)

def safe_text(ctrl):
    try:
        return (ctrl.window_text() or "").strip()
    except Exception:
        return ""

def safe_rect(ctrl):
    try:
        r = ctrl.rectangle()
        return f"({r.left},{r.top})-({r.right},{r.bottom})"
    except Exception:
        return ""

def dump_backend(backend: str, lines: list[str]) -> None:
    lines.append("=" * 100)
    lines.append(f"BACKEND: {backend}")
    lines.append("=" * 100)

    try:
        windows = Desktop(backend=backend).windows()
    except Exception as e:
        lines.append(f"ERROR listing windows: {e!r}")
        return

    matched = []
    for w in windows:
        title = safe_text(w)
        if any(t in title for t in TARGET_TITLES):
            matched.append(w)

    if not matched:
        lines.append("No matching target windows found.")
        return

    for idx, w in enumerate(matched, 1):
        title = safe_text(w)
        lines.append("")
        lines.append(f"[WINDOW {idx}] {title!r}")
        lines.append(f"RECT={safe_rect(w)}")
        try:
            info = w.element_info
            lines.append(
                "WINDOW_INFO "
                f"TYPE={getattr(info, 'control_type', '')!r} "
                f"AUTO={getattr(info, 'automation_id', '')!r} "
                f"CLASS={getattr(info, 'class_name', '')!r} "
                f"ID={getattr(info, 'control_id', '')!r}"
            )
        except Exception as e:
            lines.append(f"WINDOW_INFO_ERROR={e!r}")

        try:
            descendants = w.descendants()
        except Exception as e:
            lines.append(f"DESCENDANTS_ERROR={e!r}")
            continue

        lines.append(f"DESCENDANT_COUNT={len(descendants)}")
        for i, c in enumerate(descendants, 1):
            try:
                info = c.element_info
                text = safe_text(c)
                control_type = str(getattr(info, "control_type", "") or "")
                automation_id = str(getattr(info, "automation_id", "") or "")
                class_name = str(getattr(info, "class_name", "") or "")
                control_id = getattr(info, "control_id", "")
                rect = safe_rect(c)

                # Keep useful controls and anything with visible text / IDs.
                if not (
                    text
                    or automation_id
                    or control_id not in ("", None, 0)
                    or control_type in {
                        "Button", "Edit", "ComboBox", "List", "ListItem",
                        "DataGrid", "DataItem", "Table", "Pane", "Text"
                    }
                ):
                    continue

                lines.append(
                    f"{i:04d} "
                    f"TEXT={text!r} "
                    f"TYPE={control_type!r} "
                    f"AUTO={automation_id!r} "
                    f"CLASS={class_name!r} "
                    f"ID={control_id!r} "
                    f"RECT={rect}"
                )
            except Exception as e:
                lines.append(f"{i:04d} ERROR={e!r}")

def main():
    lines: list[str] = []
    lines.append("READ-ONLY CUSTOMER FLOW UI INSPECTION")
    lines.append("No clicks, no typing, no changes to T.V.C.")
    lines.append("")
    for backend in ("win32", "uia"):
        dump_backend(backend, lines)

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Saved: {OUT.resolve()}")
    print("ส่งไฟล์ customer_flow_inspect.txt กลับมาให้ ChatGPT ได้เลย")

if __name__ == "__main__":
    main()
