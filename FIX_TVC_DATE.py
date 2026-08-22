from pathlib import Path
import shutil
import subprocess
import sys

repo = Path(r"D:\BOT-PMS")
target = repo / "src" / "tvc_driver.py"

if not target.exists():
    print(f"ERROR: file not found: {target}")
    sys.exit(1)

text = target.read_text(encoding="utf-8")

old = """        try:
            ctrl.set_time(
                dt.year, dt.month, dt.day,
                dt.hour, dt.minute, dt.second, 0
            )
        except Exception:
"""

new = """        try:
            # DateTimePickerWrapper.set_time expects SYSTEMTIME order:
            # year, month, day_of_week, day, hour, minute, second, milliseconds.
            # Python weekday(): Monday=0 ... Sunday=6
            # Windows SYSTEMTIME: Sunday=0 ... Saturday=6.
            day_of_week = (dt.weekday() + 1) % 7
            ctrl.set_time(
                dt.year, dt.month, day_of_week, dt.day,
                dt.hour, dt.minute, dt.second, 0
            )
        except Exception:
"""

if new in text:
    print("DATE FIX: already applied. No changes made.")
    sys.exit(0)

count = text.count(old)
if count != 1:
    print(f"ERROR: expected exactly 1 old date block, found {count}.")
    print("No file was changed.")
    sys.exit(2)

backup = target.with_name(target.name + ".before-date-fix.bak")
shutil.copy2(target, backup)

updated = text.replace(old, new, 1)
target.write_text(updated, encoding="utf-8", newline="")

verify = target.read_text(encoding="utf-8")
if new not in verify:
    shutil.copy2(backup, target)
    print("ERROR: verification failed. Original file was restored.")
    sys.exit(3)

print("SUCCESS: fixed src\\tvc_driver.py")
print(f"BACKUP: {backup}")
print()
print("Changed DateTimePicker set_time argument order:")
print("  year, month, day_of_week, day, hour, minute, second, milliseconds")
print()
print("----- git diff -----")
try:
    subprocess.run(
        ["git", "-C", str(repo), "diff", "--", "src/tvc_driver.py"],
        check=False,
    )
except FileNotFoundError:
    print("git not found; run this manually:")
    print(r'git -C D:\BOT-PMS diff -- src/tvc_driver.py')
