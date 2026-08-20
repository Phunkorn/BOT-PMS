from datetime import datetime
from pathlib import Path
def stamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")
def ensure_dirs(base: Path):
    (base/"logs").mkdir(exist_ok=True)
    (base/"screenshots").mkdir(exist_ok=True)
