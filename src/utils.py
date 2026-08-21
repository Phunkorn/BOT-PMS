from datetime import datetime
import sys


def configure_utf8_stdio():
    """Make redirected process output deterministic on Windows.

    Frozen console applications inherit the machine's active code page. The
    GUI protocol is UTF-8, so configure both streams before a worker or probe
    can print localized text. Some embedded/test streams do not implement
    ``reconfigure``; those streams are intentionally left unchanged.
    """
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if not callable(reconfigure):
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def stamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")
