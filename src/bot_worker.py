"""Frozen worker entry point. Automation remains implemented only in bot.py."""

import sys

from utils import configure_utf8_stdio


# Run before importing either command implementation. This is especially
# important for a frozen console Worker started on a non-UTF-8 Windows locale.
configure_utf8_stdio()


def main():
    configure_utf8_stdio()
    if sys.argv[1:]==["--diagnose-tvc-window"]:
        from tvc_window_locator import diagnostic_main
        return diagnostic_main()
    # Internal read-only command used by the GUI so a hung Windows UIA call can
    # be terminated at the process boundary. It never enters bot automation.
    if sys.argv[1:]==["--probe-tvc"]:
        from tvc_probe import main as probe_main
        return probe_main()
    from bot import main as bot_main
    return bot_main()


if __name__=="__main__":
    raise SystemExit(main())
