"""Frozen worker entry point. Automation remains implemented only in bot.py."""

from bot import main


if __name__=="__main__":
    raise SystemExit(main())
