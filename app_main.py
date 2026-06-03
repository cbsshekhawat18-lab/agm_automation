"""
app_main.py — Entry point for the packaged AGM Final Set binary.

When run from PyInstaller (frozen):
  - Looks for Master_Input.xlsx next to the executable.
  - If missing, copies the bundled demo template next to the executable
    and exits with a "fill it in and re-run" message.
  - If present, generates output/Final_Set.docx next to the executable
    and opens it in the default viewer.

When run from source (python app_main.py) it behaves the same way,
using the project root as the work directory.
"""

import os
import shutil
import subprocess
import sys
import traceback


# Windows console (cmd.exe) defaults to cp1252 which can't print common Unicode
# characters used in our output (em-dashes, curly quotes, etc.). Reconfigure
# stdout/stderr to UTF-8 with 'replace' so any unprintable char becomes '?'
# rather than crashing the whole program with UnicodeEncodeError.
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def _is_frozen() -> bool:
    """True when running from a PyInstaller-bundled binary."""
    return bool(getattr(sys, "frozen", False))


def _app_dir() -> str:
    """Directory the user sees the binary in (NOT the PyInstaller temp dir)."""
    if _is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _bundle_dir() -> str:
    """Where PyInstaller extracted the bundled assets at runtime."""
    return getattr(sys, "_MEIPASS", _app_dir())


def _master_path(app_dir: str) -> str:
    """Where to look for Master_Input.xlsx.

    Bundled: next to the executable (app_dir/Master_Input.xlsx).
    Source : the existing dev convention (app_dir/master/Master_Input.xlsx),
             so running `python app_main.py` doesn't disturb the dev tree.
    """
    if _is_frozen():
        return os.path.join(app_dir, "Master_Input.xlsx")
    return os.path.join(app_dir, "master", "Master_Input.xlsx")


def _ist_timestamp() -> str:
    """Current time in Asia/Kolkata as YYYY-MM-DD_HH-MM-SS — filesystem-safe."""
    from datetime import datetime, timezone, timedelta
    try:
        from zoneinfo import ZoneInfo  # Python 3.9+
        now = datetime.now(ZoneInfo("Asia/Kolkata"))
    except ImportError:
        # Fallback: fixed +05:30 offset. IST has no DST so this is always correct.
        now = datetime.now(timezone(timedelta(hours=5, minutes=30)))
    return now.strftime("%Y-%m-%d_%H-%M-%S")


def _output_path(app_dir: str) -> str:
    """Where Final_Set.docx is written.

    Every run gets its own timestamped file so previous outputs are never
    overwritten — useful when iterating on the Master before the final sign-off.
    Timestamp is in IST (Asia/Kolkata), e.g.:
        output/Final_Set_2026-09-30_15-04-12_IST.docx
    """
    return os.path.join(app_dir, "output", f"Final_Set_{_ist_timestamp()}_IST.docx")


def _open_path(path: str) -> None:
    """Open `path` in the OS default viewer; quietly skip on failure."""
    try:
        if sys.platform == "darwin":
            subprocess.call(["open", path])
        elif sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.call(["xdg-open", path])
    except Exception:
        pass


def _pause(prompt: str = "\nPress Enter to exit...") -> None:
    """Hold the console open after a double-click run.

    On Windows, double-clicking a PyInstaller exe creates a console where
    `input()` immediately raises EOFError because stdin isn't a real TTY.
    Falling back to cmd.exe's `pause` command works in that environment.
    """
    if sys.platform == "win32":
        # Try input() first — works when launched from a real cmd shell.
        try:
            sys.stdout.write(prompt)
            sys.stdout.flush()
            input()
            return
        except (EOFError, OSError):
            pass
        # Fallback: cmd's own pause command — reads from the console directly
        # without needing a stdin pipe.
        try:
            os.system("pause >nul")
        except Exception:
            pass
        return

    # macOS / Linux
    try:
        input(prompt)
    except EOFError:
        pass


def _write_error_log(app_dir: str, message: str) -> None:
    """If main() crashes, write the exception to a log file next to the binary
    so we can diagnose what went wrong even when the console window vanished."""
    try:
        log_path = os.path.join(app_dir, "agm_final_set_error.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(message)
    except Exception:
        pass


def main() -> int:
    app_dir = _app_dir()
    bundle_dir = _bundle_dir()

    print("=" * 60)
    print("  AGM FINAL SET GENERATOR")
    print("=" * 60)
    print("  Privacy: This app runs 100% locally on your computer.")
    print("           No data is transmitted, uploaded, or collected.")
    print("           No internet connection required.")
    print("=" * 60)
    print(f"App folder: {app_dir}")

    # Make scripts/ importable in both source-run and frozen-run modes.
    sys.path.insert(0, os.path.join(bundle_dir, "scripts"))
    sys.path.insert(0, bundle_dir)

    master = _master_path(app_dir)
    out_path = _output_path(app_dir)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    if not os.path.exists(master):
        # Only auto-create the template when running from the bundled binary.
        # In source mode we never overwrite the dev tree.
        if not _is_frozen():
            print(f"ERROR: Master_Input.xlsx not found at {master}")
            print("       (source mode does not auto-copy a demo file)")
            _pause()
            return 1
        bundled_template = os.path.join(bundle_dir, "sample", "Master_Input_DEMO.xlsx")
        if not os.path.exists(bundled_template):
            print(f"ERROR: No Master_Input.xlsx in {app_dir}")
            print("       and no bundled demo template was found.")
            _pause()
            return 1
        os.makedirs(os.path.dirname(master), exist_ok=True)
        shutil.copy(bundled_template, master)
        print(f"Created {master}")
        print("Fill in your company data, save the file, then run this app again.")
        _open_path(master)
        _pause()
        return 0

    try:
        from loader import load_master
        from gen_final_set import build_final_set
        from validate_master import validate
    except ImportError as e:
        print(f"ERROR: Could not load generator modules: {e}")
        _pause()
        return 1

    print(f"Reading: {master}")
    try:
        data = load_master(master)
        validate(master)
    except SystemExit:
        # validate() calls sys.exit(1) when fields are missing — it already
        # printed the list; just hold the window open.
        _pause()
        return 1
    except Exception as e:
        print(f"ERROR while reading Master_Input.xlsx: {e}")
        _pause()
        return 1

    print(f"  Company: {data['company']['name']}")
    print(f"  AGM: {data['agm']['number']} on {data['agm']['date']}")
    print(f"  Directors: {len(data['directors'])}")
    print(f"  Shareholders: {len(data['shareholders'])}")
    print()
    print("Generating Final Set...")

    try:
        build_final_set(data, out_path)
    except Exception as e:
        print(f"ERROR while generating Final_Set.docx: {e}")
        _pause()
        return 1

    # build_final_set() already prints "Saved: ..." — don't repeat.
    print("=" * 60)
    print("DONE — opening the document.")
    print("=" * 60)
    _open_path(out_path)
    _pause()
    return 0


if __name__ == "__main__":
    # Wrap everything so that ANY exception (encoding error, import error,
    # permission error, etc.) is caught, printed, AND written to a log file
    # next to the binary. The pause() at the end keeps the cmd window open
    # so the user can read the error instead of watching it vanish.
    #
    # Note: main() handles validate-fail (SystemExit) internally and always
    # pauses before returning, so we don't need an outer SystemExit handler.
    try:
        rc = main()
    except KeyboardInterrupt:
        print("\nInterrupted.")
        sys.exit(130)
    except Exception as e:
        tb = traceback.format_exc()
        msg = (
            f"ERROR: {type(e).__name__}: {e}\n\n"
            f"{tb}\n"
            f"Please report this at:\n"
            f"  https://github.com/cbsshekhawat18-lab/agm_automation/issues\n"
        )
        try:
            print(msg)
        except Exception:
            # Even printing can fail on a broken console — fall through to the
            # log file write below.
            pass
        try:
            app_dir = (
                os.path.dirname(sys.executable)
                if getattr(sys, "frozen", False)
                else os.path.dirname(os.path.abspath(__file__))
            )
            _write_error_log(app_dir, msg)
            print(
                f"\nFull error details written to:\n"
                f"  {os.path.join(app_dir, 'agm_final_set_error.log')}"
            )
        except Exception:
            pass
        _pause()
        sys.exit(1)
    else:
        sys.exit(rc or 0)
