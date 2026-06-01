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
    """Hold the console open after a double-click run."""
    try:
        input(prompt)
    except EOFError:
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
    sys.exit(main())
