#!/usr/bin/env python3
"""
Install the Quilt Motion Preview & Export extension and its Python deps.

Dependencies are installed into a sidecar folder inside the Inkscape
extensions directory, so you do not need to touch Inkscape's bundled Python.
"""

from __future__ import annotations

import argparse
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
EXTENSION_FILES = (
    ROOT_DIR / "extensions" / "quilt_motion_exporter.py",
    ROOT_DIR / "extensions" / "quilt_motion_exporter.inx",
    ROOT_DIR / "extensions" / "quilt_motion_preview_app.py",
    ROOT_DIR / "extensions" / "quilt_motion_core.py",
)
REQUIREMENTS_FILE = ROOT_DIR / "requirements.txt"


def default_extension_dir() -> Path:
    override = os.environ.get("INKSCAPE_EXTENSION_DIR")
    if override:
        return Path(override).expanduser()

    system = platform.system().lower()
    if system == "windows":
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / "Inkscape" / "extensions"
        return Path.home() / "AppData" / "Roaming" / "Inkscape" / "extensions"
    if system == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "org.inkscape.Inkscape"
            / "config"
            / "inkscape"
            / "extensions"
        )

    xdg_config = os.environ.get("XDG_CONFIG_HOME")
    if xdg_config:
        return Path(xdg_config) / "inkscape" / "extensions"
    return Path.home() / ".config" / "inkscape" / "extensions"


def install_extension(dest_dir: Path, dry_run: bool) -> None:
    missing = [path for path in EXTENSION_FILES if not path.exists()]
    if missing:
        missing_list = ", ".join(str(path) for path in missing)
        raise FileNotFoundError(f"Extension source files not found: {missing_list}")

    if dry_run:
        print(f"[dry-run] Would create {dest_dir}")
    else:
        dest_dir.mkdir(parents=True, exist_ok=True)

    for src in EXTENSION_FILES:
        dest = dest_dir / src.name
        if dry_run:
            print(f"[dry-run] Would copy {src} -> {dest}")
        else:
            shutil.copy2(src, dest)
            print(f"Copied {src.name} -> {dest}")


def run_pip_install(python_exe: str, libs_dir: Path, dry_run: bool) -> None:
    if not REQUIREMENTS_FILE.exists():
        print("No requirements.txt found; skipping pip install.")
        return
    cmd = [
        python_exe,
        "-m",
        "pip",
        "install",
        "-r",
        str(REQUIREMENTS_FILE),
        "--target",
        str(libs_dir),
        "--upgrade",
    ]
    if dry_run:
        print(f"[dry-run] Would run: {' '.join(cmd)}")
        return
    try:
        libs_dir.mkdir(parents=True, exist_ok=True)
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as exc:
        print("Pip install failed; continuing without optional dependencies.")
        if platform.system().lower() == "darwin" and exc.returncode == -9:
            print("Pip install failed with SIGKILL on macOS.")
            print("This often means the Inkscape app is quarantined by Gatekeeper.")
            print("Try:")
            print('  xattr -dr com.apple.quarantine "/Applications/Inkscape.app"')
            print("Then launch Inkscape once and re-run this installer.")
        print("Some export options may be unavailable if dependencies are missing.")
        return


def check_optional_deps(libs_dir: Path) -> None:
    missing = []
    if libs_dir.exists():
        sys.path.insert(0, str(libs_dir))
    try:
        import PIL  # noqa: F401
    except Exception:
        missing.append("Pillow")

    try:
        import PySide6  # noqa: F401
    except Exception:
        missing.append("PySide6")

    if missing:
        print("Missing optional runtime dependencies:")
        for item in missing:
            print(f"  - {item}")
        print("See README.md for OS-specific install hints.")




def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Install the Quilt Motion Preview & Export Inkscape extension."
    )
    parser.add_argument(
        "--dest",
        type=Path,
        default=default_extension_dir(),
        help="Inkscape extensions directory (defaults to your user profile).",
    )
    parser.add_argument(
        "--python",
        default=sys.executable,
        help=(
            "Python executable used for pip installs (default: current interpreter)."
        ),
    )
    parser.add_argument(
        "--libs-dir",
        type=Path,
        default=None,
        help="Directory to install Python dependencies into.",
    )
    parser.add_argument(
        "--skip-pip",
        action="store_true",
        help="Skip installing pip dependencies.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show actions without making changes.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    python_exe = args.python
    libs_dir = args.libs_dir or (args.dest / "quilt_motion_exporter_libs")
    print(f"Using Python for deps: {python_exe}")
    print(f"Installing deps into: {libs_dir}")

    install_extension(args.dest, args.dry_run)
    if not args.skip_pip:
        run_pip_install(python_exe, libs_dir, args.dry_run)
    if not args.dry_run:
        check_optional_deps(libs_dir)
    print("Done. Restart Inkscape to load the extension.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
