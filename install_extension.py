#!/usr/bin/env python3
"""
Install the Quilt Motion Preview & Export extension and its Python deps.

Run this with the same Python that Inkscape uses so package installs land
in the correct environment.
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
)
REQUIREMENTS_FILE = ROOT_DIR / "requirements.txt"
DEFAULT_INKSCAPE_ENV = "INKSCAPE_PYTHON"


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


def run_pip_install(python_exe: str, dry_run: bool) -> None:
    if not REQUIREMENTS_FILE.exists():
        print("No requirements.txt found; skipping pip install.")
        return
    cmd = [python_exe, "-m", "pip", "install", "-r", str(REQUIREMENTS_FILE)]
    if dry_run:
        print(f"[dry-run] Would run: {' '.join(cmd)}")
        return
    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as exc:
        if platform.system().lower() == "darwin" and exc.returncode == -9:
            print("Pip install failed with SIGKILL on macOS.")
            print("This often means the Inkscape app is quarantined by Gatekeeper.")
            print("Try:")
            print('  xattr -dr com.apple.quarantine "/Applications/Inkscape.app"')
            print("Then launch Inkscape once and re-run this installer.")
        raise


def check_optional_deps() -> None:
    missing = []
    try:
        import PIL  # noqa: F401
    except Exception:
        missing.append("Pillow")

    try:
        import gi  # type: ignore

        gi.require_version("Gtk", "3.0")
        gi.require_foreign("cairo")
        import cairo  # noqa: F401  # type: ignore
    except Exception:
        missing.append("PyGObject (gi) + cairo")

    if missing:
        print("Missing optional runtime dependencies:")
        for item in missing:
            print(f"  - {item}")
        print("See README.md for OS-specific install hints.")


def _candidate_paths(paths: list[str]) -> list[Path]:
    return [Path(path).expanduser() for path in paths if path]


def find_inkscape_python() -> Path | None:
    override = os.environ.get(DEFAULT_INKSCAPE_ENV)
    if override:
        return Path(override).expanduser()

    system = platform.system().lower()
    candidates: list[Path] = []

    if system == "windows":
        candidates.extend(
            _candidate_paths(
                [
                    os.path.join(os.environ.get("ProgramFiles", ""), "Inkscape", "bin", "python.exe"),
                    os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Inkscape", "bin", "python.exe"),
                    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inkscape", "bin", "python.exe"),
                ]
            )
        )
    elif system == "darwin":
        candidates.extend(
            _candidate_paths(
                [
                    "/Applications/Inkscape.app/Contents/Resources/bin/python3",
                    "/Applications/Inkscape.app/Contents/Resources/bin/python",
                ]
            )
        )

    inkscape_exe = shutil.which("inkscape")
    if inkscape_exe:
        inkscape_path = Path(inkscape_exe).resolve()
        if system == "windows":
            candidates.append(inkscape_path.parent / "python.exe")
        elif system == "darwin":
            for parent in inkscape_path.parents:
                if parent.name == "Inkscape.app":
                    candidates.append(parent / "Contents" / "Resources" / "bin" / "python3")
                    candidates.append(parent / "Contents" / "Resources" / "bin" / "python")
                    break

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


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
        default=None,
        help=(
            "Python executable used for pip installs. Defaults to the detected "
            "Inkscape Python or the current interpreter."
        ),
    )
    parser.add_argument(
        "--inkscape-python",
        default=None,
        help=f"Explicit Inkscape Python path (or set {DEFAULT_INKSCAPE_ENV}).",
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
    inkscape_python = None
    if args.inkscape_python:
        inkscape_python = Path(args.inkscape_python).expanduser()
    else:
        inkscape_python = find_inkscape_python()

    python_exe = args.python or str(inkscape_python) if inkscape_python else sys.executable

    if inkscape_python:
        print(f"Using Inkscape Python: {inkscape_python}")
    else:
        print(f"Using current Python: {python_exe}")
        print("Tip: set INKSCAPE_PYTHON or --inkscape-python for bundled Inkscape Python.")

    install_extension(args.dest, args.dry_run)
    if not args.skip_pip:
        run_pip_install(python_exe, args.dry_run)
    if not args.dry_run:
        check_optional_deps()
    print("Done. Restart Inkscape to load the extension.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
