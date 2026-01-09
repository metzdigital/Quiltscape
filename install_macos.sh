#!/bin/bash
set -euo pipefail

INKSCAPE_APP="/Applications/Inkscape.app"

if [ ! -d "$INKSCAPE_APP" ]; then
  echo "Inkscape.app not found at $INKSCAPE_APP"
  echo "Edit this script or set INKSCAPE_APP to the correct path."
  exit 1
fi

if xattr -p com.apple.quarantine "$INKSCAPE_APP" >/dev/null 2>&1; then
  echo "Removing Gatekeeper quarantine from $INKSCAPE_APP"
  xattr -dr com.apple.quarantine "$INKSCAPE_APP"
fi

echo "Launching Inkscape once to complete any first-run setup..."
open -a "$INKSCAPE_APP"
echo "Once Inkscape finishes launching, close it and press Enter to continue."
read -r

echo "Running extension installer..."
python3 "$(dirname "$0")/install_extension.py"
