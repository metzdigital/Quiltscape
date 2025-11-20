# Quilt Motion Preview & Export

An Inkscape extension that turns any vector path into a long‑arm quilting motion path. It lets you:

- Preserve the draw order of the selected paths and animate them with play/pause/restart plus adjustable speed.
- Scrub through the design with a progress slider to inspect any stitch in context.
- Visualize the pattern as a multi-row pantograph (repeat count, rows, row spacing, stagger toggle, stagger percent) with a light-themed GTK preview.
- Export the resulting motion path as either a millimetre-true DXF polyline (machine-ready) or an animated GIF of the stitching motion—pick whichever fits your workflow.

> **Note:** The machine formats included here rely on open, text-based encodings of the stitch path. Every format is generated from the same normalized point stream, so the files remain easy to post-process with vendor-provided converters if needed.

## Installation

1. **Download** this repository (clone or unzip).
2. **Copy** the `extensions/` and `README.md` contents into your Inkscape user extensions folder:

   | Platform | Extension folder |
   |----------|------------------|
   | Linux    | `~/.config/inkscape/extensions` (or `%APPDATA%` equivalent for Flatpak/Snap) |
   | Windows  | `%APPDATA%\Inkscape\extensions` (e.g., `C:\Users\<you>\AppData\Roaming\Inkscape\extensions`) |
   | macOS    | `~/Library/Application Support/org.inkscape.Inkscape/config/inkscape/extensions` |

   Create the directory if it does not exist, then restart Inkscape.

3. **Ensure Gtk bindings exist** for Python:

   - **Linux (Apt)**:

     ```bash
     sudo apt install python3-gi python3-gi-cairo python3-cairo gir1.2-gtk-3.0
     ```

   - **Linux (other distros)**: install the PyGObject/Gtk runtime packages via your package manager (look for `pygobject3`, `python3-cairo`, `gtk3` introspection).

   - **Windows**: install the official Inkscape package (which bundles Python 3, Gtk, and PyGObject). If you run a custom Python, make sure `pygobject` and `pycairo` are installed through MSYS2 or the Gnome for Windows runtime.

   - **macOS**: install via Homebrew:

     ```bash
     brew install pygobject3 gtk+3 py3cairo
     ```

     or ensure the DMG build of Inkscape includes `python3-gi`. If not, use `python3 -m pip install pyobjc pygobject pycairo` within the Inkscape Python environment.

4. Start Inkscape and find the extension under `Extensions → Quilting → Quilt Motion Preview & Export`.

## Usage

1. Draw your quilting motion path in Inkscape. Combine multiple paths if you need complex patterns—the extension respects the original draw/stacking order.
2. Select the path objects you want to export.
3. Open `Extensions → Quilting → Quilt Motion Preview & Export`. The preview window opens immediately using your current selection—no extra “Apply” click required.
4. Use the preview window:
   - **Play/Pause/Restart** control the animation, while the **Preview speed** slider changes draw speed.
   - Drag the **Progress** slider to jump to any point along the stitch path.
   - Adjust the **Pantograph layout** panel to tile the design into repeated rows (control repeats, rows, row distance in mm, stagger toggle, stagger percent). The preview updates instantly to show the full pantograph layout.
   - Pick an export format and press **Export…** to write a file (DXF for machine import, Animated GIF for sharing/preview).
5. The exported files list every stitch (and jump) in document millimetres. They can be loaded directly by many quilting systems or passed through manufacturer tooling if post-processing is required.

> **Tip:** The extension runs immediately, so make sure your intended motion paths are selected before launching it. If nothing is selected, Inkscape shows an alert reminding you to pick paths first.

### Snap sandbox note

When running Inkscape from the Snap store, extensions may only write inside your home directory (and optionally removable-media if the interface is connected). The exporter now defaults the save dialog to `~/` and will display a clear message if you pick a location that the sandbox cannot access. If you need to save elsewhere, move the file afterward or connect the appropriate Snap interfaces.

## Implementation notes

- The preview window is built with PyGObject/Gtk 3, which ships with modern Inkscape builds.
- Paths are flattened via Inkscape’s `CubicSuperPath` utilities to keep Bézier curves accurate. Each sub‑path becomes a stitch segment, and travel jumps are inserted between disconnected components so long-arm controllers can raise the needle when necessary.
- Exporters live in `extensions/quilt_motion_exporter.py`. Each format is represented by a small writer function that receives the normalized motion model—adding more formats is as simple as registering another `ExportProfile`.
- Pantograph repeats are rendered purely in the preview: we offset each instance by the actual delta between its start and end nodes, optionally staggering alternate rows so you can audition complex layouts without duplicating geometry inside the SVG.

## Testing / Development

Automated verification lives under `tests/` and exercises the motion path model plus every exporter. Run it with:

```bash
python3 -m unittest discover -s tests
```

For quick smoke tests of the GTK entry point, you can also run:

```bash
cd extensions
python3 quilt_motion_exporter.py --help
```

Inside Inkscape, keep the XML editor open to inspect the produced files if you need to adjust scaling or tolerances. When modifying the preview, `journalctl -f` (on Linux) is helpful for catching Python tracebacks emitted by Inkscape.

## Roadmap ideas

- Add vendor-specific headers to the text-based formats (e.g., IQP/BQM metadata blocks).
- Support stitch sampling density controls (fixed interval vs. exact node points).
- Export machine-specific thread trims, tie-offs, and pause markers.

## Export formats

- **DXF** – AutoCAD “lightweight polyline” entities. Stitch segments go on layer `STITCH`, jump segments on `TRAVEL`, and all coordinates are exported in millimetres to match your Inkscape document.
- **Animated GIF** – A shareable preview of the motion path. The extension redraws the path over ~60 frames (using the same light theme as the preview) and shows the stitch head progressing along the pattern, making it easy to review or send to clients without requiring their quilting software.
