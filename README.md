# Quilt Motion Preview & Export
<img width="450" height="385" alt="Image" src="https://github.com/user-attachments/assets/25fafbe2-b2c5-4f67-9050-1a88f60c06cc" /><img width="500" height="385" alt="Image" src="https://github.com/user-attachments/assets/aa5bd6b5-5db5-4363-adbd-42ae2fea4126" />


An Inkscape extension that turns any vector path into a long‑arm quilting motion path. It lets you:

- Preserve the draw order of the selected paths and animate them with play/pause/restart plus adjustable speed.
- Scrub through the design with a progress slider to inspect any stitch in context, including stitched/jump coloring.
- Visualize the pattern as a multi-row pantograph (repeat count, rows, row spacing, stagger toggle + percent) with optional per-row mirroring (horizontal/vertical), global flips, and rectangular clipping so staggered rows stay flush edge-to-edge.
- See safety cues: start/end Y-axis mismatch warning label plus yellow rings on endpoints when dY > 0.1 mm.
- Optimize a stitched path to reduce self-overlaps while preserving geometry and endpoints, with a single-click toggle in the preview.
- Export either the single pattern or the entire layout (repeats/rows/stagger/mirroring/flips) using a switchback path that alternates direction row-by-row without lifting the needle, keeping the sewn path continuous.
- Export the resulting motion path as either a millimetre-true DXF polyline, a QCT-compatible DXF line file, or an animated GIF of the stitching motion—pick whichever fits your workflow.

> **Note:** The machine formats included here rely on open, text-based encodings of the stitch path. Every format is generated from the same normalized point stream, so the files remain easy to post-process with vendor-provided converters if needed.

## Installation

### macOS

1. **Download** this repository (clone or unzip).
2. **Run the installer** with your system Python:

   ```bash
   python3 install_extension.py
   ```

3. **Restart** Inkscape and open `Extensions → Quilting → Quilt Motion Preview & Export`.

Notes:
- The installer places dependencies in `quilt_motion_exporter_libs` next to the extension.
- The preview app uses the same Python used during install (saved in `quilt_motion_preview_python.txt`).

### Windows

1. **Download** this repository (clone or unzip).
2. **Run the installer** from PowerShell:

   ```powershell
   py -3 install_extension.py
   ```

3. **Restart** Inkscape and open `Extensions → Quilting → Quilt Motion Preview & Export`.

### Linux

1. **Download** this repository (clone or unzip).
2. **Run the installer**:

   ```bash
   python3 install_extension.py
   ```

3. **Restart** Inkscape and open `Extensions → Quilting → Quilt Motion Preview & Export`.

If your distro does not ship PySide6 wheels for your Python, install it via your package manager or run:

```bash
python3 -m pip install PySide6
```

### Manual install (all OSes)

1. **Download** this repository (clone or unzip).
2. **Copy** the `extensions/` and `README.md` contents into your Inkscape user extensions folder:

   | Platform | Extension folder |
   |----------|------------------|
   | Linux    | `~/.config/inkscape/extensions` (or `%APPDATA%` equivalent for Flatpak/Snap) |
   | Windows  | `%APPDATA%\Inkscape\extensions` (e.g., `C:\Users\<you>\AppData\Roaming\Inkscape\extensions`) |
   | macOS    | `~/Library/Application Support/org.inkscape.Inkscape/config/inkscape/extensions` |

3. **Install dependencies** into the sidecar folder:

   ```bash
   python3 -m pip install -r requirements.txt --target /path/to/inkscape/extensions/quilt_motion_exporter_libs
   ```

4. **Restart** Inkscape and open `Extensions → Quilting → Quilt Motion Preview & Export`.

## Usage

1. Draw your quilting motion path in Inkscape. Combine multiple paths if you need complex patterns—the extension respects the original draw/stacking order.
2. Select the path objects you want to export.
3. Open `Extensions → Quilting → Quilt Motion Preview & Export`. The preview window opens immediately using your current selection—no extra “Apply” click required.
4. Use the preview window:
   - **Play/Pause/Restart** control the animation, while the **Preview speed** slider changes draw speed.
   - Drag the **Progress** slider to jump to any point along the stitch path.
   - Adjust the **Pantograph layout** panel to tile the design into repeated rows (control repeats, rows, row distance in mm, stagger toggle, stagger percent) with per-row mirroring and global flips. Staggered rows are clipped to a rectangle and auto-filled left/right so the layout stays rectangular.
   - Enable **Mirror every other row horizontally/vertically** for variety and **Flip horizontally/vertically** for global orientation changes.
   - Click **Optimize path** to reorder stitched segments and reduce self-overlap; geometry and endpoints are preserved, and travel edges are kept intact.
   - Watch for the **Y-axis warning**: if start/end Y differ by more than 0.1 mm, a yellow label and rings appear on the endpoints.
   - Pick an export format and press **Export…** to write a file. Optionally check **Export entire layout** to bake repeats/rows/stagger/mirroring/flips into the file using a switchback, needle-down path (alternating row directions, no jumps between rows).
5. The exported files list every stitch (and jump) in document millimetres. They can be loaded directly by many quilting systems or passed through manufacturer tooling if post-processing is required.

> **Tip:** The extension runs immediately, so make sure your intended motion paths are selected before launching it. If nothing is selected, Inkscape shows an alert reminding you to pick paths first.

### Snap sandbox note

When running Inkscape from the Snap store, extensions may only write inside your home directory (and optionally removable-media if the interface is connected). The exporter now defaults the save dialog to `~/` and will display a clear message if you pick a location that the sandbox cannot access. If you need to save elsewhere, move the file afterward or connect the appropriate Snap interfaces.

## Implementation notes

- The preview window is a standalone PySide6 application launched from the extension.
- Paths are flattened via Inkscape’s `CubicSuperPath` utilities to keep Bézier curves accurate. Each sub‑path becomes a stitch segment, and travel jumps are inserted between disconnected components so long-arm controllers can raise the needle when necessary.
- Exporters live in `extensions/quilt_motion_exporter.py`. Each format is represented by a small writer function that receives the normalized motion model—adding more formats is as simple as registering another `ExportProfile`.
- Pantograph repeats are rendered purely in the preview: we offset each instance by the actual delta between its start and end nodes, optionally staggering alternate rows so you can audition complex layouts without duplicating geometry inside the SVG.

## Testing / Development

Automated verification lives under `tests/` and exercises the motion path model plus every exporter. Run it with:

```bash
python3 -m unittest discover -s tests
```

For quick smoke tests of the preview app entry point, you can also run:

```bash
cd extensions
python3 quilt_motion_preview_app.py --help
```

Inside Inkscape, keep the XML editor open to inspect the produced files if you need to adjust scaling or tolerances. When modifying the preview, `journalctl -f` (on Linux) is helpful for catching Python tracebacks emitted by Inkscape.

## Roadmap ideas

- Add vendor-specific headers to the text-based formats (e.g., IQP/BQM metadata blocks).
- Support stitch sampling density controls (fixed interval vs. exact node points).
- Export machine-specific thread trims, tie-offs, and pause markers.

## Export formats

- **DXF** – AutoCAD “lightweight polyline” entities. Stitch segments go on layer `STITCH`, jump segments on `TRAVEL`, and all coordinates are exported in millimetres to match your Inkscape document.
- **QCT DXF** – QCT-compatible DXF using per-segment `LINE` entities on layer `Layer`, with millimetre coordinates matching your document.
- **Animated GIF** – A shareable preview of the motion path. The extension redraws the path over ~60 frames (using the same light theme as the preview) and shows the stitch head progressing along the pattern, making it easy to review or send to clients without requiring their quilting software.
