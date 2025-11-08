# Quilt Motion Preview & Export

An Inkscape extension that turns any vector path into a long‑arm quilting motion path. It lets you:

- Preserve the draw order of the selected paths.
- Watch the stitching sequence in a live GTK preview (play, pause, restart, adjustable speed).
- Export the resulting motion path to several quilting machine formats (BQM, DXF, HQF, IQP, PAT, PLT, QCC, QLI, SSD, TXT).

> **Note:** The machine formats included here rely on open, text-based encodings of the stitch path. Every format is generated from the same normalized point stream, so the files remain easy to post-process with vendor-provided converters if needed.

## Installation

1. Copy the contents of this repository into Inkscape’s user extension directory. On Linux this is typically `~/.config/inkscape/extensions`. Create the folder if it does not exist.
2. Restart Inkscape. The new entry appears under `Extensions → Quilting → Quilt Motion Preview & Export`.

## Usage

1. Draw your quilting motion path in Inkscape. Combine multiple paths if you need complex patterns—the extension respects the original draw/stacking order.
2. Select the path objects you want to export.
3. Open `Extensions → Quilting → Quilt Motion Preview & Export` and click **Apply**.
4. Use the preview window:
   - **Play/Pause/Restart** control the animation.
   - Adjust the **Preview speed** slider to see the stitches rendered faster or slower.
   - The progress readout displays the stitched length and completion percentage.
   - Pick an export format and press **Export…** to write the file.
5. The exported files list every stitch (and jump) in document millimetres. They can be loaded directly by many quilting systems or passed through manufacturer tooling if post-processing is required.

## Implementation notes

- The preview window is built with PyGObject/Gtk 3, which ships with modern Inkscape builds.
- Paths are flattened via Inkscape’s `CubicSuperPath` utilities to keep Bézier curves accurate. Each sub‑path becomes a stitch segment, and travel jumps are inserted between disconnected components so long-arm controllers can raise the needle when necessary.
- Exporters live in `extensions/quilt_motion_exporter.py`. Each format is represented by a small writer function that receives the normalized motion model—adding more formats is as simple as registering another `ExportProfile`.

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
