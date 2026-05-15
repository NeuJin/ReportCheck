# Pixel Report Diff Tool

Compare two reports by rendered pixels and/or PowerPoint objects.

Recommended for PPTX report checking:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx --mode both --threshold 3
```

This creates slide screenshots, pixel-diff highlight images, and an object-level JSON diff.
If `-o` is not provided, output is saved to an `output` folder next to `pixel_report_diff.py`.
The tool intentionally refuses to write output outside the downloaded package folder.


## GUI Usage

Run the window app from the downloaded package folder:

```powershell
python report_diff_gui.py
```

Workflow:

1. Select the expected PPT/PPTX report.
2. Select the actual PPT/PPTX report.
3. Keep output as `output` so all generated files stay inside the downloaded package.
4. Click `Run compare`.
5. Watch the progress bar for current processing status and percent.\n6. Use the slide list to view only different slides.

The GUI includes inline explanations for Pixel threshold, DPI, and Allowed diff %. It highlights differences with red rectangle regions using 20% transparent fill and a strong red border. This is easier to inspect than per-pixel red noise.
## Install

Use Python 3.9+.

```powershell
pip install -r requirements.txt
```

Dependencies by input type, matched to your current environment:

- Image: `Pillow 8.1.2`
- PPT/PPTX pixel screenshot comparison: `Pillow 8.1.2`, `pywin32 224`, Microsoft PowerPoint installed
- PPTX object comparison: `python-pptx 0.6.18`
- GUI window: `PySide2 5.15.2`


## Auto Slide Alignment

Use this when two reports have inserted, removed, or shifted slides:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx --mode pixel --align-slides --min-match-score 0.82
```

The output maps each comparison pair:

```text
Expected page | Actual page | Match score | Status
```

Statuses:

- `matched`: visually similar slides are compared and overlaid.
- `extra_actual`: slide exists only in the actual report.
- `missing_actual`: slide exists only in the expected report.
- `same_index`: used when auto-align is off.

The GUI has `Auto align slides` enabled by default. Lower `Min match score` if related slides are not matching; raise it if unrelated slides are matched.
## Usage

```powershell
python pixel_report_diff.py expected.png actual.png -o output --mode pixel
```

For PPT/PPTX slide screenshot comparison:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx -o output --mode pixel --dpi 150
```

For PPTX object comparison:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx -o output --mode object
```

For best practical accuracy, combine both:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx --mode both --threshold 3
```

Useful options:

```powershell
python pixel_report_diff.py expected.png actual.png `
  -o output `
  --threshold 5 `
  --allowed-percent 0.01 `
  --highlight-color 255,0,0 `
  --alpha 51
```

Outputs:

- `page_001_overlay.png`: actual report with differences highlighted.
- `page_001_mask.png`: black/white mask of different pixels.
- `object_summary.json`: object-level differences for PPTX when using `--mode object` or `--mode both`.
- `summary.json`: machine-readable combined result summary for automation.

Exit code:

- `0`: all pages are within `--allowed-percent`.
- `1`: at least one page failed.
- `2`: input or runtime error.

## Accuracy notes

Pixel comparison catches the final visual result, including charts, images, fonts, layout shifts, and anything that appears on the slide. Small anti-aliasing or rendering differences can create noise, so use `--threshold 3` to `--threshold 8` when needed.

Object comparison catches changes in PPTX shape order, position, size, text, and embedded image content without relying on screenshot rendering. It is useful for precise automation, but it may not fully understand complex charts, SmartArt, grouped objects, or effects.

Use `--mode both` when the report quality gate matters: pixel diff verifies the visual output, and object diff explains many structural changes.






