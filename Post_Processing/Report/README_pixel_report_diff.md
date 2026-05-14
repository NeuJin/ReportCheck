# Pixel Report Diff Tool

Compare two reports by rendered pixels and/or PowerPoint objects.

Recommended for PPTX report checking:

```powershell
python pixel_report_diff.py expected.pptx actual.pptx --mode both --threshold 3
```

This creates slide screenshots, pixel-diff highlight images, and an object-level JSON diff.
If `-o` is not provided, output is saved to an `output` folder next to `pixel_report_diff.py`.
The tool intentionally refuses to write output outside the downloaded package folder.

## Install

Use Python 3.9+.

```powershell
pip install -r requirements.txt
```

Dependencies by input type, matched to your current environment:

- Image: `Pillow 8.1.2`
- PPT/PPTX pixel screenshot comparison: `Pillow 8.1.2`, `pywin32 224`, Microsoft PowerPoint installed
- PPTX object comparison: `python-pptx 0.6.18`

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
  --alpha 180
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


