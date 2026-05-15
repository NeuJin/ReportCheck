#!/usr/bin/env python3
"""Compare report files by rendered pixels and/or PowerPoint objects.

Python version: 3.9+

Supported pixel inputs:
- Images supported by Pillow: PNG, JPG, JPEG, BMP, TIFF, WEBP.
- PPT/PPTX files on Windows when Microsoft PowerPoint + pywin32 are installed.

Supported object inputs:
- PPTX files when python-pptx is installed.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
import tempfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from PIL import Image, ImageChops, ImageDraw


PPT_EXTS = {".ppt", ".pptx"}
PPTX_EXT = ".pptx"
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
EMU_PER_POINT = 12700
DEFAULT_OUTPUT_DIR = "output"


@dataclass
class PageDiff:
    page: int
    expected_page: Optional[int]
    actual_page: Optional[int]
    match_score: Optional[float]
    match_status: str
    width: int
    height: int
    compared_pixels: int
    different_pixels: int
    difference_percent: float
    max_channel_delta: int
    bbox: Optional[Tuple[int, int, int, int]]
    passed: bool
    output_overlay: str
    output_mask: str
    regions: List[Tuple[int, int, int, int]]


@dataclass
class SlideMatch:
    expected_index: Optional[int]
    actual_index: Optional[int]
    score: Optional[float]
    status: str


@dataclass
class ObjectDiff:
    slide: int
    object_index: int
    field: str
    expected: Any
    actual: Any


def parse_color(value: str) -> Tuple[int, int, int]:
    parts = value.split(",")
    if len(parts) != 3:
        raise argparse.ArgumentTypeError("Color must be formatted as R,G,B")

    try:
        rgb = tuple(int(part.strip()) for part in parts)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Color values must be integers") from exc

    if any(channel < 0 or channel > 255 for channel in rgb):
        raise argparse.ArgumentTypeError("Color values must be between 0 and 255")

    return rgb  # type: ignore[return-value]


def safe_text(value: Any) -> str:
    return "" if value is None else str(value)


def export_ppt_to_png(path: Path, dpi: int, temp_root: Path) -> List[Path]:
    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "PPT/PPTX pixel comparison requires pywin32. Install it with: pip install pywin32"
        ) from exc

    export_dir = temp_root / path.stem
    export_dir.mkdir(parents=True, exist_ok=True)

    powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
    presentation = None
    try:
        presentation = powerpoint.Presentations.Open(
            str(path.resolve()),
            WithWindow=False,
            ReadOnly=True,
        )
        width = int(round((presentation.PageSetup.SlideWidth / 72.0) * dpi))
        height = int(round((presentation.PageSetup.SlideHeight / 72.0) * dpi))
        presentation.Export(str(export_dir), "PNG", width, height)
    finally:
        if presentation is not None:
            presentation.Close()
        powerpoint.Quit()

    exported = sorted(export_dir.glob("*.PNG"), key=slide_export_sort_key)
    if not exported:
        exported = sorted(export_dir.glob("*.png"), key=slide_export_sort_key)
    if not exported:
        raise RuntimeError("PowerPoint did not export any slide images.")
    return exported


def slide_export_sort_key(path: Path) -> Tuple[int, str]:
    digits = "".join(character for character in path.stem if character.isdigit())
    number = int(digits) if digits else 0
    return number, path.name.lower()


def load_ppt_pages(path: Path, dpi: int, temp_root: Path) -> List[Image.Image]:
    image_paths = export_ppt_to_png(path, dpi, temp_root)
    pages = []
    for image_path in image_paths:
        with Image.open(image_path) as image:
            pages.append(image.convert("RGB"))
    return pages


def load_pages(path: Path, dpi: int, temp_root: Path) -> List[Image.Image]:
    suffix = path.suffix.lower()
    if suffix in PPT_EXTS:
        return load_ppt_pages(path, dpi, temp_root)
    if suffix not in IMAGE_EXTS:
        raise ValueError(
            "Unsupported input type '{}'. Use an image, PPT, or PPTX file.".format(
                path.suffix
            )
        )

    with Image.open(path) as image:
        return [image.convert("RGB")]


def fit_to_same_canvas(
    expected: Image.Image,
    actual: Image.Image,
    background: Tuple[int, int, int] = (255, 255, 255),
) -> Tuple[Image.Image, Image.Image]:
    width = max(expected.width, actual.width)
    height = max(expected.height, actual.height)

    expected_canvas = Image.new("RGB", (width, height), background)
    actual_canvas = Image.new("RGB", (width, height), background)
    expected_canvas.paste(expected.convert("RGB"), (0, 0))
    actual_canvas.paste(actual.convert("RGB"), (0, 0))

    return expected_canvas, actual_canvas


def build_threshold_mask(diff: Image.Image, threshold: int) -> Image.Image:
    channel_masks = [
        channel.point(lambda value: 255 if value > threshold else 0)
        for channel in diff.convert("RGB").split()
    ]
    return ImageChops.lighter(ImageChops.lighter(channel_masks[0], channel_masks[1]), channel_masks[2])


def count_mask_pixels(mask: Image.Image) -> int:
    histogram = mask.histogram()
    return sum(count for value, count in enumerate(histogram) if value > 0)


def mask_to_regions(mask: Image.Image, min_area: int = 16, padding: int = 3) -> List[Tuple[int, int, int, int]]:
    binary = mask.convert("1")
    width, height = binary.size
    pixels = binary.load()
    visited = set()
    regions = []

    for start_y in range(height):
        for start_x in range(width):
            point = (start_x, start_y)
            if point in visited or not pixels[start_x, start_y]:
                continue

            stack = [point]
            visited.add(point)
            min_x = max_x = start_x
            min_y = max_y = start_y
            area = 0

            while stack:
                x, y = stack.pop()
                area += 1
                min_x = min(min_x, x)
                max_x = max(max_x, x)
                min_y = min(min_y, y)
                max_y = max(max_y, y)

                for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
                    if nx < 0 or ny < 0 or nx >= width or ny >= height:
                        continue
                    neighbor = (nx, ny)
                    if neighbor in visited or not pixels[nx, ny]:
                        continue
                    visited.add(neighbor)
                    stack.append(neighbor)

            if area >= min_area:
                regions.append(
                    (
                        max(0, min_x - padding),
                        max(0, min_y - padding),
                        min(width, max_x + 1 + padding),
                        min(height, max_y + 1 + padding),
                    )
                )

    return regions


def make_overlay(
    base: Image.Image,
    regions: Sequence[Tuple[int, int, int, int]],
    highlight_color: Tuple[int, int, int],
    alpha: int,
) -> Image.Image:
    overlay = base.convert("RGBA")
    highlight_layer = Image.new("RGBA", overlay.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(highlight_layer)
    fill = (highlight_color[0], highlight_color[1], highlight_color[2], alpha)
    outline = (highlight_color[0], highlight_color[1], highlight_color[2], 255)

    for region in regions:
        draw.rectangle(region, fill=fill, outline=outline)
        for inset in range(1, 4):
            x1, y1, x2, y2 = region
            if x2 - x1 > inset * 2 and y2 - y1 > inset * 2:
                draw.rectangle((x1 + inset, y1 + inset, x2 - inset, y2 - inset), outline=outline)

    return Image.alpha_composite(overlay, highlight_layer).convert("RGB")


def compare_page(
    expected: Image.Image,
    actual: Image.Image,
    page_number: int,
    expected_page: Optional[int],
    actual_page: Optional[int],
    match_score: Optional[float],
    match_status: str,
    output_dir: Path,
    threshold: int,
    allowed_percent: float,
    highlight_color: Tuple[int, int, int],
    alpha: int,
) -> PageDiff:
    expected_canvas, actual_canvas = fit_to_same_canvas(expected, actual)
    diff = ImageChops.difference(expected_canvas, actual_canvas)
    mask = build_threshold_mask(diff, threshold)

    different_pixels = count_mask_pixels(mask)
    compared_pixels = mask.width * mask.height
    difference_percent = (different_pixels / compared_pixels) * 100 if compared_pixels else 0
    extrema = diff.getextrema()
    max_channel_delta = max(channel_max for _channel_min, channel_max in extrema)
    bbox = mask.getbbox()
    passed = difference_percent <= allowed_percent
    regions = mask_to_regions(mask)

    overlay = make_overlay(actual_canvas, regions, highlight_color, alpha)
    expected_label = expected_page if expected_page is not None else 0
    actual_label = actual_page if actual_page is not None else 0
    file_stem = "match_{:03d}_expected_{:03d}_actual_{:03d}".format(page_number, expected_label, actual_label)
    overlay_path = output_dir / "{}_overlay.png".format(file_stem)
    mask_path = output_dir / "{}_mask.png".format(file_stem)
    overlay.save(overlay_path)
    mask.save(mask_path)

    return PageDiff(
        page=page_number,
        expected_page=expected_page,
        actual_page=actual_page,
        match_score=round(match_score, 6) if match_score is not None else None,
        match_status=match_status,
        width=mask.width,
        height=mask.height,
        compared_pixels=compared_pixels,
        different_pixels=different_pixels,
        difference_percent=round(difference_percent, 6),
        max_channel_delta=max_channel_delta,
        bbox=bbox,
        passed=passed,
        output_overlay=str(overlay_path),
        output_mask=str(mask_path),
        regions=regions,
    )



def slide_signature(image: Image.Image, size: Tuple[int, int] = (64, 36)) -> Image.Image:
    return image.convert("L").resize(size)


def slide_similarity(expected: Image.Image, actual: Image.Image) -> float:
    left = slide_signature(expected)
    right = slide_signature(actual)
    diff = ImageChops.difference(left, right)
    histogram = diff.histogram()
    total = sum(histogram)
    if total == 0:
        return 1.0
    difference_sum = sum(value * count for value, count in enumerate(histogram))
    mean_difference = difference_sum / float(total)
    return max(0.0, 1.0 - (mean_difference / 255.0))


def align_slide_pages(
    expected_pages: Sequence[Image.Image],
    actual_pages: Sequence[Image.Image],
    min_match_score: float,
) -> List[SlideMatch]:
    expected_count = len(expected_pages)
    actual_count = len(actual_pages)
    if expected_count == 0 and actual_count == 0:
        return []

    scores = [
        [slide_similarity(expected_pages[i], actual_pages[j]) for j in range(actual_count)]
        for i in range(expected_count)
    ]
    gap_penalty = -0.05
    dp = [[0.0 for _ in range(actual_count + 1)] for _ in range(expected_count + 1)]
    step = [["" for _ in range(actual_count + 1)] for _ in range(expected_count + 1)]

    for i in range(1, expected_count + 1):
        dp[i][0] = dp[i - 1][0] + gap_penalty
        step[i][0] = "missing_actual"
    for j in range(1, actual_count + 1):
        dp[0][j] = dp[0][j - 1] + gap_penalty
        step[0][j] = "extra_actual"

    for i in range(1, expected_count + 1):
        for j in range(1, actual_count + 1):
            score = scores[i - 1][j - 1]
            match_gain = score - min_match_score
            candidates = (
                (dp[i - 1][j - 1] + match_gain, "matched"),
                (dp[i - 1][j] + gap_penalty, "missing_actual"),
                (dp[i][j - 1] + gap_penalty, "extra_actual"),
            )
            best_score, best_step = max(candidates, key=lambda item: item[0])
            dp[i][j] = best_score
            step[i][j] = best_step

    matches: List[SlideMatch] = []
    i = expected_count
    j = actual_count
    while i > 0 or j > 0:
        current = step[i][j]
        if current == "matched":
            score = scores[i - 1][j - 1]
            if score >= min_match_score:
                matches.append(SlideMatch(i - 1, j - 1, score, "matched"))
            else:
                matches.append(SlideMatch(i - 1, None, None, "missing_actual"))
                matches.append(SlideMatch(None, j - 1, None, "extra_actual"))
            i -= 1
            j -= 1
        elif current == "missing_actual" or j == 0:
            matches.append(SlideMatch(i - 1, None, None, "missing_actual"))
            i -= 1
        else:
            matches.append(SlideMatch(None, j - 1, None, "extra_actual"))
            j -= 1

    matches.reverse()
    return matches


def index_slide_pages(expected_pages: Sequence[Image.Image], actual_pages: Sequence[Image.Image]) -> List[SlideMatch]:
    page_count = max(len(expected_pages), len(actual_pages))
    matches: List[SlideMatch] = []
    for index in range(page_count):
        expected_index = index if index < len(expected_pages) else None
        actual_index = index if index < len(actual_pages) else None
        if expected_index is not None and actual_index is not None:
            matches.append(SlideMatch(expected_index, actual_index, None, "same_index"))
        elif expected_index is not None:
            matches.append(SlideMatch(expected_index, None, None, "missing_actual"))
        else:
            matches.append(SlideMatch(None, actual_index, None, "extra_actual"))
    return matches


def compare_pixels(args: argparse.Namespace, output_dir: Path) -> List[PageDiff]:
    temp_root = Path(tempfile.mkdtemp(prefix="report_diff_render_"))
    try:
        expected_pages = load_pages(Path(args.expected), args.dpi, temp_root / "expected")
        actual_pages = load_pages(Path(args.actual), args.dpi, temp_root / "actual")
    finally:
        shutil.rmtree(str(temp_root), ignore_errors=True)

    min_match_score = getattr(args, "min_match_score", 0.82)
    align_slides = bool(getattr(args, "align_slides", False))
    if align_slides:
        matches = align_slide_pages(expected_pages, actual_pages, min_match_score)
    else:
        matches = index_slide_pages(expected_pages, actual_pages)

    results = []
    blank = Image.new("RGB", (1, 1), (255, 255, 255))

    for compare_index, match in enumerate(matches, start=1):
        expected = expected_pages[match.expected_index] if match.expected_index is not None else blank
        actual = actual_pages[match.actual_index] if match.actual_index is not None else blank
        results.append(
            compare_page(
                expected=expected,
                actual=actual,
                page_number=compare_index,
                expected_page=match.expected_index + 1 if match.expected_index is not None else None,
                actual_page=match.actual_index + 1 if match.actual_index is not None else None,
                match_score=match.score,
                match_status=match.status,
                output_dir=output_dir,
                threshold=args.threshold,
                allowed_percent=args.allowed_percent,
                highlight_color=args.highlight_color,
                alpha=args.alpha,
            )
        )

    return results


def hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def emu_to_points(value: int) -> float:
    return round(float(value) / EMU_PER_POINT, 3)


def shape_to_record(shape: Any) -> Dict[str, Any]:
    record = {
        "name": safe_text(getattr(shape, "name", "")),
        "shape_type": safe_text(getattr(shape, "shape_type", "")),
        "left_pt": emu_to_points(getattr(shape, "left", 0)),
        "top_pt": emu_to_points(getattr(shape, "top", 0)),
        "width_pt": emu_to_points(getattr(shape, "width", 0)),
        "height_pt": emu_to_points(getattr(shape, "height", 0)),
        "text": "",
        "image_sha256": "",
    }

    if getattr(shape, "has_text_frame", False):
        record["text"] = safe_text(shape.text_frame.text)

    if hasattr(shape, "image"):
        try:
            record["image_sha256"] = hash_bytes(shape.image.blob)
        except Exception:
            record["image_sha256"] = ""

    return record


def extract_pptx_objects(path: Path) -> List[List[Dict[str, Any]]]:
    try:
        from pptx import Presentation  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError(
            "PPTX object comparison requires python-pptx. Install it with: pip install python-pptx"
        ) from exc

    presentation = Presentation(str(path))
    slides = []
    for slide in presentation.slides:
        slides.append([shape_to_record(shape) for shape in slide.shapes])
    return slides


def compare_objects(args: argparse.Namespace, output_dir: Path) -> List[ObjectDiff]:
    expected_path = Path(args.expected)
    actual_path = Path(args.actual)
    if expected_path.suffix.lower() != PPTX_EXT or actual_path.suffix.lower() != PPTX_EXT:
        raise ValueError("Object comparison currently supports .pptx files only.")

    expected_slides = extract_pptx_objects(expected_path)
    actual_slides = extract_pptx_objects(actual_path)
    slide_count = max(len(expected_slides), len(actual_slides))
    diffs = []

    for slide_index in range(slide_count):
        expected_shapes = expected_slides[slide_index] if slide_index < len(expected_slides) else []
        actual_shapes = actual_slides[slide_index] if slide_index < len(actual_slides) else []
        object_count = max(len(expected_shapes), len(actual_shapes))

        for object_index in range(object_count):
            expected = expected_shapes[object_index] if object_index < len(expected_shapes) else {}
            actual = actual_shapes[object_index] if object_index < len(actual_shapes) else {}
            fields = sorted(set(expected.keys()) | set(actual.keys()))
            for field in fields:
                if expected.get(field) != actual.get(field):
                    diffs.append(
                        ObjectDiff(
                            slide=slide_index + 1,
                            object_index=object_index + 1,
                            field=field,
                            expected=expected.get(field),
                            actual=actual.get(field),
                        )
                    )

    object_summary = output_dir / "object_summary.json"
    object_summary.write_text(
        json.dumps([asdict(diff) for diff in diffs], indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return diffs


def get_package_root() -> Path:
    return Path(__file__).resolve().parent


def resolve_output_dir(output_dir: str) -> Path:
    package_root = get_package_root()
    requested = Path(output_dir)
    if not requested.is_absolute():
        requested = package_root / requested

    resolved = requested.resolve()
    try:
        resolved.relative_to(package_root)
    except ValueError:
        raise ValueError(
            "Output directory must stay inside the tool package folder: {}".format(
                package_root
            )
        )

    return resolved


def validate_inputs(args: argparse.Namespace) -> None:
    expected_path = Path(args.expected)
    actual_path = Path(args.actual)

    if not expected_path.exists():
        raise FileNotFoundError("Expected report not found: {}".format(expected_path))
    if not actual_path.exists():
        raise FileNotFoundError("Actual report not found: {}".format(actual_path))

    if args.threshold < 0 or args.threshold > 255:
        raise ValueError("--threshold must be between 0 and 255")
    if getattr(args, "min_match_score", 0.82) < 0 or getattr(args, "min_match_score", 0.82) > 1:
        raise ValueError("--min-match-score must be between 0 and 1")
    if args.allowed_percent < 0 or args.allowed_percent > 100:
        raise ValueError("--allowed-percent must be between 0 and 100")
    if args.alpha < 0 or args.alpha > 255:
        raise ValueError("--alpha must be between 0 and 255")
    if args.dpi <= 0:
        raise ValueError("--dpi must be greater than 0")


def write_summary(output_dir: Path, pixel_results: Sequence[PageDiff], object_diffs: Sequence[ObjectDiff]) -> None:
    payload = {
        "pixel_results": [asdict(result) for result in pixel_results],
        "object_diffs": [asdict(diff) for diff in object_diffs],
    }
    (output_dir / "summary.json").write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def print_summary(pixel_results: Iterable[PageDiff], object_diffs: Sequence[ObjectDiff]) -> None:
    total_pages = 0
    failed_pages = 0

    for result in pixel_results:
        total_pages += 1
        if not result.passed:
            failed_pages += 1

        status = "PASS" if result.passed else "FAIL"
        print(
            "[{}] pair={} expected={} actual={} match={} score={} diff={:.6f}% pixels={}/{} bbox={} overlay={}".format(
                status,
                result.page,
                result.expected_page,
                result.actual_page,
                result.match_status,
                result.match_score,
                result.difference_percent,
                result.different_pixels,
                result.compared_pixels,
                result.bbox,
                result.output_overlay,
            )
        )

    if total_pages:
        print("Pages compared: {}; failed pages: {}".format(total_pages, failed_pages))

    if object_diffs:
        print("Object differences: {}".format(len(object_diffs)))
        for diff in object_diffs[:20]:
            print(
                "[OBJECT] slide={} object={} field={}".format(
                    diff.slide,
                    diff.object_index,
                    diff.field,
                )
            )
        if len(object_diffs) > 20:
            print("... {} more object differences in object_summary.json".format(len(object_diffs) - 20))
    elif total_pages == 0:
        print("Object differences: 0")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Compare report files by pixels and/or PowerPoint objects."
    )
    parser.add_argument("expected", help="Baseline report image/PPT/PPTX")
    parser.add_argument("actual", help="Report image/PPT/PPTX to check")
    parser.add_argument(
        "-o",
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for overlay images, masks, and JSON summaries",
    )
    parser.add_argument(
        "--mode",
        choices=("pixel", "object", "both"),
        default="pixel",
        help="pixel renders pages/slides; object compares PPTX shapes; both runs both checks",
    )
    parser.add_argument(
        "--align-slides",
        action="store_true",
        help="Auto-match slides by visual similarity before comparing",
    )
    parser.add_argument(
        "--min-match-score",
        type=float,
        default=0.82,
        help="Minimum visual similarity score for auto slide matching, 0.0 to 1.0",
    )
    parser.add_argument(
        "-t",
        "--threshold",
        type=int,
        default=0,
        help="Ignore per-channel pixel differences at or below this value",
    )
    parser.add_argument(
        "--allowed-percent",
        type=float,
        default=0.0,
        help="Maximum allowed percent of different pixels before a page fails",
    )
    parser.add_argument(
        "--highlight-color",
        type=parse_color,
        default=(255, 0, 0),
        help="Overlay color as R,G,B",
    )
    parser.add_argument(
        "--alpha",
        type=int,
        default=51,
        help="Highlight fill opacity from 0 to 255; 51 is 20 percent",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=150,
        help="DPI used when rendering PowerPoint slides",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        validate_inputs(args)
        output_dir = resolve_output_dir(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        pixel_results = []
        object_diffs = []

        if args.mode in ("pixel", "both"):
            pixel_results = compare_pixels(args, output_dir)
        if args.mode in ("object", "both"):
            object_diffs = compare_objects(args, output_dir)

        write_summary(output_dir, pixel_results, object_diffs)
    except Exception as exc:
        print("Error: {}".format(exc), file=sys.stderr)
        return 2

    print_summary(pixel_results, object_diffs)
    pixels_passed = all(result.passed for result in pixel_results)
    objects_passed = len(object_diffs) == 0
    return 0 if pixels_passed and objects_passed else 1


if __name__ == "__main__":
    raise SystemExit(main())






