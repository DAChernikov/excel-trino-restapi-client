from __future__ import annotations

import argparse
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "assets" / "app_icon.ico"
ICON_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def _load_pillow():
    try:
        from PIL import Image, ImageDraw, ImageFilter
    except ImportError as exc:
        raise SystemExit("Install Pillow first: pip install pillow") from exc

    return Image, ImageDraw, ImageFilter


def _save_ico(image, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output, sizes=ICON_SIZES)
    print(f"Created {output}")


def _convert_source_icon(source: Path, output: Path) -> None:
    Image, _, _ = _load_pillow()

    source_image = Image.open(source).convert("RGBA")
    canvas = Image.new("RGBA", (1024, 1024), (0, 0, 0, 0))
    source_image.thumbnail((1024, 1024))
    left = (1024 - source_image.width) // 2
    top = (1024 - source_image.height) // 2
    canvas.alpha_composite(source_image, (left, top))
    _save_ico(canvas, output)


def _draw_icon(output: Path) -> None:
    Image, ImageDraw, ImageFilter = _load_pillow()

    size = 1024
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle((58, 58, 966, 966), radius=184, fill=(0, 0, 0, 52))
    shadow = shadow.filter(ImageFilter.GaussianBlur(18))
    image.alpha_composite(shadow)

    draw.rounded_rectangle((44, 38, 980, 980), radius=190, fill=(255, 255, 255, 255))
    draw.rounded_rectangle((48, 42, 976, 976), radius=186, outline=(236, 242, 239, 255), width=3)

    # Spreadsheet tile.
    draw.rounded_rectangle((220, 353, 420, 665), radius=28, fill=(16, 129, 69, 255))
    for x in (250, 315, 380):
        draw.line((x, 385, x, 638), fill=(246, 255, 250, 255), width=14)
    for y in (448, 520, 592):
        draw.line((240, y, 405, y), fill=(246, 255, 250, 255), width=14)

    # Excel-like front mark.
    draw.rounded_rectangle((70, 303, 300, 699), radius=22, fill=(20, 146, 76, 255))
    draw.polygon(
        [(115, 420), (172, 420), (210, 485), (250, 420), (302, 420), (238, 512), (305, 626), (250, 626), (209, 550), (166, 626), (110, 626), (178, 512)],
        fill=(255, 255, 255, 255),
    )

    # Connector.
    draw.line((415, 508, 531, 508), fill=(48, 55, 78, 255), width=18)
    draw.rounded_rectangle((526, 468, 604, 548), radius=18, fill=(48, 55, 78, 255))
    draw.rectangle((594, 488, 640, 528), fill=(48, 55, 78, 255))
    draw.line((662, 508, 720, 508), fill=(48, 55, 78, 255), width=18)

    # Trino-like client mark.
    draw.arc((665, 238, 735, 476), 102, 258, fill=(210, 210, 210, 255), width=7)
    draw.arc((816, 258, 946, 498), 205, 322, fill=(210, 210, 210, 255), width=7)
    draw.ellipse((698, 260, 745, 464), fill=(226, 0, 128, 255))
    draw.ellipse((838, 300, 927, 467), fill=(226, 0, 128, 255))
    draw.rounded_rectangle((742, 448, 832, 503), radius=4, fill=(48, 48, 48, 255))
    draw.ellipse((668, 505, 884, 696), fill=(238, 250, 252, 255))
    draw.ellipse((696, 575, 718, 598), fill=(38, 50, 54, 255))
    draw.ellipse((816, 575, 838, 598), fill=(38, 50, 54, 255))
    draw.ellipse((704, 582, 710, 588), fill=(255, 255, 255, 255))
    draw.ellipse((824, 582, 830, 588), fill=(255, 255, 255, 255))
    draw.ellipse((753, 612, 788, 641), fill=(45, 52, 56, 255))
    draw.arc((738, 621, 772, 660), 20, 160, fill=(45, 52, 56, 255), width=5)
    draw.arc((770, 621, 805, 660), 20, 160, fill=(45, 52, 56, 255), width=5)
    draw.arc((690, 653, 860, 759), 30, 150, fill=(48, 48, 48, 255), width=20)
    draw.rounded_rectangle((644, 586, 674, 638), radius=10, fill=(70, 70, 70, 255))
    draw.rounded_rectangle((881, 586, 911, 638), radius=10, fill=(70, 70, 70, 255))
    draw.ellipse((668, 612, 700, 658), fill=(238, 205, 201, 210))
    draw.ellipse((849, 612, 881, 658), fill=(238, 205, 201, 210))
    draw.ellipse((914, 453, 935, 474), fill=(60, 60, 60, 255))

    _save_ico(image, output)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create Windows .ico file for Trino Excel Client.")
    parser.add_argument(
        "--source",
        type=Path,
        help="Optional source PNG/JPG image. If omitted, a built-in fallback icon is generated.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT,
        help="Output .ico path. Default: assets/app_icon.ico.",
    )
    return parser


if __name__ == "__main__":
    args = build_parser().parse_args()
    if args.source is not None:
        _convert_source_icon(args.source, args.output)
    else:
        _draw_icon(args.output)
