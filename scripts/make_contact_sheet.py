# -*- coding: utf-8 -*-
"""Build one contact sheet from export_preview.py's slide_XX.png files."""
from pathlib import Path
import argparse
import re
from PIL import Image, ImageDraw, ImageFont


def natural_key(path: Path):
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", path.name)]


def font(size: int):
    for name in ("msyh.ttc", "arial.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            pass
    return ImageFont.load_default()


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a deck contact sheet")
    parser.add_argument("preview_dir")
    parser.add_argument("-o", "--output")
    args = parser.parse_args()
    directory = Path(args.preview_dir)
    files = sorted(directory.glob("slide_*.png"), key=natural_key)
    if not files:
        raise SystemExit(f"未找到 slide_XX.png: {directory}")
    output = Path(args.output) if args.output else directory / "contact_sheet.png"
    columns = 3 if len(files) >= 7 else 2
    thumb_w, thumb_h = 480, 270
    margin, label_h = 20, 36
    rows = (len(files) + columns - 1) // columns
    sheet = Image.new("RGB", (
        margin + columns * (thumb_w + margin),
        margin + rows * (thumb_h + label_h + margin),
    ), "#E9EEF5")
    draw = ImageDraw.Draw(sheet)
    label_font = font(20)
    for index, path in enumerate(files):
        row, col = divmod(index, columns)
        x = margin + col * (thumb_w + margin)
        y = margin + row * (thumb_h + label_h + margin)
        with Image.open(path) as source:
            image = source.convert("RGB")
            image.thumbnail((thumb_w, thumb_h), Image.Resampling.LANCZOS)
            frame = Image.new("RGB", (thumb_w, thumb_h), "white")
            frame.paste(image, ((thumb_w - image.width) // 2, (thumb_h - image.height) // 2))
            sheet.paste(frame, (x, y))
        draw.rectangle((x, y, x + thumb_w - 1, y + thumb_h - 1), outline="#7BA3D8", width=1)
        draw.text((x + 4, y + thumb_h + 5), path.stem, fill="#003371", font=label_font)
    output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output)
    print(f"OK {output}（{len(files)} 页，{columns} 列）")


if __name__ == "__main__":
    main()
