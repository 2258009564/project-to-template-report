"""Trim the empty area below the last text row of a terminal screenshot.

Usage:
    python scripts/trim_console_shot.py shot-raw.png [shot.png] [options]

The dominant colour of the image is treated as the terminal background; rows
that contain at least --min-pixels pixels differing from it by more than
--threshold are kept. The result keeps --margin pixels below the last text row.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PIL import Image


def dominant_colour(image: Image.Image) -> tuple[int, int, int]:
    small = image.resize((max(1, image.width // 4), max(1, image.height // 4)))
    counts: dict[tuple[int, int, int], int] = {}
    for y in range(small.height):
        for x in range(small.width):
            pixel = small.getpixel((x, y))
            counts[pixel] = counts.get(pixel, 0) + 1
    return max(counts, key=counts.get)


def last_text_row(image: Image.Image, background: tuple[int, int, int],
                  threshold: int, min_pixels: int) -> int:
    pixels = image.load()
    last = 0
    for y in range(image.height):
        hits = 0
        for x in range(0, image.width, 2):
            r, g, b = pixels[x, y]
            if abs(r - background[0]) + abs(g - background[1]) + abs(b - background[2]) > threshold:
                hits += 1
                if hits >= min_pixels:
                    break
        if hits >= min_pixels:
            last = y
    return last


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('dest', nargs='?', type=Path)
    parser.add_argument('--margin', type=int, default=16, help='pixels kept below the last text row')
    parser.add_argument('--inset', type=int, default=0, help='pixels removed from every side first')
    parser.add_argument('--threshold', type=int, default=60, help='per-channel colour distance')
    parser.add_argument('--min-pixels', type=int, default=8, help='pixels per row that count as text')
    parser.add_argument('--min-height', type=int, default=140, help='never crop below this height')
    args = parser.parse_args(argv)

    if not args.source.exists():
        parser.error('source image not found: %s' % args.source)
    dest = args.dest or args.source

    image = Image.open(args.source).convert('RGB')
    if args.inset > 0:
        image = image.crop((args.inset, args.inset,
                            image.width - args.inset, image.height - args.inset))
    background = dominant_colour(image)
    last = last_text_row(image, background, args.threshold, args.min_pixels)
    bottom = min(image.height, last + 1 + args.margin)
    bottom = max(bottom, min(args.min_height, image.height))
    image.crop((0, 0, image.width, bottom)).save(dest)
    print('%s %dx%d -> %dx%d' % (dest, image.width, image.height, image.width, bottom))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
