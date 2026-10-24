"""Trim the empty area below the last text row of a terminal screenshot.

Usage:
    python scripts/trim_console_shot.py shot-raw.png [shot.png] [options]

The dominant colour of the image is treated as the terminal background; rows
that contain at least --min-pixels pixels differing from it by more than
--threshold are kept. The result keeps --margin pixels below the last text row.
"""
from __future__ import annotations

import argparse
from collections import Counter
import sys
from pathlib import Path

from PIL import Image


def dominant_colour(image: Image.Image) -> tuple[int, int, int]:
    # Do not interpolate colours: the estimated background must occur in the image.
    pixels = image.get_flattened_data() if hasattr(image, 'get_flattened_data') else image.getdata()
    return Counter(pixels).most_common(1)[0][0]


def last_text_row(image: Image.Image, background: tuple[int, int, int],
                  threshold: int, min_pixels: int) -> int | None:
    pixels = image.load()
    last = None
    for y in range(image.height):
        hits = 0
        for x in range(image.width):
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
    parser.add_argument('--threshold', type=int, default=60, help='sum of RGB channel differences (0..765)')
    parser.add_argument('--min-pixels', type=int, default=8, help='pixels per row that count as text')
    parser.add_argument('--min-height', type=int, default=140, help='never crop below this height')
    args = parser.parse_args(argv)

    if not args.source.exists():
        parser.error('source image not found: %s' % args.source)
    if args.margin < 0 or args.inset < 0 or args.min_pixels < 1 or args.min_height < 1 or not 0 <= args.threshold <= 765:
        parser.error('margin/inset must be nonnegative, min-pixels/min-height positive, and threshold 0..765')
    dest = args.dest or args.source.with_name(args.source.stem + '-trimmed.png')
    if args.source.resolve() == dest.resolve() or (dest.exists() and args.source.samefile(dest)):
        parser.error('destination must differ from source; keep the original run evidence')

    with Image.open(args.source) as raw:
        image = raw.convert('RGB')
    if args.inset * 2 >= min(image.size):
        parser.error('inset leaves no image content')
    if args.inset > 0:
        image = image.crop((args.inset, args.inset,
                            image.width - args.inset, image.height - args.inset))
    background = dominant_colour(image)
    if args.min_pixels > image.width:
        parser.error('min-pixels exceeds the image width')
    last = last_text_row(image, background, args.threshold, args.min_pixels)
    if last is None:
        parser.error('image is blank at the selected threshold; no crop was saved')
    bottom = min(image.height, last + 1 + args.margin)
    bottom = max(bottom, min(args.min_height, image.height))
    dest.parent.mkdir(parents=True, exist_ok=True)
    image.crop((0, 0, image.width, bottom)).save(dest, format='PNG')
    print('%s %dx%d -> %dx%d' % (dest, image.width, image.height, image.width, bottom))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
