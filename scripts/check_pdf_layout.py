#!/usr/bin/env python3
"""Flag large empty horizontal bands in rendered report pages; inspect pages too."""

import argparse
import json
from pathlib import Path

import pypdfium2 as pdfium


def check_pages(path, margins_mm, skip_first=1, max_empty_fraction=0.20):
    top, right, bottom, left = [value * 72 / 25.4 for value in margins_mm]
    results = []
    with pdfium.PdfDocument(path) as document:
        for index in range(skip_first, len(document)):
            page = document[index]
            width, height = page.get_size()
            if left + right >= width or top + bottom >= height:
                raise ValueError('Margins leave no usable page area.')
            bitmap = page.render(scale=1)
            raster = bitmap.to_pil().convert('L')
            area = raster.crop((round(left), round(top), round(width - right), round(height - bottom)))
            mask = area.point(lambda value: 255 if value < 220 else 0)
            longest = run = 0
            for row in range(mask.height):
                ink = mask.crop((0, row, mask.width, row + 1)).histogram()[255]
                run = run + 1 if ink < 3 else 0
                longest = max(longest, run)
            fraction = longest / mask.height
            results.append({'page': index + 1, 'max_empty_fraction': round(fraction, 4),
                            'passed': fraction <= max_empty_fraction})
            bitmap.close()
            page.close()
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('--margins-mm', type=float, nargs=4, default=[20, 15, 20, 25],
                        metavar=('TOP', 'RIGHT', 'BOTTOM', 'LEFT'))
    parser.add_argument('--skip-first', type=int, default=1)
    parser.add_argument('--max-empty-fraction', type=float, default=0.20)
    args = parser.parse_args()
    if (args.skip_first < 0 or any(value < 0 for value in args.margins_mm)
            or not 0 < args.max_empty_fraction < 1):
        parser.error('Use nonnegative margins/skip count and a fraction between 0 and 1.')
    pages = check_pages(args.pdf, args.margins_mm, args.skip_first, args.max_empty_fraction)
    if not pages:
        parser.error('No pages were checked; reduce --skip-first.')
    passed = all(page['passed'] for page in pages)
    print(json.dumps({'pdf': str(args.pdf.resolve()), 'passed': passed, 'pages': pages}, indent=2))
    raise SystemExit(0 if passed else 1)


if __name__ == '__main__':
    main()
