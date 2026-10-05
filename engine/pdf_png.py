# -*- coding: utf-8 -*-
"""検証用：PDF を PNG にする（Read ツールで目視するため）。

    python pdf_png.py <PDF または フォルダー> [--out DIR] [--zoom 2] [--pages 1-3]

セル内の文字切れ・図のはみ出し・意図しないページ増は、数値の検算では見つからない。
書き出した PNG を必ず目視する。
"""
import argparse
import glob
import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import pymupdf


def parse_pages(spec, n):
    if not spec:
        return range(n)
    out = []
    for part in spec.split(','):
        if '-' in part:
            a, b = part.split('-')
            out.extend(range(int(a) - 1, int(b)))
        else:
            out.append(int(part) - 1)
    return [p for p in out if 0 <= p < n]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src', help='PDFファイル、または PDF が入ったフォルダー')
    ap.add_argument('--out', help='出力先（既定：PDFと同じ場所）')
    ap.add_argument('--zoom', type=float, default=2.0)
    ap.add_argument('--pages', help='対象ページ（例 1-3）。1始まり')
    a = ap.parse_args()

    srcs = sorted(glob.glob(os.path.join(a.src, '*.pdf'))) if os.path.isdir(a.src) else [a.src]
    if not srcs:
        print('PDFがありません: %s' % a.src)
        return

    mat = pymupdf.Matrix(a.zoom, a.zoom)
    for src in srcs:
        outdir = a.out or os.path.dirname(os.path.abspath(src))
        os.makedirs(outdir, exist_ok=True)
        base = os.path.splitext(os.path.basename(src))[0]
        doc = pymupdf.open(src)
        pages = parse_pages(a.pages, doc.page_count)
        for i in pages:
            dst = os.path.join(outdir, '%s_p%d.png' % (base, i + 1))
            doc[i].get_pixmap(matrix=mat).save(dst)
            print('%s  (%dページ中 %d)' % (os.path.basename(dst), doc.page_count, i + 1))
        doc.close()


if __name__ == '__main__':
    main()
