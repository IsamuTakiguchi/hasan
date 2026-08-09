#!/usr/bin/env python3
"""生成した docx/xlsx を PDF 化し、ページ画像（PNG）に展開する目視検証ツール。

この環境には pdftoppm が無いため、PDF→画像は PyMuPDF を使う。

使い方:
  python3 scripts/verify/render_preview.py <file.docx|file.xlsx> [-d 出力ディレクトリ]

出力: <出力ディレクトリ>/<basename>.pdf と <basename>_p01.png, _p02.png, ...
既定の出力ディレクトリは入力ファイルと同じ場所。
"""
import argparse
import subprocess
import sys
from pathlib import Path

import pymupdf


def to_pdf(src: Path, outdir: Path) -> Path:
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)],
        check=True, capture_output=True, timeout=120,
    )
    pdf = outdir / (src.stem + ".pdf")
    if not pdf.exists():
        raise RuntimeError(f"PDF変換に失敗: {src}")
    return pdf


def to_images(pdf: Path, outdir: Path, dpi=110):
    doc = pymupdf.open(pdf)
    pages = []
    for i, page in enumerate(doc):
        png = outdir / f"{pdf.stem}_p{i + 1:02d}.png"
        page.get_pixmap(dpi=dpi).save(png)
        pages.append(png)
    doc.close()
    return pages


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("files", nargs="+")
    ap.add_argument("-d", "--outdir", help="出力ディレクトリ（既定: 入力と同じ場所）")
    ap.add_argument("--dpi", type=int, default=110)
    args = ap.parse_args()

    for f in args.files:
        src = Path(f)
        outdir = Path(args.outdir) if args.outdir else src.parent
        outdir.mkdir(parents=True, exist_ok=True)
        pdf = to_pdf(src, outdir)
        pages = to_images(pdf, outdir, dpi=args.dpi)
        print(f"{src.name}: {pdf.name} + {len(pages)}ページ画像 -> {outdir}")
        for p in pages:
            print(f"  {p}")


if __name__ == "__main__":
    main()
