#!/usr/bin/env python3
"""生成した docx/xlsx を PDF 化し、ページ画像（PNG）に展開する目視検証ツール。

PDF変換は soffice（LibreOffice）→ unoconvert（unoserver。Cowork のVM構成）の順で
試す。どちらも無い環境では PDF目視を省略する旨を表示して正常終了する（劣化運転。
数値整合検査 crosscheck は別途動くため、パイプラインは止めない）。
PDF→画像は PyMuPDF（pdftoppm が無い環境向け）。

使い方:
  python3 scripts/verify/render_preview.py <file.docx|file.xlsx> [-d 出力ディレクトリ]

出力: <出力ディレクトリ>/<basename>.pdf と <basename>_p01.png, _p02.png, ...
既定の出力ディレクトリは入力ファイルと同じ場所。
"""
import argparse
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:
    pymupdf = None

DEGRADED_MSG = ("PDF変換ツール（soffice/unoconvert）が無いため目視検証を省略します。"
                "生成ファイルを Word/Excel で開いて確認してください"
                "（数値整合は hasan-kit crosscheck で検査済みであること）")


def to_pdf(src: Path, outdir: Path) -> Path | None:
    pdf = outdir / (src.stem + ".pdf")
    if shutil.which("soffice"):
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)],
            check=True, capture_output=True, timeout=120,
        )
    elif shutil.which("unoconvert"):
        subprocess.run(["unoconvert", "--convert-to", "pdf", str(src), str(pdf)],
                       check=True, capture_output=True, timeout=120)
    else:
        return None
    if not pdf.exists():
        raise RuntimeError(f"PDF変換に失敗: {src}")
    return pdf


def to_images(pdf: Path, outdir: Path, dpi=110):
    if pymupdf is None:
        print(f"  PyMuPDF が無いためページ画像化を省略（PDFのみ生成: {pdf.name}）")
        return []
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

    if not shutil.which("soffice") and not shutil.which("unoconvert"):
        print(DEGRADED_MSG)
        sys.exit(0)

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
