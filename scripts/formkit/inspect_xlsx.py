#!/usr/bin/env python3
"""xlsx 書式のセルマップをダンプする（fillmap 起草・版間照合の土台）。

シートごとに、非空セルの (アドレス, 値, 結合範囲, 数式) を出力する。
裁判所配布の Excel 書式を登録するとき、どのセルがラベルでどこが記入欄かを
把握するために使う。

使い方:
  python3 scripts/formkit/inspect_xlsx.py <template.xlsx> [-o cellmap.txt]
"""
import argparse
import sys

import openpyxl


def inspect(path, out):
    wb = openpyxl.load_workbook(path, data_only=False)
    for ws in wb.worksheets:
        merges = list(ws.merged_cells.ranges)
        print(f"=== sheet: {ws.title} (dim={ws.dimensions}, "
              f"merged={len(merges)}, 印刷範囲={ws.print_area or '-'})", file=out)
        if merges:
            print("merged: " + " ".join(str(m) for m in merges), file=out)
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                kind = "f" if isinstance(c.value, str) and str(c.value).startswith("=") else "v"
                print(f"{c.coordinate}\t{kind}\t{c.value}", file=out)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("xlsx")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()
    out = open(args.output, "w", encoding="utf-8") if args.output else sys.stdout
    inspect(args.xlsx, out)
    if args.output:
        out.close()
        print(f"-> {args.output}")


if __name__ == "__main__":
    main()
