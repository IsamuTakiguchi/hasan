#!/usr/bin/env python3
"""fillmap の全アンカーが対応テンプレートに解決できるかの検査（登録の合格条件）。

diff_fillmap.py と同じ解決ロジックを、同版の fillmap×template に対して走らせる。
登録直後は全フィールドが OK になるはず（REBOUND/BROKEN が出たら fillmap の記載ミス）。

使い方:
  python3 scripts/formkit/check_anchors.py courts/.../v4.0/fillmap.yaml courts/.../v4.0/template.docx
"""
import argparse
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
from diff_fillmap import diff, load_template_index  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("fillmap")
    ap.add_argument("template")
    args = ap.parse_args()

    fillmap = yaml.safe_load(open(args.fillmap, encoding="utf-8"))
    index = load_template_index(args.template)
    _, results = diff(fillmap, index)

    counts = {"OK": 0, "REBOUND": 0, "BROKEN": 0}
    for r in results:
        counts[r["status"]] += 1
        if r["status"] != "OK":
            print(f"  {r['status']}  {r['field']} — {r['note']}")
    print(f"OK {counts['OK']} / REBOUND {counts['REBOUND']} / BROKEN {counts['BROKEN']}")
    sys.exit(0 if counts["BROKEN"] == 0 and counts["REBOUND"] == 0 else 1)


if __name__ == "__main__":
    main()
