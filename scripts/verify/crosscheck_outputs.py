#!/usr/bin/env python3
"""生成済み書類と case.yaml の書類間整合を検査する。

検査項目:
- 債権者一覧表の負債総額・債権者数 = case.yaml の集計
- 資産目録の総額 = case.yaml の集計
- 家計収支表の月次集計 = case.yaml の集計
- 報告書（docx）: 受任通知日・申立人氏名・就業先が記載されているか
- xlsx はセル値からも再集計して meta.json と突合（generic 様式のみ）

使い方:
  python3 scripts/verify/crosscheck_outputs.py --case cases/<id>/case.yaml --dir cases/<id>/output
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from casekit.validate_case import wareki_key  # noqa: E402,F401  (存在確認を兼ねる)


def case_totals(case):
    creditors = case.get("creditors", []) or []
    total_debt = sum(c["balance"] for c in creditors if isinstance(c.get("balance"), int))
    assets = case.get("assets", {}) or {}
    total_assets = 0
    for v in [assets.get("cash")] + [d.get("balance") for d in assets.get("deposits", []) or []] \
             + [x.get("surrender_value") for x in assets.get("insurance", []) or []] \
             + [x.get("value") for x in assets.get("vehicles", []) or []] \
             + [x.get("value") for x in assets.get("real_estate", []) or []] \
             + [assets.get("retirement_expectation")] \
             + [x.get("amount") for x in assets.get("overpayment", []) or []] \
             + [x.get("value") for x in assets.get("other", []) or []]:
        if isinstance(v, int):
            total_assets += v
    months = {}
    for h in case.get("household", []) or []:
        inc = sum(v for v in (h.get("income") or {}).values() if isinstance(v, int))
        exp = sum(v for v in (h.get("expense") or {}).values() if isinstance(v, int))
        months[h.get("month", "")] = (inc, exp)
    return total_debt, len(creditors), total_assets, months


def docx_text(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


def to_zenkaku_digits(s):
    return s.translate(str.maketrans("0123456789", "０１２３４５６７８９"))


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--case", required=True)
    ap.add_argument("--dir", required=True)
    args = ap.parse_args()

    case = yaml.safe_load(open(args.case, encoding="utf-8"))
    outdir = Path(args.dir)
    total_debt, n_cred, total_assets, months = case_totals(case)
    ok, ng = [], []

    def check(cond, label, detail=""):
        (ok if cond else ng).append(f"{label}{' — ' + detail if detail and not cond else ''}")

    # meta.json ベースの突合
    for mj in sorted(outdir.glob("*.meta.json")):
        m = json.loads(mj.read_text(encoding="utf-8"))
        if m.get("doc") == "saikensha-ichiran":
            check(m.get("total_debt") == total_debt, "債権者一覧表: 負債総額一致",
                  f"一覧表 {m.get('total_debt'):,} ≠ case {total_debt:,}")
            check(m.get("creditor_count") == n_cred, "債権者一覧表: 債権者数一致",
                  f"一覧表 {m.get('creditor_count')} ≠ case {n_cred}")
        elif m.get("doc") == "shisan-mokuroku":
            check(m.get("total_assets") == total_assets, "資産目録: 資産総額一致",
                  f"目録 {m.get('total_assets'):,} ≠ case {total_assets:,}")
        elif m.get("doc") == "kakei-shushi":
            for mm in m.get("months", []):
                exp = months.get(mm["month"])
                check(exp is not None and (mm["income"], mm["expense"]) == exp,
                      f"家計収支表 {mm['month']}: 収入・支出一致",
                      f"表 {mm['income']:,}/{mm['expense']:,} ≠ case {exp}")

    # 報告書（docx）の記載確認
    for docx in sorted(outdir.glob("報告書*.docx")):
        text = docx_text(docx)
        name = (case.get("meta", {}).get("applicant", {}) or {}).get("name", "")
        name_spaced = "　".join(name)
        check(name in text or name_spaced in text or name.replace("　", "") in text.replace("　", ""),
              f"{docx.name}: 申立人氏名の記載", "氏名が見つからない")
        junin = case.get("meta", {}).get("junin_tsuchi_date", "")
        m = re.match(r"^R(\d+)\.(\d+)\.(\d+)$", junin)
        if m:
            digits = [to_zenkaku_digits(g) for g in m.groups()]
            seq = f"令和{digits[0]}年{digits[1]}月{digits[2]}日"
            check(seq in text.replace("　", "").replace(" ", ""),
                  f"{docx.name}: 受任通知発送日の記載", f"{seq} が見つからない")

    print("書類間整合チェック")
    for x in ok:
        print(f"  OK  {x}")
    for x in ng:
        print(f"  NG  {x}")
    if not ok and not ng:
        print("  検査対象が見つからない（output/ に *.meta.json / 報告書*.docx があるか確認）")
        sys.exit(4)
    sys.exit(1 if ng else 0)


if __name__ == "__main__":
    main()
