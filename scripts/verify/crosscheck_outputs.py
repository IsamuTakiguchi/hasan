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
    """docx の本文テキスト。壊れた・偽の docx は None（呼び出し側で NG 扱い）。"""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8")
    except Exception:
        return None
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

    # 出力形式の監査: 書式名を含むファイルの拡張子が本来の形式（registry の doc_type）と
    # 一致しているか。Excel 書式が Word で「自作」された事故（Cowork 等でビルダーを
    # 経由しなかった場合）を検出する。キーワードは前方一致優先（財産目録 vs 添付目録等）。
    FORM_EXT = [
        ("資産及び負債一覧表", ".xlsx"), ("債権者一覧表", ".xlsx"),
        ("添付目録", ".docx"), ("疎明資料目録", ".docx"), ("引継資料一覧表", ".docx"),
        ("財産目録", ".xlsx"), ("資産目録", ".xlsx"), ("家計収支表", ".xlsx"),
        ("リース物件", ".xlsx"), ("訴訟", ".xlsx"), ("処分行為", ".xlsx"),
        ("公租公課チェック表", ".xlsx"), ("チェック表", ".xlsx"),
        ("標準資料一覧表", ".xlsx"),
        ("自由財産拡張", ".docx"), ("申立書", ".docx"), ("報告書", ".docx"),
    ]
    for f in sorted(outdir.iterdir()):
        if f.suffix not in (".docx", ".xlsx"):
            continue
        for kw, ext in FORM_EXT:
            if kw in f.name:
                check(f.suffix == ext, f"{f.name}: 出力形式が書式どおり（{ext}）",
                      f"この書式の正本は {ext}（裁判所配布テンプレート）なのに {f.suffix} で"
                      f"出力されている。builder/fill-docx を経由せず自作された可能性 — "
                      f"正規の生成手順で作り直すこと")
                break

    # meta.json ベースの突合
    metas = {}
    for mj in sorted(outdir.glob("*.meta.json")):
        m = json.loads(mj.read_text(encoding="utf-8"))
        metas[m.get("doc")] = m

    # 債権者一覧表: 一般用（B1105）＋公租公課用（B1106）の合算で case と突合
    if "saikensha-ichiran" in metas:
        g = metas["saikensha-ichiran"]
        k = metas.get("saikensha-ichiran-kouso", {})
        got_total = g.get("total_debt", 0) + k.get("total_debt", 0)
        got_count = g.get("creditor_count", 0) + k.get("creditor_count", 0)
        label = "債権者一覧表（一般＋公租公課）" if k else "債権者一覧表"
        check(got_total == total_debt, f"{label}: 負債総額一致",
              f"一覧表計 {got_total:,} ≠ case {total_debt:,}")
        check(got_count == n_cred, f"{label}: 債権者数一致",
              f"一覧表計 {got_count} ≠ case {n_cred}")
        n_kouso = sum(1 for c in case.get("creditors", []) if c.get("kind") == "公租公課")
        if n_kouso and not k and not g.get("kouso_excluded") is None:
            check(False, "公租公課用一覧表の生成", f"公租公課 {n_kouso} 件があるのに B1106 が未生成")

    # 管財書類間の突合（builder の meta.json 同士）
    if "saikensha-kanzai" in metas:
        m = metas["saikensha-kanzai"]
        check(m.get("total_debt") == total_debt, "債権者一覧表（管財）: 負債総額一致",
              f"一覧表 {m.get('total_debt'):,} ≠ case {total_debt:,}")
        check(m.get("creditor_count") == n_cred, "債権者一覧表（管財）: 債権者数一致",
              f"一覧表 {m.get('creditor_count')} ≠ case {n_cred}")
    if "shisan-fusai" in metas and "shisan-kanzai" in metas:
        f, s = metas["shisan-fusai"], metas["shisan-kanzai"]
        check(f.get("total_assets") == s.get("total_assets"),
              "資産及び負債一覧表: 資産計＝財産目録計",
              f"{f.get('total_assets'):,} ≠ {s.get('total_assets'):,}")
        check(f.get("total_recovery") == s.get("total_recovery"),
              "資産及び負債一覧表: 回収見込計＝財産目録回収見込計",
              f"{f.get('total_recovery'):,} ≠ {s.get('total_recovery'):,}")
    if "shisan-fusai" in metas and "saikensha-kanzai" in metas:
        f, k = metas["shisan-fusai"], metas["saikensha-kanzai"]
        check(f.get("total_debt") == k.get("total_debt"),
              "資産及び負債一覧表: 負債計＝債権者一覧表計",
              f"{f.get('total_debt'):,} ≠ {k.get('total_debt'):,}")

    for m in metas.values():
        if m.get("doc") == "shisan-mokuroku":
            check(m.get("total_assets") == total_assets, "資産目録: 資産総額一致",
                  f"目録 {m.get('total_assets'):,} ≠ case {total_assets:,}")
        elif m.get("doc") == "kakei-shushi":
            for mm in m.get("months", []):
                exp = months.get(mm["month"])
                check(exp is not None and (mm["income"], mm["expense"]) == exp,
                      f"家計収支表 {mm['month']}: 収入・支出一致",
                      f"表 {mm['income']:,}/{mm['expense']:,} ≠ case {exp}")

    def parse_man(s):
        """「２，２０４万５，０００」形式（全角は正規化済み前提）→ 円の整数。"""
        s = s.replace(",", "").strip()
        if "万" in s:
            man, _, rest = s.partition("万")
            return int(man or 0) * 10000 + int(rest or 0)
        return int(s)

    # 申立書（docx）: 申立ての理由の債権者数・債務総額が case と一致するか
    priority_debt = sum(c["balance"] for c in case.get("creditors", []) or []
                        if isinstance(c.get("balance"), int)
                        and c.get("kind") in ("労働債権", "公租公課"))
    for docx in sorted(outdir.glob("*申立書*.docx")):
        text = docx_text(docx)
        if text is None:
            check(False, f"{docx.name}: docx として読める", "壊れているか docx 形式でない")
            continue
        text = text.translate(str.maketrans("０１２３４５６７８９，", "0123456789,"))
        m = re.search(r"債権者(\d+)人に対し[,，]金([\d,]+)円", text)
        km = re.search(r"一般破産債権総額([\d,万]+)円（債権者\s*(\d+)\s*人）", text)
        if m:  # 同時廃止（B1102）
            check(int(m.group(1)) == n_cred, f"{docx.name}: 債権者数一致",
                  f"申立書 {m.group(1)} ≠ case {n_cred}")
            check(int(m.group(2).replace(",", "")) == total_debt, f"{docx.name}: 債務総額一致",
                  f"申立書 {m.group(2)} ≠ case {total_debt:,}")
        elif km:  # 管財（0203）: 一般＋優先で case と突合
            ippan = parse_man(km.group(1))
            n_ippan = int(km.group(2))
            ym = re.search(r"優先的破産債権及び財団債権総額([\d,万]+)円（債権者\s*(\d+)\s*人）", text)
            yusen = parse_man(ym.group(1)) if ym else 0
            n_yusen = int(ym.group(2)) if ym else 0
            check(ippan + yusen == total_debt, f"{docx.name}: 債務総額一致（一般＋優先）",
                  f"申立書 {ippan + yusen:,} ≠ case {total_debt:,}")
            check(yusen == priority_debt, f"{docx.name}: 優先的破産債権・財団債権額一致",
                  f"申立書 {yusen:,} ≠ case {priority_debt:,}")
            check(n_ippan + n_yusen == n_cred, f"{docx.name}: 債権者数一致（一般＋優先）",
                  f"申立書 {n_ippan + n_yusen} ≠ case {n_cred}")
            rm = re.search(r"回収見込額合計\s*([\d,万]+)円", text)
            if rm and "shisan-kanzai" in metas:
                got = parse_man(rm.group(1))
                want = metas["shisan-kanzai"].get("total_recovery")
                check(got == want, f"{docx.name}: 回収見込額合計＝財産目録回収見込計",
                      f"申立書 {got:,} ≠ 財産目録 {want:,}")
        else:
            check(False, f"{docx.name}: 申立ての理由の数値", "債権者数・金額が読み取れない（未記入の可能性）")

    # 報告書（docx）の記載確認 — 自然人用（B1110/0205）の欄に対する検査。
    # 法人用報告書（0104）には氏名・受任通知日の該当欄が無いためスキップする
    proc_type = (case.get("meta", {}) or {}).get("proc_type", "")
    for docx in sorted(outdir.glob("報告書*.docx")) if proc_type != "管財（法人）" else []:
        text = docx_text(docx)
        if text is None:
            check(False, f"{docx.name}: docx として読める", "壊れているか docx 形式でない")
            continue
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
