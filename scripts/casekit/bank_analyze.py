#!/usr/bin/env python3
"""通帳明細（正規化CSV）から特異な取引を検出し、case.yaml の bank_analysis に
貼り込める YAML を出力する。

入力CSV（hasan-nyushukkin-bunseki スキルが通帳PDFから起こす）:
  date,account,description,in,out
  R7.1.15,戊田銀行 堺支店,ｺｳﾔﾏｶｰﾄﾞ,0,35000
  - date: 短縮元号（R7.1.15）または YYYY-MM-DD
  - in/out: 円（整数、なければ0）

検出規則:
  大口出金/大口入金     閾値以上（--threshold、既定200,000円）
  偏頗弁済疑い          支払不能時期（meta.shiharai_funo.date）以降の債権者宛て弁済
  受任通知後弁済        受任通知日（meta.junin_tsuchi_date）以降の債権者宛て弁済
  使途不明              閾値以上の現金引出（ATM・現金等の摘要）
  ギャンブル疑い        摘要のキーワード（競馬・競艇・パチンコ・宝くじ等）
  現金化疑い            金券・チケット・買取等のキーワード
  資産処分              支払不能時期の前後の大口入金（解約・売却の可能性）

使い方:
  python3 scripts/casekit/bank_analyze.py --csv cases/<id>/work/bank.csv \
      --case cases/<id>/case.yaml [--threshold 200000] [-o cases/<id>/work/bank_analysis.yaml]

出力は機械検出の「候補」。採否・説明の聴取は弁護士（スキル側）が判断し、
確認済みのものに resolved: true を付ける。
"""
import argparse
import csv
import re
import sys
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
from validate_case import wareki_key  # noqa: E402

GAMBLE_KW = ["競馬", "競輪", "競艇", "パチンコ", "パチスロ", "宝くじ", "ロト", "toto", "カジノ", "ｹｲﾊﾞ", "ﾊﾟﾁﾝｺ"]
CASHING_KW = ["金券", "チケット", "買取", "質屋", "ｷﾝｹﾝ", "ﾁｹｯﾄ", "ｶｲﾄﾘ"]
CASH_KW = ["ATM", "ＡＴＭ", "現金", "引出", "ｹﾞﾝｷﾝ", "CD", "ＣＤ"]
INCOME_KW = ["給与", "給料", "賞与", "年金", "ｷｭｳﾖ", "ｷｭｳﾘｮｳ", "ｼｮｳﾖ", "ﾈﾝｷﾝ"]  # 定期収入は検出対象外


def norm_date(s):
    s = (s or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
    if m:
        y, mo, d = map(int, m.groups())
        # 西暦→令和のみ簡易対応（2019以降）
        if y >= 2019:
            return f"R{y - 2018}.{mo}.{d}"
        return s
    return s


def contains_any(text, kws):
    t = (text or "")
    return any(k.lower() in t.lower() for k in kws)


def creditor_tokens(case):
    """債権者名の照合用トークン（社名の中核部分）を作る。"""
    toks = []
    for c in case.get("creditors", []) or []:
        name = c.get("name", "")
        core = re.sub(r"(株式会社|有限会社|合同会社|信用金庫|信用保証協会|銀行|カード|市|町|村|（.*?）)", "", name)
        for t in filter(None, [name, core]):
            if len(t) >= 2:
                toks.append((t, name))
    return toks


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--csv", required=True)
    ap.add_argument("--case", required=True)
    ap.add_argument("--threshold", type=int, default=200000)
    ap.add_argument("-o", "--output", help="YAML出力先（省略時は標準出力）")
    args = ap.parse_args()

    case = yaml.safe_load(open(args.case, encoding="utf-8"))
    meta = case.get("meta", {}) or {}
    junin = wareki_key(meta.get("junin_tsuchi_date", ""))
    funo = wareki_key(((meta.get("shiharai_funo") or {}).get("date") or "") + ".1"
                      if meta.get("shiharai_funo") else "")
    toks = creditor_tokens(case)

    findings = []
    with open(args.csv, encoding="utf-8-sig") as f:
        for i, row in enumerate(csv.DictReader(f)):
            date = norm_date(row.get("date", ""))
            dkey = wareki_key(date)
            desc = row.get("description", "")
            acct = row.get("account", "")
            inflow = int(row.get("in") or 0)
            outflow = int(row.get("out") or 0)

            def add(flag, note, amount, direction):
                findings.append({"date": date, "account": acct, "amount": amount,
                                 "direction": direction, "flag": flag,
                                 "note": f"{desc} — {note}" if note else desc})

            hit_cred = next((full for t, full in toks if t and t in desc), None)
            if outflow and hit_cred:
                if junin and dkey and dkey >= junin:
                    add("受任通知後弁済", f"債権者「{hit_cred}」宛て。受任通知後の弁済は偏頗行為", outflow, "出金")
                elif funo and dkey and dkey >= funo:
                    add("偏頗弁済疑い", f"債権者「{hit_cred}」宛て。支払不能後の弁済", outflow, "出金")
            if outflow >= args.threshold:
                if contains_any(desc, GAMBLE_KW):
                    add("ギャンブル疑い", "摘要にギャンブル関連", outflow, "出金")
                elif contains_any(desc, CASHING_KW):
                    add("現金化疑い", "摘要に換金関連", outflow, "出金")
                elif contains_any(desc, CASH_KW):
                    add("使途不明", f"{args.threshold:,}円以上の現金引出。使途の聴取が必要", outflow, "出金")
                elif not hit_cred:
                    add("大口出金", f"{args.threshold:,}円以上", outflow, "出金")
            elif outflow and (contains_any(desc, GAMBLE_KW) or contains_any(desc, CASHING_KW)):
                add("ギャンブル疑い" if contains_any(desc, GAMBLE_KW) else "現金化疑い", "", outflow, "出金")
            if inflow >= args.threshold and not contains_any(desc, INCOME_KW):
                note = "解約・売却等による資産処分の可能性。原資の聴取が必要" \
                    if funo and dkey and dkey >= funo else f"{args.threshold:,}円以上"
                add("資産処分" if (funo and dkey and dkey >= funo) else "大口入金", note, inflow, "入金")

    out = yaml.dump({"bank_analysis": findings}, allow_unicode=True, sort_keys=False, width=100)
    if args.output:
        Path(args.output).write_text(out, encoding="utf-8")
        print(f"{len(findings)} 件の検出 -> {args.output}")
    else:
        print(out)


if __name__ == "__main__":
    main()
