#!/usr/bin/env python3
"""case.yaml（事件モデル）の検証。書式生成前の必須ゲート。

1. スキーマ検証（schema/case.schema.yaml, JSON Schema）
2. 算術・整合検査:
   - 債権者の残高・番号の妥当性、負債総額の集計
   - 家計収支の月次集計（収入計・支出計・収支）
   - 受任通知日と bank_analysis の「受任通知後弁済」フラグの整合
3. 出所（sources）検査: 主要な値に出所が付いているか（警告）

使い方:
  python3 scripts/casekit/validate_case.py cases/<事件ID>/case.yaml [--schema schema/case.schema.yaml]

exit 0=合格（警告があっても可）, 1=スキーマ違反または致命的な矛盾。
集計値（負債総額・資産総額・月次収支）を標準出力に出すので、書類間突合の基準にする。
"""
import argparse
import sys
from pathlib import Path

import yaml

try:
    import jsonschema
except ImportError:
    jsonschema = None

REPO = Path(__file__).resolve().parent.parent.parent


def wareki_key(s):
    """'R7.3.10' 等をおおまかな比較キーに変換（曖昧表記は None）。"""
    if not isinstance(s, str):
        return None
    import re
    m = re.match(r"^([MTSHR])(\d+)(?:\.(\d+))?(?:\.(\d+))?$", s)
    if not m:
        return None
    base = {"M": 1867, "T": 1911, "S": 1925, "H": 1988, "R": 2018}[m.group(1)]
    return (base + int(m.group(2)), int(m.group(3) or 0), int(m.group(4) or 0))


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("case_yaml")
    ap.add_argument("--schema", default=str(REPO / "schema" / "case.schema.yaml"))
    args = ap.parse_args()

    case = yaml.safe_load(open(args.case_yaml, encoding="utf-8"))
    errors, warnings = [], []

    # 1. スキーマ検証
    if jsonschema is None:
        warnings.append("jsonschema 未インストールのためスキーマ検証をスキップ")
    else:
        schema = yaml.safe_load(open(args.schema, encoding="utf-8"))
        for e in jsonschema.Draft202012Validator(schema).iter_errors(case):
            errors.append(f"schema: {'/'.join(map(str, e.absolute_path))}: {e.message[:120]}")

    # 2. 債権者
    creditors = case.get("creditors", []) or []
    total_debt = 0
    unknown_balance = []
    for i, c in enumerate(creditors):
        b = c.get("balance")
        if b is None:
            unknown_balance.append(c.get("name", f"#{i}"))
        elif b < 0:
            errors.append(f"creditors[{i}] {c.get('name')}: 残高が負 ({b})")
        else:
            total_debt += b
    nos = [c.get("no") for c in creditors if c.get("no") is not None]
    if len(nos) != len(set(nos)):
        errors.append("creditors: 債権者番号 no が重複している")

    # 3. 家計収支
    months = []
    for i, h in enumerate(case.get("household", []) or []):
        inc = sum(v for v in (h.get("income") or {}).values() if isinstance(v, int))
        exp = sum(v for v in (h.get("expense") or {}).values() if isinstance(v, int))
        months.append((h.get("month", f"#{i}"), inc, exp, inc - exp))

    # 4. 受任通知日と入出金分析の整合
    junin = wareki_key((case.get("meta") or {}).get("junin_tsuchi_date", ""))
    for i, b in enumerate(case.get("bank_analysis", []) or []):
        if b.get("flag") == "受任通知後弁済":
            d = wareki_key(b.get("date", ""))
            if junin and d and d < junin:
                errors.append(f"bank_analysis[{i}]: 受任通知後弁済フラグだが日付 {b.get('date')} が受任通知日より前")

    # 4.5 手続種別
    meta = case.get("meta", {}) or {}
    proc = meta.get("proc_type")
    if not proc:
        warnings.append("meta.proc_type（同時廃止/管財（自然人）/管財（法人））が未設定 → intake で最初に確認すること")
    elif proc == "管財（法人）":
        if not case.get("corporation"):
            warnings.append("管財（法人）事件だが corporation（債務者法人情報）が空")
        for sec in ("household", "family"):
            if case.get(sec):
                warnings.append(f"管財（法人）事件で {sec}（自然人用セクション）が記入されている → 種別の取り違えでないか確認")
    elif case.get("corporation"):
        warnings.append(f"{proc} 事件で corporation（法人セクション）が記入されている → 種別の取り違えでないか確認")

    # 5. 出所検査（主要セクション）
    sources = case.get("sources", {}) or {}
    def has_source(prefix):
        return any(k == prefix or k.startswith(prefix + ".") or k.startswith(prefix + "[") for k in sources)
    for i, c in enumerate(creditors):
        if not has_source(f"creditors[{i}]"):
            warnings.append(f"creditors[{i}] {c.get('name')}: 出所（sources）がない → 生成時は要確認扱い")
    for sec in ("career", "household", "timeline"):
        if case.get(sec) and not has_source(sec) and not any(k.startswith(sec) for k in sources):
            warnings.append(f"{sec}: 出所（sources）がない")

    # 6. 資産集計
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

    # ---- レポート ----
    print(f"事件: {(case.get('meta') or {}).get('case_id', '?')}")
    print(f"負債総額（残高判明分）: {total_debt:,} 円 / 債権者 {len(creditors)} 名"
          + (f"（残高不明: {', '.join(unknown_balance)}）" if unknown_balance else ""))
    print(f"資産総額（判明分）: {total_assets:,} 円")
    for m, inc, exp, bal in months:
        print(f"家計 {m}: 収入 {inc:,} / 支出 {exp:,} / 収支 {bal:+,}")
    oq = case.get("open_questions", []) or []
    if oq:
        print(f"未確認事項: {len(oq)} 件（questions.md を参照）")

    if warnings:
        print("\n警告:")
        for w in warnings:
            print(f"  - {w}")
    if errors:
        print("\nエラー:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    print("\nOK")


if __name__ == "__main__":
    main()
