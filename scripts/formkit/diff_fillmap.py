#!/usr/bin/env python3
"""旧版 fillmap を新版テンプレートに照合し、アンカーの生存状況を判定する
（書式改訂追従の中核）。

判定:
  OK      paraId がそのまま一致（同一ファイルの微修正時のみ起こる）
  REBOUND paraId は変わったが、文言＋構造パスで一意に解決できた → 自動再束縛
          （裁判所の再配布ファイルは paraId が総入替えになるのが普通で、これが主経路）
  BROKEN  一意に解決できない（文言変更・欄の増減・表構造変更）→ 人間/Claude が
          新テンプレートを読んで fillmap を修正する

underline_longtext の absorb_paraIds と timeline_rows の extra_paraIds は、
再束縛したアンカー段落に後続する同種の空欄段落（下線付き空欄／空段落、sectPr除外）を
旧版と同数だけ拾い直す。

使い方:
  python3 scripts/formkit/diff_fillmap.py \
      --old-fillmap courts/.../v4.0/fillmap.yaml \
      --new-template courts/.../v4.1/template.docx \
      [--write courts/.../v4.1/fillmap.yaml]   # 再束縛済み fillmap を書き出す
      [--report courts/.../v4.1/diff-from-v4.0.md]

exit 0=全て OK/REBOUND, 1=BROKEN あり。
"""
import argparse
import copy
import sys
import tempfile
import zipfile
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
from anchorlib import W, ParaIndex, load_document, normalize  # noqa: E402


def load_template_index(docx_path):
    with tempfile.TemporaryDirectory() as td:
        with zipfile.ZipFile(docx_path) as z:
            z.extract("word/document.xml", td)
        tree, root = load_document(Path(td) / "word" / "document.xml")
    return ParaIndex(root)


def para_flags(index, para):
    """空欄段落の分類: 'EU'=空+下線, 'E'=空, その他=None。sectPrはNone。"""
    p = para.elem
    if para.text:
        return None
    if p.find(f"{W}pPr/{W}sectPr") is not None:
        return None
    for rpr in p.iter(f"{W}rPr"):
        u = rpr.find(f"{W}u")
        if u is not None and u.get(f"{W}val", "single") != "none":
            return "EU"
    return "E"


def following_blanks(index, para, kind, count):
    """para の直後に続く同一パスの空欄段落を count 個まで返す。"""
    out = []
    started = False
    for q in index.paras:
        if q is para:
            started = True
            continue
        if not started:
            continue
        if len(out) >= count:
            break
        if q.path != para.path:
            break
        if para_flags(index, q) in (kind, "EU" if kind == "E" else kind):
            out.append(q)
        elif q.text:
            break
    return out


def resolve_one(index, anchor, results, fid, label=""):
    p, status, note = index.resolve(anchor)
    results.append({"field": fid + (f":{label}" if label else ""), "status": status, "note": note})
    return p, status


def diff(old_fillmap, index):
    new_map = copy.deepcopy(old_fillmap)
    results = []
    for field in new_map["fields"]:
        fid = field["id"]
        if "option_anchors" in field:
            for opt, spec in field["option_anchors"].items():
                anchor = {k: spec.get(k) for k in ("paraId", "context", "path", "idx")}
                p, status = resolve_one(index, anchor, results, fid, opt)
                if p is not None:
                    spec["paraId"] = p.para_id
                    spec["path"] = p.path
                    spec["idx"] = p.idx
                    if spec.get("context"):
                        spec["context"] = p.text
            continue

        p, status = resolve_one(index, field["anchor"], results, fid)
        if p is None:
            continue
        field["anchor"]["paraId"] = p.para_id
        field["anchor"]["path"] = p.path
        field["anchor"]["idx"] = p.idx
        if field["anchor"].get("context"):
            field["anchor"]["context"] = p.text

        if field["kind"] == "underline_longtext" and field.get("absorb_paraIds"):
            blanks = following_blanks(index, p, "EU", len(field["absorb_paraIds"]))
            if len(blanks) < len(field["absorb_paraIds"]):
                results.append({"field": f"{fid}:absorb", "status": "BROKEN",
                                "note": f"後続の下線空欄が{len(blanks)}件しか見つからない（旧{len(field['absorb_paraIds'])}件）"})
            field["absorb_paraIds"] = [b.para_id for b in blanks]
        if field["kind"] == "timeline_rows" and field.get("extra_paraIds"):
            blanks = following_blanks(index, p, "E", len(field["extra_paraIds"]))
            if len(blanks) < len(field["extra_paraIds"]):
                results.append({"field": f"{fid}:rows", "status": "REBOUND",
                                "note": f"後続の空欄行が{len(blanks)}件（旧{len(field['extra_paraIds'])}件）。不足分は生成時に複製される"})
            field["extra_paraIds"] = [b.para_id for b in blanks]
    return new_map, results


def find_uncovered(index, new_map):
    """新テンプレートのチェックボックス・下線段落のうち、どのフィールドからも
    参照されていないもの（新設欄の候補）。"""
    covered = set()
    for f in new_map["fields"]:
        if "anchor" in f:
            covered.add(f["anchor"].get("paraId"))
        for spec in (f.get("option_anchors") or {}).values():
            covered.add(spec.get("paraId"))
        for pid in (f.get("absorb_paraIds") or []) + (f.get("extra_paraIds") or []):
            covered.add(pid)
    out = []
    for q in index.paras:
        if q.para_id in covered:
            continue
        if "□" in q.text or "☑" in q.text:
            out.append((q.para_id, q.text[:50]))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--old-fillmap", required=True)
    ap.add_argument("--new-template", required=True)
    ap.add_argument("--write", help="再束縛済み fillmap の書き出し先")
    ap.add_argument("--report", help="差分レポート（Markdown）の書き出し先")
    args = ap.parse_args()

    old_map = yaml.safe_load(open(args.old_fillmap, encoding="utf-8"))
    index = load_template_index(args.new_template)
    new_map, results = diff(old_map, index)

    counts = {"OK": 0, "REBOUND": 0, "BROKEN": 0}
    for r in results:
        counts[r["status"]] += 1
    uncovered = find_uncovered(index, new_map)

    print(f"照合結果: OK {counts['OK']} / REBOUND {counts['REBOUND']} / BROKEN {counts['BROKEN']}")
    for r in results:
        if r["status"] == "BROKEN":
            print(f"  BROKEN  {r['field']} — {r['note']}")
    if uncovered:
        print(f"未参照のチェック欄（新設候補）: {len(uncovered)} 件")

    if args.report:
        lines = [f"# fillmap 版間照合レポート", "",
                 f"- 旧 fillmap: `{args.old_fillmap}`",
                 f"- 新テンプレート: `{args.new_template}`",
                 f"- 判定: OK {counts['OK']} / REBOUND {counts['REBOUND']} / BROKEN {counts['BROKEN']}", ""]
        if counts["BROKEN"]:
            lines += ["## BROKEN（要修正）", ""]
            lines += [f"- **{r['field']}** — {r['note']}" for r in results if r["status"] == "BROKEN"]
            lines += [""]
        rebound = [r for r in results if r["status"] == "REBOUND"]
        if rebound:
            lines += ["## REBOUND（自動再束縛済み）", ""]
            lines += [f"- {r['field']} — {r['note']}" for r in rebound]
            lines += [""]
        if uncovered:
            lines += ["## 未参照のチェック欄（新設欄の候補。fillmap への追加を検討）", ""]
            lines += [f"- `{pid}` {text}" for pid, text in uncovered]
            lines += [""]
        Path(args.report).write_text("\n".join(lines), encoding="utf-8")
        print(f"レポート -> {args.report}")

    if args.write:
        if counts["BROKEN"]:
            print(f"注意: BROKEN が残った状態の fillmap を書き出す（要修正マーク付き）")
            broken_ids = {r["field"].split(":")[0] for r in results if r["status"] == "BROKEN"}
            for f in new_map["fields"]:
                if f["id"] in broken_ids:
                    f["FIXME"] = "BROKEN — 新テンプレートを確認してアンカーを修正すること"
        with open(args.write, "w", encoding="utf-8") as fp:
            yaml.dump(new_map, fp, allow_unicode=True, sort_keys=False, width=100)
        print(f"再束縛済み fillmap -> {args.write}")

    sys.exit(1 if counts["BROKEN"] else 0)


if __name__ == "__main__":
    main()
