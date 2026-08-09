#!/usr/bin/env python3
"""fillmap 駆動の docx 記入エンジン。

裁判所配布の白紙テンプレートをコピーし、fillmap.yaml のアンカー定義に従って
values.yaml の値を書き込む。テンプレートのレイアウトは一切変更しない
（既存 run のテキスト置換・空セルへの run 挿入・□→☑ の文字置換のみ）。

値の組み立て（case.yaml からの転記ルール・短縮元号化・文章の起案）は
スキル（Claude）側の仕事。このスクリプトは values.yaml を機械的に反映する。

使い方:
  python3 scripts/build/fill_docx.py \
      --template courts/osaka-sakai/forms/houkokusho/v4.0/template.docx \
      --fillmap  courts/osaka-sakai/forms/houkokusho/v4.0/fillmap.yaml \
      --values   cases/<id>/work/houkokusho_values.yaml \
      --output   cases/<id>/output/報告書_記入済み.docx

values.yaml: {fields: {<field_id>: <値>}}。値の無い field は空欄のまま残り、
実行後に「未記入」として報告される。

fillmap の kind:
  set_text          段落全体のテキストを値で置き換え（先頭 run の書式を維持）
  insert_cell_text  空セル段落の </w:p> 直前に run を挿入（format.sz 指定可、既定20=10pt）
  checkbox          値（文字列 or リスト）に対応する「□ラベル」を「☑ラベル」に置換。
                    options: {値: リテラル} / option_anchors: {値: {paraId, literal}}
  underline_fill    段落内の下線 run（format.slot 番目、既定0）のテキストを値にする
  underline_slots   段落内の下線 run 列に値リストを順に入れる（"" はスキップ）
  underline_longtext 値を先頭下線 run に入れ、absorb_paraIds の空欄段落を削除して
                    自然折返しにする（行分割固定を避ける）
  timeline_rows     値 [{date, text}] を空欄段落列（anchor + extra_paraIds）に
                    「日付run + tab + 内容run」で書き込む。足りなければ末尾段落を複製
"""
import argparse
import copy
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import yaml
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from formkit.anchorlib import W, W14, ParaIndex, load_document, normalize  # noqa: E402

DOCX_SKILL_SCRIPTS = os.environ.get("DOCX_SKILL_SCRIPTS", "/root/.claude/skills/docx/scripts")


def make_run(text, sz=None, underline=False, preserve=True):
    run = etree.Element(f"{W}r")
    rpr = etree.SubElement(run, f"{W}rPr")
    fonts = etree.SubElement(rpr, f"{W}rFonts")
    fonts.set(f"{W}hint", "eastAsia")
    if sz:
        s = etree.SubElement(rpr, f"{W}sz")
        s.set(f"{W}val", str(sz))
    if underline:
        u = etree.SubElement(rpr, f"{W}u")
        u.set(f"{W}val", "single")
    t = etree.SubElement(run, f"{W}t")
    t.text = text
    if preserve:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def runs_of(p):
    return p.findall(f"{W}r")


def underlined_runs(p):
    out = []
    for r in runs_of(p):
        rpr = r.find(f"{W}rPr")
        if rpr is not None:
            u = rpr.find(f"{W}u")
            if u is not None and u.get(f"{W}val", "single") != "none":
                out.append(r)
    return out


def set_run_text(run, text):
    """run 内の最初の w:t のテキストを差し替える（tab 等の他要素は維持）。"""
    t = run.find(f"{W}t")
    if t is None:
        t = etree.SubElement(run, f"{W}t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


class Filler:
    def __init__(self, root):
        self.index = ParaIndex(root)
        self.filled = []
        self.skipped = []
        self.warnings = []

    def warn(self, fid, msg):
        self.warnings.append(f"{fid}: {msg}")

    def resolve(self, fid, anchor):
        p, status, note = self.index.resolve(anchor)
        if status == "BROKEN" or p is None:
            self.warn(fid, f"アンカー解決失敗 ({note})")
            return None
        if status == "REBOUND":
            self.warn(fid, f"アンカー再束縛 ({note})")
        return p

    # --- kinds -----------------------------------------------------------
    def fill_set_text(self, field, value):
        p = self.resolve(field["id"], field["anchor"])
        if p is None:
            return False
        rs = runs_of(p.elem)
        if rs:
            set_run_text(rs[0], str(value))
            for r in rs[1:]:
                p.elem.remove(r)
        else:
            sz = (field.get("format") or {}).get("sz")
            p.elem.append(make_run(str(value), sz=sz))
        return True

    def fill_insert_cell_text(self, field, value):
        p = self.resolve(field["id"], field["anchor"])
        if p is None:
            return False
        if normalize("".join(t.text or "" for t in p.elem.iter(f"{W}t"))):
            self.warn(field["id"], "空セルのはずが既にテキストあり。追記します")
        sz = (field.get("format") or {}).get("sz", 20)
        p.elem.append(make_run(str(value), sz=sz))
        return True

    def fill_checkbox(self, field, value):
        values = value if isinstance(value, list) else [value]
        ok_any = False
        for v in values:
            v = str(v)
            if "option_anchors" in field:
                spec = field["option_anchors"].get(v)
                if spec is None:
                    self.warn(field["id"], f"未知の選択肢「{v}」")
                    continue
                anchor = {**field.get("anchor", {}), **{k: spec[k] for k in ("paraId", "context", "path") if k in spec}}
                literal = spec["literal"]
            else:
                anchor = field["anchor"]
                literal = (field.get("options") or {}).get(v)
                if literal is None:
                    self.warn(field["id"], f"未知の選択肢「{v}」")
                    continue
            p = self.resolve(field["id"], anchor)
            if p is None:
                continue
            replaced = False
            for t in p.elem.iter(f"{W}t"):
                if t.text and literal in t.text:
                    t.text = t.text.replace(literal, literal.replace("□", "☑", 1), 1)
                    replaced = True
                    break
            if replaced:
                ok_any = True
            else:
                self.warn(field["id"], f"リテラル「{literal}」が段落内に見つからない（run分割の可能性）")
        return ok_any

    def fill_underline_fill(self, field, value):
        p = self.resolve(field["id"], field["anchor"])
        if p is None:
            return False
        urs = underlined_runs(p.elem)
        slot = (field.get("format") or {}).get("slot", 0)
        if slot >= len(urs):
            self.warn(field["id"], f"下線runが{len(urs)}個しかない (slot={slot})")
            return False
        set_run_text(urs[slot], str(value))
        return True

    def fill_underline_slots(self, field, value):
        p = self.resolve(field["id"], field["anchor"])
        if p is None:
            return False
        urs = underlined_runs(p.elem)
        vals = value if isinstance(value, list) else [value]
        if len(vals) > len(urs):
            self.warn(field["id"], f"値{len(vals)}個に対し下線runが{len(urs)}個")
        ok = False
        for r, v in zip(urs, vals):
            if v is None or v == "":
                continue
            set_run_text(r, str(v))
            ok = True
        return ok

    def fill_underline_longtext(self, field, value):
        p = self.resolve(field["id"], field["anchor"])
        if p is None:
            return False
        urs = underlined_runs(p.elem)
        if not urs:
            self.warn(field["id"], "下線runが見つからない")
            return False
        set_run_text(urs[0], str(value))
        for pid in field.get("absorb_paraIds", []):
            q = self.index.by_id.get(pid)
            if q is None or q.elem.getparent() is None:
                self.warn(field["id"], f"統合対象の空欄段落 {pid} が見つからない")
            elif q.elem.find(f"{W}pPr/{W}sectPr") is not None:
                self.warn(field["id"], f"段落 {pid} はsectPr付きのため削除しない（fillmapを修正すること）")
            else:
                q.elem.getparent().remove(q.elem)
        return True

    def fill_timeline_rows(self, field, value):
        if not isinstance(value, list):
            self.warn(field["id"], "値は [{date, text}] のリストで指定する")
            return False
        pids = [field["anchor"]["paraId"]] + list(field.get("extra_paraIds", []))
        paras = []
        for pid in pids:
            q = self.index.by_id.get(pid)
            if q is None:
                self.warn(field["id"], f"行段落 {pid} が見つからない")
            else:
                paras.append(q.elem)
        if not paras:
            return False
        # sectPr（セクション区切り）を含む段落は行として使わない（複製すると
        # セクションが増殖し版面が壊れる）。fillmap 登録ミスへの防御
        safe = [p for p in paras
                if p.find(f"{W}pPr") is None or p.find(f"{W}pPr/{W}sectPr") is None]
        if len(safe) < len(paras):
            self.warn(field["id"], f"sectPr付き段落{len(paras) - len(safe)}件を行から除外した（fillmapを修正すること）")
        paras = safe
        if not paras:
            return False
        # 全行にアンカー段落の pPr（タブ位置・ぶら下げインデント）を適用して
        # 体裁を揃える（書式によっては後半の空欄行に pPr が無いため）
        anchor_ppr = paras[0].find(f"{W}pPr")
        for p in paras[1:]:
            own = p.find(f"{W}pPr")
            if own is not None:
                p.remove(own)
            if anchor_ppr is not None:
                p.insert(0, copy.deepcopy(anchor_ppr))
        # 行が足りなければ最後の空欄段落を複製して増やす
        while len(paras) < len(value):
            clone = copy.deepcopy(paras[-1])
            for attr in (f"{W14}paraId", f"{W14}textId"):
                clone.attrib.pop(attr, None)
            paras[-1].addnext(clone)
            paras.append(clone)
        for p, row in zip(paras, value):
            for r in runs_of(p):
                p.remove(r)
            date_run = make_run(str(row.get("date", "")))
            tab_run = etree.Element(f"{W}r")
            etree.SubElement(tab_run, f"{W}tab")
            text_run = make_run(str(row.get("text", "")))
            for el in (date_run, tab_run, text_run):
                p.append(el)
        return True

    KINDS = {
        "set_text": fill_set_text,
        "insert_cell_text": fill_insert_cell_text,
        "checkbox": fill_checkbox,
        "underline_fill": fill_underline_fill,
        "underline_slots": fill_underline_slots,
        "underline_longtext": fill_underline_longtext,
        "timeline_rows": fill_timeline_rows,
    }

    def apply(self, fillmap, values):
        vals = values.get("fields", {})
        for field in fillmap["fields"]:
            fid = field["id"]
            if fid not in vals or vals[fid] is None or vals[fid] == "" or vals[fid] == []:
                self.skipped.append(fid)
                continue
            fn = self.KINDS.get(field["kind"])
            if fn is None:
                self.warn(fid, f"未知の kind: {field['kind']}")
                continue
            if fn(self, field, vals[fid]):
                self.filled.append(fid)
        unknown = set(vals) - {f["id"] for f in fillmap["fields"]}
        for fid in sorted(unknown):
            self.warn(fid, "fillmap に存在しない field_id（値は無視された）")


def merge_runs(unpacked: Path):
    script = Path(DOCX_SKILL_SCRIPTS) / "merge_runs.py"
    if not script.exists():
        print(f"note: merge_runs.py が見つからないためスキップ ({script})", file=sys.stderr)
        return
    subprocess.run([sys.executable, str(script), str(unpacked)],
                   cwd=DOCX_SKILL_SCRIPTS, check=True, capture_output=True)


def rezip(unpacked: Path, out: Path):
    names = []
    for p in sorted(unpacked.rglob("*")):
        if p.is_file():
            names.append(p.relative_to(unpacked).as_posix())
    # [Content_Types].xml を先頭に
    names.sort(key=lambda n: (n != "[Content_Types].xml", n))
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.write(unpacked / n, n)


def validate(out: Path, original: Path):
    script = Path(DOCX_SKILL_SCRIPTS) / "office" / "validate.py"
    if not script.exists():
        return "note: validate.py が見つからないためスキップ"
    r = subprocess.run([sys.executable, str(script), str(out.resolve()), "--original", str(original.resolve())],
                       cwd=DOCX_SKILL_SCRIPTS, capture_output=True, text=True)
    return (r.stdout + r.stderr).strip() or f"validate exit={r.returncode}"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template", required=True)
    ap.add_argument("--fillmap", required=True)
    ap.add_argument("--values", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--no-validate", action="store_true")
    args = ap.parse_args()

    fillmap = yaml.safe_load(open(args.fillmap, encoding="utf-8"))
    values = yaml.safe_load(open(args.values, encoding="utf-8"))

    with tempfile.TemporaryDirectory() as td:
        unpacked = Path(td) / "unpacked"
        unpacked.mkdir()
        with zipfile.ZipFile(args.template) as z:
            z.extractall(unpacked)
        merge_runs(unpacked)

        doc_xml = unpacked / "word" / "document.xml"
        tree, root = load_document(doc_xml)
        filler = Filler(root)
        filler.apply(fillmap, values)
        tree.write(str(doc_xml), xml_declaration=True, encoding="UTF-8", standalone=True)

        rezip(unpacked, Path(args.output))

    print(f"出力: {args.output}")
    print(f"記入済み: {len(filler.filled)} 項目")
    if filler.skipped:
        print(f"未記入（値なし・空欄のまま）: {', '.join(filler.skipped)}")
    if filler.warnings:
        print("警告:")
        for w in filler.warnings:
            print(f"  - {w}")
    if not args.no_validate:
        print("validate:", validate(Path(args.output), Path(args.template)))
    # 警告があっても出力は生成する（目視検証で判断）。BROKEN は exit code で伝える
    broken = [w for w in filler.warnings if "アンカー解決失敗" in w]
    sys.exit(2 if broken else 0)


if __name__ == "__main__":
    main()
