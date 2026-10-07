#!/usr/bin/env python3
"""fillmap 駆動の docx 記入エンジン。

裁判所配布の白紙テンプレートをコピーし、fillmap.yaml のアンカー定義に従って
values.yaml の値を書き込む。テンプレートのレイアウトは一切変更しない
（既存 run のテキスト置換・空セルへの run 挿入・□→☑ の文字置換のみ）。

値の組み立て（case.yaml からの転記ルール・短縮元号化・文章の起案）は
スキル（Claude）側の仕事。このスクリプトは values.yaml を機械的に反映する。

使い方:
  python3 scripts/build/fill_docx.py \
      --template courts/osaka/forms/houkokusho/v4.0/template.docx \
      --fillmap  courts/osaka/forms/houkokusho/v4.0/fillmap.yaml \
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
import re
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


def set_run_text(run, text, drop_tabs=False):
    """run 内のテキストを差し替える（tab 等の他要素は維持）。

    run が複数の <w:t> を持つ場合（merge_runs 後もあり得る）は、先頭に全文を
    入れて残りは削除する（置換残りを防ぐ）。

    drop_tabs=True のときは run 内の <w:tab/> も削除する。merge_runs 後の
    段落では空欄用タブが値 run に混在しており、全文置換でテキストだけ
    差し替えるとタブが値の後ろに連続残留してレイアウトが崩れる（B1102 で実例）。

    drop_tabs=False（下線記入等）では、値の w:t を**タブより前**に移す。
    統合後の下線 run は [tab, tab, 値] の並びになりがちで、そのままでは
    値がタブ位置（右端）へ飛んで折り返す。値を先頭に置けば値は下線の起点
    （ラベル直後）に印字され、残るタブが罫線をタブ位置まで伸ばす。
    """
    ts = run.findall(f"{W}t")
    if not ts:
        t = etree.SubElement(run, f"{W}t")
    else:
        t = ts[0]
        for extra in ts[1:]:
            run.remove(extra)
    if drop_tabs:
        for tab in run.findall(f"{W}tab"):
            run.remove(tab)
    else:
        tabs = run.findall(f"{W}tab")
        if tabs and run.index(t) > run.index(tabs[0]):
            run.remove(t)
            run.insert(run.index(tabs[0]), t)
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


_CJK_SPACE = re.compile(r"(?<=[^\x00-\x7f]) +(?=[^\x00-\x7f])")


def clean_fill_value(value):
    """記入値の清浄化。(清浄化後の値, 実施内容リスト) を返す。

    YAML の折返し（>- 等）で組み立てた値には改行や和文中の半角スペースが
    混入しやすく、裁判所書式の字面を崩す。改行は除去して連結し、
    全角文字に挟まれた半角スペースも除去する（意図的な半角空白は
    英数字の前後にしか現れない前提）。
    """
    notes = []
    if not isinstance(value, str):
        return value, notes
    if "\n" in value or "\r" in value:
        value = value.replace("\r", "").replace("\n", "")
        notes.append("改行を除去")
    cleaned = _CJK_SPACE.sub("", value)
    if cleaned != value:
        notes.append("和文中の半角スペースを除去")
        value = cleaned
    return value, notes


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
        # フィールドコード（fldChar/instrText。EQ重ね打ちラベル等）を含む run は
        # 温存する。消すと begin/end が欠けて以降の文書全体がフィールド命令扱いになる
        rs = [
            r for r in rs
            if r.find(f"{W}fldChar") is None and r.find(f"{W}instrText") is None
        ]
        # 段落全文置換では空欄用タブも不要（値が行全体を構成する）。残すと
        # タブが値の後ろに連続残留してレイアウトが崩れる。format.keep_tabs で温存可
        keep_tabs = (field.get("format") or {}).get("keep_tabs", False)
        if rs:
            set_run_text(rs[0], str(value), drop_tabs=not keep_tabs)
            for r in rs[1:]:
                p.elem.remove(r)
        else:
            sz = (field.get("format") or {}).get("sz")
            p.elem.append(make_run(str(value), sz=sz))
        self._warn_orphan_tabs(field["id"], p.elem)
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
            if self._check_literal(p.elem, literal):
                ok_any = True
            else:
                self.warn(field["id"], f"リテラル「{literal}」が段落内に見つからない")
        return ok_any

    # チェック記号 → 記入後の記号（疎明資料目録等の ◇=必須・○=該当時 にも対応）
    BOX_MARKS = {"□": "☑", "◇": "◆", "○": "●"}

    @classmethod
    def _check_literal(cls, p_elem, literal):
        """段落内の「□ラベル」を「☑ラベル」にする。空白のゆらぎと run 分割を許容する。

        段落の連結テキスト上で（空白を無視して）リテラルの位置を探し、その記号が
        段落内で何個目かを数えて、該当する w:t の記号だけを置換する。
        リテラル先頭の記号が □ なら ☑、◇ なら ◆、○ なら ● になる。
        """
        import re as _re
        box = literal[0] if literal else ""
        mark = cls.BOX_MARKS.get(box)
        if mark is None:  # 対応記号で始まらないリテラルは想定外
            return False
        ts = [t for t in p_elem.iter(f"{W}t")]
        full = "".join(t.text or "" for t in ts)
        pattern = "[\\s　]*".join(_re.escape(ch) for ch in literal if not ch.isspace() and ch != "　")
        m = _re.search(pattern, full)
        if not m:
            return False
        nth = full[: m.start()].count(box)  # 手前にある同記号の数 = 対象記号の序数
        seen = 0
        for t in ts:
            if not t.text:
                continue
            boxes = t.text.count(box)
            if seen + boxes > nth:
                i = -1
                for _ in range(nth - seen + 1):
                    i = t.text.index(box, i + 1)
                t.text = t.text[:i] + mark + t.text[i + 1:]
                return True
            seen += boxes
        return False

    def fill_underline_fill(self, field, value):
        p = self.resolve(field["id"], field["anchor"])
        if p is None:
            return False
        urs = underlined_runs(p.elem)
        fmt = field.get("format") or {}
        slot = fmt.get("slot", 0)
        if slot >= len(urs):
            self.warn(field["id"], f"下線runが{len(urs)}個しかない (slot={slot})")
            return False
        # 下線 run のタブはタブ位置まで罫線を伸ばす役割があるため既定で温存。
        # 値が長くタブで溢れる書式だけ fillmap の format.drop_tabs: true で除去する
        set_run_text(urs[slot], str(value), drop_tabs=fmt.get("drop_tabs", False))
        self._warn_orphan_tabs(field["id"], p.elem)
        return True

    def fill_underline_slots(self, field, value):
        p = self.resolve(field["id"], field["anchor"])
        if p is None:
            return False
        urs = underlined_runs(p.elem)
        vals = value if isinstance(value, list) else [value]
        if len(vals) > len(urs):
            self.warn(field["id"], f"値{len(vals)}個に対し下線runが{len(urs)}個")
        drop = (field.get("format") or {}).get("drop_tabs", False)
        ok = False
        for r, v in zip(urs, vals):
            if v is None or v == "":
                continue
            set_run_text(r, str(v), drop_tabs=drop)
            ok = True
        if ok:
            self._warn_orphan_tabs(field["id"], p.elem)
        return ok

    def _warn_orphan_tabs(self, fid, p_elem):
        """記入後の段落で、最後のテキストの後に下線なしの連続タブが残っていたら警告。

        下線 run の末尾タブは罫線をタブ位置まで伸ばす正当な用法なので対象外。
        下線のない残留タブはタブ位置まで行を伸ばすだけでレイアウトを崩す
        （B1102 の実崩れの原因）。将来の同種事故の検知網。
        """
        underlined = set(underlined_runs(p_elem))
        seq = []  # 段落内の run 子要素を文書順に (種別, 下線か) で並べる
        for r in runs_of(p_elem):
            for e in r:
                if e.tag == f"{W}t" and (e.text or "").strip():
                    seq.append(("t", r in underlined))
                elif e.tag == f"{W}tab":
                    seq.append(("tab", r in underlined))
        last_t = max((i for i, (k, _) in enumerate(seq) if k == "t"), default=-1)
        plain_tabs = sum(1 for k, u in seq[last_t + 1:] if k == "tab" and not u)
        if plain_tabs >= 2:
            self.warn(fid, f"記入後の段落末尾に下線なしタブが{plain_tabs}個残っている"
                           "（レイアウト崩れの恐れ。fillmap の kind/keep_tabs を見直す）")

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
            if fn(self, field, self._clean(fid, vals[fid])):
                self.filled.append(fid)
        unknown = set(vals) - {f["id"] for f in fillmap["fields"]}
        for fid in sorted(unknown):
            self.warn(fid, "fillmap に存在しない field_id（値は無視された）")

    def _clean(self, fid, value):
        """文字列値の清浄化（リスト・dict は要素ごと）。実施したら警告で報告する。"""
        if isinstance(value, str):
            cleaned, notes = clean_fill_value(value)
            if notes:
                self.warn(fid, f"値を清浄化した: {'・'.join(notes)}")
            return cleaned
        if isinstance(value, list):
            return [self._clean(fid, v) for v in value]
        if isinstance(value, dict):
            return {k: self._clean(fid, v) for k, v in value.items()}
        return value


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
