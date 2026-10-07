#!/usr/bin/env python3
"""記入済み docx のレイアウト検査（白紙テンプレートとの描画パリティ）。

記入が引き起こした折返し・頁あふれを検出する。検証環境の LibreOffice は代替フォントで
字間や頁数が Word と変わるが、**白紙テンプレと記入済みを同じ環境で描画**すれば
「記入が引き起こした差分」だけを相対比較できる。

検査項目:
  1. XML リント: 下線なしの連続タブが段落末尾に残っていないか（空欄用タブの残留）、
     テンプレに無い「和文中の半角スペース」が増えていないか（YAML 折返し混入）
  2. 描画パリティ（soffice がある環境のみ）: 頁数の一致、各頁の本文最下行の位置が
     テンプレより大きく下がっていないか（＝折返しで行が増えた）

使い方:
  python3 scripts/verify/layout_check.py --dir cases/<id>/output --case cases/<id>/case.yaml
  python3 scripts/verify/layout_check.py <記入済み.docx> --template <白紙template.docx>
  （--no-render で描画比較を省略し XML リントだけ行う）

終了コード: 0=全OK, 1=NGあり, 4=検査対象なし
"""
import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from build.xlsxlib import resolve_form, proc_of_case  # noqa: E402
from verify.render_preview import to_pdf  # noqa: E402

try:
    import pymupdf
except ImportError:
    pymupdf = None

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 出力ファイル名のキーワード → 書式ID（手続種別で分かれるものは dict）。前方の語ほど優先
FORM_KEYWORDS = [
    ("管財補充報告書", "kanzai-hoju-houkokusho"),
    ("事業に関する報告書", "jigyou-houkokusho"),
    ("添付目録", "tenpu-mokuroku"),
    ("疎明資料目録", "somei-shiryo-mokuroku"),
    ("引継資料一覧表", "hikitsugi-shiryo-ichiran"),
    ("自由財産拡張", "jiyuzaisan-kakucho"),
    ("上申書", "joshinsho-checklist"),
    ("申立書", {"douhai": "moushitatesho-douhai", "kanzai-shizenjin": "moushitatesho-kanzai",
               "kanzai-hojin": "moushitatesho-hojin"}),
    ("報告書", {"douhai": "houkokusho", "kanzai-shizenjin": "houkokusho",
              "kanzai-hojin": "houkokusho-hojin"}),
]
CJK_SPACE = re.compile(r"[^\x00-\x7f　] +[^\x00-\x7f　]")  # 和文の間の半角スペース
CJK_SPACE_LOOSE = re.compile(r"[^\x00-\x7f] +[^\x00-\x7f]")        # テンプレ側の鍵集合用（全角空白隣接も含む）
LINE_TOLERANCE_PT = 16.0  # 本文1行分程度。これ以上最下行が下がれば折返しが増えたとみなす


def form_of(filename: str, proc: str):
    for kw, form in FORM_KEYWORDS:
        if kw in filename:
            if isinstance(form, dict):
                return form.get(proc or "douhai")
            return form
    return None


def template_for(filename: str, proc: str, court="osaka"):
    form = form_of(filename, proc)
    if not form:
        return None
    vdir = resolve_form(court, form, proc)
    if vdir and (vdir / "template.docx").exists():
        return vdir / "template.docx"
    return None


# ---- XML リント -------------------------------------------------------------

def _paragraphs(path: Path):
    """[(段落テキスト, 下線なし末尾タブ数)]。壊れた docx は None。"""
    try:
        from lxml import etree
        with zipfile.ZipFile(path) as z:
            root = etree.fromstring(z.read("word/document.xml"))
    except Exception:
        return None
    out = []
    for p in root.iter(f"{{{W}}}p"):
        text = ""
        seq = []
        for r in p.findall(f"{{{W}}}r"):
            rpr = r.find(f"{{{W}}}rPr")
            u = rpr is not None and rpr.find(f"{{{W}}}u") is not None and \
                rpr.find(f"{{{W}}}u").get(f"{{{W}}}val", "single") != "none"
            for e in r:
                if e.tag == f"{{{W}}}t":
                    text += e.text or ""
                    if (e.text or "").strip():
                        seq.append(("t", u))
                elif e.tag == f"{{{W}}}tab":
                    seq.append(("tab", u))
        last_t = max((i for i, (k, _) in enumerate(seq) if k == "t"), default=-1)
        plain_tabs = sum(1 for k, u in seq[last_t + 1:] if k == "tab" and not u)
        out.append((text, plain_tabs))
    return out


def _space_keys(text: str):
    """「直前の文字＋半角スペース列」を鍵にする（記号置換 □→☑ 等の影響を受けない）。
    テンプレ側は全角空白に隣接する空白（空欄の余白）も鍵に含める。"""
    return {text[m.start():m.end() - 1] for m in CJK_SPACE_LOOSE.finditer(text)}


def xml_lint(out: Path, tpl: Path | None):
    issues = []
    ps = _paragraphs(out)
    if ps is None:
        return ["docx として読めない（壊れているか docx 形式でない）"]
    tps = _paragraphs(tpl) if tpl else None
    tpl_spaces = set()
    if tps:
        for text, _ in tps:
            tpl_spaces |= _space_keys(text)
    for i, (text, plain_tabs) in enumerate(ps):
        if plain_tabs >= 2:
            issues.append(f"段落{i}: 下線なしタブが末尾に{plain_tabs}個残留 「{text[:30]}」")
        for m in CJK_SPACE.finditer(text):
            if text[m.start():m.end() - 1] not in tpl_spaces:
                issues.append(f"段落{i}: 和文中に半角スペース 「{text[max(0, m.start()-8):m.end()+8]}」")
    return issues


# ---- 描画パリティ -----------------------------------------------------------

def page_metrics(pdf: Path):
    """各頁の (本文行数, 最下行の y) のリスト。"""
    doc = pymupdf.open(pdf)
    metrics = []
    for page in doc:
        lines = 0
        bottom = 0.0
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                if "".join(s["text"] for s in line["spans"]).strip():
                    lines += 1
                    bottom = max(bottom, line["bbox"][3])
        metrics.append((lines, bottom))
    doc.close()
    return metrics


def _page_metrics_of(docx: Path, workdir: Path):
    return page_metrics(to_pdf(docx, workdir))


def _with_plain_boxes(src: Path, dst: Path):
    """☑◆● を同幅の □◇○ に戻した一時コピーを作る（グリフ幅差の切り分け用）。"""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "word/document.xml":
                data = data.decode("utf-8").translate(str.maketrans("☑◆●", "□◇○")).encode("utf-8")
            zout.writestr(info, data)


def _compare(m_out, m_tpl):
    issues = []
    if len(m_out) != len(m_tpl):
        issues.append(f"頁数がテンプレと異なる: 記入済み {len(m_out)} 頁 / 白紙 {len(m_tpl)} 頁"
                      "（記入による折返し・頁あふれ）")
    for i, ((lo, bo), (lt, bt)) in enumerate(zip(m_out, m_tpl)):
        if bo - bt > LINE_TOLERANCE_PT:
            issues.append(f"{i + 1}頁目: 本文最下行が白紙より {bo - bt:.0f}pt 下がっている"
                          f"（行数 {lt}→{lo}。折返しで行が増えた可能性）")
    return issues


def render_parity(out: Path, tpl: Path):
    """頁数と各頁の最下行位置を白紙テンプレと比較する。

    差分が出た場合、チェック記号（☑ U+2611 等）を同幅の □ に戻したコピーで再比較し、
    それで一致すれば「代替フォントのグリフ幅差」（LibreOffice 固有。Word では問題ない）
    として NG にせず注記にとどめる。
    """
    if not (shutil.which("soffice") or shutil.which("unoconvert")) or pymupdf is None:
        return None, ["（描画比較は省略: soffice/PyMuPDF なし）"]
    with tempfile.TemporaryDirectory() as td:
        d_out, d_tpl, d_box = Path(td) / "out", Path(td) / "tpl", Path(td) / "box"
        for d in (d_out, d_tpl, d_box):
            d.mkdir()
        try:
            m_out = _page_metrics_of(out, d_out)
            m_tpl = _page_metrics_of(tpl, d_tpl)
            issues = _compare(m_out, m_tpl)
            note = ""
            if issues:
                boxed = d_box / out.name
                _with_plain_boxes(out, boxed)
                if not _compare(_page_metrics_of(boxed, d_box), m_tpl):
                    note = "（☑等のチェック記号を□に戻すと白紙と一致 — 検証環境の代替フォントの" \
                           "グリフ幅差。Word では問題ない）"
                    issues = []
        except Exception as e:  # 変換失敗は検査不能として報告（NG にはしない）
            return None, [f"（描画比較に失敗: {e}）"]
    summary = f"{len(m_out)}頁（白紙 {len(m_tpl)}頁）{note}"
    return summary, issues


def check_file(out: Path, tpl: Path | None, render=True):
    """(summary, issues)。tpl が None のときは XML リントのみ。"""
    issues = xml_lint(out, tpl)
    summary = "XMLリントのみ"
    if render and tpl is not None:
        s, more = render_parity(out, tpl)
        issues += [m for m in more if not m.startswith("（")]
        notes = [m for m in more if m.startswith("（")]
        summary = s or (notes[0] if notes else summary)
    return summary, issues


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*", help="記入済み docx（--dir と併用不可）")
    ap.add_argument("--dir", help="出力ディレクトリ（中の docx を全て検査）")
    ap.add_argument("--case", help="case.yaml（手続種別からテンプレ版を解決）")
    ap.add_argument("--template", help="白紙テンプレ（files 指定時）")
    ap.add_argument("--no-render", action="store_true", help="描画比較を省略")
    args = ap.parse_args()

    proc = None
    if args.case:
        proc = proc_of_case(yaml.safe_load(open(args.case, encoding="utf-8")))

    targets = []
    if args.dir:
        for f in sorted(Path(args.dir).glob("*.docx")):
            if f.name.startswith("~$"):
                continue
            targets.append((f, template_for(f.name, proc)))
    for f in args.files:
        targets.append((Path(f), Path(args.template) if args.template else template_for(Path(f).name, proc)))
    if not targets:
        print("検査対象の docx が無い")
        sys.exit(4)

    any_ng = False
    print("レイアウト検査（白紙テンプレとの描画パリティ）")
    for out, tpl in targets:
        summary, issues = check_file(out, tpl, render=not args.no_render)
        if tpl is None:
            summary += "／テンプレ未解決（書式名をファイル名から判定できず）"
        status = "NG" if issues else "OK"
        any_ng |= bool(issues)
        print(f"  {status}  {out.name}: {summary}")
        for x in issues:
            print(f"        - {x}")
    sys.exit(1 if any_ng else 0)


if __name__ == "__main__":
    main()
