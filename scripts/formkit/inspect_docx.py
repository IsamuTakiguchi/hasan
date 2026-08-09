#!/usr/bin/env python3
"""docx 書式のアンカー一覧をダンプする（fillmap 起草・版間照合の土台）。

裁判所配布の .docx から word/document.xml を読み、段落ごとに
  index / w14:paraId / 構造パス（見出し系列・表位置）/ 正規化テキスト / フラグ
を1行ずつ出力する。run 断片化の影響を受けないよう、テキストは段落内の
<w:t> を連結してから正規化する。

使い方:
  python3 scripts/formkit/inspect_docx.py <template.docx> [-o anchors.txt]

出力列（タブ区切り）:
  idx  paraId  path  flags  text
  flags: T=表内 E=空 C=チェックボックス(□/☑あり) U=下線runあり
         S=セクション区切り(pPr内にsectPr。行複製・削除の対象にしてはならない)
"""
import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}
W = NS["w"]
W14 = NS["w14"]

# 見出しとして構造パスに積む段落の判定（報告書系書式の実態に合わせた素朴な規則）
HEADING_RE = [
    re.compile(r"^第[０-９0-9一二三四五六七八九十]+"),          # 第１ 第２ …
    re.compile(r"^[０-９0-9]+[　 ．.]"),                        # １　職歴 など
]


def normalize(text: str) -> str:
    """照合用の正規化: 連続空白を1つの全角スペースに、前後を刈る。"""
    return re.sub(r"[\s　]+", "　", text).strip("　 ")


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))


def has_underline_run(p) -> bool:
    for rpr in p.iter(f"{{{W}}}rPr"):
        u = rpr.find(f"{{{W}}}u")
        if u is not None and u.get(f"{{{W}}}val", "single") != "none":
            return True
    return False


def iter_block_paragraphs(body):
    """body 直下と表内の段落を文書順に (p, table_path) で列挙する。

    table_path は表外なら ""、表内なら "table[i]>row[j]>cell[k]"。
    ネストした表は親のパスを引き継ぐ。
    """
    table_counter = [0]

    def walk(elem, prefix):
        for child in elem:
            tag = child.tag
            if tag == f"{{{W}}}p":
                yield child, prefix
            elif tag == f"{{{W}}}tbl":
                ti = table_counter[0]
                table_counter[0] += 1
                for ri, row in enumerate(child.findall(f"{{{W}}}tr")):
                    for ci, cell in enumerate(row.findall(f"{{{W}}}tc")):
                        base = f"table[{ti}]>row[{ri}]>cell[{ci}]"
                        path = f"{prefix}>{base}" if prefix else base
                        yield from walk(cell, path)

    yield from walk(body, "")


def heading_trail_update(trail, text):
    """見出し系列を更新する。第○は深さ0、番号見出しは深さ1。"""
    if HEADING_RE[0].match(text):
        return [text[:20]]
    if HEADING_RE[1].match(text):
        return (trail[:1] if trail else []) + [text[:20]]
    return trail


def inspect(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(f"{{{W}}}body")
    rows = []
    trail = []
    for idx, (p, tpath) in enumerate(iter_block_paragraphs(body)):
        raw = para_text(p)
        text = normalize(raw)
        if not tpath and text:
            trail = heading_trail_update(trail, text)
        flags = ""
        flags += "T" if tpath else ""
        flags += "E" if not text else ""
        flags += "C" if ("□" in text or "☑" in text) else ""
        flags += "U" if has_underline_run(p) else ""
        ppr = p.find(f"{{{W}}}pPr")
        flags += "S" if (ppr is not None and ppr.find(f"{{{W}}}sectPr") is not None) else ""
        heading = ">".join(trail)
        path = f"{heading}>{tpath}" if (heading and tpath) else (tpath or heading)
        rows.append({
            "idx": idx,
            "paraId": p.get(f"{{{W14}}}paraId", ""),
            "path": path,
            "flags": flags,
            "text": text,
        })
    return rows


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("docx")
    ap.add_argument("-o", "--output", help="出力先（省略時は標準出力）")
    ap.add_argument("--full", action="store_true", help="テキストを切り詰めない")
    args = ap.parse_args()

    rows = inspect(args.docx)
    out = open(args.output, "w", encoding="utf-8") if args.output else sys.stdout
    print("idx\tparaId\tpath\tflags\ttext", file=out)
    for r in rows:
        text = r["text"] if args.full else r["text"][:60]
        print(f"{r['idx']}\t{r['paraId']}\t{r['path']}\t{r['flags']}\t{text}", file=out)
    if args.output:
        out.close()
        print(f"{len(rows)} paragraphs -> {args.output}")


if __name__ == "__main__":
    main()
