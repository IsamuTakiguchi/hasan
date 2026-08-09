"""fillmap の多層アンカー（paraId → 文言 → 構造パス）を docx の段落に解決する共通ライブラリ。

fill_docx.py（記入）と diff_fillmap.py（版間照合）の両方が使う。
解決結果は3段階:
  OK      — paraId が一致
  REBOUND — paraId 不一致だが context（＋path）で一意に解決できた
  BROKEN  — 一意に解決できない
"""
import re
from dataclasses import dataclass

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W = "{%s}" % W_NS
W14 = "{%s}" % W14_NS

HEADING_RE = [
    re.compile(r"^第[０-９0-9一二三四五六七八九十]+"),
    re.compile(r"^[０-９0-9]+[　 ．.]"),
]


def normalize(text: str) -> str:
    return re.sub(r"[\s　]+", "　", text).strip("　 ")


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


@dataclass
class Para:
    elem: object
    idx: int
    para_id: str
    path: str      # 見出し系列 + 表位置
    text: str      # 正規化済み


def index_paragraphs(root):
    """document.xml のルートから全段落を文書順に Para のリストで返す。"""
    body = root.find(f"{W}body")
    paras = []
    trail = []
    table_counter = [0]

    def walk(elem, prefix):
        for child in elem:
            if child.tag == f"{W}p":
                yield child, prefix
            elif child.tag == f"{W}tbl":
                ti = table_counter[0]
                table_counter[0] += 1
                for ri, row in enumerate(child.findall(f"{W}tr")):
                    for ci, cell in enumerate(row.findall(f"{W}tc")):
                        base = f"table[{ti}]>row[{ri}]>cell[{ci}]"
                        yield from walk(cell, f"{prefix}>{base}" if prefix else base)

    for idx, (p, tpath) in enumerate(walk(body, "")):
        text = normalize(para_text(p))
        if not tpath and text:
            if HEADING_RE[0].match(text):
                trail = [text[:20]]
            elif HEADING_RE[1].match(text):
                trail = (trail[:1] if trail else []) + [text[:20]]
        heading = ">".join(trail)
        path = f"{heading}>{tpath}" if (heading and tpath) else (tpath or heading)
        paras.append(Para(p, idx, p.get(f"{W14}paraId", ""), path, text))
    return paras


class ParaIndex:
    def __init__(self, root):
        self.paras = index_paragraphs(root)
        self.by_id = {p.para_id: p for p in self.paras if p.para_id}

    def resolve(self, anchor: dict):
        """anchor 辞書 {paraId, context, path, idx} を (Para|None, status, note) で返す。

        解決の優先順:
          1. paraId 一致（文言も矛盾しない）→ OK
          2. 文言（完全一致を部分一致より優先）＋構造パスで一意 → REBOUND
          3. 空欄アンカーはパス一致の空欄段落のうち文書順で最初のもの → REBOUND
          4. なお複数残るときは idx（旧版での文書内位置）に最も近いもの → REBOUND
        """
        pid = anchor.get("paraId", "")
        ctx = normalize(anchor.get("context", "") or "")
        path = anchor.get("path", "") or ""
        old_idx = anchor.get("idx")

        def pick_nearest(cands, how):
            if not cands:
                return None
            if len(cands) == 1:
                return cands[0], how
            if old_idx is not None:
                best = min(cands, key=lambda q: abs(q.idx - old_idx))
                return best, f"{how},idx近接"
            return None

        p = self.by_id.get(pid)
        if p is not None and (not ctx or ctx in p.text or p.text == ctx):
            return p, "OK", ""
        # paraId を持たない書式（旧型式docx）では文言＋構造解決が正規経路 → OK 扱い
        rebound = "REBOUND" if pid else "OK"

        # 第2層: context で候補を集める（完全一致を優先）
        subs = [q for q in self.paras if ctx and ctx in q.text]
        exact = [q for q in subs if q.text == ctx]
        for cands, label in ((exact, "context完全一致"), (subs, "context")):
            if not cands:
                continue
            if path:
                hit = pick_nearest([q for q in cands if q.path == path], f"{label}+path")
                if hit:
                    return hit[0], rebound, f"paraId {pid} -> {hit[0].para_id} ({hit[1]})"
                # path 完全一致がなければ表位置だけ（見出し文言の微修正に強い）
                tail = path.split(">")[-2:]
                hit = pick_nearest([q for q in cands if q.path.split(">")[-2:] == tail], f"{label}+path末尾")
                if hit:
                    return hit[0], rebound, f"paraId {pid} -> {hit[0].para_id} ({hit[1]})"
            hit = pick_nearest(cands, label)
            if hit:
                return hit[0], rebound, f"paraId {pid} -> {hit[0].para_id} ({hit[1]})"

        # 空欄アンカー（空セル・空欄行）: パス一致の空欄段落のうち最初のもの
        if not ctx and path:
            blanks = [q for q in self.paras if q.path == path and q.text == ""]
            if blanks:
                q = blanks[0]
                note = "path,空欄先頭" if len(blanks) > 1 else "path,空欄"
                return q, rebound, f"paraId {pid} -> {q.para_id} ({note})"

        if p is not None:
            # paraId 一致を最後の拠り所として使う（文言不一致の警告付き）
            return p, "REBOUND", f"paraId一致だが文言不一致: 期待「{ctx}」実際「{p.text[:30]}」"
        return None, "BROKEN", f"解決不能: paraId={pid} context「{ctx[:30]}」候補{len(subs)}件"


def load_document(path_docx_xml):
    """word/document.xml のパスから (tree, root) を返す。"""
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(path_docx_xml), parser)
    return tree, tree.getroot()
