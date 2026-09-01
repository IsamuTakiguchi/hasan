"""事件フォルダの変更検知。前回処理時からの差分（新規・変更・削除）を報告する。

update スキル（事件更新オーケストレータ）の入口。状態は作業フォルダ
（作業ファイル/ または work/）の state.json にファイルごとの SHA1 で記録し、
--write で現在の状態を保存する。

使い方:
  python3 scan_changes.py --dir <事件フォルダ> [--write]

対象:
  - 原資料: 受領資料/ または input/（再帰）
  - 事件モデル: case.yaml・questions.md
  - 聴取シート・面談メモ等: 作業ファイル/*.docx または work/*.docx
  - 申立書類: 申立書類/ または output/

申立書類の docx/xlsx は --write 時に内容スナップショット（段落・セル値）を
<作業フォルダ>/snapshots/ に保存し、次回スキャンで変更ファイルの
**どの記載がどう変わったか**（段落・セル単位の diff）まで表示する。
"""
import argparse
import difflib
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path

INPUT_DIRS = ["受領資料", "input"]
WORK_DIRS = ["作業ファイル", "work"]
OUTPUT_DIRS = ["申立書類", "output"]
SKIP_NAMES = {".DS_Store", "Thumbs.db", "state.json"}


def work_dir(base: Path) -> Path:
    """作業フォルダ（中間物置き場）。既存のものを優先し、無ければ
    フォルダの流儀（受領資料=日本語構成）に合わせて選ぶ。"""
    for name in WORK_DIRS:
        if (base / name).is_dir():
            return base / name
    return base / ("作業ファイル" if (base / "受領資料").is_dir() else "work")


SNAP_DIRNAME = "snapshots"
MAX_DIFF_LINES = 40  # 1ファイルあたりの内容diff表示上限


def docx_paragraphs(path: Path):
    """docx の本文を段落テキストのリストで返す。壊れたファイルは None。"""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("word/document.xml").decode("utf-8")
    except Exception:
        return None
    paras = []
    for chunk in re.split(r"</w:p>", xml):
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", chunk))
        if text.strip():
            paras.append(text)
    return paras


def xlsx_cells(path: Path):
    """xlsx を {シート名: {セル番地: 値}} で返す（数式は数式文字列）。壊れたら None。"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    except Exception:
        return None
    sheets = {}
    try:
        for ws in wb.worksheets:
            cells = {}
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        cells[c.coordinate] = str(c.value)
            sheets[ws.title] = cells
    finally:
        wb.close()
    return sheets


def snapshot_of(path: Path):
    if path.suffix == ".docx":
        paras = docx_paragraphs(path)
        return None if paras is None else {"type": "docx", "paras": paras}
    if path.suffix == ".xlsx":
        sheets = xlsx_cells(path)
        return None if sheets is None else {"type": "xlsx", "sheets": sheets}
    return None


def snap_path(wd: Path, rel_in_output: str) -> Path:
    return wd / SNAP_DIRNAME / (rel_in_output.replace("/", "__") + ".json")


def print_content_diff(out_dir: Path, wd: Path, rel_in_output: str):
    """変更された申立書類の内容diff（段落・セル単位）を表示する。"""
    sp = snap_path(wd, rel_in_output)
    cur_path = out_dir / rel_in_output
    if cur_path.suffix not in (".docx", ".xlsx"):
        return
    if not sp.is_file():
        print("        （内容diffなし — 前回のスナップショット未保存）")
        return
    try:
        old = json.loads(sp.read_text(encoding="utf-8"))
    except Exception:
        return
    new = snapshot_of(cur_path)
    if new is None or new.get("type") != old.get("type"):
        print("        （内容を読み取れないため diff 省略 — ファイルが壊れている可能性）")
        return

    lines = []
    if new["type"] == "docx":
        for d in difflib.unified_diff(old.get("paras", []), new["paras"],
                                      lineterm="", n=0):
            if d.startswith("+++") or d.startswith("---") or d.startswith("@@"):
                continue
            mark = "＋" if d.startswith("+") else "－"
            lines.append(f"        {mark} {d[1:].strip()}")
    else:
        old_sheets, new_sheets = old.get("sheets", {}), new["sheets"]
        for sheet in sorted(set(old_sheets) | set(new_sheets)):
            o, n = old_sheets.get(sheet, {}), new_sheets.get(sheet, {})
            for addr in sorted(set(o) | set(n), key=lambda a: (len(a), a)):
                ov, nv = o.get(addr), n.get(addr)
                if ov != nv:
                    lines.append(f"        {sheet}!{addr}: "
                                 f"{ov if ov is not None else '（空欄）'} → "
                                 f"{nv if nv is not None else '（空欄）'}")
    if not lines:
        print("        （本文・セル値の変更なし — 書式設定のみの変更）")
        return
    for ln in lines[:MAX_DIFF_LINES]:
        print(ln)
    if len(lines) > MAX_DIFF_LINES:
        print(f"        …ほか {len(lines) - MAX_DIFF_LINES} 件の変更")


def write_snapshots(out_dir: Path, wd: Path, output_files: dict, base: Path):
    """申立書類の docx/xlsx の内容スナップショットを保存し、消えた分を掃除する。"""
    snap_dir = wd / SNAP_DIRNAME
    keep = set()
    for rel_from_base in output_files:
        p = base / rel_from_base
        if p.suffix not in (".docx", ".xlsx"):
            continue
        rel_in_output = str(p.relative_to(out_dir))
        snap = snapshot_of(p)
        if snap is None:
            continue
        sp = snap_path(wd, rel_in_output)
        sp.parent.mkdir(parents=True, exist_ok=True)
        sp.write_text(json.dumps(snap, ensure_ascii=False), encoding="utf-8")
        keep.add(sp.name)
    if snap_dir.is_dir():
        for f in snap_dir.glob("*.json"):
            if f.name not in keep:
                f.unlink()


def sha1(path: Path) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def collect(base: Path) -> dict:
    """領域名 → {相対パス: sha1}"""
    areas = {}

    def add(area, root, recursive=True, pattern="*"):
        d = {}
        if root.is_dir():
            it = root.rglob(pattern) if recursive else root.glob(pattern)
            for p in sorted(it):
                if p.is_file() and p.name not in SKIP_NAMES and not p.name.startswith("~$"):
                    d[str(p.relative_to(base))] = sha1(p)
        areas[area] = d

    for name in INPUT_DIRS:
        if (base / name).is_dir():
            add("受領資料（原資料）", base / name)
            break
    else:
        areas["受領資料（原資料）"] = {}

    model = {}
    for name in ("case.yaml", "questions.md"):
        p = base / name
        if p.is_file():
            model[name] = sha1(p)
    areas["事件モデル"] = model

    wd = work_dir(base)
    add(f"聴取シート・面談メモ等（{wd.name}）", wd, recursive=False, pattern="*.docx")

    for name in OUTPUT_DIRS:
        if (base / name).is_dir():
            add("申立書類", base / name)
            break
    else:
        areas["申立書類"] = {}
    return areas


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dir", default=".", help="事件フォルダ（既定: カレント）")
    ap.add_argument("--write", action="store_true",
                    help="現在の状態を作業フォルダの state.json に保存")
    args = ap.parse_args()

    base = Path(args.dir).resolve()
    wd = work_dir(base)
    state_path = wd / "state.json"
    out_dir = None
    for name in OUTPUT_DIRS:
        if (base / name).is_dir():
            out_dir = base / name
            break
    current = collect(base)

    previous = None
    if state_path.exists():
        try:
            previous = json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            previous = None

    any_change = False
    if previous is None:
        print("前回の状態記録なし（初回スキャン — 全ファイルを未処理として扱う）")
        for area, files in current.items():
            if files:
                print(f"[{area}]")
                for f in files:
                    print(f"  新規  {f}")
                any_change = True
    else:
        for area, files in current.items():
            prev = previous.get(area, {})
            added = sorted(set(files) - set(prev))
            removed = sorted(set(prev) - set(files))
            changed = sorted(f for f in set(files) & set(prev) if files[f] != prev[f])
            if added or removed or changed:
                any_change = True
                print(f"[{area}]")
                for f in added:
                    print(f"  新規  {f}")
                for f in changed:
                    print(f"  変更  {f}")
                    if area == "申立書類" and out_dir is not None:
                        print_content_diff(out_dir, wd,
                                           str((base / f).relative_to(out_dir)))
                for f in removed:
                    print(f"  削除  {f}")
        if not any_change:
            print("前回処理時から変更はありません")

    if args.write:
        state_path.parent.mkdir(parents=True, exist_ok=True)
        state_path.write_text(json.dumps(current, ensure_ascii=False, indent=1),
                              encoding="utf-8")
        if out_dir is not None:
            write_snapshots(out_dir, wd, current.get("申立書類", {}), base)
        print(f"状態を保存: {state_path.relative_to(base)}")

    sys.exit(0 if not any_change else 2)  # 2=差分あり（エラーではない）


if __name__ == "__main__":
    main()
