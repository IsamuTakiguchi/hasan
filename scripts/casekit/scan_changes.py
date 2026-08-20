"""事件フォルダの変更検知。前回処理時からの差分（新規・変更・削除）を報告する。

update スキル（事件更新オーケストレータ）の入口。状態は work/state.json に
ファイルごとの SHA1 で記録し、--write で現在の状態を保存する。

使い方:
  python3 scan_changes.py --dir <事件フォルダ> [--write]

対象:
  - 原資料: 受領資料/ または input/（再帰）
  - 事件モデル: case.yaml・questions.md
  - 聴取シート・面談メモ等: work/*.docx
  - 申立書類: 申立書類/ または output/
"""
import argparse
import hashlib
import json
import sys
from pathlib import Path

STATE_REL = "work/state.json"
INPUT_DIRS = ["受領資料", "input"]
OUTPUT_DIRS = ["申立書類", "output"]
SKIP_NAMES = {".DS_Store", "Thumbs.db", "state.json"}


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

    add("聴取シート・面談メモ等（work）", base / "work", recursive=False, pattern="*.docx")

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
    ap.add_argument("--write", action="store_true", help="現在の状態を work/state.json に保存")
    args = ap.parse_args()

    base = Path(args.dir).resolve()
    state_path = base / STATE_REL
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
                for f in removed:
                    print(f"  削除  {f}")
        if not any_change:
            print("前回処理時から変更はありません")

    if args.write:
        state_path.parent.mkdir(parents=True, exist_ok=True)
        state_path.write_text(json.dumps(current, ensure_ascii=False, indent=1),
                              encoding="utf-8")
        print(f"状態を保存: {state_path.relative_to(base)}")

    sys.exit(0 if not any_change else 2)  # 2=差分あり（エラーではない）


if __name__ == "__main__":
    main()
