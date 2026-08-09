"""xlsx 書式の記入・生成の共通ライブラリ。

2つのモードを支える:
1. テンプレートモード — 裁判所配布 xlsx（courts/ に登録済み）を fillmap 駆動で記入。
   数式・結合セル・印刷設定を保持するため、既存ブックを開いて値だけを書く。
2. 汎用ドラフトモード — 書式未登録のあいだの事務所内ドラフトを openpyxl で新規生成。
   builder スクリプトが章立てを組む。

fillmap（xlsx用）の kind:
  set_cell    anchor: {sheet, cell} に値を書く。label_check: {アドレス: 期待文言} で
              近傍ラベルを検証（不一致は警告＝書式改訂の兆候）
  table_rows  anchor: {sheet}, first_data_row, columns: {列レター: キー}, 値は辞書のリスト。
              max_rows を超えたら行挿入（直前行の書式をコピー）
"""
from pathlib import Path

import yaml
from copy import copy as style_copy

import openpyxl

REPO = Path(__file__).resolve().parent.parent.parent


def resolve_form(court: str, form: str):
    """registry.yaml から現行版のディレクトリを返す。未登録なら None。"""
    reg_path = REPO / "courts" / court / "forms" / form / "registry.yaml"
    if not reg_path.exists():
        return None
    reg = yaml.safe_load(reg_path.read_text(encoding="utf-8"))
    cur = reg.get("current")
    if not cur:
        return None
    vdir = reg_path.parent / cur
    if not (vdir / "fillmap.yaml").exists():
        return None
    return vdir


def registration_guidance(court: str, form: str, name: str) -> str:
    return (
        f"裁判所書式「{name}」が未登録です（courts/{court}/forms/{form}/）。\n"
        f"裁判所配布の Excel ファイルを入手したら hasan-register-form スキルで登録してください\n"
        f"（登録するまでは --generic で事務所内ドラフト様式を生成できます）。"
    )


class XlsxFiller:
    def __init__(self, template_path):
        self.wb = openpyxl.load_workbook(template_path)
        self.warnings = []
        self.filled = 0

    def _ws(self, name):
        if name in self.wb.sheetnames:
            return self.wb[name]
        self.warnings.append(f"シート「{name}」が無い（先頭シートで代替）")
        return self.wb.worksheets[0]

    def set_cell(self, field, value):
        ws = self._ws(field["anchor"]["sheet"])
        for addr, expect in (field.get("label_check") or {}).items():
            actual = ws[addr].value
            if actual != expect:
                self.warnings.append(
                    f"{field['id']}: ラベル検証不一致 {addr}=「{actual}」期待「{expect}」（書式改訂の疑い）")
        ws[field["anchor"]["cell"]] = value
        self.filled += 1

    def table_rows(self, field, rows):
        ws = self._ws(field["anchor"]["sheet"])
        first = field["first_data_row"]
        cols = field["columns"]
        max_rows = field.get("max_rows")
        if max_rows and len(rows) > max_rows:
            # 最終データ行の直後に必要数を挿入し、書式をコピー
            insert_at = first + max_rows
            n = len(rows) - max_rows
            ws.insert_rows(insert_at, n)
            for i in range(n):
                for col in cols:
                    src = ws[f"{col}{insert_at - 1}"]
                    dst = ws[f"{col}{insert_at + i}"]
                    dst.font = style_copy(src.font)
                    dst.border = style_copy(src.border)
                    dst.alignment = style_copy(src.alignment)
                    dst.number_format = src.number_format
        for i, row in enumerate(rows):
            for col, key in cols.items():
                v = row.get(key)
                if v is not None:
                    ws[f"{col}{first + i}"] = v
        self.filled += len(rows)

    def apply(self, fillmap, values):
        vals = values.get("fields", {})
        for field in fillmap["fields"]:
            v = vals.get(field["id"])
            if v is None or v == "" or v == []:
                continue
            if field["kind"] == "set_cell":
                self.set_cell(field, v)
            elif field["kind"] == "table_rows":
                self.table_rows(field, v)
            else:
                self.warnings.append(f"{field['id']}: 未知の kind {field['kind']}")

    def save(self, out):
        Path(out).parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(out)


# ---- 汎用ドラフトモードの体裁ヘルパー ----------------------------------
THIN = openpyxl.styles.Side(style="thin")
BORDER = openpyxl.styles.Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = openpyxl.styles.PatternFill("solid", fgColor="EEEEEE")
FONT = openpyxl.styles.Font(name="游ゴシック", size=10)
FONT_TITLE = openpyxl.styles.Font(name="游ゴシック", size=14, bold=True)


def style_header(ws, row, ncols, start_col=1):
    for i in range(ncols):
        c = ws.cell(row=row, column=start_col + i)
        c.border = BORDER
        c.fill = HEAD_FILL
        c.font = openpyxl.styles.Font(name="游ゴシック", size=10, bold=True)
        c.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_data(ws, row, ncols, start_col=1):
    for i in range(ncols):
        c = ws.cell(row=row, column=start_col + i)
        c.border = BORDER
        c.font = FONT
        c.alignment = openpyxl.styles.Alignment(vertical="center", wrap_text=True)


def draft_note(ws, row, text="※ 事務所内ドラフト様式（裁判所書式は未登録）。提出前に裁判所書式へ転記・登録すること。"):
    c = ws.cell(row=row, column=1, value=text)
    c.font = openpyxl.styles.Font(name="游ゴシック", size=9, italic=True, color="AA0000")
