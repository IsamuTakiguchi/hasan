---
name: hasan-saikensha-ichiran
description: >-
  破産申立ての債権者一覧表（Excel）を事件モデル case.yaml から生成するスキル。
  「債権者一覧表を作って」「債権者リストを作成して」「一覧表に転記して」
  などのフレーズがトリガー。破産申立ての債権者一覧表の作成では必ず使用すること。
  xlsxスキルと併用すること。
---

# 債権者一覧表 生成スキル

## 手順

1. `python3 scripts/casekit/validate_case.py cases/<事件ID>/case.yaml` を通す。
   債権者の残高不明・出所なしの警告が出たら、生成は続行しつつ最後に報告する。
2. 生成:
   ```
   python3 scripts/build/build_saikensha_ichiran.py \
     --case cases/<事件ID>/case.yaml \
     --output cases/<事件ID>/output/債権者一覧表_<姓>.xlsx
   ```
   - **裁判所書式が未登録なら exit 3 で案内が出る**。利用者に裁判所配布の Excel 書式の
     提供を依頼し（hasan-register-form で登録）、それまでの間は利用者の了解を得て
     `--generic`（事務所内ドラフト様式）で出力する。
3. `python3 scripts/verify/crosscheck_outputs.py --case ... --dir cases/<事件ID>/output`
   で負債総額・債権者数の突合が OK になることを確認する。
4. `scripts/verify/render_preview.py` でPDF・画像化し目視（列のはみ出し・文字切れ）。

## 記載ルール

- 並び順はビルダーが処理する: 一般債権者を case.yaml の記載順で先に、**公租公課は末尾**。
  番号は自動で振り直される（case.yaml の no は取込順の参考値）。
- 残高不明の債権者は金額空欄のまま載せる（合計に含まれない旨が注記される）。
  勝手に推計しない。
- 保証人有無が null（未確認）の債権者は空欄になる → questions.md で確認を促す。
- 裁判所書式版で列構成が異なる場合は fillmap.yaml の `columns` 対応を確認する
  （ビルダーの値キー: no/name/address/kind/origin_date/use/principal/balance/
  last_payment/guarantor/note）。
