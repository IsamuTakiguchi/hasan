---
name: saikensha-ichiran
description: >-
  破産申立ての債権者一覧表（Excel）を事件モデル case.yaml から生成するスキル。
  「債権者一覧表を作って」「債権者リストを作成して」「一覧表に転記して」
  などのフレーズがトリガー。破産申立ての債権者一覧表の作成では必ず使用すること。
  xlsxスキルと併用すること。
---

# 債権者一覧表 生成スキル

## 出力形式の鉄則（必ず守る）

- この書式の出力は**必ずテンプレートと同じ形式**にする（Excel書式→.xlsx／Word書式→.docx）。
- 生成は hasan-kit（builder / fill-docx）経由のみ。**Word や Excel の文書を新規作成して
  代替することは禁止**（裁判所書式のレイアウト・数式・チェック欄が失われ、提出できない
  書面になる）。case.yaml の値をチャット上で整形して文書化するのも同様に禁止。
- hasan-kit の実体は**この SKILL.md の2階層上**にある `scripts/hasan-kit`
  （SKILL.md はプラグインルート直下の skills/ 配下にある。PATH に hasan-kit が無い環境では
  必ずこの絶対パスで呼ぶ。例: SKILL.md が /path/to/hasan/skills/generate/SKILL.md なら
  実体は /path/to/hasan/scripts/hasan-kit）。
- hasan-kit が見つからない・実行に失敗した場合は**そこで停止**し、エラー全文と
  `hasan-kit doctor` の結果を利用者に報告する（書類の自作でしのがない）。


## 手順

1. `hasan-kit validate-case cases/<事件ID>/case.yaml` を通す。
   債権者の残高不明・出所なしの警告が出たら、生成は続行しつつ最後に報告する。
2. 生成（裁判所書式 B1105 登録済み。fillmap 駆動で記入される）:
   ```
   hasan-kit saikensha \
     --case cases/<事件ID>/case.yaml \
     --output cases/<事件ID>/output/債権者一覧表_<姓>.xlsx
   ```
   - **公租公課の債権者がいる場合は必ず続けて公租公課用（B1106）も生成する**:
     `... --kouso --output cases/<事件ID>/output/債権者一覧表（公租公課用）_<姓>.xlsx`
   - 17件以上は書式の枠数超過（警告が出る）。2枚目はテンプレートのコピーに手動転記
     し、利用者に報告する。
   - 他庁書式など未登録の書式が必要なら register-form で登録
     （暫定は `--generic` の事務所内ドラフト）。
3. `hasan-kit crosscheck --case ... --dir cases/<事件ID>/output`
   で負債総額・債権者数の突合が OK になることを確認する。
4. `hasan-kit preview` でPDF・画像化し目視（列のはみ出し・文字切れ）。

## 記載ルール

- 並び順はビルダーが処理する: 一般債権者を case.yaml の記載順で先に、**公租公課は末尾**。
  番号は自動で振り直される（case.yaml の no は取込順の参考値）。
- 残高不明の債権者は金額空欄のまま載せる（合計に含まれない旨が注記される）。
  勝手に推計しない。
- 保証人有無が null（未確認）の債権者は空欄になる → questions.md で確認を促す。
- 裁判所書式版で列構成が異なる場合は fillmap.yaml の `columns` 対応を確認する
  （ビルダーの値キー: no/name/address/kind/origin_date/use/principal/balance/
  last_payment/guarantor/note）。
