---
name: jigyou-houkokusho
description: >-
  破産申立ての「事業に関する報告書」（個人事業者用・大阪地裁B1112様式）を
  事件モデル case.yaml の business 区分から作成・記入するスキル。
  「事業に関する報告書を作って」「事業報告書を記入して」「B1112を埋めて」
  「個人事業の報告書」などのフレーズがトリガー。個人事業者（過去に営んでいた場合を
  含む）の破産申立てでは必ずこのスキルを使用すること。docxスキルと併用すること。
---

# 事業に関する報告書（B1112）作成スキル

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


**適用範囲**: 同時廃止セットの書式。管財セットでの要否は管財書式の受領時に判断。

書式: `$(hasan-kit courts)/osaka/forms/jigyou-houkokusho/`（registry.yaml の current が現行版）。

## 対象の判断（重要）

- **個人事業者**（現在または過去に屋号で事業を営んでいた）の事件で作成する。
  申立書（B1102）参考事項1が「申立前６か月以内」「現在」の場合は必須。
- **法人代表者のみ**（個人事業なし）の場合はこの書式は不要。法人の破綻経緯は
  報告書（B1110）第２に記載する。乙山ストア架空事例はこのパターン（生成対象外）。
- 迷う場合（法人成り前の個人事業期間がある等）は questions.md で確認する。

## 手順

1. case.yaml の `business` 区分を確認（無ければ intake で聴取・追記から）。
2. `cases/<事件ID>/work/jigyou_values.yaml` を組み立てる
   （見本: `$(hasan-kit courts)/osaka/forms/jigyou-houkokusho/v4.1/examples/values.sample.yaml`）。
3. `hasan-kit fill-docx` で記入 → `render_preview.py` でPDF目視。

## values 組み立てルール

- `atesaki`／`dairinin`／`houkoku_date` は報告書（houkokusho）と同じ要領。
- 年度別営業状況（nendo1〜3）: 古い年度から3年分。金額は「6,200,000円」・
  人数は「0人」の形式（「円」「人」ごと置換）。確定申告書・帳簿を出所とする。
  資料が無い年度は空欄にして questions.md へ。
- `jigyo_kikan`: 「平成２８年４月～令和７年１月」形式で全文置換
  （継続中は「～現在」ではなく終期を書かず questions.md で表記を確認）。
- `tax_total`（公租公課滞納額合計）は **creditors の kind=公租公課 残高合計**を使い、
  債権者一覧表（公租公課用 B1106）と一致させる。`tax_umu` も同じ根拠で有/無。
- `rental_arrears_detail`: 下線2スロット [か月数, 合計額]。
- 有無チェックは資料・聴取で確認できたものだけ。未確認は値を渡さない。

## 検証

- validate PASS ＋ PDF目視（有無チェック位置・表の3年分・下線欄の折返し）。
- 公租公課の額が B1106・case.yaml と一致していること（crosscheck）。
