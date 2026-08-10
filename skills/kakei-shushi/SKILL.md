---
name: kakei-shushi
description: >-
  破産申立ての家計収支表（Excel）を事件モデル case.yaml から計算ミスなく生成するスキル。
  「家計収支表を作って」「家計簿をまとめて」「収支表を作成して」
  などのフレーズがトリガー。破産申立ての家計収支表の作成では必ず使用すること。
  xlsxスキルと併用すること。
---

# 家計収支表 生成スキル

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

1. `hasan-kit validate-case cases/<事件ID>/case.yaml` を通す
   （月次の収入計・支出計・収支が表示されるので妥当性を見る）。
2. 生成:
   ```
   hasan-kit kakei \
     --case cases/<事件ID>/case.yaml \
     --output cases/<事件ID>/output/家計収支表_<姓>.xlsx
   ```
   裁判所書式（B1111 家計収支表）登録済み。費目は fillmap の正規名で case.yaml に
   記録されている前提（intake の規約）。正規名に無い費目は「その他」行に回り、
   実行ログに出る。書式内の数式（繰越・合計）は保持される。
3. hasan-kit crosscheck で月次集計の突合、hasan-kit preview で目視。

## 記載ルール

- 対象月は通常直近2か月（裁判所の運用に従う）。case.yaml の household の月をそのまま使う。
- 費目名は case.yaml のキーがそのまま行になる。裁判所書式版では書式の費目に合わせて
  intake 段階でキーを揃えておく（fillmap 登録時に費目対応を確認）。
- 収支が大きくプラス（可処分が多い）・マイナス（生活費不足）の月は、報告書第２の
  記載と整合しているか確認し、必要なら questions.md へ。
- 合計行は数式（=SUM）で入るので、Excel 上でも再計算で検算できる。
