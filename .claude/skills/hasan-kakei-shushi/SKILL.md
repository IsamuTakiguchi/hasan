---
name: hasan-kakei-shushi
description: >-
  破産申立ての家計収支表（Excel）を事件モデル case.yaml から計算ミスなく生成するスキル。
  「家計収支表を作って」「家計簿をまとめて」「収支表を作成して」
  などのフレーズがトリガー。破産申立ての家計収支表の作成では必ず使用すること。
  xlsxスキルと併用すること。
---

# 家計収支表 生成スキル

## 手順

1. `python3 scripts/casekit/validate_case.py cases/<事件ID>/case.yaml` を通す
   （月次の収入計・支出計・収支が表示されるので妥当性を見る）。
2. 生成:
   ```
   python3 scripts/build/build_kakei_shushi.py \
     --case cases/<事件ID>/case.yaml \
     --output cases/<事件ID>/output/家計収支表_<姓>.xlsx
   ```
   裁判所書式が未登録なら exit 3 で案内が出る。利用者に書式提供を依頼し
   （hasan-register-form で登録）、それまでは了解を得て `--generic` で出力する。
3. crosscheck_outputs.py で月次集計の突合、render_preview.py で目視。

## 記載ルール

- 対象月は通常直近2か月（裁判所の運用に従う）。case.yaml の household の月をそのまま使う。
- 費目名は case.yaml のキーがそのまま行になる。裁判所書式版では書式の費目に合わせて
  intake 段階でキーを揃えておく（fillmap 登録時に費目対応を確認）。
- 収支が大きくプラス（可処分が多い）・マイナス（生活費不足）の月は、報告書第２の
  記載と整合しているか確認し、必要なら questions.md へ。
- 合計行は数式（=SUM）で入るので、Excel 上でも再計算で検算できる。
