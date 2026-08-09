---
name: hasan-shisan-mokuroku
description: >-
  破産申立ての資産目録（Excel）を事件モデル case.yaml から生成するスキル。
  「資産目録を作って」「財産目録を作成して」「資産をまとめて」
  などのフレーズがトリガー。破産申立ての資産目録の作成では必ず使用すること。
  xlsxスキルと併用すること。
---

# 資産目録 生成スキル

## 手順

1. `python3 scripts/casekit/validate_case.py cases/<事件ID>/case.yaml` を通す。
2. 生成:
   ```
   python3 scripts/build/build_shisan_mokuroku.py \
     --case cases/<事件ID>/case.yaml \
     --output cases/<事件ID>/output/資産目録_<姓>.xlsx
   ```
   裁判所書式が未登録なら exit 3 で案内が出る。利用者に書式提供を依頼し
   （hasan-register-form で登録）、それまでは了解を得て `--generic` で出力する。
3. crosscheck_outputs.py で資産総額の突合、render_preview.py で目視。

## 記載ルール

- 区分はビルダーが case.yaml の assets から組み立てる:
  現金／預貯金（口座ごと・基準日付き）／保険（解約返戻金見込）／自動車（所有権留保
  ローン残を備考に）／不動産（被担保債権残を備考に）／退職金見込／過払金等／その他。
- 評価額不明（null）は空欄のまま。0円資産（退職金見込0等）は0と明記する。
- 通帳残高の基準日（as_of）が受任時点と大きくずれていたら questions.md へ
  （直近残高証明の取り付けを促す）。
