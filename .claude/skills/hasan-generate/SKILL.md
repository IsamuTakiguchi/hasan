---
name: hasan-generate
description: >-
  破産申立資料一式（債権者一覧表・報告書・資産目録・家計収支表）を事件モデル
  case.yaml から一括生成するオーケストレータスキル。個別書式のスキル
  （hasan-houkokusho 等）を正しい順序で実行し、書類間の数値整合検査と
  PDF目視検証まで行う。「破産申立書類一式を作って」「申立資料をまとめて作成して」
  「破産の書類を全部作って」「一式生成して」などのフレーズがトリガー。
  破産申立資料をまとめて作成する場面では必ずこのスキルを使用すること。
---

# 破産申立資料 一括生成スキル（オーケストレータ）

## 前提

- `cases/<事件ID>/case.yaml` が存在すること（なければ先に hasan-intake）。
- 通帳があるのに `bank_analysis` が空なら、先に hasan-nyushukkin-bunseki を促す
  （免責関係の記載が変わり得るため）。

## 実行順序

数値の親（債権者・資産・家計の生データ）→ 子（それを引用する報告書）の順で作る。

1. **生成前ゲート**: `python3 scripts/casekit/validate_case.py cases/<事件ID>/case.yaml`
   - エラーなら中断して利用者に報告。警告（出所なし等）は控えて最後にまとめる。
   - 表示される負債総額・資産総額・月次収支を控える（後の突合の基準）。
2. **債権者一覧表**: hasan-saikensha-ichiran の手順で
   `build_saikensha_ichiran.py` を実行。
3. **資産目録**: hasan-shisan-mokuroku の手順で `build_shisan_mokuroku.py` を実行。
4. **家計収支表**: hasan-kakei-shushi の手順で `build_kakei_shushi.py` を実行。
5. **報告書**: hasan-houkokusho の手順で values.yaml を組み立てて `fill_docx.py` を実行。
   第２・第３の記載が 1〜4 の数値・bank_analysis と矛盾しないよう組み立てる。
6. **書類間整合**:
   `python3 scripts/verify/crosscheck_outputs.py --case ... --dir cases/<事件ID>/output`
   が全 OK になること。NG が出たら原因（case.yaml と生成物のどちらが古いか）を
   特定して作り直す。生成物を手で直して辻褄を合わせてはならない。
7. **目視検証**: `python3 scripts/verify/render_preview.py cases/<事件ID>/output/*.docx
   cases/<事件ID>/output/*.xlsx` で全書類をPDF・画像化し、各スキルのチェックリストで
   目視する。
8. 途中で失敗した書式があっても残りは続行し、最後にまとめて報告する。

## 出力命名

`cases/<事件ID>/output/<書式名>_<申立人姓>[_記入済み].docx|.xlsx`
（例: `債権者一覧表_甲野.xlsx`, `報告書_甲野_記入済み.docx`）

## 最終報告（必須）

- 生成した書類と各書類の要点（債権者数・負債総額・資産総額・月次収支）
- 空欄のまま残した項目とその理由（根拠資料なし等）
- 推測を含む記載・資料間の矛盾 → questions.md の確認事項（新規分を明示）
- 裁判所書式が未登録で --generic（事務所内ドラフト様式）になった書類
  → 裁判所書式の入手と hasan-register-form での登録を案内
