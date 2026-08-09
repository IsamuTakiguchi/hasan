---
name: hasan-register-form
description: >-
  裁判所から提供された破産申立て書式（Word/Excel、非公開・随時改訂）をこのリポジトリに
  登録・更新するスキル。新しい書式ファイルを courts/ に版として取り込み、記入位置の
  アンカー定義（fillmap.yaml）を作成し、旧版からの自動再束縛と差分レポートで
  書式改訂に追従する。「新しい書式を登録して」「書式が改訂された」「書式を更新して」
  「裁判所の書式を取り込んで」「新版の報告書書式が来た」などのフレーズがトリガー。
  裁判所書式ファイルの登録・改訂対応では必ずこのスキルを使用すること。
---

# 裁判所書式の登録・更新スキル

裁判所書式は非公開・随時改訂される。このスキルは書式を **版として保存**し、生成ロジックが
束縛する **fillmap（多層アンカー定義）** を整備して、改訂されても最小の手間で追従できる
ようにする。原本は絶対に編集しない。

## A. 初回登録（この書式を初めて取り込む）

1. **配置**: 利用者から受け取ったファイルを
   `courts/<裁判所ID>/forms/<書式ID>/v<版>/template.docx（.xlsx）` に置く。
   - 書式IDはローマ字ケバブ（例 saikensha-ichiran）。版は書式に記載の版数
     （なければ受領日 vYYYYMM）。
   - `registry.yaml` を作成し `current: "v<版>"` と受領情報を書く
     （既存書式の registry を見本に）。
2. **機械ダンプ**:
   - docx: `python3 scripts/formkit/inspect_docx.py <template> -o <版dir>/anchors.txt --full`
     フラグの意味: T=表内 E=空 C=チェック U=下線 **S=sectPr（行複製・削除禁止）**
   - xlsx: `python3 scripts/formkit/inspect_xlsx.py <template> -o <版dir>/cellmap.txt`
3. **実物を目視**: `python3 scripts/verify/render_preview.py <template>` でPDF・画像化し、
   ダンプと突き合わせながら書式の構成を把握する。
4. **fillmap.yaml の起草**: ダンプを根拠に記入欄を定義する。
   - docx の anchor は **paraId + context（正規化文言）+ path（構造パス）+ idx** の
     4層を必ず埋める（版更新時の自動再束縛の材料）。kind は
     `scripts/build/fill_docx.py` のdocstring参照（set_text / insert_cell_text /
     checkbox / underline_fill / underline_slots / underline_longtext / timeline_rows）。
   - 既存の `courts/osaka-sakai/forms/houkokusho/v4.0/fillmap.yaml` が完全な見本。
     規模が大きい書式は houkokusho の登録で使ったように、anchors.txt から fillmap を
     機械生成する使い捨てスクリプトを書くとよい。
   - xlsx の anchor は `{sheet, cell}` + `label_check`（近傍ラベル検証）。
     繰返し行は `table_rows`（first_data_row / columns / max_rows）。
   - **sectPr 段落（Sフラグ）を absorb_paraIds / extra_paraIds に入れないこと。**
5. **構造メモ**: `<版dir>/structure.md` に書式の構成・注意点・未定義領域を書く
   （houkokusho の structure.md が見本）。
6. **アンカー検査**: `python3 scripts/formkit/check_anchors.py <fillmap> <template>` が
   **全件 OK**（REBOUND/BROKEN ゼロ）になること。
7. **スモークフィル**: 架空事例の記入値の見本 `<版dir>/examples/values.sample.yaml` を
   作り（`schema/examples/case.sample.yaml` から組み立てる）、fill_docx.py／該当
   builder で実際に生成 → render_preview.py で全ページ目視（欄ズレ・折返し・チェック位置）。
8. コミット（テンプレート・fillmap・structure.md・anchors.txt・examples）。

## B. 版更新（改訂された書式を受け取った）

1. 新ファイルを `v<新版>/template.docx` に配置（旧版は残す）。
2. **自動照合**:
   ```
   python3 scripts/formkit/diff_fillmap.py \
     --old-fillmap courts/.../v<旧>/fillmap.yaml \
     --new-template courts/.../v<新>/template.docx \
     --write  courts/.../v<新>/fillmap.yaml \
     --report courts/.../v<新>/diff-from-v<旧>.md
   ```
   判定: **OK**=paraId一致 / **REBOUND**=文言・構造パス・文書位置で自動再束縛
   （裁判所の再配布は paraId 総入替えが普通なので REBOUND が主）/ **BROKEN**=要修正。
3. **BROKEN の修理**: レポートの BROKEN と「未参照のチェック欄（新設候補）」を見ながら、
   新テンプレートの anchors.txt（inspect_docx.py で作る）を根拠に fillmap の該当
   フィールドを修正・追加する。fillmap 内の `FIXME` マークを消す。
   欄が廃止されていたらフィールドを削除し、structure.md に記録する。
4. `check_anchors.py` 全件 OK → **スモークフィル**（A-7 と同じ）→ 目視合格。
5. `registry.yaml` の `current:` を新版に更新し、structure.md の変更点を追記してコミット。
6. 利用者へ報告: 変更された欄・新設欄・廃止欄、体裁ルールへの影響
   （セル幅変更で短縮表記の要否が変わる等）。

## 注意

- テンプレート原本は**編集禁止**（記入は必ずコピーに対して行う）。
- 書式は非公開資料。リポジトリは private を維持し、外部に送信しない。
- 旧版は削除しない（過去事件の再現・比較のため）。
- 模擬改訂テスト（v4.0 の paraId を全部入れ替えて2箇所の文言を変えた試験）で、
  報告書の全229アンカーが BROKEN ゼロで自動再束縛されることを確認済み。この水準を
  他の書式でも維持する（4層アンカーを省略しない）。
