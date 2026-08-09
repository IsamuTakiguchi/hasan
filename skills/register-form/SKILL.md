---
name: register-form
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

**プラグイン運用での前提**: 書式の登録・更新は**プラグインのソースリポジトリの
クローン**（git 管理下の hasan リポジトリ）で行う。インストールされたプラグイン
ディレクトリへの書き込みは更新で消えるため不可。作業ディレクトリが git クローンで
ない場合は、クローンの場所を利用者に確認してそこで作業する。登録後は
コミット・プッシュし、必要なら `claude plugin tag --push` で版タグを付けて配布する
（各利用環境はプラグイン更新で新書式を受け取る）。
以下のパス表記 `courts/...` `scripts/...` はそのクローンのリポジトリルート相対。
登録完了後、Cowork 利用者がいる場合は `hasan-kit package` で新しい配布zipを作成して
再アップロードを案内する（plugin.json の version を上げてから）。

## A. 初回登録（この書式を初めて取り込む）

1. **手続種別セットの確認**: この書式がどのセットに属すかを利用者に確認する —
   **douhai（同時廃止）／kanzai-shizenjin（管財・自然人）／kanzai-hojin（管財・法人）／
   共用（複数セット）**。判断材料: 書式の表題・版表記（同時廃止一式は ver.4.1 世代）・
   利用者の説明。共用の典型例: 報告書（自然人用）が同時廃止と管財（自然人）で
   同一書式の場合。
2. **配置**: 利用者から受け取ったファイルを
   `courts/<裁判所ID>/forms/<書式ID>/v<版>/template.docx（.xlsx）` に置く。
   - 書式IDはローマ字ケバブ（例 saikensha-ichiran）。**管財版は `-kanzai`、
     法人用は `-hojin` サフィックス**（例 moushitatesho-kanzai, houkokusho-hojin。
     既存書式と同一内容の共用なら新IDを作らず registry の procs に追記するだけ）。
     版は書式に記載の版数（なければ受領日 vYYYYMM）。
   - `registry.yaml` を作成し `current: "v<版>"`・受領情報・**`procs: [<セット>]`** を書く
     （既存書式の registry を見本に）。
   - **`courts/<裁判所ID>/procs.yaml` の該当セットの forms に追記**する
     （生成順の位置・via（builder か skill か）・when 条件を利用者に確認。
     生成対象外の保管書式は keep_only へ）。
   - 管財（法人）の書式で case.yaml に無い情報（法人の債権者分類・労働債権・
     担保権等）が必要になったら、schema/case.schema.yaml の corporation ほかを
     拡張し、サンプル事例と intake スキルにも反映する。
2. **機械ダンプ**:
   - docx: `hasan-kit inspect-docx <template> -o <版dir>/anchors.txt --full`
     フラグの意味: T=表内 E=空 C=チェック U=下線 **S=sectPr（行複製・削除禁止）**
   - xlsx: `hasan-kit inspect-xlsx <template> -o <版dir>/cellmap.txt`
3. **実物を目視**: `hasan-kit preview <template>` でPDF・画像化し、
   ダンプと突き合わせながら書式の構成を把握する。
4. **fillmap.yaml の起草**: ダンプを根拠に記入欄を定義する。
   - docx の anchor は **paraId + context（正規化文言）+ path（構造パス）+ idx** の
     4層を必ず埋める（版更新時の自動再束縛の材料）。kind は
     `hasan-kit fill-docx` のdocstring参照（set_text / insert_cell_text /
     checkbox / underline_fill / underline_slots / underline_longtext / timeline_rows）。
   - 既存の `courts/osaka/forms/houkokusho/v4.0/fillmap.yaml` が完全な見本。
     規模が大きい書式は houkokusho の登録で使ったように、anchors.txt から fillmap を
     機械生成する使い捨てスクリプトを書くとよい。
   - xlsx の anchor は `{sheet, cell}` + `label_check`（近傍ラベル検証）。
     繰返し行は `table_rows`（first_data_row / columns / max_rows）。
   - **sectPr 段落（Sフラグ）を absorb_paraIds / extra_paraIds に入れないこと。**
5. **構造メモ**: `<版dir>/structure.md` に書式の構成・注意点・未定義領域を書く
   （houkokusho の structure.md が見本）。
6. **アンカー検査**: `hasan-kit check-anchors <fillmap> <template>` が
   **全件 OK**（REBOUND/BROKEN ゼロ）になること。
7. **スモークフィル**: 架空事例の記入値の見本 `<版dir>/examples/values.sample.yaml` を
   作り（`schema/examples/case.sample.yaml` から組み立てる）、hasan-kit fill-docx／該当
   builder で実際に生成 → hasan-kit preview で全ページ目視（欄ズレ・折返し・チェック位置）。
8. コミット（テンプレート・fillmap・structure.md・anchors.txt・examples）。

## B. 版更新（改訂された書式を受け取った）

1. 新ファイルを `v<新版>/template.docx` に配置（旧版は残す）。
2. **自動照合**:
   ```
   hasan-kit diff-fillmap \
     --old-fillmap courts/.../v<旧>/fillmap.yaml \
     --new-template courts/.../v<新>/template.docx \
     --write  courts/.../v<新>/fillmap.yaml \
     --report courts/.../v<新>/diff-from-v<旧>.md
   ```
   判定: **OK**=paraId一致 / **REBOUND**=文言・構造パス・文書位置で自動再束縛
   （裁判所の再配布は paraId 総入替えが普通なので REBOUND が主）/ **BROKEN**=要修正。
3. **BROKEN の修理**: レポートの BROKEN と「未参照のチェック欄（新設候補）」を見ながら、
   新テンプレートの anchors.txt（hasan-kit inspect-docx で作る）を根拠に fillmap の該当
   フィールドを修正・追加する。fillmap 内の `FIXME` マークを消す。
   欄が廃止されていたらフィールドを削除し、structure.md に記録する。
4. `hasan-kit check-anchors` 全件 OK → **スモークフィル**（A-7 と同じ）→ 目視合格。
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
