# hasan — 破産申立資料AI自動作成システム

個人（自然人）破産・管財事件の申立資料（報告書・債権者一覧表・資産目録・家計収支表）を、
事件フォルダに置かれた原資料（調査票・通帳PDF・面談メモ等）から自動作成するリポジトリ。
裁判所配布の非公開書式（Word/Excel）を版管理し、書式改訂に追従できることが最大の特徴。

**このリポジトリは Claude Code プラグイン**（`.claude-plugin/plugin.json`）でもある。
この CLAUDE.md は**ソースリポジトリでの開発・書式登録作業用**の規約。利用者としての
使い方（インストール・作業フォルダ）は README.md を参照。スキルは `skills/`（プラグイン
形式・`/hasan:<スキル名>`）、スクリプト呼出しは `bin/hasan-kit` に統一されている。
事件データ（cases/）は利用時は**作業フォルダ側**に置く（このリポジトリ内の cases/ は
開発時の検証用の受け皿で、従来どおり git 管理外）。

## 全体ワークフロー

```
0. 面談準備   hearing-sheet        聴取シート生成（2回目以降は questions.md の差分シート）
   面談記録   menmemo              走り書き・書き起こし → 面談メモ docx（input/ の原資料に）
1. 書式登録   register-form        裁判所配布の新書式 → courts/<裁判所>/forms/<書式>/v<版>/
                                   （template + fillmap.yaml + structure.md + anchors.txt）
2. 事件取込   intake               cases/<事件>/input/ の原資料 → case.yaml + questions.md
3. 入出金分析 nyushukkin-bunseki   通帳 → work/bank.csv → case.yaml の bank_analysis/menseki
4. 一式生成   generate             case.yaml + fillmap → output/ の記入済み docx/xlsx
                                   （個別書式は houkokusho 等の書式別スキル）
5. 検証       hasan-kit crosscheck/preview  書類間数値突合 + PDF目視
6. 依頼連絡   line-message         questions.md → 依頼者向けLINE文面（手動送信）
```

## 手続種別（最初に選択）

事件は `meta.proc_type` で **同時廃止／管財（自然人）／管財（法人）** に分かれ、
書式セット・生成順は `courts/osaka/procs.yaml` が定義する（registry.yaml の `procs:` が
所属セット）。管財2セットは書式未登録の受け皿（受領後 register-form で登録・追記）。
管財（法人）は債務者＝法人（case.yaml の `corporation` を使い、自然人用セクションは使わない）。

## 登録済み書式（courts/osaka/ — 大阪地裁 本庁・堺支部・岸和田支部共通様式、同時廃止セット）

| 書式ID | 書式（原番号） | 形式 | 生成 |
|---|---|---|---|
| moushitatesho-douhai | 破産手続開始申立書・同時廃止用（B1102） | docx | hasan-moushitatesho |
| houkokusho | 報告書・自然人用 ver.4.1（B1110） | docx | hasan-houkokusho |
| jigyou-houkokusho | 事業に関する報告書（B1112） | docx | hasan-jigyou-houkokusho（個人事業者のみ） |
| saikensha-ichiran | 債権者一覧表（B1105） | xlsx | build_saikensha_ichiran.py |
| saikensha-ichiran-kouso | 債権者一覧表・公租公課用（B1106） | xlsx | 同 --kouso |
| shisan-mokuroku | 財産目録（B1109） | xlsx | build_shisan_mokuroku.py |
| kakei-shushi | 家計収支表（B1111） | xlsx | build_kakei_shushi.py |
| joshinsho-checklist | チェックリストに関する上申書（B1113） | docx | fill_docx.py（弁護士指示時のみ） |
| checklist-douhai | 同時廃止チェックリスト（B1114） | docx | 生成対象外（保管のみ） |
| hyojun-shiryo-ichiran | 標準資料一覧表（B1103） | xlsx | 生成対象外（保管のみ） |

宛先支部は事件ごとに case.yaml の `meta.court_branch`（第６民事部/堺支部/岸和田支部）で指定。

## ディレクトリ規約

- `courts/<裁判所ID>/forms/<書式ID>/` — 裁判所書式（版別）。`registry.yaml` の `current:` が現行版。
  - `v*/template.docx|template.xlsx` は裁判所配布の**白紙原本。絶対に編集しない**。
  - `fillmap.yaml` = 記入位置の機械可読アンカー定義。生成ロジックは必ずこれ経由で欄を特定する。
- `cases/<事件ID>/` — 事件ワークスペース。**依頼者情報を含むため git 管理外**（cases/.gitignore）。
  - `input/`（原資料・読取専用扱い）, `work/`（中間物）, `case.yaml`（事件モデル＝唯一の真実）,
    `questions.md`（弁護士への確認事項）, `output/`（記入済み書面＋検証用PDF/画像）。
- `schema/` — case.yaml のスキーマと架空サンプル事例。
- `scripts/` — formkit（書式登録・照合）/ casekit（検証・正規化）/ build（xlsx生成）/ verify（検証）。

## 鉄則

1. **cases/ 配下は絶対にコミットしない**（依頼者の実名・口座情報を含む）。依頼者情報を
   cases/<事件ID>/ の外（scratchpad含む一時領域は可、リポジトリ内の他所は不可）へ複製しない。
2. **courts/ のテンプレート原本は編集禁止**。記入は必ずコピーに対して行う。
   書式は非公開資料のため、このリポジトリは private を維持する。
3. **根拠のない記載はしない**。case.yaml の値には `sources` で出所（どの資料の何頁か）を付け、
   出所のない値は書式生成時に空欄扱いとし questions.md に回す。
4. **書類間の数値整合**は case.yaml を単一の真実として担保する。生成器が原資料を直接読むことは禁止。
5. 生成した docx/xlsx は必ず `scripts/verify/render_preview.py` でPDF・ページ画像化して目視検証する。

## 技術メモ

- docx 記入は「配布docxを unzip → docxスキルの merge_runs.py → word/document.xml を直接編集 →
  zip 再パック → validate.py --original → soffice.py でPDF」。新規生成はしない（レイアウト完全保持）。
- 段落の一意特定は fillmap の多層アンカー（w14:paraId → 文言 → 構造パス → 文書内位置idx）で
  行う。B1102等の旧型式docxは paraId が無く、文言以下の層だけで解決する。
- xlsx 記入は openpyxl（scripts/build/）。数式・結合セル・印刷設定を壊さないこと。
- この環境に pdftoppm は無い。PDF→画像は PyMuPDF（`python3 -m scripts.verify.render_preview`）を使う。
- 日付は裁判所書式の実態に合わせ「R4後半ころ」等の曖昧文字列も許容。短縮元号（H22.6）等の
  表記変換は fillmap の `date_style` 指定が担う。

## 書式が改訂されたら

新しい配布ファイルを受け取ったら「新しい書式を登録して」と依頼する（hasan-register-form）。
手順の詳細は docs/form-registration.md。旧版は削除せず courts/ に残す（過去事件の再現用）。
