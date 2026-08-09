---
name: moushitatesho
description: >-
  破産手続開始申立書（大阪地裁）を事件モデル case.yaml から作成・記入するスキル。
  同時廃止用（B1102）と管財・自然人用（0203）の両書式に対応し、事件の手続種別
  meta.proc_type で自動的に使い分ける。裁判所配布Word書式のXMLを直接編集し
  レイアウトを保持する。「破産の申立書を作って」「開始申立書を記入して」
  「同時廃止の申立書」「管財の申立書」「B1102を埋めて」「0203を埋めて」
  などのフレーズがトリガー。破産手続開始申立書の作成では
  必ずこのスキルを使用すること。docxスキルと併用すること。
---

# 破産手続開始申立書 作成スキル

**書式の使い分け**（meta.proc_type で自動選択）:

| proc_type | 書式 | ディレクトリ |
|---|---|---|
| 同時廃止 | 破産手続開始申立書・同時廃止用（B1102） | `$(hasan-kit courts)/osaka/forms/moushitatesho-douhai/` |
| 管財（自然人） | 破産申立書（自然人・管財事件用）（0203） | `$(hasan-kit courts)/osaka/forms/moushitatesho-kanzai/` |
| 管財（法人） | 破産申立書（法人用）（0103） | `$(hasan-kit courts)/osaka/forms/moushitatesho-hojin/` |

いずれも registry.yaml の current が現行版。構造・記入欄は各ディレクトリの
structure.md / fillmap.yaml 参照。法人用の values 組み立ての詳細は
skills/houkokusho-hojin と同様 corporation セクションから（下記は自然人2書式のルール）。

## 手順

1. `hasan-kit validate-case` を通し、負債総額・債権者数を控える。
2. case.yaml から `cases/<事件ID>/work/moushitatesho_values.yaml` を組み立てる
   （完全な見本: 各版 `examples/values.sample.yaml`）。
3. `hasan-kit fill-docx --template ... --fillmap ... --values ... --output
   cases/<事件ID>/output/破産申立書_<姓>.docx`
4. `hasan-kit preview` でPDF・目視（チェック位置・下線欄・申立ての理由の数値）。

## values 組み立てルール（B1102・同時廃止）

- `atesaki`: meta.court_branch（第６民事部／堺支部／岸和田支部）。
- `dairinin`: 「　乙　川　一　郎　　印」形式（下線runを「印」ごと置換）。
- `soutatsu_basho`: 「〒590-0000　住所　事務所名」（attorneys[0].address + office）。
- `tel_fax`: 「Tel（072）000－0000　Fax（072）000－0001」で全文置換。
- `furigana`: 「(ふりがな)　こうの　たろう」で全文置換（kana はひらがな）。
- `shimei_kyusei`: リスト [氏名（1字ずつ全角スペース割付）, 旧姓]。旧姓が無ければ ""。
- `nenrei_seinengappi`: 「年　　　齢　（３７歳）（昭和６３年３月１０日生）」で全文置換。
  年齢は**申立（予定）日時点**で計算。元号は生年から確定させ、使わない元号は書かない。
- `honseki`: 「住民票のとおり」または「国籍」＋ `honseki_kokuseki_text` に国名。
- `jusho`: 住民票どおりなら「住　居　所　☑〒(590-0000)　住民票記載のとおり」で全文置換。
  異なる場合は `jusho_betsu` の下線に住所（このとき jusho は置換しない）。
- **`riyu`（申立ての理由）**: 全文置換。数値は必ず case.yaml から計算する:
  - 債権者数 = 全債権者数（公租公課含む）
  - 債務総額 = 残高判明分の合計（hasan-kit validate-case の表示と一致）
  - 除外後額 = 総額 −（kind=保証・求償の残高）−（使途が住宅ローンの残高）
  - 金額は全角数字＋全角カンマ（例 金２２，１７０，０００円）
  - **債権者一覧表（B1105/B1106）と数値が一致すること**（crosscheckで検査される）
- 参考事項:
  - `sanko1_kojin_jigyo`: business（個人事業）の有無から。法人代表者のみなら「否」
  - `sanko2_hojin_daihyo`: 現在（登記簿上含む）法人代表か。career の現職が法人代表者
    でなくても**退任登記が未了なら「当」**（登記の確認を questions.md へ）
  - `sanko3_seikatsu_hogo`: meta.seikatsu_hogo
  - `sanko4_saisei`: menseki.past_saisei_ninka.exists（有/無）
  - `sanko5_menseki`: menseki.past_menseki.exists — **報告書第３の５①と一致必須**
- `denshi_nofu`: meta.denshi_nofu_code があるときのみ。
- 有無・チェックは資料で確認できるものだけ。未確認は値を渡さず questions.md へ。

## values 組み立てルール（0203・管財（自然人）固有）

共通欄（dairinin / soutatsu_basho / tel_fax / honseki）は B1102 と同じ要領。相違点:

- `atesaki_shibu`: **支部事件のみ**「支部」を渡す（☑のみ入る。支部名は提出時に
  手書き — 完了報告で必ず案内）。本庁（第６民事部）は値を渡さない。
- `shimei`: リスト [氏名, 旧姓・通称・屋号]。**ふりがなはルビのため機械記入不可**
  （提出時に手書き案内を完了報告に含める）。
- 金額欄（`ippan_saiken`・`yusen_saiken`・`kaishu_mikomi`）: 下線が「＿万＿」の形なので
  値も「２，２０４万５，０００」のように**万込み**で渡す（円は印字済み）。
  - 一般破産債権 = 公租公課・労働債権を除く債権の総額と件数
  - 優先的破産債権及び財団債権 = 公租公課＋労働債権等の総額と件数
  - 回収見込額合計 = 財産目録（0218総括表）の回収見込額合計と一致
  - **債権者一覧表（0209-0216）・財産目録（0218-0234）と一致すること**（crosscheck対象）
- `hikitsugi_genkin`: 管財人への引継予定現金（「３００，０００」形式。**必ず記載**の欄。
  未定なら弁護士に確認＝questions.md へ）。
- `hojin_daihyo`: リスト（例 [当, 法人破産有]）。当のときは `hojin_jiken`（係属裁判所
  5スロット）・`hojin_shinko`・`hojin_kanzainin`・`hojin_yotei`・`hojin_yotei_jiki` も検討。
- `kojin_jigyo`: 当／否。`haigusha_moushitate`: 有／無（有なら `haigusha_jiken`）。
- 申立ての趣旨の自由財産拡張（■）は財産目録側の欄。この書式では記入不要
  （自由財産拡張申立書 0240 と財産目録の■欄が対応書類）。

## 注意

- どちらの書式も w14:paraId の無い旧型式docx。アンカーは文言＋位置で解決される
  （fillmap の context を変えないこと）。
- 0203 の年齢・住居所・連絡先のラベルは EQ フィールド。fill_docx が保護するが、
  fillmap を編集するときもフィールド run に触れる kind を割り当てないこと。
- 印紙・郵券・受領印の枠（ページ下部のフレーム）は触らない。
- 署名欄・氏名欄の折返しはPDF代替フォントでは実際とずれる。Word実機での最終確認を案内する。
