---
name: houkokusho-hojin
description: >-
  破産管財（法人）事件の「報告書（法人用）」（大阪地裁0104様式）を事件モデル
  case.yaml の corporation セクションから作成・記入するスキル。管財手続の希望・
  営業内容・事業用物件・従業員・支払停止・許認可・売掛金等の18項目と、
  破産原因の事情（別紙）を扱う。「法人の報告書を作って」「報告書（法人用）を記入して」
  「0104を埋めて」「会社の破産の報告書」などのフレーズがトリガー。
  法人破産の報告書の作成では必ずこのスキルを使用すること。docxスキルと併用すること。
---

# 報告書（法人用）（0104）作成スキル

**適用範囲**: 管財（法人）事件のみ（meta.proc_type=管財（法人））。
債務者は法人 — 記入値は case.yaml の **corporation** セクションから組み立てる
（自然人用の applicant・career・family・menseki は使わない）。

書式: `$(hasan-kit courts)/osaka/forms/houkokusho-hojin/`（registry の current が現行版）。
構造・記入欄は structure.md / fillmap.yaml、見本は examples/values.sample.yaml。

## 手順

1. `hasan-kit validate-case cases/<事件ID>/case.yaml` を通す。
2. corporation から `cases/<事件ID>/work/houkokusho_hojin_values.yaml` を組み立てる。
3. `hasan-kit fill-docx --template ... --fillmap ... --values ...
   --output cases/<事件ID>/output/報告書_<商号略>.docx`
4. **項目18の別紙**: corporation.hasan_gennin（破産原因が生じた事情・粉飾決算の有無）を
   時系列の経緯書として `hasan-kit memo-docx` で生成し、報告書に添付する
   （書式内に記入欄が無いため）。timeline があれば時系列を再利用する。
5. `hasan-kit preview` でPDF・目視。

## values 組み立てルール

- `dairinin`: 「　乙　川　一　郎　」— **「印」を含めない**（別runのため）。
- `daihyosha_sign`: 会社代表者の氏名（準自己破産は申立人）。「甲　野　太　郎」形式。
- **`kanzai_kibou`（項目1）は弁護士の手続選択** — case.yaml に無ければ値を渡さず
  questions.md へ。
- 項目3 事業用物件: corporation.offices[]（[0]=本店）から bukken1_*/bukken2_*。
  3件以上は別紙（手書き案内）。賃借の返戻金は財産目録⑫賃借保証金・敷金と一致させる。
- 項目4 従業員: corporation.employees / union から。手続の完了チェック
  （離職票・雇用保険・社会保険・源泉・住民税）は**控え等の資料で確認できたものだけ**。
  労働組合が無い場合は kahansu_daihyo（従業員過半数代表者）を記入。
- 項目5〜7: corporation.shiharai_teishi / licenses、リースは creditors kind=リース。
- 項目8・9: lawsuits / pre_bankruptcy_disposals の有無と一致させ、有なら
  0133-0134 一覧表を必ず生成する。
- 項目10: corporation.bank_accounts_closed（相殺予定以外の口座解約）。
- 項目11〜13: assets.accounts_receivable / loans_receivable / inventory・machinery・
  fixtures から。財産目録（0116-0131）の記載と一致させる。
- 項目15・16: 受任後業務3項目の□は**実施済みのものだけ**「確認した」。代表者連絡先は
  corporation.representative_address/tel/mobile（管財人が直ちに連絡できる状態が前提 —
  不明なら questions.md へ）。
- 有無・チェックは資料で確認できるものだけ。未確認は値を渡さず questions.md へ。

## 注意

- 申立書（0103）・資産及び負債一覧表（0106）・債権者一覧表（0107-0114）・
  財産目録（0116-0131）と数値・チェックの整合をとる（crosscheck が突合）。
- 代表者個人の破産を同時申立てする場合は、代表者の管財（自然人）事件を別事件として
  intake し、0103 の「代表者の破産申立てをしたか」＝有・同時申立と整合させる。
