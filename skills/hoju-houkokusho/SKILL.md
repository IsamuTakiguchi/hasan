---
name: hoju-houkokusho
description: >-
  破産管財（自然人）事件の「管財補充報告書」（大阪地裁0204様式）を事件モデル
  case.yaml から作成・記入するスキル。管財手続の希望・自由財産拡張・郵便取扱局・
  個人事業の清算状況（従業員・売掛金・リース・継続的契約）等の管財固有25項目を扱う。
  「管財補充報告書を作って」「補充報告書を記入して」「0204を埋めて」
  などのフレーズがトリガー。管財補充報告書の作成では必ずこのスキルを使用すること。
  docxスキルと併用すること。
---

# 管財補充報告書（0204）作成スキル

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


**適用範囲**: 管財（自然人）事件のみ（meta.proc_type=管財（自然人））。
同時廃止には無い書式。法人事件では使わない（法人は報告書 0104 に一本化）。

書式: `$(hasan-kit courts)/osaka/forms/kanzai-hoju-houkokusho/`（registry の current が現行版）。
構造・記入欄は structure.md / fillmap.yaml、記入値の見本は examples/values.sample.yaml。

## 手順

1. `hasan-kit validate-case cases/<事件ID>/case.yaml` を通す。
2. case.yaml から `cases/<事件ID>/work/hoju_values.yaml` を組み立てる。
3. `hasan-kit fill-docx --template ... --fillmap ... --values ...
   --output cases/<事件ID>/output/管財補充報告書_<姓>.docx`
4. `hasan-kit preview` でPDF・目視（チェック位置・下線欄）。

## values 組み立てルール

- `jiken_bango`: 事件番号は申立て前は未付番 → 値を渡さず空欄のまま。
- `saimusha`: 「債務者　甲野太郎」形式。`dairinin`: 「　乙　川　一　郎　」
  （**「印」を含めない** — この書式は印が別run）。
- **`kanzai_kibou`（項目1）は弁護士の手続選択そのもの**。case.yaml に無ければ
  値を渡さず、questions.md に「一般管財（招集型/非招集型）・個別管財のいずれを
  希望するか」を必ず載せる。
- 有無チェック（4〜8, 15, 17, 18, 20, 23, 25）: case.yaml の対応データから。
  **裏付けの無い「無」を書かない** — 資料で確認できないものは値を渡さず questions.md へ。
  - 4 相続財産: assets の相続関係／家族聴取
  - 5 訴訟: lawsuits[]（0236 一覧表と連動。有なら sosho_ichiran_check も）
  - 6 倒産直前の処分: bank_analysis / menseki.henpa / asset_disposal
    （有なら 0237 一覧表と tosan_ichiran_check）
  - 7 公租公課滞納: creditors の kind=公租公課（有なら 0216 と kouso_ichiran_check）
- 9 居住物件: household.housing（賃貸→自己所有でない）。自己所有なら明渡し見込みを
  弁護士に確認。
- 10〜12 郵便取扱局: **住所地の集配郵便局**。下線は局名のみ（「堺」など）。
  不明なら空欄＋questions.md（管財人への郵便転送に使われる重要欄）。
- 13 自由財産拡張: 財産目録の■欄・自由財産拡張申立書（0240）と整合させる。
  99万円超なら別紙が必要（jiyuzaisan_gaku_besshi）。
- 15〜25 は**個人事業者（現在または申立て前6か月以内）のみ**。非該当なら一切値を
  渡さない。business セクションから記入し、解雇日・社会保険等の手続状況は
  資料（離職票控え等）で確認できたものだけチェック。
- 日付3スロット（解雇通知・解雇予定・手形不渡・閉店廃業）はリスト ["７", "２", "２８"]。
- 21・22 の回収困難な理由、24 の解約未了の場所・番号は書式に下線が無く
  機械記入できない → 該当時は完了報告で手書き箇所として案内。

## 注意

- 事件番号欄・宛先は付番後に手書き/再生成で補う運用。
- 報告書（0205）・申立書（0203）・財産目録・債権者一覧表と記載の整合をとる
  （特に 5〜7 の有無と各一覧表の添付有無）。
