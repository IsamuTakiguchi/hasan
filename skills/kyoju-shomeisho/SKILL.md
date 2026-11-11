---
name: kyoju-shomeisho
description: >-
  破産申立てで申立人が親族・知人所有の家屋に無償で住み賃貸借契約書がないとき、所有者の居住証明書（署名押印用）を事務所書式から作り、事件データと確認事項に反映する。
  「居住証明書を作って」「賃貸借契約書がない」「義父の家に住んでいる」「親族所有の家に無償で住んでいる」
  「知人の家に住んでいる場合の書類」「住居の疎明資料」などのフレーズがトリガー。
---

# 賃貸借契約書がないときの居住証明書（破産申立て）

> パス表記: この手順は事件フォルダ方式（`受領資料/`・`作業ファイル/`・`申立書類/`・`システム/`）で書いている。
> 作業フォルダ方式・ソースリポジトリでは `cases/<事件ID>/` 配下の input/・work/・output/、`hasan/`（プラグインルート）に読み替える（対応表は事件フォルダの CLAUDE.md）。

申立人が義父・親族・知人の所有する家屋に無償で住んでいて賃貸借契約書がない場合に、所有者作成の居住証明書で代替する。
（大阪地裁B系列 同時廃止チェックリスト：「友人の賃借マンションに無償で住んでいるようなときは，その友人名義の賃貸借契約書と同人作成の居住証明書を提出」）

## 前提・鉄則
- 事件フォルダの CLAUDE.md に従う。`システム/` は読み取り専用、依頼者情報はフォルダの外へ出さない。
- 書類の記入は `システム/scripts/hasan-kit fill-docx` で行う（Word を自作しない）。失敗したら止めて報告する。
- 作業前に `hasan-kit scan --dir .` で手修正・新資料がないか確認し、あれば先に反映する（先生の手入力を上書きしない）。

## 1. 事実の確認
- 建物の所有者・所在（地番）: 建物登記事項証明書
- 住居表示: 住民票（申立人の住所）
- 所有者と申立人の関係（義父等）、家賃の有無（無償か）、駐車場等の別契約の有無
- 所有者自身が賃借している物件なら、所有者名義の賃貸借契約書も別途必要（questions.md に追加）
- 不明点は推測で埋めず questions.md へ。

## 2. 事務所書式の準備
- 書式は先生から受け取った事務所書式（例: `受領資料/参考書式/居住証明書（事務所書式）.doc`）を使う。無ければ添付を依頼する。
- 原本は編集しない。作業用にコピーして docx 化し、段落に paraId を付ける（fill-docx のアンカー用）:
  ```bash
  soffice --headless --convert-to docx 書式.doc   # 作業用ディレクトリで
  python3 - <<'PY'
  import zipfile,re
  zin=zipfile.ZipFile('書式.docx'); zout=zipfile.ZipFile('居住証明書_template.docx','w',zipfile.ZIP_DEFLATED); n=0
  for it in zin.infolist():
      d=zin.read(it.filename)
      if it.filename=='word/document.xml':
          def rep(m):
              global n; n+=1
              return '<w:p w14:paraId="%08X" w14:textId="%08X"'%(0x4B590000+n,0x4B5A0000+n)
          d=re.sub(r'<w:p(?=[ >])',rep,d.decode()).encode()
      zout.writestr(it,d)
  zout.close()
  PY
  ```
- 出来た `作業ファイル/居住証明書_template.docx` を `hasan-kit inspect-docx <template> --full` で確認し、本文段落の paraId・idx・context を控える。

## 3. fillmap・values を作って記入
- `作業ファイル/kyoju_shomei_fillmap.yaml`: 本文段落を `kind: set_text` で1項目（anchor に paraId・path:''・idx・context を全部書く）。
- `作業ファイル/kyoju_shomei_values.yaml`: 本文は書式の文言を崩さず、氏名と所在だけ差し替える。
  - 例「  （申立人氏名）が、私所有の（所在地番 例：○○県○○市○○町○丁目○○番地○○）の建物（住居表示　○○県○○市○○町○丁目○番○号）に居住していることを証明いたします。」
  - 書式が地番表記なら地番で書き、建物を特定するため住居表示を括弧で添える。数字は全角。
  - 日付・所有者の住所・氏名（押印）は所有者の自署のため空欄のまま。
- 実行:
  ```bash
  export DOCX_SKILL_SCRIPTS=$HOME/docx_scripts   # 端末に合わせる（作業ファイル/README.md 参照）
  システム/scripts/hasan-kit fill-docx --template 作業ファイル/居住証明書_template.docx \
    --fillmap 作業ファイル/kyoju_shomei_fillmap.yaml --values 作業ファイル/kyoju_shomei_values.yaml \
    --output "申立書類/居住証明書_<申立人>（<所有者>様 署名押印用）.docx"
  システム/scripts/hasan-kit preview "申立書類/居住証明書_….docx" -d 作業ファイル/_prevN
  ```
- 画像で体裁が書式どおりか目視する（フォント崩れは pdf-layout-fix スキルを参照）。

## 4. 事件データ・確認事項への反映
- case.yaml `family.housing`: kind・rent 0・note（所有者・関係・所在・無償）に加え `lease_doc` に「賃貸借契約書はないため所有者の居住証明書で代替」と取付けの段取りを書く。sources に出所（先生指示・登記・住民票）を記録 → `hasan-kit validate-case`。
- 報告書第１の５「申立人以外の者（氏名，申立人との関係）」が未記入なら補助 fillmap（houkokusho_extra_fillmap.yaml 方式）で記入し、レイアウト検査を通す。
- questions.md: 「居住証明書の取付け（所有者の署名押印）」を要確認に、不足資料と「手書き箇所（日付・住所・氏名）」に追加。取り付けられない場合の代替（登記事項証明書＋報告書の説明）も書く。
- 聴取シート（hearing_spec.yaml）に取付け確認の項目を入れ、`hasan-kit memo-docx` で差分シートを作り直す。
- 作業ファイル/README.md に生成コマンドを残す。最後に `hasan-kit scan --dir . --write`。

## 報告
作った書面の場所、本文の文言（地番・住居表示のどちらで書いたか）、所有者に書いてもらう欄、残る確認事項を簡潔に伝える。
