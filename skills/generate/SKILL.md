---
name: generate
description: >-
  破産申立資料一式（債権者一覧表・報告書・資産目録・家計収支表）を事件モデル
  case.yaml から一括生成するオーケストレータスキル。個別書式のスキル
  （houkokusho 等）を正しい順序で実行し、書類間の数値整合検査と
  PDF目視検証まで行う。「破産申立書類一式を作って」「申立資料をまとめて作成して」
  「破産の書類を全部作って」「一式生成して」などのフレーズがトリガー。
  破産申立資料をまとめて作成する場面では必ずこのスキルを使用すること。
---

# 破産申立資料 一括生成スキル（オーケストレータ）

## 前提

- `cases/<事件ID>/case.yaml` が存在すること（なければ先に intake）。
  `cases/` は**現在の作業フォルダ**側（プラグインのインストール先には置かない）。
- **手続種別で書式セットを解決する**: `meta.proc_type`（同時廃止／管財（自然人）／
  管財（法人））を読み、`$(hasan-kit courts)/osaka/procs.yaml` の該当セットの
  forms（生成順・via・when 条件）に従って生成する。下の「実行順序」は
  **同時廃止セット**の展開形。
  - proc_type 未設定 → intake に戻って利用者に確認（勝手に既定しない）
  - 該当セットの forms が空（**書式が未登録**）→ 生成せず、裁判所配布の
    書式一式の提供と register-form での登録を案内して終了する
  - when 条件の判断: 「公租公課の債権者あり」= creditors に kind=公租公課 が存在／
    「個人事業者（過去含む）」= business.exists が true 又は申立書参考事項1が事業者／
    「リース債権あり」= creditors に kind=リース／「lawsuits または
    pre_bankruptcy_disposals あり」= 各配列が非空／「自由財産拡張の申立てをする場合」=
    assets に jiyuzaisan: true の項目がある（無ければ弁護士に確認）
- **環境準備**: `hasan-kit` が PATH に無い環境（Cowork 等）では、プラグインの
  インストールディレクトリを探して絶対パスで使う
  （例: `ls ~/.claude/plugins/*/*/scripts/hasan-kit` や `find ~/.claude -name hasan-kit`）。
  初回は `hasan-kit doctor` で環境診断（依存パッケージは初回実行時に自動導入される）。
  PDF変換ツールが無い環境では preview が目視省略の劣化運転になる（crosscheck は動く）。
- 通帳があるのに `bank_analysis` が空なら、先に nyushukkin-bunseki を促す
  （免責関係の記載が変わり得るため）。

## 実行順序（同時廃止セット）

数値の親（債権者・資産・家計の生データ）→ 子（それを引用する申立書・報告書）の順で作る。
（管財セットが登録されたら procs.yaml の順序に従う。原則は同じ「親→子」）

1. **生成前ゲート**: `hasan-kit validate-case cases/<事件ID>/case.yaml`
   - エラーなら中断して利用者に報告。警告（出所なし等）は控えて最後にまとめる。
   - 表示される負債総額・資産総額・月次収支を控える（後の突合の基準）。
2. **債権者一覧表**: `hasan-kit saikensha`（一般用）。公租公課の債権者がいる
   場合は続けて `--kouso` で公租公課用（B1106）も生成。
3. **資産目録**: `hasan-kit shisan`。
4. **家計収支表**: `hasan-kit kakei`。
5. **申立書**: moushitatesho の手順で values を組み立てて `hasan-kit fill-docx`。
   申立ての理由の債権者数・総額・除外後額は 2 の数値と一致させる。
6. **報告書**: houkokusho の手順で values を組み立てて `hasan-kit fill-docx`。
   第２・第３の記載が 2〜4 の数値・bank_analysis と矛盾しないよう組み立てる。
7. **事業に関する報告書**: 個人事業者（申立書参考事項1が事業者）の場合のみ、
   jigyou-houkokusho の手順で生成。
8. **書類間整合**:
   `hasan-kit crosscheck --case ... --dir cases/<事件ID>/output`
   が全 OK になること。NG が出たら原因（case.yaml と生成物のどちらが古いか）を
   特定して作り直す。生成物を手で直して辻褄を合わせてはならない。
9. **目視検証**: `hasan-kit preview cases/<事件ID>/output/*.docx
   cases/<事件ID>/output/*.xlsx` で全書類をPDF・画像化し、各スキルのチェックリストで
   目視する。
10. 途中で失敗した書式があっても残りは続行し、最後にまとめて報告する。
   （チェックリストに関する上申書 B1113 は、チェックできない項目があると弁護士が
   判断したときのみ指示を受けて作成する）

## 管財（自然人）セットの追加手順

procs.yaml の kanzai-shizenjin の forms 順で生成する。同時廃止との違い:

1. xlsx は `hasan-kit kanzai-xlsx --doc <書式>`（債権者一覧表8シート・被課税公租公課
   チェック表・財産目録17シート・資産及び負債一覧表・リース物件・訴訟/処分一覧表）。
   版・書式は case.yaml の proc_type から自動で管財版が選ばれる。
2. docx は moushitatesho（0203）→ houkokusho（0205版）→ hoju-houkokusho（0204）の
   順で values を組み立てて fill-docx。数値は xlsx の meta.json と一致させる
   （crosscheck が 資産計・負債計・回収見込計・申立書の数値を突合する）。
3. **目録3点は最後に**（生成結果が確定してから）:
   - tenpu-mokuroku（0207）: 実際に生成・添付した書類だけ「添付」を values に渡す
   - somei-shiryo-mokuroku（0238）: 添付する疎明資料に応じて◇○（通数は手書き案内）
   - hikitsugi-shiryo-ichiran（0239）: 財産目録で「あり」の区分＝◇、引継資料の□は
     **弁護士の判断事項** — 未確認分は questions.md へ
4. jiyuzaisan-kakucho（0240）は自由財産拡張をする場合のみ（財産目録の■・
   管財補充報告書13項と3点セットで整合させる）。
5. **手書き案内を最終報告に必ず含める**: 報告書（0205）の宛先☑（本庁の□は
   箇条書き記号のため機械記入不可）／0203のふりがな（ルビ）／疎明資料目録の通数／
   支部名の記入。

## 出力命名

`cases/<事件ID>/output/<書式名>_<申立人姓>[_記入済み].docx|.xlsx`
（例: `債権者一覧表_甲野.xlsx`, `報告書_甲野_記入済み.docx`）

## 最終報告（必須）

- 生成した書類と各書類の要点（債権者数・負債総額・資産総額・月次収支）
- 空欄のまま残した項目とその理由（根拠資料なし等）
- 推測を含む記載・資料間の矛盾 → questions.md の確認事項（新規分を明示）
- 裁判所書式が未登録で --generic（事務所内ドラフト様式）になった書類
  → 裁判所書式の入手と register-form での登録を案内
- 依頼者への確認・資料依頼が必要な項目があるとき → line-message スキルで
  LINE文面を作れる旨を案内（次回面談前の準備は hearing-sheet の差分シート）
