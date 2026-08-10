# 破産事件フォルダ（hasan・1事件用）

このフォルダは **1つの破産事件の作業フォルダ**。破産事件に関する依頼を受けたら、
**必ず `hasan/skills/` の該当スキル（SKILL.md）を読み、その手順に厳密に従うこと**。
スキルを読まずに自己流で書類を作ることを禁止する。

## このフォルダの構成（スキルのパス読み替え）

事件データは**このフォルダ直下**に置く。スキル内の `cases/<事件ID>/` という表記は、
すべて**このフォルダ直下**に読み替えること:

| スキル内の表記 | このフォルダでの実体 |
|---|---|
| cases/<事件ID>/input/ | ./input/（原資料: 調査票・通帳PDF・面談メモ等） |
| cases/<事件ID>/case.yaml | ./case.yaml（事件モデル＝唯一の真実） |
| cases/<事件ID>/questions.md | ./questions.md（弁護士への確認事項） |
| cases/<事件ID>/work/ | ./work/（中間物） |
| cases/<事件ID>/output/ | ./output/（記入済み書面＋検証用PDF/画像） |

## 依頼 → 使うスキル（hasan/skills/<名前>/SKILL.md）

| 依頼の例 | スキル |
|---|---|
| 破産申立書類一式を作成して／全部作って | **generate**（オーケストレータ。最初にこれを読む） |
| 資料を読み込んで／事件を取り込んで | intake |
| 聴取シート・ヒアリングシートを作って | hearing-sheet |
| 面談メモにまとめて・清書して | menmemo |
| 通帳を分析して | nyushukkin-bunseki |
| 申立書を作って | moushitatesho |
| 報告書を作って（自然人） | houkokusho ／ 管財補充報告書は hoju-houkokusho |
| 報告書を作って（法人） | houkokusho-hojin |
| 事業に関する報告書 | jigyou-houkokusho |
| 債権者一覧表・財産目録・家計収支表（個別） | saikensha-ichiran / shisan-mokuroku / kakei-shushi |
| LINEで送る文面を作って | line-message |

## 鉄則（スキルより先に必ず把握する）

1. **裁判所書式の出力は必ずテンプレートと同じ形式**（Excel書式→.xlsx／Word書式→.docx）。
   債権者一覧表・財産目録・資産及び負債一覧表・家計収支表・各種チェック表は **Excel**、
   申立書・報告書・目録類は **Word**（各書式の形式は
   `hasan/courts/osaka/forms/<書式>/registry.yaml` の doc_type）。
2. **生成は必ず `hasan/scripts/hasan-kit` 経由**（builder / fill-docx）。
   Word・Excel・Markdown を自作して代替することは禁止（裁判所書式でなくなる）。
   面談メモ・聴取シートも `hasan-kit memo-docx` による **Word（.docx）** で作る。
3. hasan-kit が失敗したら**書類を自作せず停止**し、エラー全文と
   `hasan/scripts/hasan-kit doctor` の結果を利用者に報告する。
4. 手続種別（同時廃止／管財（自然人）／管財（法人））を最初に確認する（intake 参照）。
5. 生成後は `hasan-kit crosscheck --case case.yaml --dir output`（数値・出力形式の監査）と
   `hasan-kit preview`（PDF目視）まで必ず実施する。
6. このフォルダは依頼者情報を含む。フォルダの外へ複製しない・共有しない。

## 環境メモ

- hasan-kit の実体: `hasan/scripts/hasan-kit`（このフォルダからの相対パス。
  依存パッケージは初回実行時に自動導入。診断は `hasan/scripts/hasan-kit doctor`）
- hasan/ はエンジン・書式一式（読み取り専用扱い。編集しない。書式改訂は
  ソースリポジトリ側で行い、hasan/ を差し替える）
