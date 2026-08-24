# 破産申立資料 作業フォルダ（hasan）

このフォルダは **hasan（破産申立資料AI自動作成システム）の作業フォルダ**。
破産事件に関する依頼を受けたら、**必ず `hasan/skills/` の該当スキル（SKILL.md）を
読み、その手順に厳密に従うこと**。スキルを読まずに自己流で書類を作ることを禁止する。

## 依頼 → 使うスキル（hasan/skills/<名前>/SKILL.md）

| 依頼の例 | スキル |
|---|---|
| （指示が具体的でない・続き・資料や記入済み聴取シートを追加した） | **update**（差分検知→質問→反映。既定の入口） |
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
| 新しい書式を登録して | register-form |
| LINEで送る文面を作って | line-message |

## 鉄則（スキルより先に必ず把握する）

1. **裁判所書式の出力は必ずテンプレートと同じ形式**（Excel書式→.xlsx／Word書式→.docx）。
   どの書式が何形式かは `hasan/courts/osaka/forms/<書式>/registry.yaml` の doc_type。
   債権者一覧表・財産目録・資産及び負債一覧表・家計収支表・各種チェック表は **Excel**、
   申立書・報告書・目録類は **Word**。
2. **生成は必ず `hasan/scripts/hasan-kit` 経由**（builder / fill-docx）。
   Word・Excel・Markdown を自作して代替することは禁止（裁判所書式でなくなる）。
   面談メモ・聴取シートも `hasan-kit memo-docx` による **Word（.docx）** で作る。
3. hasan-kit が失敗したら**書類を自作せず停止**し、エラー全文と
   `hasan/scripts/hasan-kit doctor` の結果を利用者に報告する。
4. 事件データは**このフォルダの `cases/<事件ID>/`** に置く（hasan/ の中には置かない）。
   case.yaml が唯一の真実。生成器が原資料を直接読むことは禁止。
5. 生成後は `hasan-kit crosscheck`（数値・出力形式の監査）と `hasan-kit preview`
   （PDF目視）まで必ず実施する。
6. input/ の資料でファイル名が分かりにくいもの（記号・数字だけ・機器既定名）は、
   内容確認後に内容が分かる名前へリネームし、`work/資料名変更履歴.md` に
   元の名前を記録する（規約は intake スキル参照）。

## 環境メモ

- hasan-kit の実体: `hasan/scripts/hasan-kit`（このフォルダからの相対パス。
  依存パッケージは初回実行時に自動導入。診断は `hasan/scripts/hasan-kit doctor`）
- 手続種別（同時廃止／管財（自然人）／管財（法人））を最初に確認する（intake 参照）。
  書式セットは `hasan/courts/osaka/procs.yaml`。
- cases/ は依頼者情報を含む。このフォルダの外へ複製しない。
