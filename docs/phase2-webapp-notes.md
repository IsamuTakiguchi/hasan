# Phase 2（Webアプリ化）に向けたメモ

Phase 1（本リポジトリ）は Claude Code / Cowork 上のスキル＋スクリプト構成だが、
判断ロジックを次の**機械可読資産**に寄せてあるため、Webアプリ（ぱらりこさん型の
アップロード→ドラフト生成UI）への載せ替えは資産の再利用で済む。

## そのまま再利用できる資産

- `schema/case.schema.yaml` — 事件モデルの契約。フロントの入力フォーム／APIの型定義の源泉
- `courts/**/fillmap.yaml` + `template.*` — 書式資産。サーバーからそのまま参照
- `scripts/build/fill_docx.py` / `xlsxlib.py` / `build_*.py` — 生成エンジン。CLI を
  そのまま関数呼び出しに変えるだけ（依存: lxml, openpyxl, PyYAML）
- `scripts/formkit/*` — 書式登録・更新パイプライン（管理画面の「書式更新」機能になる）
- `scripts/verify/*` — 検証（soffice + PyMuPDF が動くコンテナが必要）

## AI が担っている部分（API 呼び出しに置換する箇所）

- intake: 原資料 → case.yaml（Claude API・構造化出力。sources 必須のプロンプトは
  hasan-intake/SKILL.md を流用）
- 報告書の values.yaml 組み立て（体裁ルールは hasan-houkokusho/SKILL.md がそのまま
  システムプロンプトになる）
- 入出金分析の候補吟味（機械検出 bank_analyze.py の後段）

## アーキテクチャ上の留意

- 「弁護士が必ず最終確認する」フローを UI に組み込む（questions.md 相当の確認キュー）
- 依頼者情報の分離: cases/ 相当は事件ごとに暗号化ストレージへ。書式（courts/）とは
  ライフサイクルを分ける
- 書式は非公開資産のためテナント外に出さない。AI 推論に渡すのは必要最小限
- 学習利用不可の API 契約（Anthropic の商用 API は入力を学習に使わない）を明記
