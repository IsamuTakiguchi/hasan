# hasan — 破産申立資料AI自動作成システム（Claude Code プラグイン）

裁判所配布の書式（Word/Excel・非公開・随時改訂）に、事件資料から抽出した情報を
AIが転記して破産申立資料一式を作成する **Claude Code プラグイン**です。
スキル10個・生成エンジン（Python）・裁判所書式（版管理付き）・事件モデルスキーマを
1パッケージに同梱しています。

> **重要**: 裁判所の非公開書式（courts/）を同梱しているため、このリポジトリと
> マーケットプレイスは **private を維持**してください。第三者への配布・販売は
> エンジンと書式パックの分離＋匿名化（skill-anonymizer）が前提です。

## インストール

Claude Code / Cowork で:

```
/plugin marketplace add IsamuTakiguchi/hasan
/plugin install hasan@hasan-marketplace
```

ローカルのクローンから使う場合（書式登録もするメイン環境向け）:

```
git clone https://github.com/IsamuTakiguchi/hasan
/plugin marketplace add ./hasan
/plugin install hasan@hasan-marketplace
```

依存: Python3 + `pip install lxml openpyxl PyYAML jsonschema PyMuPDF`、
PDF検証に LibreOffice（writer/calc）。

## 使い方

事件用の**作業フォルダ**（プラグインとは別の場所）を作り、そこで Claude に依頼します:

```
~/bankruptcy-cases/           # 作業フォルダ（例）
├── .gitignore                # 「cases/」を必ず無視（依頼者情報）
└── cases/
    └── 2026-008-yamada/
        └── input/            # 調査票・通帳PDF・面談メモ等を置く
```

1. 「資料を読み込んで」（/hasan:intake）→ case.yaml と確認事項リストができる
2. 「通帳を分析して」（/hasan:nyushukkin-bunseki）→ 偏頗弁済等の検出
3. 「破産申立書類一式を作って」（/hasan:generate）→ 申立書・報告書・債権者一覧表
   （一般＋公租公課）・財産目録・家計収支表を裁判所書式で生成し、数値整合検査・
   PDF目視まで実施

個別書式のみの依頼も可: 「報告書を作って」「債権者一覧表を作って」等。
スクリプトを直接使う場合は `hasan-kit --help`（プラグインの bin/ が PATH に入ります）。

## 書式が改訂されたら

「新しい書式を登録して」（/hasan:register-form）。登録は**このリポジトリのクローン**で
行い（インストール先への書き込みは更新で消えるため不可）、コミット・プッシュ後、
配布する場合は `claude plugin tag --push` で版タグを付けます。各利用環境は
プラグイン更新で新書式を受け取ります。手順詳細は [docs/form-registration.md](docs/form-registration.md)。

## ドキュメント

- 開発者向け（ソースリポジトリでの作業規約）: [CLAUDE.md](CLAUDE.md)
- 書式の登録・更新手順: [docs/form-registration.md](docs/form-registration.md)
- 将来のWebアプリ化メモ: [docs/phase2-webapp-notes.md](docs/phase2-webapp-notes.md)
