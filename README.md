# hasan — 破産申立資料AI自動作成システム（Claude Code プラグイン）

裁判所配布の書式（Word/Excel・非公開・随時改訂）に、事件資料から抽出した情報を
AIが転記して破産申立資料一式を作成する **Claude Code プラグイン**です。
同時廃止・管財（自然人）・管財（法人）の3書式セット、スキル15個・生成エンジン
（Python）・裁判所書式（版管理付き）・事件モデルスキーマを1パッケージに同梱しています。

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

最初に**手続種別**を選びます（intake が確認します）:
**同時廃止**／**管財（自然人）**／**管財（法人）**。以後の聴取シート・生成書式は
選んだ種別のセット（`courts/osaka/procs.yaml`）で動きます。
**3セットとも書式登録済み**です — 同時廃止（B1102〜B1114の10書式）、
管財・自然人（0201〜0240 の一式: 申立書・報告書・管財補充報告書・債権者一覧表8シート・
財産目録17シート・資産及び負債一覧表・各種目録・自由財産拡張申立書）、
管財・法人（0101〜0136 の一式: 法人用申立書・報告書・財産目録16シート等）。
同じ書式でも手続種別で配布版が違うもの（報告書・家計収支表等）は自動で正しい版が選ばれます。

0. 「聴取シートを作って」（/hasan:hearing-sheet）→ 面談用ヒアリングシート
   （2回目以降は未確認事項だけの差分シート）。面談後は「面談メモにまとめて」
   （/hasan:menmemo）で走り書きを清書 → そのまま原資料に
1. 「資料を読み込んで」（/hasan:intake）→ case.yaml と確認事項リストができる
2. 「通帳を分析して」（/hasan:nyushukkin-bunseki）→ 偏頗弁済等の検出
3. 「破産申立書類一式を作って」（/hasan:generate）→ 申立書・報告書・債権者一覧表
   （一般＋公租公課）・財産目録・家計収支表を裁判所書式で生成し、数値整合検査・
   PDF目視まで実施
4. 「LINEで送る文面を作って」（/hasan:line-message）→ 不足資料・確認事項を
   依頼者向けの平易な文面に（送信は手動）

個別書式のみの依頼も可: 「報告書を作って」「債権者一覧表を作って」等。
スクリプトを直接使う場合は `scripts/hasan-kit --help`（インストール先の
`scripts/hasan-kit` を絶対パスで呼ぶ。スキルは自動で実体を探します）。

## Claude Cowork で使う

Cowork（デスクトップアプリ）でも同じプラグインが使えます。

1. 配布zipを作る（このリポジトリのクローンで）: `scripts/hasan-kit package`
   → `hasan-plugin-v1.0.0.zip` ができる
2. Cowork タブ → **Customize** → **Install** → zipを**アップロード**
   （個人プランは private GitHub マーケットプレイスから直接インストールできないため、
   zipアップロードが正規ルート。Team/Enterprise は管理画面の Private Marketplace 連携で
   private リポジトリから配布可能）
3. 事件フォルダは Cowork の「フォルダを接続」で作業フォルダ（cases/ を含む）を渡す

Cowork の実行環境（Linux VM）には Python・LibreOffice が入っており、不足する
pip パッケージは hasan-kit が初回実行時に自動導入します（診断は `hasan-kit doctor`）。

既知の注意:
- タブ間（Chat/Cowork/Code）のスキル同期は遅延あり。反映されないときはアプリ再起動
- `hasan-kit` が PATH に無い環境では、スキルが
  `~/.claude/plugins/.../scripts/hasan-kit` を絶対パスで探すフォールバックを内蔵
  （claude.ai ホスト型プラグインは bin/ 同梱不可のため scripts/ に配置している）
- 書式改訂時は plugin.json の version を上げて zip を再アップロード

## 書式が改訂されたら

「新しい書式を登録して」（/hasan:register-form）。登録は**このリポジトリのクローン**で
行い（インストール先への書き込みは更新で消えるため不可）、コミット・プッシュ後、
配布する場合は `claude plugin tag --push` で版タグを付けます。各利用環境は
プラグイン更新で新書式を受け取ります。手順詳細は [docs/form-registration.md](docs/form-registration.md)。

## ドキュメント

- 開発者向け（ソースリポジトリでの作業規約）: [CLAUDE.md](CLAUDE.md)
- 書式の登録・更新手順: [docs/form-registration.md](docs/form-registration.md)
- 将来のWebアプリ化メモ: [docs/phase2-webapp-notes.md](docs/phase2-webapp-notes.md)
