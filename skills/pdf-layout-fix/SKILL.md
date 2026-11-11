---
name: pdf-layout-fix
description: >-
  破産申立書類などWordでは収まっているのにPDF（hasan-kit preview・統合版PDF）でレイアウトが崩れるとき、原因のフォント代替を直してPDFを作り直す。
  「PDFが崩れている」「PDFだと2頁になる」「Wordでは1頁なのに」「プレビューのレイアウトがおかしい」
  「行間が広がる」「統合版PDFを作り直して」などのフレーズがトリガー。
---

# PDFのレイアウト崩れの修正（フォント代替）

> パス表記: この手順は事件フォルダ方式（`受領資料/`・`作業ファイル/`・`申立書類/`・`システム/`）で書いている。
> 作業フォルダ方式・ソースリポジトリでは `cases/<事件ID>/` 配下の input/・work/・output/、`hasan/`（プラグインルート）に読み替える（対応表は事件フォルダの CLAUDE.md）。

## 症状と原因
- Word では1頁に収まる申立書が、PDFでは2頁にはみ出す／行間が広がる／印紙欄・受領印欄が2頁目にも出る。
- 原因はほぼ常に、作業用の仮想環境（Linux・LibreOffice）に ＭＳ明朝・ＭＳゴシック・Century が無く、Noto Serif CJK・DejaVu 等で代わりに描いていること（字の高さ・幅が大きく行が膨らむ）。Word ファイル自体は壊れていない。

## 1. 診断
```bash
python3 -c "import pymupdf,sys;d=pymupdf.open(sys.argv[1]);print(len(d),'頁');print([f[3] for f in d[0].get_fonts()])" 対象.pdf
fc-match "ＭＳ 明朝"; fc-match Century
```
Noto・DejaVu が出ていればフォント代替が原因。

## 2. 字幅互換フォントを入れる（仮想環境のホームに・sudo不要）
```bash
mkdir -p $HOME/fontdl && cd $HOME/fontdl
apt-get download fonts-ipafont-mincho fonts-ipafont-gothic
for d in *.deb; do dpkg-deb -x "$d" ex; done
mkdir -p $HOME/.local/share/fonts && find ex -name '*.ttf' -exec cp {} $HOME/.local/share/fonts/ \;
mkdir -p $HOME/.config/fontconfig && cat > $HOME/.config/fontconfig/fonts.conf <<'XML'
<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <dir>~/.local/share/fonts</dir>
  <alias binding="same"><family>ＭＳ 明朝</family><prefer><family>IPAMincho</family></prefer></alias>
  <alias binding="same"><family>MS Mincho</family><prefer><family>IPAMincho</family></prefer></alias>
  <alias binding="same"><family>ＭＳ Ｐ明朝</family><prefer><family>IPAPMincho</family></prefer></alias>
  <alias binding="same"><family>ＭＳ ゴシック</family><prefer><family>IPAGothic</family></prefer></alias>
  <alias binding="same"><family>MS Gothic</family><prefer><family>IPAGothic</family></prefer></alias>
  <alias binding="same"><family>ＭＳ Ｐゴシック</family><prefer><family>IPAPGothic</family></prefer></alias>
  <alias binding="same"><family>Century</family><prefer><family>C059</family></prefer></alias>
</fontconfig>
XML
fc-cache -f; fc-match "ＭＳ 明朝"   # → IPAMincho になること
```
- apt が使えない環境では、その旨を報告し、Word からのPDF出力を案内する（別ルートでの取得を試み続けない）。

## 3. PDFを作り直して確認
- `hasan-kit preview <書類...> -d 作業ファイル/_prevN`（既存の _preview は削除できないことがあるので新しいフォルダへ）。
- `hasan-kit layout-check --dir 申立書類 --case case.yaml` で白紙テンプレと頁数が一致することを確認。
- 申立書は1頁、各書類の頁数を先生のWord表示と突き合わせ、崩れの出ていた頁は画像で目視する。
- 古い崩れたプレビューは削除せず `作業ファイル/_old_preview_<日付>/` へ移し、新しいPDFを元の場所へコピーする（削除は先生の許可が要る）。

## 4. 統合版PDF（書類一式を1つにまとめたPDF）の場合
- まず統合版がどう作られたか（どのPDF・画像をどの順で結合したか、作業ファイル/README.md や結合スクリプト）を確認する。
- 構成PDFを手順3で作り直し、同じ順序で pymupdf（`doc.insert_pdf`）で結合し直す。画像資料はA4頁に貼り込む。
- 結合後、全頁の頁数・順序・崩れを確認し、旧版は退避する。
- 別事件のフォルダはこの会話で接続されていないと開けない。`device_request_folder_access` がフォルダ名の全角スペース等で使えない場合は、Claudeデスクトップアプリの「＋」→「フォルダを追加」で追加してもらう。

## 5. 記録と注意
- 作業ファイル/README.md に「プレビューのフォント設定」と再設定手順を書き残す（仮想環境が作り直されると設定は消える）。
- IPA明朝は字幅互換だが本物のＭＳ明朝ではない。**裁判所に出すPDFは Word の「名前を付けて保存 → PDF」で作るのが確実**、と必ず添える。
- 最後に `hasan-kit scan --dir . --write`。
