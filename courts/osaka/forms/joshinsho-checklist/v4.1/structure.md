# チェックリストに関する上申書（B1113）v4.1 — メモ

「破産同時廃止申立てチェックリストを確認しました」の定型上申書（15段落）。
本文（idx14）は印字済みで記入不要。記入欄は日付・宛先・代理人名・自由記載の下線1行
（idx11。チェックできない項目がある場合の補足等に使う。内容は弁護士が生成時に指示）。

生成: fill_docx.py + fillmap で記入（専用スキルなし）。例:
  python3 scripts/build/fill_docx.py \
    --template courts/osaka/forms/joshinsho-checklist/v4.1/template.docx \
    --fillmap  courts/osaka/forms/joshinsho-checklist/v4.1/fillmap.yaml \
    --values   <values.yaml> --output <出力先>
