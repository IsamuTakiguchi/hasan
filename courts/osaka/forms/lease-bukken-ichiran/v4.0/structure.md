# リース物件一覧表（0235＝0132）v4.0 — 書式構造メモ

1シート・A1:H20。B債権者名/Cリース物件/D所在地/E契約書等□/F返還済□。
リース債権がある場合には必ず提出。creditors（kind=リース）の lease 情報から生成
（build_kanzai_xlsx --doc lease）。**法人用0132とセル単位で完全同一**のため
自然人・法人で共用（registry procs: [kanzai-shizenjin, kanzai-hojin]）。
債権者一覧表213リース債権シート・管財補充報告書20（取戻権）と整合させる。
