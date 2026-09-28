# 業務ポータル

共有ファイルサーバー上の共有JSONを中心に、少人数で利用する業務管理ツールです。

## 配布物

- `source/light/` — 最新のライト版HTMLソース
- `source/dark/` — 最新のダーク版HTMLソース
- `release/` — v10.1ライト版・ダーク版の作成VBSと配布用ZIP
- `docs/HANDOVER_DESIGN_V10.1.md` — 保守・改修を別プロジェクトへ引き継ぐための設計書
- `checksums/SHA256SUMS.txt` — 配布物のSHA-256値
- `outlook/` — 従来版Outlook向けの端末用連係ヘルパー、初回インストーラー、削除用VBS、導入手順
- `release/Outlook連携インストーラー_latest.zip` — 上記の端末用ファイルをまとめたZIP
- `release/業務ポータル_latest_HTML_導入一式.zip` — 最新ライト／ダークHTML、サンプル、ドキュメント、Outlook連係をまとめた導入用ZIP

最新HTMLソースには、KPIとタスク一覧を一体化したホーム、全件検索できる共有活動ログ、見やすさを高めたライト／ダークテーマが含まれます。`release/` のVBSはv10.1安定版です。

最新HTML版は画面上部のボタンでライト／ダークを切り替えられます。従来版Outlookを使うPCでは、`outlook/README.md` に沿って一度だけ端末用連係を設定すると、上部の「Outlookへ反映」から自分の未完了・期限ありタスクを予定表へ反映できます。Outlook予定は本人のメールボックス側にも保存され得ます。安定版の作成VBSにはこれらの機能は含まれません。

実行時は、各テーマの `index.html` と `icon.ico` を同じフォルダーに置き、Microsoft Edgeから開いてください。
