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
- `release/業務ポータル_一括インストーラー_latest.vbs` — 最新HTML版とOutlook連係ファイルを内包する単体のVBSインストーラー。実行すると `%LOCALAPPDATA%\WorkPortal` に導入し、必要なら従来版Outlookの設定へ進みます。

最新HTMLソースには、KPIとタスク一覧を一体化したホーム、全件検索できる共有活動ログ、見やすさを高めたライト／ダークテーマが含まれます。`release/` のVBSはv10.1安定版です。

最新HTML版は画面上部のボタンでライト／ダークを切り替えられます。従来版Outlookを使うPCでは、`outlook/README.md` に沿って一度だけ端末用連係を設定すると、上部の「Outlookへ反映」から自分の未完了・期限ありタスクを予定表へ反映できます。Outlook予定は本人のメールボックス側にも保存され得ます。安定版の作成VBSにはこれらの機能は含まれません。

実行時は、各テーマの `index.html` と `icon.ico` を同じフォルダーに置き、Microsoft Edgeから開いてください。

一括インストーラーは現在のWindows利用者に導入します。既存のアプリ一式は別フォルダーに退避し、共有JSONとブラウザー内の個人データには手を加えません。共有JSONへの接続は画面から行い、Outlook連係の設定時は既存の共有JSONのフルパスを入力してください。生成元は `scripts/build_full_installer.py` です。Windows/Edge/Outlookでの実機確認は導入先で行ってください。

導入・Outlook連携の設定・同期結果などの確認ダイアログは日本語で表示します。
