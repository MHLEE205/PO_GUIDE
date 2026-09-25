# POFaxHelper（P/O FAX自動入力ツール）

PO_GUIDEで作成したPDFを複合機のFAXプリンターへ送り、「ファクス送信の設定／確認」画面に
FAX番号・宛先名を自動入力する常駐ツール。**「送信開始」は利用者が内容を確認して押す**（自動では押さない）。

## 利用者向け
1. PO_GUIDEのFAX欄にある「ダウンロード」から `POFaxHelper.exe` を取得して実行（初回のみ）
   - 以後はWindows起動時に自動起動し、タスクトレイに常駐（終了・自動起動の解除はトレイメニュー）
   - SmartScreenの警告が出た場合は「詳細情報」→「実行」
2. PO_GUIDEで「P/O自動作成」→ FAX欄に「🟢 FAX自動入力ツール接続中」と出ていれば
   「📠 FAX自動入力」を押す → FAX画面が開き番号・宛先名が入力される → 確認して「送信開始」

## 管理者向け
- 新しい形のFAX画面への対応: `profiles.json` に画面タイトルと入力手順を追加（exe再配布不要）
- exeのビルド: `fax_helper.py` の `VERSION` を上げてコミットし、`fax-helper-v<VERSION>` タグをpush
  → GitHub Actions（`.github/workflows/fax-helper.yml`）がビルドしてReleasesに公開
- ログ: `%APPDATA%\POFaxHelper\helper.log`、設定（選択プリンター）: `%APPDATA%\POFaxHelper\config.json`
- 仕様元: `sample/FAX送信実装マニュアル.docx`
