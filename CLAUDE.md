# CLAUDE.md — PO_GUIDE プロジェクト

## プロジェクト概要
- **ツール名**: P/O申請ガイド
- **GitHub**: https://github.com/MHLEE205/PO_GUIDE
- **公開URL**: https://mhlee205.github.io/PO_GUIDE/
- **用途**: 社内向けピックアップオーダー申請ガイド（船社・港別）

## Azure / COMPASS 設定
- **アプリ登録名**: LEENAI-ShippingSchedule
- **Client ID**: f071a165-5e9b-44bf-b6a1-baba654524db
- **Tenant ID**: a56988ff-c2b8-4df8-9727-889ad4205198
- **Dataverse URL**: https://orgde512c6f.crm7.dynamics.com
- **Redirect URI**: https://mhlee205.github.io/PO_GUIDE/
- **認証方式**: PKCE OAuth 2.0 (ライブラリなし、純粋JS)

## GitHub設定
- **リポジトリ**: MHLEE205/PO_GUIDE
- **ブランチ**: main
- **PAT**: (ローカル環境の認証情報マネージャ等で別途管理。このファイルには記載しない)
- **PAT スコープ**: repo / workflow

## ファイル構成
```
PO_GUIDE/
└── index.html   # サイト本体（HTML/CSS/JS全て含む）
```

## 主要機能
1. **COMPASS BKG検索**: BKG Noを入力 → 船社・POL自動取得
2. **カードハイライト**: 該当船社・港のカードを自動ハイライト
3. **ピックアップ希望日自動計算**: CY OPEN前1営業日（土日祝遡及）
4. **港フィルター**: 大阪 / 神戸 / すべて
5. **船社名検索**: キーワードでカード絞り込み
6. **SharePointリンク**: 各船社フォームへ直リンク

## 対応船社一覧
### 大阪港 (OSAKA)
| 船社 | 方法 | 宛先 |
|---|---|---|
| YANG MING | メール | export_osaka_c9@mitsui-koun.co.jp |
| INTERASIA | メール | osaka_empty@kamigumi.co.jp |
| HEUNG A | FAX | 06-6612-6581 (辰巳商会) |
| KMTC | FAX | 06-6612-6581 (辰巳商会) |
| SINOKOR | FAX | 06-6612-6581 (辰巳商会) |
| NAMSUNG | FAX | 06-6612-6581 (辰巳商会) |
| DONG YOUNG | FAX | 06-6612-6581 (辰巳商会) |
| PAN OCEAN | FAX | 06-6612-6581 (辰巳商会) |

### 神戸港 (KOBE)
| 船社 | 方法 | 宛先 |
|---|---|---|
| CNC / CMA CGM | FAX | 078-306-3920 |
| KMTC | FAX | 078-230-6108 (NX日本通運) |
| SINOKOR | FAX | 078-304-1227 (住友倉庫KICT) |
| NAMSUNG | FAX | PO EXCEL.xlsx参照 |
| DONG YOUNG | FAX | PO EXCEL.xlsx参照 |
| TS LINE | メール | PO EXCEL.xlsx参照 |

## SharePointフォルダ構成
```
PICK UP/
├── OSAKA/    # 大阪港フォームPDF・Excel
├── KOBE/     # 神戸港フォームPDF・Excel
└── PO EXCEL.xlsx  # 船社設定マスタ
```
※ SharePointリンクは index.html 内の各カードのhrefを直接編集

## コード修正時の注意事項
- `REDIRECT_URI` は `https://mhlee205.github.io/PO_GUIDE/` 固定
- `CLIENT_ID` / `TENANT_ID` は変更禁止
- `$expand` は絶対使用しない → bare query + 個別マスター照会
- カード追加時は `data-port` / `data-carrier` / `data-carrier-key` を必ず設定
- 新船社追加後は `highlightCards()` のマッチングロジックを確認

## GitHub push方法
```bash
# ファイル編集後、以下でpush
git add index.html
git commit -m "修正内容を記述"
git push origin main
```

## 作成日
2026-07-02

✦ Powered by LEENAI Automation System | LEENEAR CORPORATION
