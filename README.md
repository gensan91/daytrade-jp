# daytrade-jp

日本株ギャップアップ・デイトレードの半自動化システム（Excel VBA + マーケットスピードII RSS）。

現在は**紙トレード検証フェーズ**。ライブ運用前に統計的にクリーンなデータセットを作ることが目的。

## 現在地

→ **[STATE.md](STATE.md)** を読むこと。ここが唯一の現在地。

## 構成

| ファイル | 役割 | 更新タイミング |
|---|---|---|
| `STATE.md` | 現在地。毎回上書き | 取引日の終了後 |
| `KNOWLEDGE.md` | 技術知見。追記のみ | 新しい実測が取れたとき |
| `src/*.bas` | 標準モジュール（VBEからエクスポート） | VBA を変更した直後 |
| `src/ThisWorkbook.cls` | ドキュメントモジュール（**貼り付け専用**） | 同上 |
| `src/frmController.frm.txt` | UserForm のイベントコードのみ | 同上 |
| `data/trades.csv` | Trades シートの書き出し（22列） | 取引日の終了後 |
| `デイトレ用.xlsm` | ブック本体（バイナリ・**唯一の完全バックアップ**） | 週1回 |

## src/ の取り扱い

- `.bas`（標準モジュール）は **VBE で先に「解放（削除）」してから File > Import File**。
  そのまま Import すると `modConfig1` のような別モジュールが増え、元のコードが生き続ける。
- `ThisWorkbook.cls` は **Import File では置き換えられない**。コードペインを全選択→削除→貼り付け。
  そのため `Attribute VB_...` 行は最初から除いてある。
- `frmController.frm.txt` は**イベントコードだけ**。ボタンの配置・サイズ・キャプションといった
  レイアウト情報（`.frx`）は含まれない。
  **`frmController` を復元できるのは `デイトレ用.xlsm` だけ。** src/ はコードのバックアップであって
  システムのバックアップではない。
- 文字コードは **Shift-JIS(CP932) / CRLF**。UTF-8 だと VBE で文字化けする。

## 前提環境

- Excel 365
- マーケットスピードII RSS（楽天証券）
- JPX `data_j.xls` は各自で配布元から取得し、`%USERPROFILE%\OneDrive\株・資産管理\銘柄表\` に置く
  （リポジトリには含めない。`.gitignore` で除外）
