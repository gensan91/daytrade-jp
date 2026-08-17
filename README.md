# daytrade-jp

日本株ギャップアップ・デイトレードの半自動化システム（Excel VBA + マーケットスピードII RSS）。

現在は**紙トレード検証フェーズ**。ライブ運用前に統計的にクリーンなデータセットを作ることが目的。

## 現在地

→ **[STATE.md](STATE.md)** を読むこと。ここが唯一の現在地。

## 構成

| ファイル | 役割 |
|---|---|
| `STATE.md` | 現在地。毎回上書き |
| `KNOWLEDGE.md` | 技術知見。追記のみ |
| `ROUTINE.md` | 毎朝の手順 |
| `src/*.bas` | VBAモジュール（VBEからエクスポート） |
| `data/trades.csv` | Trades シートの書き出し |

## 前提環境

- Excel 365
- マーケットスピードII RSS（楽天証券）
- `.bas` は Shift-JIS(CP932) / CRLF
