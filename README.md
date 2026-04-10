# Dropbox Excel 日付更新ツール

> Dropbox 内の Excel ファイルへ、指定セルに日付を自動入力するデスクトップツール（Windows）

[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?logo=python)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-lightgrey?logo=windows)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

---

## 概要

施設の入居者一覧などを管理する Excel ファイルに対して、毎日・毎月行う日付入力作業を自動化するツールです。

- **1 施設**への書き込みも、**複数施設の一括処理**も、同じ画面から操作できます
- 設定は**パターンとして保存**でき、次回から呼び出すだけで再利用できます
- 実行履歴は**ログとして自動記録**されます

---

## 機能一覧

| タブ | 機能 |
|------|------|
| ① 単体実行 | ファイル・シート・セルを選択して 1 施設に日付を書き込む |
| ② 一括実行 | 保存済みパターンを使って複数施設を同じ日付でまとめて処理する |
| ③ 履歴・パターン管理 | パターンの保存・複写・削除、実行ログの確認・エクスポート |

### 主な特徴

- 🗓 **日付形式を自由設定**：`M/D` / `YYYY/MM/DD` / `YYYY年MM月DD日` など書式記号で自由に定義
- 📐 **セル範囲表記に対応**：`H3:H20` や `I2, H3:H20, C5` のような混在指定が可能
- ❌ **除外セル**：対象セルの中から特定のセルだけスキップ
- 🖱 **Excel 連携セル選択**：実際に Excel を開いてセルを選択し、番地をツールに取り込む（pywin32）
- 💾 **パターン保存・複写・インポート/エクスポート**：設定を再利用・他 PC へ移行
- 🔍 **Dropbox フォルダスキャン**：ファイル名・シート名で AND/OR 検索して施設を追加
- 📋 **実行ログ（CSV）**：日時・ユーザー名・施設名・書込日付・結果を自動記録
- ✅ **3 択確認ダイアログ**：実行前に「はい／いいえ／ファイルを開く」を選択

---

## クイックスタート

```bash
# 1. リポジトリを取得
git clone https://github.com/your-username/dropbox-excel-tool.git
cd dropbox-excel-tool

# 2. ライブラリをインストール
pip install -r requirements.txt

# 3. 設定ファイルを作成
copy config.xlsx_date_targets.sample.json config.xlsx_date_targets.json
# → config.xlsx_date_targets.json を編集して施設情報を入力

# 4. 起動
python main.py
```

詳細な手順は **[INSTALL.md](INSTALL.md)** を参照してください。

---

## 動作要件

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11 |
| Python | 3.11 以上 |
| Dropbox | デスクトップアプリ（`%USERPROFILE%\Dropbox` が存在すること） |

---

## ファイル構成

```
main.py                              # メインUI（3タブ構成）
config_loader.py                     # JSON/CSV 設定マスタ読み込み
excel_updater.py                     # Excel への日付書き込み（除外セル対応）
pattern_store.py                     # パターン保存管理（patterns.json）
log_manager.py                       # 実行ログ管理（execution_log.csv）
cell_selector.py                     # pywin32 による Excel セル選択連携
scanner.py                           # Dropbox フォルダスキャンと AND/OR 検索
requirements.txt                     # 依存パッケージ一覧
config.xlsx_date_targets.sample.json # 施設設定サンプル → コピーして使用
patterns.sample.json                 # パターンサンプル
```

> `config.xlsx_date_targets.json` / `patterns.json` / `execution_log.csv` は  
> 環境固有のため `.gitignore` で Git 管理対象外です。

---

## ドキュメント

| ドキュメント | 内容 |
|-------------|------|
| [INSTALL.md](INSTALL.md) | 詳細なインストール・セットアップ手順 |
| [FAQ.md](FAQ.md) | よくある質問とトラブル対応 |
| [CHANGELOG.md](CHANGELOG.md) | バージョンごとの更新履歴 |

---

## ライセンス

[MIT License](LICENSE)
