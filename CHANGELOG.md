# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [v2.1.0] - 2026-04-09

### Added

#### ① 単体実行タブ
- タイトル行右側に「全てクリア」ボタンを追加（全フィールドを一括リセット）
- 施設名・対象セル・除外セルの各フィールド右側に「✕」個別クリアボタンを追加
- 日付形式に `⚙設定` ダイアログを追加
  - リアルタイムプレビュー表示
  - プリセット形式ボタン（M/D / YYYY/MM/DD / YYYY年MM月DD日 など）
  - デフォルトを `M/D` に変更
- セル範囲表記 `H3:H20` に対応
  - 実行時に内部で H3〜H20 へ自動展開
  - 単独セルとの混在（`I2, H3:H20, C5`）も可能
  - 除外セルも同様の範囲表記に対応
  - 保存・表示時は連続セルを範囲形式に自動圧縮

#### ② 一括実行タブ
- 「追加」ボタンを 2 択化
  - **新規パターン設定から追加**：パターン設定ダイアログを開いて新規登録
  - **保存済みパターンから選択**：一覧から複数選択してリストへ追加
- ③タブの「一括実行タブへ追加」ボタンから直接リストへ追加可能に

#### ③ 履歴・パターン管理タブ
- 保存済みパターン一覧に AND/OR 検索を追加（パターン名・施設名・実行日時・ステータス対象）
- 実行ログ一覧に AND/OR 検索を追加（日時・ユーザー・施設名・書込日付・結果対象）
- 「複写」ボタンを追加（選択パターンを `_コピー` という名前で複製し編集ダイアログを開く）
- パターン編集ダイアログに「別名で保存」ボタンを追加（元パターンを変更せず別名保存）
- 「一括実行タブへ追加」ボタンを追加

### Changed
- `PatternEditDialog` に `allow_rename_save` オプションを追加
- `compress_cells()` 関数を追加（連続セルリストを範囲表記に圧縮）
- `format_date_preview()` 関数を追加（YYYY/MM/DD/M/D 書式のリアルタイムプレビュー）
- `search_rows()` 汎用 AND/OR 検索関数を追加
- `center_window()` ヘルパーを追加（Toplevel ウィンドウの中央配置を共通化）
- ウィンドウサイズを `780×660` → `820×700` に拡張

---

## [v2.0.0] - 2026-04-09

### Added
- **3タブ構成 UI**（`main.py` を全面刷新）
  - ① 単体実行タブ
  - ② 一括実行タブ（F案：左リスト＋右詳細）
  - ③ 履歴・パターン管理タブ
- `pattern_store.py`：施設パターンの JSON 永続管理
- `log_manager.py`：実行ログの CSV 自動記録（日時・Windows ログインユーザー名・施設名・結果）
- `cell_selector.py`：pywin32 による Excel 連携セル選択（実ファイルを開いて選択→取得）
- `scanner.py`：Dropbox フォルダスキャンと AND/OR キーワード検索
- 確認ダイアログを 3 択化（はい／いいえ／ファイルを開く）
- 対象セル・除外セル機能（`excel_updater.py` に `exclude_cells` 引数を追加）
- 戻る・進むボタン（一括実行タブで施設を順番にナビゲート）
- パターンのエクスポート／インポート（JSON 形式）
- Dropbox スキャン機能（フォルダ指定 + サブフォルダ対応 + AND/OR 検索）
- `config.xlsx_date_targets.sample.json` / `patterns.sample.json` サンプルファイル
- `requirements.txt`
- `README.md`

### Changed
- `excel_updater.py`：`update_dates()` に `exclude_cells` 引数を追加
- `config_loader.py`：既存コードをそのまま継承・流用

### Removed
- `dropbox_file_opener.py`（機能を `main.py` の単体実行タブに統合）

---

## [v1.0.0] - 2026-03-xx *(初期リリース)*

### Added
- `main.py`：単一ウィンドウ UI（施設選択・日付入力・実行）
- `config_loader.py`：JSON / CSV 形式の設定マスタ読み込み
- `excel_updater.py`：openpyxl による Excel セル・列への日付書き込み
- `dropbox_file_opener.py`：Dropbox ファイル選択補助ダイアログ
- `config.xlsx_date_targets.json` / `.csv`：施設設定マスタ
