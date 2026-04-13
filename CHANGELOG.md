# Changelog

このプロジェクトのすべての変更を記録します。  
フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.0.0/) に準拠し、
バージョニングは [Semantic Versioning](https://semver.org/lang/ja/) に従います。

---

## [v1.2.0] - 2026-04-14

### Fixed
- exe化時にパターン・ログが起動のたびにリセットされる問題を修正
  - PyInstaller環境では `__file__` が一時フォルダを指すため、`sys.frozen` で判定し `sys.executable` のフォルダを使用するよう変更
  - 対象: `pattern_store.py` / `log_manager.py` / `config_loader.py`

### Changed
- `build/` `dist/` `*.spec` を `.gitignore` に追加（バイナリをgit管理外へ）
- マニュアル（docx/pdf）を `docs/` フォルダに整理
- 不要な旧ログ・旧JSONファイルを削除

---

## [v1.1.0] - 2026-04-09

### Added
- 3タブ構成のGUIを全面的に作り直し（単体実行 / 一括実行 / 履歴・パターン管理）
- パターン保存・読み込み機能（`pattern_store.py`）
- 実行ログのCSV記録機能（`log_manager.py`）
- Dropboxフォルダスキャン・AND/OR検索機能（`scanner.py`）
- pywin32によるExcelセル選択連携（`cell_selector.py`）
- 確認ダイアログ（はい／いいえ／ファイルを開く）
- パターンのJSONエクスポート／インポート機能
- HowToUseマニュアル（docx/pdf）追加

### Changed
- モジュール構成を全面的に分離・整理
  - `config_loader.py` / `excel_updater.py` / `pattern_store.py` / `log_manager.py` / `cell_selector.py` / `scanner.py`

---

## [v1.0.0] - 2026-03-29

### Added
- 初回リリース
- Dropbox内のExcelファイルへ指定セルに日付を書き込む基本機能
- tkinter によるGUI実装
- openpyxl によるExcel操作
- PyInstaller によるexe化対応（`main.spec`）
