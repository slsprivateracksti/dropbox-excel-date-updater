# Release 前 最終チェックリスト
（dropbox-excel-date-updater）

## 1. バージョン・ファイル名

- [ ] 今回のバージョン番号を決めた（例: `v2.0.0`）。
- [ ] Release 用 zip のファイル名を決めた  
      例: `dropbox-excel-date-updater-v2.0.0-source.zip`

## 2. リポジトリ内のファイル構成

- [ ] 次のファイルがリポジトリ直下に存在する

  - [ ] `main.py`
  - [ ] `config_loader.py`
  - [ ] `excel_updater.py`
  - [ ] `pattern_store.py`
  - [ ] `log_manager.py`
  - [ ] `cell_selector.py`
  - [ ] `scanner.py`
  - [ ] `requirements.txt`
  - [ ] `README.md`
  - [ ] `LICENSE`
  - [ ] `CHANGELOG.md`
  - [ ] `INSTALL.md`
  - [ ] `FAQ.md`
  - [ ] `.gitignore`
  - [ ] `config.xlsx_date_targets.sample.json`

- [ ] 説明書・資料が `docs/` に入っている

  - [ ] `docs/HowToUse_DropboxExcelTool.pdf`
  - [ ] `docs/HowToUse_DropboxExcelTool.docx`（任意）

- [ ] 画面キャプチャが `screenshots/` に入っている（名前は例）

  - [ ] `screenshots/main-overview.png`
  - [ ] `screenshots/single-run-tab.png`
  - [ ] `screenshots/batch-run-tab.png`
  - [ ] `screenshots/pattern-management-tab.png`

## 3. 公開しないファイルの確認（.gitignore で除外）

- [ ] `.venv/` や `venv/` はコミットされていない
- [ ] `__pycache__/` はコミットされていない
- [ ] `config.xlsx_date_targets.json` はコミットされていない
- [ ] `config.xlsx_date_targets.csv` はコミットされていない
- [ ] `patterns.json` はコミットされていない
- [ ] `execution_log.csv` はコミットされていない
- [ ] 実際の Dropbox 内 Excel ファイルは入っていない
- [ ] `.vscode/settings.json` はコミットされていない
- [ ] `Thumbs.db`, `Desktop.ini` はコミットされていない

## 4. 動作確認

- [ ] Python 3.11 以上の環境で `pip install -r requirements.txt` が成功する
- [ ] `main.py` を実行して、ツールが起動する
- [ ] 単体実行タブでテスト用の Excel に対して日付書き込みができる
- [ ] 一括実行タブでテスト用パターンを使った一括実行が動作する
- [ ] 履歴・パターン管理タブでログ・パターンが正常に扱える

（テストには実運用の機密データではなく、テスト用の Excel ファイルを使う）

## 5. ドキュメント内容の確認

- [ ] `README.md` に以下が書かれている
  - [ ] ツール概要
  - [ ] 主な機能
  - [ ] 動作環境（Windows / Dropbox 前提 / Python バージョン）
  - [ ] インストール手順（`pip install -r requirements.txt` など）
  - [ ] 起動方法（`main.py` を実行）
  - [ ] 基本的な使い方（単体実行 / 一括実行 / パターン管理）
  - [ ] 説明書（PDF）やスクリーンショットへのリンク

- [ ] `INSTALL.md` に、より詳しい導入手順が書かれている
- [ ] `FAQ.md` に、よくあるエラーやトラブル対応が整理されている
- [ ] `CHANGELOG.md` に今回バージョンの変更内容が追記されている
- [ ] `LICENSE` に公開条件が明記されている

## 6. Release 用 zip の作成

1. 新しいフォルダを作る（例）

   ```text
   dropbox-excel-date-updater-v2.0.0-source/
   ```

2. 次のファイル・フォルダだけをその中にコピーする

   - `main.py`
   - `config_loader.py`
   - `excel_updater.py`
   - `pattern_store.py`
   - `log_manager.py`
   - `cell_selector.py`
   - `scanner.py`
   - `requirements.txt`
   - `README.md`
   - `INSTALL.md`
   - `FAQ.md`
   - `LICENSE`
   - `config.xlsx_date_targets.sample.json`
   - `docs/HowToUse_DropboxExcelTool.pdf`

3. 上記フォルダを zip 化する

   - [ ] zip ファイル名が `dropbox-excel-date-updater-v2.0.0-source.zip` になっている

4. zip の中身を確認する

   - [ ] 不要なファイル（`.venv`, `__pycache__`, `.git` など）が入っていない
   - [ ] 実運用データ（`config.xlsx_date_targets.json`, `patterns.json`, `execution_log.csv`）が入っていない
   - [ ] PDF 説明書が入っている

## 7. GitHub Release 作成時の確認

- [ ] 対応するタグ（例: `v2.0.0`）を作成する
- [ ] Release タイトル（例: `Dropbox Excel Date Updater v2.0.0`）を入力する
- [ ] Release 説明欄に、主な変更点と注意事項を書く
- [ ] Assets に `dropbox-excel-date-updater-v2.0.0-source.zip` をアップロードする
- [ ] 自動で付く `Source code (zip)` / `Source code (tar.gz)` はそのままでOK
- [ ] 「Publish release」を押す前に内容を再確認する