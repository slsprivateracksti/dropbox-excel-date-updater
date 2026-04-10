# インストール手順 / Installation Guide

## 動作要件

| 項目 | 要件 |
|------|------|
| OS | Windows 10 / 11（64bit） |
| Python | 3.11 以上 |
| Dropbox | デスクトップアプリがインストール済みであること |
| Dropbox パス | `%USERPROFILE%\Dropbox` が存在すること |

---

## Step 1 — Python のインストール確認

コマンドプロンプトを開き、以下を実行してください。

```cmd
python --version
```

`Python 3.11.x` 以上が表示されれば OK です。  
表示されない場合は [python.org](https://www.python.org/downloads/) からインストールしてください。

> **ヒント**: インストール時に「Add Python to PATH」にチェックを入れてください。

---

## Step 2 — リポジトリの取得

### Git を使う場合

```cmd
git clone https://github.com/your-username/dropbox-excel-tool.git
cd dropbox-excel-tool
```

### ZIP でダウンロードする場合

GitHub ページ右上の **Code → Download ZIP** でダウンロードし、任意のフォルダに展開してください。

---

## Step 3 — ライブラリのインストール

ツールのフォルダで以下を実行します。

```cmd
pip install -r requirements.txt
```

### インストールされるライブラリ

| ライブラリ | 用途 | 必須 |
|-----------|------|------|
| `openpyxl` | Excel ファイルの読み書き | ✅ 必須 |
| `tkcalendar` | カレンダーポップアップ | 推奨 |
| `pywin32` | Excel 連携セル選択（Windows のみ） | 推奨 |

> `pywin32` は Windows 専用です。インストールしなくてもツールは動作しますが、  
> 「Excelで選択」ボタンが使えなくなります（手入力で代替可）。

---

## Step 4 — 設定ファイルの作成

サンプルファイルをコピーして設定ファイルを作成します。

```cmd
copy config.xlsx_date_targets.sample.json config.xlsx_date_targets.json
```

テキストエディタ（メモ帳など）で `config.xlsx_date_targets.json` を開き、  
施設名・ファイルパス・シート名を実際の環境に合わせて編集してください。

### 設定ファイルの例

```json
{
  "facilities": [
    {
      "name": "A施設",
      "file_path": "D薬局\\①入居者一覧・セット表\\入居者一覧（Hグループ）.xlsx",
      "sheet_name": "A施設",
      "date_format": "M/D",
      "targets": [
        { "type": "cell", "address": "I2" }
      ]
    }
  ]
}
```

> **`file_path` の注意点**  
> `%USERPROFILE%\Dropbox` からの相対パスを記述します。  
> バックスラッシュは `\\` と二重にしてください。

---

## Step 5 — 起動

```cmd
python main.py
```

メインウィンドウが表示されれば完了です。

---

## アップデート方法

### Git を使っている場合

```cmd
git pull origin main
pip install -r requirements.txt
```

### ZIP でインストールした場合

1. 最新 ZIP をダウンロードして展開
2. `config.xlsx_date_targets.json` と `patterns.json` は上書きしないよう注意
3. それ以外のファイルを新しいものに置き換える
4. `pip install -r requirements.txt` を再実行

---

## アンインストール方法

ツールのフォルダを削除するだけです。レジストリへの書き込みは行っていません。

ライブラリも削除したい場合：

```cmd
pip uninstall openpyxl tkcalendar pywin32
```

---

## トラブルシューティング

詳しいトラブル対応は [FAQ.md](FAQ.md) を参照してください。

| 症状 | 確認事項 |
|------|---------|
| 起動時に `ModuleNotFoundError` | `pip install -r requirements.txt` を再実行 |
| 「Dropboxフォルダが見つかりません」 | Dropbox アプリの起動と同期完了を確認 |
| 「Excelで選択」が動作しない | `pip install pywin32` を実行 |
| Excel ファイルが開けない | Dropbox の同期完了を確認、ファイルパスを再選択 |
