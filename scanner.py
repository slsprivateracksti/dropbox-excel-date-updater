"""
scanner.py
Dropbox フォルダ内の Excel ファイルをスキャンし、
ファイル名・フォルダ名・シート名で検索する。
"""

from __future__ import annotations

import os
from typing import Generator

try:
    import openpyxl
    _OPENPYXL_AVAILABLE = True
except ImportError:
    _OPENPYXL_AVAILABLE = False

_EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


class ScanResult:
    """スキャン結果の1件。"""
    __slots__ = ("file_path", "rel_path", "folder_name", "file_name", "sheet_names")

    def __init__(
        self,
        file_path: str,
        rel_path: str,
        folder_name: str,
        file_name: str,
        sheet_names: list[str],
    ) -> None:
        self.file_path   = file_path      # フルパス
        self.rel_path    = rel_path       # Dropboxルートからの相対パス
        self.folder_name = folder_name    # 親フォルダ名
        self.file_name   = file_name      # ファイル名（拡張子なし）
        self.sheet_names = sheet_names    # シート名リスト

    def display_label(self, sheet: str) -> str:
        return f"{self.folder_name} / {self.file_name} > {sheet}"

    def to_dict(self, sheet: str) -> dict:
        return {
            "file_path"  : self.rel_path,
            "sheet_name" : sheet,
            "file_name"  : self.file_name,
            "folder_name": self.folder_name,
            "label"      : self.display_label(sheet),
        }


class DropboxScanner:
    """
    Parameters
    ----------
    dropbox_root : str
        Dropbox フォルダの絶対パス。
    """

    def __init__(self, dropbox_root: str) -> None:
        self.dropbox_root = dropbox_root

    def get_top_folders(self) -> list[str]:
        """Dropboxルート直下のフォルダ名一覧を返す。"""
        try:
            return sorted([
                d for d in os.listdir(self.dropbox_root)
                if os.path.isdir(os.path.join(self.dropbox_root, d))
            ])
        except Exception:
            return []

    def scan_folder(
        self,
        rel_folder: str,
        include_subdirs: bool = True,
        progress_callback=None,
    ) -> list[ScanResult]:
        """
        指定フォルダ（Dropboxルートからの相対パス）をスキャンして
        ScanResult のリストを返す。

        progress_callback(current_file: str) を渡すと進捗通知できる。
        """
        root = os.path.join(self.dropbox_root, rel_folder)
        results: list[ScanResult] = []

        for file_path in self._iter_excel_files(root, include_subdirs):
            if progress_callback:
                progress_callback(os.path.basename(file_path))

            sheet_names = self._get_sheet_names(file_path)
            rel_path    = os.path.relpath(file_path, self.dropbox_root)
            folder_name = os.path.basename(os.path.dirname(file_path))
            file_name   = os.path.splitext(os.path.basename(file_path))[0]

            results.append(ScanResult(
                file_path   = file_path,
                rel_path    = rel_path,
                folder_name = folder_name,
                file_name   = file_name,
                sheet_names = sheet_names,
            ))

        return results

    def search(
        self,
        results: list[ScanResult],
        keyword: str,
        mode: str = "OR",
    ) -> list[dict]:
        """
        スキャン結果からキーワード検索する。

        Parameters
        ----------
        results : list[ScanResult]
        keyword : str
            スペース区切りで複数ワード指定可（AND/OR は mode で切り替え）。
        mode : str
            "AND" または "OR"

        Returns
        -------
        list[dict]
            各要素は ScanResult.to_dict(sheet) の形式 + マッチ理由。
        """
        words = [w.strip() for w in keyword.split() if w.strip()]
        if not words:
            # キーワードなし → 全件返す（シート展開）
            return [r.to_dict(s) for r in results for s in r.sheet_names]

        matched: list[dict] = []
        for r in results:
            for sheet in r.sheet_names:
                targets_str = [
                    r.folder_name,
                    r.file_name,
                    sheet,
                ]
                if mode == "AND":
                    hit = all(
                        any(w.lower() in t.lower() for t in targets_str)
                        for w in words
                    )
                else:  # OR
                    hit = any(
                        any(w.lower() in t.lower() for t in targets_str)
                        for w in words
                    )
                if hit:
                    matched.append(r.to_dict(sheet))

        return matched

    # ── 内部処理 ─────────────────────────────

    def _iter_excel_files(
        self,
        root: str,
        include_subdirs: bool,
    ) -> Generator[str, None, None]:
        if not os.path.isdir(root):
            return
        if include_subdirs:
            for dirpath, _dirnames, filenames in os.walk(root):
                for fname in filenames:
                    if os.path.splitext(fname)[1].lower() in _EXCEL_EXTENSIONS:
                        yield os.path.join(dirpath, fname)
        else:
            for fname in os.listdir(root):
                fpath = os.path.join(root, fname)
                if os.path.isfile(fpath) and os.path.splitext(fname)[1].lower() in _EXCEL_EXTENSIONS:
                    yield fpath

    @staticmethod
    def _get_sheet_names(file_path: str) -> list[str]:
        if not _OPENPYXL_AVAILABLE:
            return []
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception:
            return []
