"""
pattern_store.py
施設パターン（ファイル・シート・対象セル・除外セル）の保存・読み込み・管理。
保存先: patterns.json（スクリプトと同じフォルダ）
"""

import json
import os
import sys
from datetime import datetime
from typing import Any

if getattr(sys, "frozen", False):
    # PyInstaller でexe化された場合: exeファイルのあるフォルダを使う
    _HERE = os.path.dirname(sys.executable)
else:
    # 通常の.py実行時
    _HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PATTERN_PATH = os.path.join(_HERE, "patterns.json")


class PatternStore:
    """
    パターンの構造:
    {
      "pattern_name": str,        # 例: "GHひらばり_通常"
      "facility_name": str,       # 施設名
      "file_path": str,           # Dropboxルートからの相対パス
      "sheet_name": str,          # シート名
      "target_cells": list[str],  # 対象セル ["I2", "C5"]
      "exclude_cells": list[str], # 除外セル ["D5"]
      "date_format": str,         # "YYYY/MM/DD"
      "last_executed": str,       # ISO datetime or ""
      "last_status": str,         # "成功" | "失敗" | ""
    }
    """

    def __init__(self, path: str = DEFAULT_PATTERN_PATH) -> None:
        self._path = path
        self._patterns: list[dict[str, Any]] = []
        self._load()

    # ── 読み込み ────────────────────────────

    def _load(self) -> None:
        if not os.path.exists(self._path):
            self._patterns = []
            return
        try:
            with open(self._path, encoding="utf-8") as f:
                self._patterns = json.load(f)
        except Exception:
            self._patterns = []

    def reload(self) -> None:
        self._load()

    # ── 保存 ────────────────────────────────

    def _save(self) -> None:
        with open(self._path, "w", encoding="utf-8") as f:
            json.dump(self._patterns, f, ensure_ascii=False, indent=2)

    # ── 公開メソッド ─────────────────────────

    def get_all(self) -> list[dict[str, Any]]:
        return list(self._patterns)

    def get_pattern_names(self) -> list[str]:
        return [p["pattern_name"] for p in self._patterns]

    def get_by_name(self, pattern_name: str) -> dict[str, Any] | None:
        for p in self._patterns:
            if p["pattern_name"] == pattern_name:
                return dict(p)
        return None

    def get_by_facility(self, facility_name: str) -> list[dict[str, Any]]:
        return [p for p in self._patterns if p.get("facility_name") == facility_name]

    def save_pattern(self, pattern: dict[str, Any]) -> bool:
        """
        パターンを追加または上書き保存する。
        pattern_name が既存と一致する場合は上書き。
        Returns True on success.
        """
        name = pattern.get("pattern_name", "").strip()
        if not name:
            return False

        # 必須フィールドのデフォルト補完
        pattern.setdefault("facility_name", "")
        pattern.setdefault("file_path", "")
        pattern.setdefault("sheet_name", "")
        pattern.setdefault("target_cells", [])
        pattern.setdefault("exclude_cells", [])
        pattern.setdefault("date_format", "YYYY/MM/DD")
        pattern.setdefault("last_executed", "")
        pattern.setdefault("last_status", "")

        for i, p in enumerate(self._patterns):
            if p["pattern_name"] == name:
                self._patterns[i] = pattern
                self._save()
                return True

        self._patterns.append(pattern)
        self._save()
        return True

    def update_execution_result(self, pattern_name: str, success: bool) -> None:
        """実行後にlast_executed / last_status を更新する。"""
        for p in self._patterns:
            if p["pattern_name"] == pattern_name:
                p["last_executed"] = datetime.now().isoformat(timespec="seconds")
                p["last_status"]   = "成功" if success else "失敗"
                self._save()
                return

    def delete_pattern(self, pattern_name: str) -> bool:
        before = len(self._patterns)
        self._patterns = [p for p in self._patterns if p["pattern_name"] != pattern_name]
        if len(self._patterns) < before:
            self._save()
            return True
        return False

    # ── エクスポート／インポート ──────────────

    def export_to(self, path: str) -> bool:
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._patterns, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    def import_from(self, path: str, overwrite: bool = False) -> tuple[int, int]:
        """
        Returns (added, skipped).
        overwrite=True の場合は同名パターンを上書き。
        """
        try:
            with open(path, encoding="utf-8") as f:
                imported: list[dict] = json.load(f)
        except Exception:
            return 0, 0

        added = skipped = 0
        existing_names = self.get_pattern_names()

        for p in imported:
            name = p.get("pattern_name", "")
            if name in existing_names and not overwrite:
                skipped += 1
            else:
                self.save_pattern(p)
                added += 1

        return added, skipped
