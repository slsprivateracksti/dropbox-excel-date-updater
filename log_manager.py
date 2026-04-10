"""
log_manager.py
実行ログを CSV に記録・読み込みする。
"""

import csv
import os
from datetime import datetime
from getpass import getuser
from typing import Any

_HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_LOG_PATH = os.path.join(_HERE, "execution_log.csv")

_FIELDNAMES = [
    "datetime",
    "user",
    "pattern_name",
    "facility_name",
    "file_path",
    "sheet_name",
    "target_cells",
    "exclude_cells",
    "write_date",
    "status",
    "detail",
]


class LogManager:
    def __init__(self, path: str = DEFAULT_LOG_PATH) -> None:
        self._path = path
        self._ensure_file()

    def _ensure_file(self) -> None:
        if not os.path.exists(self._path):
            with open(self._path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=_FIELDNAMES)
                writer.writeheader()

    def append(
        self,
        pattern_name: str,
        facility_name: str,
        file_path: str,
        sheet_name: str,
        target_cells: list[str],
        exclude_cells: list[str],
        write_date: str,
        success: bool,
        detail: str = "",
    ) -> None:
        try:
            user = getuser()
        except Exception:
            user = os.environ.get("USERNAME", "unknown")

        row = {
            "datetime"     : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user"         : user,
            "pattern_name" : pattern_name,
            "facility_name": facility_name,
            "file_path"    : file_path,
            "sheet_name"   : sheet_name,
            "target_cells" : ", ".join(target_cells),
            "exclude_cells": ", ".join(exclude_cells),
            "write_date"   : write_date,
            "status"       : "成功" if success else "失敗",
            "detail"       : detail,
        }
        with open(self._path, "a", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=_FIELDNAMES)
            writer.writerow(row)

    def read_all(self) -> list[dict[str, Any]]:
        try:
            with open(self._path, encoding="utf-8-sig", newline="") as f:
                return list(csv.DictReader(f))
        except Exception:
            return []

    def export_to(self, dest_path: str) -> bool:
        import shutil
        try:
            shutil.copy2(self._path, dest_path)
            return True
        except Exception:
            return False
