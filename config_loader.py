"""
config_loader.py
設定マスタ（JSON / CSV）を読み込み、施設ごとのターゲット情報を提供する。
"""

import csv
import json
import os
import sys
from typing import Any

FacilityConfig = dict[str, Any]
TargetInfo = dict[str, Any]

if getattr(sys, "frozen", False):
    _HERE = os.path.dirname(sys.executable)
else:
    _HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_JSON_PATH = os.path.join(_HERE, "config.xlsx_date_targets.json")
DEFAULT_CSV_PATH  = os.path.join(_HERE, "config.xlsx_date_targets.csv")


class ConfigLoader:
    def __init__(self, config_path: str = DEFAULT_JSON_PATH) -> None:
        self._config_path: str = config_path
        self._data: dict[str, FacilityConfig] = {}
        self._loaded: bool = False
        self._error: str = ""

    def load(self) -> bool:
        self._data = {}
        self._error = ""
        ext = os.path.splitext(self._config_path)[1].lower()
        try:
            if ext == ".json":
                self._load_json()
            elif ext == ".csv":
                self._load_csv()
            else:
                raise ValueError(f"未対応の拡張子です: {ext}")
            self._loaded = True
            return True
        except FileNotFoundError:
            self._error = f"設定ファイルが見つかりません: {self._config_path}"
        except json.JSONDecodeError as e:
            self._error = f"JSON の解析に失敗しました: {e}"
        except Exception as e:
            self._error = f"設定ファイルの読み込みに失敗しました: {e}"
        self._loaded = False
        return False

    def get_facility_names(self) -> list[str]:
        return list(self._data.keys())

    def get_targets_for_facility(self, facility_name: str) -> list[TargetInfo]:
        facility = self._data.get(facility_name)
        if facility is None:
            return []
        results: list[TargetInfo] = []
        for target in facility.get("targets", []):
            results.append({
                "file_path"  : facility.get("file_path", ""),
                "sheet_name" : facility.get("sheet_name", ""),
                "date_format": facility.get("date_format", "YYYY/MM/DD"),
                "type"       : target.get("type", "cell"),
                "address"    : target.get("address"),
                "col"        : target.get("col"),
                "start_row"  : target.get("start_row"),
                "end_row"    : target.get("end_row"),
            })
        return results

    @property
    def is_loaded(self) -> bool:
        return self._loaded

    @property
    def last_error(self) -> str:
        return self._error

    def _load_json(self) -> None:
        with open(self._config_path, encoding="utf-8") as f:
            raw: dict = json.load(f)
        for facility in raw.get("facilities", []):
            name = facility.get("name", "").strip()
            if not name:
                continue
            self._data[name] = facility

    def _load_csv(self) -> None:
        with open(self._config_path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row.get("facility_name", "").strip()
                if not name:
                    continue
                if name not in self._data:
                    self._data[name] = {
                        "name"       : name,
                        "file_path"  : row.get("file_path", "").strip(),
                        "sheet_name" : row.get("sheet_name", "").strip(),
                        "date_format": row.get("date_format", "YYYY/MM/DD").strip(),
                        "targets"    : [],
                    }
                target_type = row.get("target_type", "cell").strip()
                target: dict[str, Any] = {"type": target_type}
                if target_type == "cell":
                    target["address"] = row.get("cell_address", "").strip()
                elif target_type == "column":
                    target["col"] = row.get("col", "").strip()
                    try:
                        target["start_row"] = int(row.get("start_row", 0))
                        target["end_row"]   = int(row.get("end_row",   0))
                    except ValueError:
                        target["start_row"] = None
                        target["end_row"]   = None
                self._data[name]["targets"].append(target)
