"""
cell_selector.py
pywin32 を使って、現在 Excel で選択中のセル番地を取得する。

使い方:
    1. Excel でファイルを開き、セルを選択する
    2. このモジュールの get_selected_cells() を呼ぶ
    3. 選択中のセル番地リストが返る（例: ["I2", "C5", "D5"]）

依存: pywin32 (pip install pywin32)
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass


def is_available() -> bool:
    """pywin32 が利用可能かチェック。"""
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


def get_selected_cells(file_path: str | None = None) -> list[str]:
    """
    現在 Excel で選択中のセル番地リストを返す。

    Parameters
    ----------
    file_path : str | None
        特定のファイルに限定する場合はフルパスを渡す。
        None の場合はアクティブなブックの選択セルを取得。

    Returns
    -------
    list[str]
        例: ["I2", "C5", "D5"]
        取得できない場合は空リスト。
    """
    try:
        import win32com.client
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return []

    try:
        if file_path:
            # 指定ファイルのブックを探す
            wb = None
            import os
            target_name = os.path.basename(file_path).lower()
            for book in excel.Workbooks:
                if book.Name.lower() == target_name:
                    wb = book
                    break
            if wb is None:
                return []
            ws = wb.ActiveSheet
        else:
            ws = excel.ActiveSheet

        selection = excel.Selection
        cells: list[str] = []

        # 複数セル選択（Areas）に対応
        for area in selection.Areas:
            for cell in area.Cells:
                addr = _rc_to_addr(cell.Row, cell.Column)
                if addr not in cells:
                    cells.append(addr)

        return cells

    except Exception:
        return []


def open_file_in_excel(full_path: str) -> bool:
    """
    指定したファイルを Excel で開く。
    既に開いている場合はアクティブにするだけ。

    Returns True if succeeded.
    """
    try:
        import win32com.client
        import os
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel = win32com.client.Dispatch("Excel.Application")

        excel.Visible = True

        # 既に開いていないか確認
        target_name = os.path.basename(full_path).lower()
        for book in excel.Workbooks:
            if book.Name.lower() == target_name:
                book.Activate()
                excel.WindowState = -4137  # xlMaximized
                return True

        excel.Workbooks.Open(full_path)
        excel.WindowState = -4137
        return True

    except Exception:
        # pywin32 未インストール時は os.startfile にフォールバック
        try:
            import os
            os.startfile(full_path)
            return True
        except Exception:
            return False


def _rc_to_addr(row: int, col: int) -> str:
    """行列インデックス（1始まり）をセル番地文字列に変換（例: 2,9 → "I2"）。"""
    col_str = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return f"{col_str}{row}"
