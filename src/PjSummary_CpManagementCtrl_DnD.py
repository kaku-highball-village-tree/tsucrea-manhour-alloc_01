# -*- coding: utf-8 -*-
"""
SellGeneralAdminCost_Allocation_DnD.py

ドラッグ＆ドロップで工数TSVと損益計算書TSVを受け取り、
SellGeneralAdminCost_Allocation_Cmd_0002.py を実行するGUI。

使い方:
  ウィンドウに工数TSVと損益計算書TSVをドラッグ＆ドロップする。

仕様:
  - 入力は以下の2種類のみ:
      工数_yyyy年mm月_step0014_各プロジェクトの計上カンパニー名_工数_カンパニーの工数.tsv
      工数_yyyy年mm月_step15_各プロジェクトの工数.tsv
      損益計算書_yyyy年mm月_A∪B_プロジェクト名_C∪D_vertical.tsv
  - yyyy年mm月 が一致する工数/損益計算書の組み合わせのみ有効。
  - 有効な組み合わせは yyyy年mm月 の連続した範囲のみ採用する。
    (例: 2025年07月〜2025年10月 が連続していれば有効)
  - 採用された連続範囲はテキストファイルに記録する。
    (例: 採用範囲: 2025年07月〜2025年10月)
  - 有効な組み合わせのみを Cmd 版に渡して実行する。
"""

from __future__ import annotations

import ctypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import tkinter as tk
from ctypes import wintypes
from typing import Dict, List, Optional, Tuple

import win32api
import win32con
import win32gui

if not hasattr(win32con, "DEFAULT_GUI_FONT"):
    win32con.DEFAULT_GUI_FONT = 17


class DRAWITEMSTRUCT(ctypes.Structure):
    _fields_ = [
        ("CtlType", wintypes.UINT),
        ("CtlID", wintypes.UINT),
        ("itemID", wintypes.UINT),
        ("itemAction", wintypes.UINT),
        ("itemState", wintypes.UINT),
        ("hwndItem", wintypes.HWND),
        ("hDC", wintypes.HDC),
        ("rcItem", wintypes.RECT),
        ("itemData", ctypes.c_void_p),
    ]

BUTTON_ID_BASE: int = 1001
BUTTON_HEIGHT: int = 28
BUTTON_MARGIN: int = 10
g_last_output_directory: Optional[str] = None
g_action_button_handles: List[int] = []
g_action_button_brush_handle: Optional[int] = None
g_default_gui_font_handle: Optional[int] = None
g_right_button_down_handle: Optional[int] = None
g_main_window_handle: Optional[int] = None
ACTION_BUTTON_COLOR = (0xC0, 0xE8, 0xFF)

BUTTON_LABELS: Tuple[str, ...] = (
    "期間",
    "カンパニー実績",
    "カンパニー利益率順位",
    "プロジェクト別損益",
    "グループ別損益",
    "カンパニー別損益",
    "CP別経営管理用(計上カンパニー)",
    "CP別経営管理用(計上グループ)",
    "PJ別損益計算書",
)


def show_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    iOwnerWindowHandle: int = g_main_window_handle or win32gui.GetForegroundWindow()
    iMessageBoxType: int = (
        win32con.MB_OK
        | win32con.MB_ICONINFORMATION
        | win32con.MB_TASKMODAL
        | win32con.MB_SETFOREGROUND
    )
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


def show_error_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    iOwnerWindowHandle: int = g_main_window_handle or win32gui.GetForegroundWindow()
    iMessageBoxType: int = (
        win32con.MB_OK
        | win32con.MB_ICONERROR
        | win32con.MB_TASKMODAL
        | win32con.MB_SETFOREGROUND
    )
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


def report_exception(pszContext: str, exc: Exception) -> None:
    pszTraceback = traceback.format_exc()
    append_error_log(
        "\n".join(
            [
                f"[Error] {pszContext}",
                f"Detail: {exc}",
                pszTraceback,
            ]
        )
    )
    show_error_message_box(
        f"Error: {pszContext}. Detail = {exc}\n(See error log for traceback.)",
        "SellGeneralAdminCost_Allocation_DnD",
    )


def append_error_log(pszMessage: str) -> None:
    pszOutputPath: str = os.path.join(
        os.path.dirname(__file__),
        "SellGeneralAdminCost_Allocation_DnD_error.txt",
    )
    with open(pszOutputPath, "a", encoding="utf-8", newline="") as objFile:
        objFile.write(pszMessage + "\n")


def get_temp_output_directory() -> str:
    pszBaseDirectory: str = os.path.dirname(__file__)
    pszOutputDirectory: str = os.path.join(pszBaseDirectory, "temp")
    os.makedirs(pszOutputDirectory, exist_ok=True)
    return pszOutputDirectory


def get_manhour_temp_output_directory() -> str:
    pszBaseDirectory: str = os.path.dirname(__file__)
    pszOutputDirectory: str = os.path.join(pszBaseDirectory, "temp", "工数系")
    os.makedirs(pszOutputDirectory, exist_ok=True)
    return pszOutputDirectory


def set_last_output_directory(pszDirectory: Optional[str]) -> None:
    global g_last_output_directory
    if pszDirectory is None:
        return
    if not os.path.isdir(pszDirectory):
        return
    g_last_output_directory = pszDirectory


def open_last_output_directory() -> None:
    if g_last_output_directory is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    if not os.path.isdir(g_last_output_directory):
        show_error_message_box(
            "Error: 出力フォルダーが見つかりません。\n" + g_last_output_directory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(g_last_output_directory)


def open_script_directory() -> None:
    pszScriptDirectory = os.path.abspath(os.path.dirname(__file__))
    if not os.path.isdir(pszScriptDirectory):
        os.makedirs(pszScriptDirectory, exist_ok=True)
    try:
        os.startfile(pszScriptDirectory)
    except Exception as exc:
        show_error_message_box(
            "Error: フォルダーを開けませんでした。\n" + str(exc),
            "SellGeneralAdminCost_Allocation_DnD",
        )


def extract_project_code_from_file_name(pszFileName: str) -> Optional[str]:
    pszPrefix = "PJサマリ_単・累計_"
    if not pszFileName.startswith(pszPrefix):
        return None
    pszBase = pszFileName
    if pszBase.endswith(".xlsx"):
        pszBase = pszBase[: -len(".xlsx")]
    pszBody = pszBase[len(pszPrefix) :]
    if pszBody == "":
        return None
    return pszBody.split("_", 1)[0]


def is_valid_project_code(pszCode: str) -> bool:
    if re.fullmatch(r"P\d{5}", pszCode) is not None:
        return True
    return re.fullmatch(r"[A-OQ-Z]\d{3}", pszCode) is not None


def choose_project_pl_code(
    pszProjectDirectory: str,
) -> Optional[str]:
    objCandidates = [
        pszName
        for pszName in os.listdir(pszProjectDirectory)
        if pszName.startswith("PJサマリ_単・累計_") and pszName.endswith(".xlsx")
    ]
    objCandidates.sort()
    objResult: Dict[str, Optional[str]] = {"code": None}

    def on_select(event: tk.Event) -> None:
        objSelection = objListBox.curselection()
        if not objSelection:
            return
        pszFileName = objListBox.get(objSelection[0])
        pszCode = extract_project_code_from_file_name(pszFileName)
        if pszCode:
            objEntryVar.set(pszCode)

    def on_confirm() -> None:
        pszCode = objEntryVar.get().strip().upper()
        if not is_valid_project_code(pszCode):
            show_error_message_box(
                "Error: PJコードの形式が正しくありません。\n"
                + "P00000 または P 以外の英大文字 + 3桁を入力してください。",
                "SellGeneralAdminCost_Allocation_DnD",
            )
            return
        objResult["code"] = pszCode
        objWindow.grab_release()
        objWindow.destroy()

    def on_cancel() -> None:
        objWindow.grab_release()
        objWindow.destroy()

    objWindow = tk.Tk()
    objWindow.title("プロジェクト損益: PJコード選択")
    objWindow.geometry("640x400")
    objWindow.resizable(True, True)
    objWindow.attributes("-topmost", True)
    objWindow.grab_set()
    objWindow.focus_force()
    objWindow.lift()
    objWindow.protocol("WM_DELETE_WINDOW", on_cancel)

    objFrame = tk.Frame(objWindow)
    objFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    objLabel = tk.Label(
        objFrame,
        text="PJコードを入力するか、一覧から選択してください。",
        anchor="w",
    )
    objLabel.pack(fill=tk.X)

    objListFrame = tk.Frame(objFrame)
    objListFrame.pack(fill=tk.BOTH, expand=True, pady=5)

    objScrollBar = tk.Scrollbar(objListFrame)
    objScrollBar.pack(side=tk.RIGHT, fill=tk.Y)

    objListBox = tk.Listbox(
        objListFrame,
        yscrollcommand=objScrollBar.set,
        selectmode=tk.SINGLE,
    )
    objListBox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    objScrollBar.config(command=objListBox.yview)

    for pszName in objCandidates:
        objListBox.insert(tk.END, pszName)

    objListBox.bind("<<ListboxSelect>>", on_select)

    objEntryVar = tk.StringVar()
    objEntry = tk.Entry(objFrame, textvariable=objEntryVar)
    objEntry.pack(fill=tk.X, pady=5)
    objEntry.focus_set()

    objButtonFrame = tk.Frame(objFrame)
    objButtonFrame.pack(fill=tk.X, pady=5)

    objOkButton = tk.Button(objButtonFrame, text="OK", width=12, command=on_confirm)
    objOkButton.pack(side=tk.RIGHT, padx=5)
    objCancelButton = tk.Button(
        objButtonFrame, text="Cancel", width=12, command=on_cancel
    )
    objCancelButton.pack(side=tk.RIGHT, padx=5)

    objWindow.mainloop()
    return objResult["code"]


def find_latest_execution_root_directory() -> Optional[str]:
    pszBaseDirectory = os.path.abspath(os.path.dirname(__file__))
    objCandidates: List[str] = []
    for pszName in os.listdir(pszBaseDirectory):
        if not pszName.endswith("_損益工数実行表"):
            continue
        pszPath = os.path.join(pszBaseDirectory, pszName)
        if os.path.isdir(pszPath):
            objCandidates.append(pszPath)
    if not objCandidates:
        return None
    objCandidates.sort()
    return objCandidates[-1]


def handle_period_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszPeriodDirectory = os.path.join(pszExecutionRoot, "期間")
    objCandidateFileNames: List[str] = [
        "SellGeneralAdminCost_Allocation_Cmd_SelectedRange_And_AccountPeriodRange.txt",
        "SellGeneralAdminCost_Allocation_Cmd_SelectedRange.txt",
        "SellGeneralAdminCost_Allocation_DnD_SelectedRange.txt",
    ]
    pszTargetPath: Optional[str] = None
    for pszFileName in objCandidateFileNames:
        pszCandidatePath: str = os.path.join(pszPeriodDirectory, pszFileName)
        if os.path.isfile(pszCandidatePath):
            pszTargetPath = pszCandidatePath
            break
    if pszTargetPath is None:
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + os.path.join(
                pszPeriodDirectory,
                objCandidateFileNames[0],
            ),
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszTargetPath)


def handle_period_left_double_click() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszPeriodDirectory = os.path.join(pszExecutionRoot, "期間")
    pszTargetPath = os.path.join(
        pszPeriodDirectory,
        "SellGeneralAdminCost_Allocation_Cmd_AccountPeriodRange.txt",
    )
    if not os.path.isfile(pszTargetPath):
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszTargetPath)

def handle_period_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszPeriodDirectory = os.path.join(pszExecutionRoot, "期間")
    if not os.path.isdir(pszPeriodDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszPeriodDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszPeriodDirectory)


def handle_company_results_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszCompanyDirectory = os.path.join(pszExecutionRoot, "カンパニー実績")
    pszTargetPath = os.path.join(
        pszCompanyDirectory,
        "PJサマリ_PJ別_売上・売上原価・販管費・利益率.xlsx",
    )
    if not os.path.isfile(pszTargetPath):
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszTargetPath)


def handle_company_results_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszCompanyDirectory = os.path.join(pszExecutionRoot, "カンパニー実績")
    if not os.path.isdir(pszCompanyDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszCompanyDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszCompanyDirectory)


def handle_company_margin_rank_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszMarginDirectory = os.path.join(pszExecutionRoot, "カンパニー利益率順位")
    pszTargetPath = os.path.join(
        pszMarginDirectory,
        "PJサマリ_単月・累計_粗利金額ランキング.xlsx",
    )
    if not os.path.isfile(pszTargetPath):
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszTargetPath)


def handle_company_margin_rank_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszMarginDirectory = os.path.join(pszExecutionRoot, "カンパニー利益率順位")
    if not os.path.isdir(pszMarginDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszMarginDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszMarginDirectory)


def handle_project_pl_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszProjectDirectory = os.path.join(pszExecutionRoot, "プロジェクト損益")
    if not os.path.isdir(pszProjectDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszProjectDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszProjectCode = choose_project_pl_code(pszProjectDirectory)
    if pszProjectCode is None:
        return
    pszPrefix = f"PJサマリ_単・累計_{pszProjectCode}"
    objCandidates = [
        pszName
        for pszName in os.listdir(pszProjectDirectory)
        if pszName.startswith(pszPrefix) and pszName.endswith(".xlsx")
    ]
    if not objCandidates:
        pszTargetPath = os.path.join(
            pszProjectDirectory,
            f"PJサマリ_単・累計_{pszProjectCode}.xlsx",
        )
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    objCandidates.sort()
    pszTargetPath = os.path.join(pszProjectDirectory, objCandidates[0])
    os.startfile(pszTargetPath)


def handle_project_pl_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszProjectDirectory = os.path.join(pszExecutionRoot, "プロジェクト損益")
    if not os.path.isdir(pszProjectDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszProjectDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszProjectDirectory)


def choose_pj_income_statement_file(
    pszTargetDirectory: str,
) -> Optional[str]:
    objCandidates = [
        pszName
        for pszName in os.listdir(pszTargetDirectory)
        if pszName.startswith("販管費配賦後_損益計算書_") and pszName.endswith(".xlsx")
    ]
    objCandidates.sort()
    if not objCandidates:
        return None

    objResult: Dict[str, Optional[str]] = {"file": None}

    def on_select(event: tk.Event) -> None:
        objSelection = objListBox.curselection()
        if not objSelection:
            return
        objEntryVar.set(objListBox.get(objSelection[0]))

    def on_confirm() -> None:
        pszFileName = objEntryVar.get().strip()
        if pszFileName == "":
            return
        objResult["file"] = pszFileName
        objWindow.grab_release()
        objWindow.destroy()

    def on_cancel() -> None:
        objWindow.grab_release()
        objWindow.destroy()

    objWindow = tk.Tk()
    objWindow.title("PJ別損益計算書: ファイル選択")
    objWindow.geometry("900x600")
    objWindow.resizable(True, True)
    objWindow.attributes("-topmost", True)
    objWindow.grab_set()
    objWindow.focus_force()
    objWindow.lift()
    objWindow.protocol("WM_DELETE_WINDOW", on_cancel)

    objFrame = tk.Frame(objWindow)
    objFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    objLabel = tk.Label(
        objFrame,
        text="ファイルを一覧から選択してください。",
        anchor="w",
    )
    objLabel.pack(fill=tk.X)

    objListFrame = tk.Frame(objFrame)
    objListFrame.pack(fill=tk.BOTH, expand=True, pady=5)

    objScrollBar = tk.Scrollbar(objListFrame)
    objScrollBar.pack(side=tk.RIGHT, fill=tk.Y)

    objListBox = tk.Listbox(
        objListFrame,
        yscrollcommand=objScrollBar.set,
        selectmode=tk.SINGLE,
    )
    objListBox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    objScrollBar.config(command=objListBox.yview)

    for pszName in objCandidates:
        objListBox.insert(tk.END, pszName)

    objListBox.bind("<<ListboxSelect>>", on_select)

    objEntryVar = tk.StringVar()
    objEntry = tk.Entry(objFrame, textvariable=objEntryVar)
    objEntry.pack(fill=tk.X, pady=5)

    objButtonFrame = tk.Frame(objFrame)
    objButtonFrame.pack(fill=tk.X, pady=5)

    objOkButton = tk.Button(objButtonFrame, text="OK", width=12, command=on_confirm)
    objOkButton.pack(side=tk.RIGHT, padx=5)
    objCancelButton = tk.Button(
        objButtonFrame, text="Cancel", width=12, command=on_cancel
    )
    objCancelButton.pack(side=tk.RIGHT, padx=5)

    objWindow.mainloop()
    return objResult["file"]


def handle_pj_income_statement_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszTargetDirectory = os.path.join(pszExecutionRoot, "PJ別損益計算書")
    if not os.path.isdir(pszTargetDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszTargetDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return

    def parse_selected_range_file(
        pszPath: str,
    ) -> Optional[Tuple[Tuple[int, int], Tuple[int, int]]]:
        try:
            with open(pszPath, "r", encoding="utf-8", newline="") as objFile:
                pszContent: str = objFile.read()
        except OSError:
            return None

        objMatch = re.search(r"開始:\s*(\d{4})/(\d{2}).*?終了:\s*(\d{4})/(\d{2})", pszContent, re.DOTALL)
        if objMatch is None:
            objMatch = re.search(r"採用範囲:\s*(\d{4})年(\d{1,2})月〜(\d{4})年(\d{1,2})月", pszContent)
            if objMatch is None:
                return None

        iStartYear: int = int(objMatch.group(1))
        iStartMonth: int = int(objMatch.group(2))
        iEndYear: int = int(objMatch.group(3))
        iEndMonth: int = int(objMatch.group(4))
        if not (1 <= iStartMonth <= 12 and 1 <= iEndMonth <= 12):
            return None
        return (iStartYear, iStartMonth), (iEndYear, iEndMonth)

    pszPeriodDirectory = os.path.join(pszExecutionRoot, "期間")
    objRangeFileNames: List[str] = [
        "SellGeneralAdminCost_Allocation_Cmd_SelectedRange.txt",
        "SellGeneralAdminCost_Allocation_DnD_SelectedRange.txt",
    ]

    pszTargetPath: Optional[str] = None
    for pszRangeFileName in objRangeFileNames:
        pszRangePath: str = os.path.join(pszPeriodDirectory, pszRangeFileName)
        if not os.path.isfile(pszRangePath):
            continue
        objSelectedRange = parse_selected_range_file(pszRangePath)
        if objSelectedRange is None:
            continue
        _, (iEndYear, iEndMonth) = objSelectedRange
        pszEndLabel: str = f"{iEndYear}年{iEndMonth:02d}月"
        pszTargetPath = os.path.join(
            pszTargetDirectory,
            f"販管費配賦後_損益計算書_{pszEndLabel}_A∪B_プロジェクト名_C∪D_両方.xlsx",
        )
        break

    if pszTargetPath is None:
        show_error_message_box(
            "Error: 対象期間を取得できませんでした。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return

    if not os.path.isfile(pszTargetPath):
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return

    os.startfile(pszTargetPath)


def handle_pj_income_statement_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszTargetDirectory = os.path.join(pszExecutionRoot, "PJ別損益計算書")
    if not os.path.isdir(pszTargetDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszTargetDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszTargetDirectory)


def handle_group_pl_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszGroupDirectory = os.path.join(pszExecutionRoot, "グループ別損益")
    pszTargetPath = os.path.join(
        pszGroupDirectory,
        "PJサマリ_グループ別合計.xlsx",
    )
    if not os.path.isfile(pszTargetPath):
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszTargetPath)


def handle_group_pl_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszGroupDirectory = os.path.join(pszExecutionRoot, "グループ別損益")
    if not os.path.isdir(pszGroupDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszGroupDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszGroupDirectory)


def handle_company_pl_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszCompanyDirectory = os.path.join(pszExecutionRoot, "カンパニー別損益")
    pszTargetPath = os.path.join(
        pszCompanyDirectory,
        "PJサマリ_カンパニー別合計.xlsx",
    )
    if not os.path.isfile(pszTargetPath):
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszTargetPath)


def handle_company_pl_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszCompanyDirectory = os.path.join(pszExecutionRoot, "カンパニー別損益")
    if not os.path.isdir(pszCompanyDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszCompanyDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszCompanyDirectory)


def handle_cp_management_company_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszCompanyDirectory = os.path.join(pszExecutionRoot, "CP別経営管理表_計上カンパニー")
    if not os.path.isdir(pszCompanyDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszCompanyDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszPrefix = "CP別経営管理_計上カンパニー_累計_"
    objCandidates = [
        pszName
        for pszName in os.listdir(pszCompanyDirectory)
        if pszName.startswith(pszPrefix) and pszName.endswith(".xlsx")
    ]
    if not objCandidates:
        pszTargetPath = os.path.join(
            pszCompanyDirectory,
            "CP別経営管理_計上カンパニー_累計_yyyy年mm月-yyyy年mm月.xlsx",
        )
        show_error_message_box(
            "Error: ファイルが見つかりません。\n" + pszTargetPath,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    objCandidates.sort()
    pszTargetPath = os.path.join(pszCompanyDirectory, objCandidates[-1])
    os.startfile(pszTargetPath)


def handle_cp_management_company_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszCompanyDirectory = os.path.join(pszExecutionRoot, "CP別経営管理表_計上カンパニー")
    if not os.path.isdir(pszCompanyDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszCompanyDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszCompanyDirectory)


def handle_cp_management_group_left_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return

    def parse_selected_range_file(
        pszPath: str,
    ) -> Optional[Tuple[Tuple[int, int], Tuple[int, int]]]:
        try:
            with open(pszPath, "r", encoding="utf-8", newline="") as objFile:
                pszContent: str = objFile.read()
        except OSError:
            return None

        objMatch = re.search(r"開始:\s*(\d{4})/(\d{2}).*?終了:\s*(\d{4})/(\d{2})", pszContent, re.DOTALL)
        if objMatch is None:
            objMatch = re.search(r"採用範囲:\s*(\d{4})年(\d{1,2})月〜(\d{4})年(\d{1,2})月", pszContent)
            if objMatch is None:
                return None

        iStartYear: int = int(objMatch.group(1))
        iStartMonth: int = int(objMatch.group(2))
        iEndYear: int = int(objMatch.group(3))
        iEndMonth: int = int(objMatch.group(4))
        if not (1 <= iStartMonth <= 12 and 1 <= iEndMonth <= 12):
            return None
        return (iStartYear, iStartMonth), (iEndYear, iEndMonth)

    pszGroupDirectory = os.path.join(pszExecutionRoot, "CP別経営管理表_計上グループ")
    pszPeriodDirectory = os.path.join(pszExecutionRoot, "期間")
    objRangeFileNames: List[str] = [
        "SellGeneralAdminCost_Allocation_Cmd_SelectedRange.txt",
        "SellGeneralAdminCost_Allocation_DnD_SelectedRange.txt",
    ]

    pszTargetPath: Optional[str] = None
    for pszRangeFileName in objRangeFileNames:
        pszRangePath: str = os.path.join(pszPeriodDirectory, pszRangeFileName)
        if not os.path.isfile(pszRangePath):
            continue
        objSelectedRange = parse_selected_range_file(pszRangePath)
        if objSelectedRange is None:
            continue
        (iStartYear, iStartMonth), (iEndYear, iEndMonth) = objSelectedRange
        pszStartLabel: str = f"{iStartYear}年{iStartMonth:02d}月"
        pszEndLabel: str = f"{iEndYear}年{iEndMonth:02d}月"
        pszTargetPath = os.path.join(
            pszGroupDirectory,
            f"CP別経営管理_計上グループ_累計_{pszStartLabel}-{pszEndLabel}.xlsx",
        )
        break

    if pszTargetPath is not None and os.path.isfile(pszTargetPath):
        os.startfile(pszTargetPath)
        return

    objFallbackPaths: List[str] = []
    if os.path.isdir(pszGroupDirectory):
        for pszName in os.listdir(pszGroupDirectory):
            if (
                pszName.startswith("CP別経営管理_計上グループ_累計_")
                and pszName.endswith(".xlsx")
            ):
                objFallbackPaths.append(os.path.join(pszGroupDirectory, pszName))

    if objFallbackPaths:
        objFallbackPaths.sort(
            key=lambda pszPath: (
                os.path.getmtime(pszPath),
                os.path.basename(pszPath),
            ),
        )
        os.startfile(objFallbackPaths[-1])
        return

    if pszTargetPath is None:
        pszTargetPath = os.path.join(
            pszGroupDirectory,
            "CP別経営管理_計上グループ_累計_*.xlsx",
        )
    show_error_message_box(
        "Error: ファイルが見つかりません。\n" + pszTargetPath,
        "SellGeneralAdminCost_Allocation_DnD",
    )


def handle_cp_management_group_right_down() -> None:
    pszExecutionRoot = find_latest_execution_root_directory()
    if pszExecutionRoot is None:
        show_error_message_box(
            "Error: 出力フォルダーがまだ作成されていません。",
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    pszGroupDirectory = os.path.join(pszExecutionRoot, "CP別経営管理表_計上グループ")
    if not os.path.isdir(pszGroupDirectory):
        show_error_message_box(
            "Error: フォルダーが見つかりません。\n" + pszGroupDirectory,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    os.startfile(pszGroupDirectory)


def update_action_button_layout(iWindowHandle: int) -> None:
    if not g_action_button_handles:
        return
    objClientRect = win32gui.GetClientRect(iWindowHandle)
    iButtonWidth: int = objClientRect[2] - BUTTON_MARGIN * 2
    if iButtonWidth < 120:
        iButtonWidth = 120
    iButtonX: int = BUTTON_MARGIN
    iButtonSpacing: int = BUTTON_MARGIN
    iTotalButtonsHeight: int = (
        BUTTON_HEIGHT * len(g_action_button_handles)
        + iButtonSpacing * (len(g_action_button_handles) - 1)
    )
    iButtonY: int = objClientRect[3] - BUTTON_MARGIN - iTotalButtonsHeight
    if iButtonY < BUTTON_MARGIN:
        iButtonY = BUTTON_MARGIN
    for iButtonHandle in g_action_button_handles:
        win32gui.MoveWindow(
            iButtonHandle,
            iButtonX,
            iButtonY,
            iButtonWidth,
            BUTTON_HEIGHT,
            True,
        )
        iButtonY += BUTTON_HEIGHT + iButtonSpacing


def handle_action_button_left_click(iButtonId: int) -> None:
    if iButtonId == BUTTON_ID_BASE + 0:
        handle_period_left_down()
    elif iButtonId == BUTTON_ID_BASE + 1:
        handle_company_results_left_down()
    elif iButtonId == BUTTON_ID_BASE + 2:
        handle_company_margin_rank_left_down()
    elif iButtonId == BUTTON_ID_BASE + 3:
        handle_project_pl_left_down()
    elif iButtonId == BUTTON_ID_BASE + 4:
        handle_group_pl_left_down()
    elif iButtonId == BUTTON_ID_BASE + 5:
        handle_company_pl_left_down()
    elif iButtonId == BUTTON_ID_BASE + 6:
        handle_cp_management_company_left_down()
    elif iButtonId == BUTTON_ID_BASE + 7:
        handle_cp_management_group_left_down()
    elif iButtonId == BUTTON_ID_BASE + 8:
        handle_pj_income_statement_left_down()


def handle_action_button_right_click(iButtonId: int) -> None:
    if iButtonId == BUTTON_ID_BASE + 0:
        handle_period_right_down()
    elif iButtonId == BUTTON_ID_BASE + 1:
        handle_company_results_right_down()
    elif iButtonId == BUTTON_ID_BASE + 2:
        handle_company_margin_rank_right_down()
    elif iButtonId == BUTTON_ID_BASE + 3:
        handle_project_pl_right_down()
    elif iButtonId == BUTTON_ID_BASE + 4:
        handle_group_pl_right_down()
    elif iButtonId == BUTTON_ID_BASE + 5:
        handle_company_pl_right_down()
    elif iButtonId == BUTTON_ID_BASE + 6:
        handle_cp_management_company_right_down()
    elif iButtonId == BUTTON_ID_BASE + 7:
        handle_cp_management_group_right_down()
    elif iButtonId == BUTTON_ID_BASE + 8:
        handle_pj_income_statement_right_down()


def set_right_button_down_handle(iControlHandle: Optional[int]) -> None:
    global g_right_button_down_handle
    g_right_button_down_handle = iControlHandle


def is_right_button_down(iControlHandle: int) -> bool:
    return g_right_button_down_handle == iControlHandle


def get_low_word(iValue: int) -> int:
    return iValue & 0xFFFF


def get_high_word(iValue: int) -> int:
    return (iValue >> 16) & 0xFFFF


def ensure_default_gui_font_handle() -> Optional[int]:
    global g_default_gui_font_handle
    if g_default_gui_font_handle is None:
        iFontId = getattr(win32con, "DEFAULT_GUI_FONT", 17)
        g_default_gui_font_handle = win32gui.GetStockObject(iFontId)
    return g_default_gui_font_handle


def create_action_buttons(iWindowHandle: int) -> None:
    global g_action_button_handles
    g_action_button_handles = []
    iFontHandle = ensure_default_gui_font_handle()
    for iIndex, pszLabel in enumerate(BUTTON_LABELS):
        iButtonHandle: int = win32gui.CreateWindowEx(
            win32con.WS_EX_NOPARENTNOTIFY,
            "BUTTON",
            pszLabel,
            win32con.WS_TABSTOP
            | win32con.WS_VISIBLE
            | win32con.WS_CHILD
            | win32con.BS_OWNERDRAW,
            0,
            0,
            120,
            BUTTON_HEIGHT,
            iWindowHandle,
            BUTTON_ID_BASE + iIndex,
            win32api.GetModuleHandle(None),
            None,
        )
        iFontHandle = win32gui.GetStockObject(win32con.DEFAULT_GUI_FONT)
        win32gui.SendMessage(
            iButtonHandle,
            win32con.WM_SETFONT,
            iFontHandle,
            True,
        )
        g_action_button_handles.append(iButtonHandle)
        if iFontHandle:
            win32gui.SendMessage(
                iButtonHandle,
                win32con.WM_SETFONT,
                iFontHandle,
                True,
            )
        win32gui.ShowWindow(iButtonHandle, win32con.SW_SHOWNORMAL)
    update_action_button_layout(iWindowHandle)


def ensure_action_button_brush() -> Optional[int]:
    global g_action_button_brush_handle
    if g_action_button_brush_handle is None:
        iRed, iGreen, iBlue = ACTION_BUTTON_COLOR
        g_action_button_brush_handle = win32gui.CreateSolidBrush(
            win32api.RGB(iRed, iGreen, iBlue)
        )
    return g_action_button_brush_handle


def build_unique_temp_path(pszDirectory: str, pszFileName: str) -> str:
    return os.path.join(pszDirectory, pszFileName)


def move_output_files_to_temp(pszStdOut: str) -> List[str]:
    pszTempDirectory: str = get_temp_output_directory()
    pszCmdDirectory: str = os.path.dirname(__file__)
    objMoved: List[str] = []
    objCostReportPatterns: List[re.Pattern[str]] = [
        re.compile(r"^累計_製造原価報告書_.*\.tsv$"),
        re.compile(r"^製造原価報告書_.*\.tsv$"),
    ]
    objIncomeStatementPatterns: List[re.Pattern[str]] = [
        re.compile(r"^累計_損益計算書_.*\.tsv$"),
    ]
    for pszLine in pszStdOut.splitlines():
        pszLineText: str = pszLine.strip()
        if not pszLineText.startswith("Output:"):
            continue
        pszOutputPath: str = pszLineText.replace("Output:", "", 1).strip()
        if pszOutputPath == "" or not os.path.isfile(pszOutputPath):
            continue
        pszBaseName: str = os.path.basename(pszOutputPath)
        pszTargetDirectory: str = pszTempDirectory
        if any(objPattern.match(pszBaseName) for objPattern in objCostReportPatterns):
            pszTargetDirectory = os.path.join(pszTempDirectory, "製造原価報告書系")
            os.makedirs(pszTargetDirectory, exist_ok=True)
        elif any(objPattern.match(pszBaseName) for objPattern in objIncomeStatementPatterns):
            pszTargetDirectory = os.path.join(pszTempDirectory, "損益計算書系")
            os.makedirs(pszTargetDirectory, exist_ok=True)
        pszTargetPath: str = build_unique_temp_path(pszTargetDirectory, pszBaseName)
        shutil.move(pszOutputPath, pszTargetPath)
        objMoved.append(pszTargetPath)
        pszBaseName = os.path.basename(pszTargetPath)
        if pszBaseName.startswith("累計_製造原価報告書_"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszBaseName)
            shutil.copy2(pszTargetPath, pszCopyPath)
        if (
            pszBaseName.startswith("損益計算書_")
            and pszBaseName.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv")
            and "販管費配賦_step" not in pszBaseName
        ):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszBaseName)
            shutil.copy2(pszTargetPath, pszCopyPath)

    pszIncomeStatementDirectory: str = os.path.join(pszTempDirectory, "損益計算書系")
    os.makedirs(pszIncomeStatementDirectory, exist_ok=True)
    for pszFileName in sorted(os.listdir(pszTempDirectory)):
        if not any(objPattern.match(pszFileName) for objPattern in objIncomeStatementPatterns):
            continue
        pszSourcePath: str = os.path.join(pszTempDirectory, pszFileName)
        if not os.path.isfile(pszSourcePath):
            continue
        pszDestinationPath: str = os.path.join(pszIncomeStatementDirectory, pszFileName)
        if os.path.exists(pszDestinationPath):
            os.remove(pszDestinationPath)
        shutil.move(pszSourcePath, pszDestinationPath)
        objMoved.append(pszDestinationPath)

    return objMoved


def parse_year_month_from_pl_csv(pszFilePath: str) -> Optional[Tuple[int, int]]:
    pszBaseName: str = os.path.basename(pszFilePath)
    objMatch = re.search(r"(\d{2})\.(\d{1,2})\.csv$", pszBaseName)
    if objMatch is None:
        return None
    iYear: int = 2000 + int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    if iMonth < 1 or iMonth > 12:
        return None
    return iYear, iMonth


def move_pl_outputs_to_temp(pszCsvPath: str) -> None:
    objYearMonth = parse_year_month_from_pl_csv(pszCsvPath)
    if objYearMonth is None:
        return
    iYear, iMonth = objYearMonth
    pszMonth: str = f"{iMonth:02d}"
    objPrefixes: List[str] = [
        f"損益計算書_{iYear}年{pszMonth}月",
        f"製造原価報告書_{iYear}年{pszMonth}月",
    ]
    pszSourceDirectory: str = os.path.dirname(pszCsvPath)
    pszTempDirectory: str = get_temp_output_directory()
    pszCmdDirectory: str = os.path.dirname(__file__)
    try:
        objEntries: List[str] = os.listdir(pszSourceDirectory)
    except OSError:
        return
    for pszEntry in objEntries:
        if not pszEntry.endswith(".tsv"):
            continue
        if not any(pszEntry.startswith(pszPrefix) for pszPrefix in objPrefixes):
            continue
        pszSourcePath: str = os.path.join(pszSourceDirectory, pszEntry)
        if not os.path.isfile(pszSourcePath):
            continue
        pszTargetPath: str = build_unique_temp_path(pszTempDirectory, pszEntry)
        shutil.move(pszSourcePath, pszTargetPath)
        if pszEntry.startswith("損益計算書_") and pszEntry.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)
        if pszEntry.startswith("製造原価報告書_") and pszEntry.endswith("_A∪B_プロジェクト名_C∪D.tsv"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)
        if pszEntry.startswith("製造原価報告書_") and pszEntry.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)


def move_manhour_outputs_to_temp(pszCsvPath: str) -> None:
    objYearMonth = parse_year_month_from_pl_csv(pszCsvPath)
    if objYearMonth is None:
        return
    iYear, iMonth = objYearMonth
    pszMonth: str = f"{iMonth:02d}"
    pszPrefix: str = f"工数_{iYear}年{pszMonth}月"
    pszSourceDirectory: str = os.path.dirname(pszCsvPath)
    pszTempDirectory: str = get_manhour_temp_output_directory()
    pszCmdDirectory: str = os.path.dirname(__file__)
    try:
        objEntries: List[str] = os.listdir(pszSourceDirectory)
    except OSError:
        return
    for pszEntry in objEntries:
        if not pszEntry.endswith(".tsv"):
            continue
        if not pszEntry.startswith(pszPrefix):
            continue
        pszSourcePath: str = os.path.join(pszSourceDirectory, pszEntry)
        if not os.path.isfile(pszSourcePath):
            continue
        pszTargetPath: str = build_unique_temp_path(pszTempDirectory, pszEntry)
        shutil.move(pszSourcePath, pszTargetPath)
        if pszEntry.startswith("工数_") and pszEntry.endswith("_step0014_各プロジェクトの計上カンパニー名_工数_カンパニーの工数.tsv"):
            pszCopyPath = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)


def build_pl_tsv_base_name(iYear: int, iMonth: int) -> str:
    pszMonth: str = f"{iMonth:02d}"
    return f"損益計算書_{iYear}年{pszMonth}月_A∪B_プロジェクト名_C∪D_vertical.tsv"


def find_pl_tsv_paths_for_year_months(objYearMonthTexts: List[str]) -> List[str]:
    if not objYearMonthTexts:
        return []
    pszBaseDirectory: str = os.path.dirname(__file__)
    pszTempDirectory: str = get_temp_output_directory()
    objFound: List[str] = []
    objSeen: set[str] = set()
    objDirectories: List[str] = [pszBaseDirectory, pszTempDirectory]
    for pszYearMonthText in objYearMonthTexts:
        objValue = parse_year_month_value(pszYearMonthText)
        if objValue is None:
            continue
        iYear, iMonth = objValue
        pszBaseName: str = build_pl_tsv_base_name(iYear, iMonth)
        for pszDirectory in objDirectories:
            pszCandidate: str = os.path.join(pszDirectory, pszBaseName)
            if not os.path.isfile(pszCandidate):
                continue
            if pszCandidate in objSeen:
                continue
            objFound.append(pszCandidate)
            objSeen.add(pszCandidate)
    return objFound


def parse_year_month_from_name(pszBaseName: str) -> Optional[str]:
    iPrefixIndex: int = pszBaseName.find("_")
    if iPrefixIndex < 0:
        return None
    iSecondIndex: int = pszBaseName.find("_", iPrefixIndex + 1)
    if iSecondIndex < 0:
        return None
    pszYearMonth: str = pszBaseName[iPrefixIndex + 1 : iSecondIndex]
    if "年" not in pszYearMonth or "月" not in pszYearMonth:
        return None
    return pszYearMonth


def parse_year_month_value(pszYearMonth: str) -> Optional[Tuple[int, int]]:
    try:
        iYearText: str = pszYearMonth.split("年", 1)[0]
        iMonthText: str = pszYearMonth.split("年", 1)[1].split("月", 1)[0]
        iYear: int = int(iYearText)
        iMonth: int = int(iMonthText)
    except (ValueError, IndexError):
        return None
    if iMonth < 1 or iMonth > 12:
        return None
    return iYear, iMonth


def is_pl_csv_file(pszBaseName: str) -> bool:
    pszNormalized: str = pszBaseName.lower()
    return re.fullmatch(r"損益計算書\d{2}\.\d{1,2}\.csv", pszNormalized) is not None


def is_manhour_csv_file(pszBaseName: str) -> bool:
    pszNormalized: str = pszBaseName.lower()
    return re.fullmatch(r"工数\d{2}\.\d{1,2}\.csv", pszNormalized) is not None


def is_step14_tsv_file(pszBaseName: str) -> bool:
    if not pszBaseName.startswith("工数_"):
        return False
    if pszBaseName.endswith("_step0014_各プロジェクトの計上カンパニー名_工数_カンパニーの工数.tsv"):
        return True
    return pszBaseName.endswith("_step15_各プロジェクトの工数.tsv")


def is_pl_tsv_file(pszBaseName: str) -> bool:
    return pszBaseName.startswith("損益計算書_") and pszBaseName.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv")


def is_consecutive_months(objYearMonths: List[Tuple[int, int]]) -> bool:
    if not objYearMonths:
        return False
    for iIndex in range(1, len(objYearMonths)):
        iPrevYear, iPrevMonth = objYearMonths[iIndex - 1]
        iNextYear, iNextMonth = objYearMonths[iIndex]
        iExpectedYear: int = iPrevYear
        iExpectedMonth: int = iPrevMonth + 1
        if iExpectedMonth == 13:
            iExpectedMonth = 1
            iExpectedYear += 1
        if iNextYear != iExpectedYear or iNextMonth != iExpectedMonth:
            return False
    return True


def collect_valid_pairs(
    objFilePaths: List[str],
) -> List[Tuple[str, str, Tuple[int, int], str]]:
    objManhourMap: Dict[str, str] = {}
    objPlMap: Dict[str, str] = {}
    for pszFilePath in objFilePaths:
        pszBaseName: str = os.path.basename(pszFilePath)
        if pszBaseName.startswith("工数_"):
            pszYearMonth: Optional[str] = parse_year_month_from_name(pszBaseName)
            if pszYearMonth is None:
                return []
            objManhourMap[pszYearMonth] = pszFilePath
            continue
        if pszBaseName.startswith("損益計算書_"):
            pszYearMonth = parse_year_month_from_name(pszBaseName)
            if pszYearMonth is None:
                return []
            objPlMap[pszYearMonth] = pszFilePath
            continue
        return []

    objPairs: List[Tuple[str, str, Tuple[int, int], str]] = []
    for pszYearMonth, pszManhourPath in objManhourMap.items():
        if pszYearMonth not in objPlMap:
            continue
        objValue: Optional[Tuple[int, int]] = parse_year_month_value(pszYearMonth)
        if objValue is None:
            continue
        objPairs.append((pszManhourPath, objPlMap[pszYearMonth], objValue, pszYearMonth))
    return objPairs


def select_consecutive_pairs(
    objPairs: List[Tuple[str, str, Tuple[int, int], str]],
) -> List[Tuple[str, str, Tuple[int, int], str]]:
    if not objPairs:
        return []
    objPairsSorted = sorted(objPairs, key=lambda objItem: objItem[2])
    objYearMonths: List[Tuple[int, int]] = [objItem[2] for objItem in objPairsSorted]
    if not is_consecutive_months(objYearMonths):
        return []
    return objPairsSorted


def build_cmd_args(objPairs: List[Tuple[str, str, Tuple[int, int], str]]) -> List[str]:
    objManhourFiles: List[str] = [objItem[0] for objItem in objPairs]
    objPlFiles: List[str] = [objItem[1] for objItem in objPairs]
    objArgs: List[str] = []
    objArgs.extend(objManhourFiles)
    objArgs.extend(objPlFiles)
    return objArgs


def write_selected_range_file(
    objPairs: List[Tuple[str, str, Tuple[int, int], str]],
) -> Optional[str]:
    if not objPairs:
        return None
    iStartYear: int = objPairs[0][2][0]
    iStartMonth: int = objPairs[0][2][1]
    iEndYear: int = objPairs[-1][2][0]
    iEndMonth: int = objPairs[-1][2][1]
    pszStartText: str = f"{iStartYear:04d}/{iStartMonth:02d}"
    pszEndText: str = f"{iEndYear:04d}/{iEndMonth:02d}"
    objLines: List[str] = [
        "採用範囲:",
        f"開始: {pszStartText}",
        f"終了: {pszEndText}",
    ]
    pszOutputDirectory: str = os.path.dirname(os.path.abspath(__file__))
    pszOutputFileName: str = "SellGeneralAdminCost_Allocation_DnD_SelectedRange.txt"
    pszOutputPath: str = os.path.join(pszOutputDirectory, pszOutputFileName)
    with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutputFile:
        objOutputFile.write("\n".join(objLines) + "\n")
    return pszOutputPath


def run_allocation_with_pairs(
    objPairs: List[Tuple[str, str, Tuple[int, int], str]],
) -> int:
    if not objPairs:
        return 1

    pszRangePath: Optional[str] = write_selected_range_file(objPairs)
    objArgs: List[str] = build_cmd_args(objPairs)
    pszScriptPath: str = os.path.join(os.path.dirname(__file__), "SellGeneralAdminCost_Allocation_Cmd_0002.py")
    objCommand: List[str] = [sys.executable, pszScriptPath]
    objCommand.extend(objArgs)

    try:
        objResult = subprocess.run(
            objCommand,
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as exc:  # noqa: BLE001
        pszErrorMessage: str = (
            "Error: unexpected exception while running SellGeneralAdminCost_Allocation_Cmd_0002.py. Detail = "
            + str(exc)
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    if objResult.returncode != 0:
        pszStdErr: str = objResult.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        pszErrorMessage = (
            "Error: SellGeneralAdminCost_Allocation_Cmd_0002.py exited with non-zero return code.\n\n"
            + "Return code = "
            + str(objResult.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return objResult.returncode

    pszStdOut: str = objResult.stdout
    objMoved = move_output_files_to_temp(pszStdOut)
    if objMoved:
        set_last_output_directory(os.path.dirname(objMoved[0]))
    if pszStdOut.strip() != "":
        print(pszStdOut)
    pszStdOut = "成功しました！"
    if pszRangePath is not None:
        pszStdOut += "\n\n採用範囲を記録しました: " + pszRangePath
    show_message_box(pszStdOut, "SellGeneralAdminCost_Allocation_DnD")
    return 0


def run_pl_csv_to_tsv(
    objCsvFiles: List[str],
) -> int:
    pszScriptPath: str = os.path.join(os.path.dirname(__file__), "PL_CsvToTsv_Cmd_0002.py")
    if not os.path.exists(pszScriptPath):
        pszErrorMessage: str = (
            "Error: PL_CsvToTsv_Cmd_0002.py not found. Path = " + pszScriptPath
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    objCommand: List[str] = [sys.executable, pszScriptPath] + objCsvFiles
    append_error_log("Running: " + " ".join(objCommand))
    try:
        objResult = subprocess.run(
            objCommand,
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as exc:  # noqa: BLE001
        pszErrorMessage: str = (
            "Error: unexpected exception while running PL_CsvToTsv_Cmd_0002.py. Detail = "
            + str(exc)
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    if objResult.returncode != 0:
        pszStdErr: str = objResult.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        pszErrorMessage = (
            "Error: PL_CsvToTsv_Cmd_0002.py exited with non-zero return code.\n\n"
            + "Return code = "
            + str(objResult.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return objResult.returncode

    pszStdOut: str = objResult.stdout.strip()
    if pszStdOut != "":
        print(pszStdOut)
        move_output_files_to_temp(pszStdOut)

    for pszCsvPath in objCsvFiles:
        move_pl_outputs_to_temp(pszCsvPath)

    pszMessage: str = "PL_CsvToTsv_Cmd_0002.py finished successfully."
    if pszStdOut != "":
        pszMessage = pszStdOut
    show_message_box(pszMessage, "SellGeneralAdminCost_Allocation_DnD")
    return 0


def run_manhour_csv_to_sheet(
    objCsvFiles: List[str],
) -> int:
    pszScriptPath: str = os.path.join(
        os.path.dirname(__file__),
        "make_manhour_to_sheet8_01_0003.py",
    )
    if not os.path.exists(pszScriptPath):
        pszErrorMessage: str = (
            "Error: make_manhour_to_sheet8_01_0003.py not found. Path = "
            + pszScriptPath
        )
        append_error_log(pszErrorMessage)
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    objMessages: List[str] = []
    for pszCsvPath in objCsvFiles:
        objCommand: List[str] = [sys.executable, pszScriptPath, pszCsvPath]
        try:
            objResult = subprocess.run(
                objCommand,
                check=False,
                capture_output=True,
                text=True,
            )
        except Exception as exc:  # noqa: BLE001
            pszErrorMessage: str = (
                "Error: unexpected exception while running make_manhour_to_sheet8_01_0003.py. Detail = "
                + str(exc)
            )
            append_error_log(pszErrorMessage)
            show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
            return 1

        if objResult.returncode != 0:
            pszStdErr: str = objResult.stderr
            if pszStdErr.strip() == "":
                pszStdErr = "Process exited with non-zero return code and no stderr output."
            pszErrorMessage = (
                "Error: make_manhour_to_sheet8_01_0003.py exited with non-zero return code.\n\n"
                + "Return code = "
                + str(objResult.returncode)
                + "\n\n"
                + "stderr:\n"
                + pszStdErr
            )
            append_error_log(pszErrorMessage)
            show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
            return objResult.returncode

        pszStdOut: str = objResult.stdout.strip()
        if pszStdOut != "":
            print(pszStdOut)
            move_output_files_to_temp(pszStdOut)
        move_manhour_outputs_to_temp(pszCsvPath)

    pszMessage: str = "make_manhour_to_sheet8_01_0003.py finished successfully."
    show_message_box(pszMessage, "SellGeneralAdminCost_Allocation_DnD")
    return 0


def run_step10_tsv_only(
    objStep10Files: List[str],
) -> int:
    pszScriptPath: str = os.path.join(
        os.path.dirname(__file__),
        "make_manhour_to_sheet8_01_0003.py",
    )
    if not os.path.exists(pszScriptPath):
        pszErrorMessage: str = (
            "Error: make_manhour_to_sheet8_01_0003.py not found. Path = "
            + pszScriptPath
        )
        append_error_log(pszErrorMessage)
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    iExitCode: int = 0
    for pszStep10Path in objStep10Files:
        objCommand: List[str] = [sys.executable, pszScriptPath, pszStep10Path]
        try:
            objResult = subprocess.run(
                objCommand,
                check=False,
                capture_output=True,
                text=True,
            )
        except Exception as exc:  # noqa: BLE001
            pszErrorMessage: str = (
                "Error: unexpected exception while running make_manhour_to_sheet8_01_0003.py. Detail = "
                + str(exc)
            )
            append_error_log(pszErrorMessage)
            show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
            iExitCode = 1
            continue

        if objResult.returncode != 0:
            pszStdErr: str = objResult.stderr
            if pszStdErr.strip() == "":
                pszStdErr = "Process exited with non-zero return code and no stderr output."
            pszErrorMessage = (
                "Error: make_manhour_to_sheet8_01_0003.py exited with non-zero return code.\n\n"
                + "Return code = "
                + str(objResult.returncode)
                + "\n\n"
                + "stderr:\n"
                + pszStdErr
            )
            append_error_log(pszErrorMessage)
            show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
            iExitCode = 1
            continue

        pszStdOut: str = objResult.stdout.strip()
        if pszStdOut != "":
            print(pszStdOut)
            move_output_files_to_temp(pszStdOut)
        move_manhour_outputs_to_temp(pszStep10Path)

    if iExitCode == 0:
        pszMessage: str = "Manhour TSV only flow finished successfully."
        show_message_box(pszMessage, "SellGeneralAdminCost_Allocation_DnD")
    return iExitCode


def draw_instruction_text(
    iWindowHandle: int,
) -> None:
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(
        iWindowHandle,
    )
    objClientRect = win32gui.GetClientRect(
        iWindowHandle,
    )
    iFontHandle = ensure_default_gui_font_handle()
    iPreviousFontHandle = None
    if iFontHandle:
        iPreviousFontHandle = win32gui.SelectObject(
            iDeviceContextHandle,
            iFontHandle,
        )

    iMargin: int = 5
    iBottomReserved: int = 0
    objClientRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin - iBottomReserved,
    )

    pszInstructionText: str = (
        "工数TSVと損益計算書TSVを、このウィンドウにドラッグ＆ドロップしてください。\n"
        "有効な年月の連続範囲のみ処理されます。\n"
        "採用された年月範囲はテキストファイルに記録します。"
    )

    iDrawTextFormat: int = win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK
    win32gui.DrawText(
        iDeviceContextHandle,
        pszInstructionText,
        -1,
        objClientRect,
        iDrawTextFormat,
    )
    if iPreviousFontHandle:
        win32gui.SelectObject(
            iDeviceContextHandle,
            iPreviousFontHandle,
        )
    win32gui.EndPaint(
        iWindowHandle,
        objPaintStruct,
    )


def window_proc(
    iWindowHandle: int,
    iMessage: int,
    iWparam: int,
    iLparam: int,
) -> int:
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(
            iWindowHandle,
            True,
        )
        if not g_action_button_handles:
            create_action_buttons(iWindowHandle)
        return 0

    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        iFileCount: int = win32api.DragQueryFile(
            iDropHandle,
            -1,
        )

        objFiles: List[str] = []
        for iIndex in range(iFileCount):
            pszFilePath: str = win32api.DragQueryFile(
                iDropHandle,
                iIndex,
            )
            objFiles.append(pszFilePath)

        win32api.DragFinish(iDropHandle)

        objCsvFiles: List[str] = []
        objManhourCsvFiles: List[str] = []
        objStep14TsvFiles: List[str] = []
        objPlTsvFiles: List[str] = []
        objUnexpectedFiles: List[str] = []
        bAllCsv: bool = True
        bAllManhourCsv: bool = True
        for pszFilePath in objFiles:
            pszBaseName: str = os.path.basename(pszFilePath)
            if is_pl_csv_file(pszBaseName):
                objCsvFiles.append(pszFilePath)
            else:
                bAllCsv = False
            if is_manhour_csv_file(pszBaseName):
                objManhourCsvFiles.append(pszFilePath)
            else:
                bAllManhourCsv = False
            if is_step14_tsv_file(pszBaseName):
                objStep14TsvFiles.append(pszFilePath)
            elif is_pl_tsv_file(pszBaseName):
                objPlTsvFiles.append(pszFilePath)
            elif not (is_pl_csv_file(pszBaseName) or is_manhour_csv_file(pszBaseName)):
                objUnexpectedFiles.append(pszFilePath)

        if objUnexpectedFiles:
            pszErrorMessage = "Error: unexpected or mixed file types detected.\n"
            pszErrorMessage += "\n".join(objUnexpectedFiles)
            show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
            return 0

        objManhourTsvFiles: List[str] = objStep14TsvFiles

        if bAllCsv and objCsvFiles and not (objStep14TsvFiles or objPlTsvFiles):
            run_pl_csv_to_tsv(objCsvFiles)
            return 0
        if bAllManhourCsv and objManhourCsvFiles and not (objStep14TsvFiles or objPlTsvFiles):
            run_manhour_csv_to_sheet(objManhourCsvFiles)
            return 0

        if objManhourTsvFiles and not (objPlTsvFiles or objCsvFiles or objManhourCsvFiles):
            run_step10_tsv_only(objManhourTsvFiles)
            return 0

        if objManhourTsvFiles and objPlTsvFiles and not (objCsvFiles or objManhourCsvFiles):
            objPairs = collect_valid_pairs(objManhourTsvFiles + objPlTsvFiles)
            objPairs = select_consecutive_pairs(objPairs)
            if not objPairs:
                pszErrorMessage = (
                    "Error: dropped Step0014 TSV and PL TSV files are invalid or not consecutive by year/month."
                )
                show_error_message_box(
                    pszErrorMessage,
                    "SellGeneralAdminCost_Allocation_DnD",
                )
                return 0
            run_allocation_with_pairs(objPairs)
            return 0

        if objManhourTsvFiles and objCsvFiles and not objPlTsvFiles and not objManhourCsvFiles:
            iPlExitCode: int = run_pl_csv_to_tsv(objCsvFiles)
            if iPlExitCode != 0:
                return 0
            objYearMonthsText: List[str] = []
            for pszManhourPath in objManhourTsvFiles:
                pszBaseName = os.path.basename(pszManhourPath)
                pszYearMonth = parse_year_month_from_name(pszBaseName)
                if pszYearMonth is None:
                    continue
                objYearMonthsText.append(pszYearMonth)
            objPlCandidates: List[str] = find_pl_tsv_paths_for_year_months(objYearMonthsText)
            if not objPlCandidates:
                pszErrorMessage = (
                    "Error: PL TSV files generated from CSV were not found for the dropped manhour files."
                )
                show_error_message_box(
                    pszErrorMessage,
                    "SellGeneralAdminCost_Allocation_DnD",
                )
                return 0
            objPairs = collect_valid_pairs(objManhourTsvFiles + objPlCandidates)
            objPairs = select_consecutive_pairs(objPairs)
            if not objPairs:
                pszErrorMessage = (
                    "Error: dropped manhour TSV and generated PL TSV files are invalid or not consecutive by year/month."
                )
                show_error_message_box(
                    pszErrorMessage,
                    "SellGeneralAdminCost_Allocation_DnD",
                )
                return 0
            run_allocation_with_pairs(objPairs)
            return 0

        objPairs = collect_valid_pairs(objFiles)
        objPairs = select_consecutive_pairs(objPairs)
        if not objPairs:
            pszErrorMessage: str = (
                "Error: dropped files are invalid or not consecutive by year/month."
            )
            show_error_message_box(
                pszErrorMessage,
                "SellGeneralAdminCost_Allocation_DnD",
            )
            return 0

        run_allocation_with_pairs(objPairs)
        return 0

    if iMessage == win32con.WM_COMMAND:
        iButtonId = get_low_word(iWparam)
        iNotifyCode = get_high_word(iWparam)
        if iNotifyCode == win32con.BN_CLICKED:
            handle_action_button_left_click(iButtonId)
            return 0

    if iMessage == win32con.WM_DRAWITEM:
        objDrawItem = DRAWITEMSTRUCT.from_address(iLparam)
        iDeviceContextHandle = objDrawItem.hDC
        iLeft = objDrawItem.rcItem.left
        iTop = objDrawItem.rcItem.top
        iRight = objDrawItem.rcItem.right
        iBottom = objDrawItem.rcItem.bottom
        iItemState = objDrawItem.itemState
        pszButtonText = win32gui.GetWindowText(objDrawItem.hwndItem)
        iFontHandle = ensure_default_gui_font_handle()
        iPreviousFontHandle = None
        if iFontHandle:
            iPreviousFontHandle = win32gui.SelectObject(
                iDeviceContextHandle,
                iFontHandle,
            )
        objBrushHandle = ensure_action_button_brush()
        if objBrushHandle is not None:
            win32gui.FillRect(
                iDeviceContextHandle,
                (iLeft, iTop, iRight, iBottom),
                objBrushHandle,
            )
        if iItemState & win32con.ODS_SELECTED or is_right_button_down(objDrawItem.hwndItem):
            win32gui.DrawEdge(
                iDeviceContextHandle,
                (iLeft, iTop, iRight, iBottom),
                win32con.EDGE_SUNKEN,
                win32con.BF_RECT,
            )
            iLeft += 2
            iTop += 2
        else:
            win32gui.DrawEdge(
                iDeviceContextHandle,
                (iLeft, iTop, iRight, iBottom),
                win32con.EDGE_RAISED,
                win32con.BF_RECT,
            )
        if pszButtonText == "実行":
            win32gui.SetTextColor(iDeviceContextHandle, win32api.RGB(255, 255, 255))
        else:
            win32gui.SetTextColor(iDeviceContextHandle, win32api.RGB(0, 0, 0))
        win32gui.SetBkMode(iDeviceContextHandle, win32con.TRANSPARENT)
        win32gui.DrawText(
            iDeviceContextHandle,
            pszButtonText,
            -1,
            (iLeft, iTop, iRight, iBottom),
            win32con.DT_CENTER | win32con.DT_VCENTER | win32con.DT_SINGLELINE,
        )
        if iPreviousFontHandle:
            win32gui.SelectObject(
                iDeviceContextHandle,
                iPreviousFontHandle,
            )
        return 1

    if iMessage == win32con.WM_CONTEXTMENU:
        iControlHandle: int = iWparam
        if iControlHandle in g_action_button_handles:
            iButtonId = win32gui.GetDlgCtrlID(iControlHandle)
            handle_action_button_right_click(iButtonId)
            return 0

    if iMessage == win32con.WM_PARENTNOTIFY:
        iEvent = get_low_word(iWparam)
        if iEvent == win32con.WM_LBUTTONDBLCLK:
            iControlHandle: int = iLparam
            if iControlHandle in g_action_button_handles:
                iButtonId = win32gui.GetDlgCtrlID(iControlHandle)
                if iButtonId == BUTTON_ID_BASE + 0:
                    handle_period_left_double_click()
                    return 0
        if iEvent in (win32con.WM_RBUTTONDOWN, win32con.WM_RBUTTONUP):
            iControlHandle: int = iLparam
            if iControlHandle in g_action_button_handles:
                if iEvent == win32con.WM_RBUTTONDOWN:
                    set_right_button_down_handle(iControlHandle)
                else:
                    set_right_button_down_handle(None)
                win32gui.InvalidateRect(iControlHandle, None, True)
                return 0

    if iMessage == win32con.WM_SIZE:
        update_action_button_layout(iWindowHandle)
        return 0

    if iMessage == win32con.WM_CTLCOLORBTN:
        iDeviceContextHandle: int = iWparam
        iControlHandle: int = iLparam
        if iControlHandle in g_action_button_handles:
            win32gui.SetBkMode(iDeviceContextHandle, win32con.TRANSPARENT)
            win32gui.SetTextColor(iDeviceContextHandle, win32api.RGB(0, 0, 0))
            objBrushHandle = ensure_action_button_brush()
            if objBrushHandle is not None:
                return objBrushHandle
    if iMessage == win32con.WM_DRAWITEM:
        objDrawItem = win32gui.PyDRAWITEMSTRUCT(iLparam)
        if objDrawItem.hwndItem in g_action_button_handles:
            iDeviceContextHandle = objDrawItem.hDC
            objRect = objDrawItem.rcItem
            objBrushHandle = ensure_action_button_brush()
            if objBrushHandle is not None:
                win32gui.FillRect(iDeviceContextHandle, objRect, objBrushHandle)
            win32gui.SetBkMode(iDeviceContextHandle, win32con.TRANSPARENT)
            win32gui.SetTextColor(iDeviceContextHandle, win32api.RGB(0, 0, 0))
            iFontHandle = win32gui.GetStockObject(win32con.DEFAULT_GUI_FONT)
            iPreviousFont = win32gui.SelectObject(iDeviceContextHandle, iFontHandle)
            pszText: str = win32gui.GetWindowText(objDrawItem.hwndItem)
            win32gui.DrawText(
                iDeviceContextHandle,
                pszText,
                -1,
                objRect,
                win32con.DT_CENTER | win32con.DT_VCENTER | win32con.DT_SINGLELINE,
            )
            if iPreviousFont:
                win32gui.SelectObject(iDeviceContextHandle, iPreviousFont)
            if objDrawItem.itemState & win32con.ODS_FOCUS:
                win32gui.DrawFocusRect(iDeviceContextHandle, objRect)
            return 1

    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(
            iWindowHandle,
        )
        return 0

    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0

    return win32gui.DefWindowProc(
        iWindowHandle,
        iMessage,
        iWparam,
        iLparam,
    )


def register_window_class(
    pszWindowClassName: str,
) -> int:
    iInstanceHandle: int = win32api.GetModuleHandle(None)

    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1

    iClassAtom: int = win32gui.RegisterClass(objWndClass)
    return iClassAtom


def create_main_window(
    pszWindowClassName: str,
    pszWindowTitle: str,
) -> int:
    global g_main_window_handle
    iInstanceHandle: int = win32api.GetModuleHandle(None)

    iWindowStyle: int = (
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX
    )
    iWindowExStyle: int = win32con.WS_EX_ACCEPTFILES

    iWindowPosX: int = win32con.CW_USEDEFAULT
    iWindowPosY: int = win32con.CW_USEDEFAULT
    iDesktopWidth: int = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    iDesktopHeight: int = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    iWindowWidth: int = int(iDesktopWidth * 0.5)
    iWindowHeight: int = int(iDesktopHeight * 0.5)
    iParentWindowHandle: int = 0
    iMenuHandle: int = 0

    iWindowHandle: int = win32gui.CreateWindowEx(
        iWindowExStyle,
        pszWindowClassName,
        pszWindowTitle,
        iWindowStyle,
        iWindowPosX,
        iWindowPosY,
        iWindowWidth,
        iWindowHeight,
        iParentWindowHandle,
        iMenuHandle,
        iInstanceHandle,
        None,
    )
    g_main_window_handle = iWindowHandle

    if not g_action_button_handles:
        create_action_buttons(iWindowHandle)

    win32gui.ShowWindow(
        iWindowHandle,
        win32con.SW_SHOWNORMAL,
    )
    win32gui.UpdateWindow(iWindowHandle)
    update_action_button_layout(iWindowHandle)
    win32gui.InvalidateRect(iWindowHandle, None, True)

    iHwndInsertAfter: int = win32con.HWND_TOPMOST
    iFlags: int = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    win32gui.SetWindowPos(
        iWindowHandle,
        iHwndInsertAfter,
        0,
        0,
        0,
        0,
        iFlags,
    )

    win32gui.DragAcceptFiles(
        iWindowHandle,
        True,
    )

    return iWindowHandle


def main() -> None:
    pszWindowClassName: str = "SellGeneralAdminCostAllocationDndWindowClass"
    pszWindowTitle: str = "SellGeneralAdminCost Allocation (Drag & Drop)"

    try:
        register_window_class(pszWindowClassName)
    except Exception as exc:
        report_exception("failed to register window class", exc)
        return

    try:
        create_main_window(
            pszWindowClassName,
            pszWindowTitle,
        )
    except Exception as exc:
        report_exception("failed to create main window", exc)
        return

    try:
        win32gui.PumpMessages()
    except Exception as exc:
        report_exception("unexpected exception in message loop", exc)
        return
    return


if __name__ == "__main__":
    main()
