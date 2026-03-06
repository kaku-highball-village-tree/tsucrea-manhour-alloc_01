# -*- coding: utf-8 -*-
"""
make_manhour_to_sheet8_01_0002.py

役割:
  単一のジョブカン工数 CSV を入力として、
  工数_yyyy年mm月.tsv を同一フォルダに生成する。

実行例:
  python make_manhour_to_sheet8_01_0002.py manhour_xxxxxx.csv
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import sys
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from pandas import DataFrame


def write_error_text_utf8(pszErrorFilePath: str, pszText: str) -> None:
    with open(pszErrorFilePath, mode="a", encoding="utf-8") as objFile:
        objFile.write(pszText)


def write_debug_error(pszMessage: str, objBaseDirectoryPath: Path | None = None) -> None:
    pszFileName: str = "make_manhour_to_sheet8_01_0002_error.txt"
    objErrorPath: Path = (
        objBaseDirectoryPath / pszFileName if objBaseDirectoryPath is not None else Path(pszFileName)
    )
    with open(objErrorPath, mode="a", encoding="utf-8") as objFile:
        objFile.write(pszMessage + "\n")


def get_target_year_month_from_filename(pszInputFilePath: str) -> Tuple[int, int]:
    pszBaseName: str = os.path.basename(pszInputFilePath)
    objMatch: re.Match[str] | None = re.search(r"(\d{2})\.(\d{1,2})\.csv$", pszBaseName)
    if objMatch is None:
        raise ValueError("入力ファイル名から対象年月を取得できません。")
    iYearTwoDigits: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    iYear: int = 2000 + iYearTwoDigits
    return iYear, iMonth


def split_by_fiscal_boundary(
    objStartMonth: Tuple[int, int],
    objEndMonth: Tuple[int, int],
    iFiscalEndMonth: int,
) -> List[Tuple[Tuple[int, int], Tuple[int, int]]]:
    if objStartMonth > objEndMonth:
        return []

    objRanges: List[Tuple[Tuple[int, int], Tuple[int, int]]] = []
    iCurrentYear, iCurrentMonth = objStartMonth

    while (iCurrentYear, iCurrentMonth) <= objEndMonth:
        iRangeEndYear: int = iCurrentYear
        if iCurrentMonth > iFiscalEndMonth:
            iRangeEndYear += 1
        iRangeEndMonth: Tuple[int, int] = (iRangeEndYear, iFiscalEndMonth)
        if iRangeEndMonth > objEndMonth:
            iRangeEndMonth = objEndMonth

        objRanges.append(((iCurrentYear, iCurrentMonth), iRangeEndMonth))

        iNextYear, iNextMonth = iRangeEndMonth
        if iNextMonth == 12:
            iCurrentYear, iCurrentMonth = iNextYear + 1, 1
        else:
            iCurrentYear, iCurrentMonth = iNextYear, iNextMonth + 1

    return objRanges


def build_cumulative_ranges_including_previous_terms(
    objStartMonth: Tuple[int, int],
    objEndMonth: Tuple[int, int],
) -> List[Tuple[Tuple[int, int], Tuple[int, int]]]:
    objFiscalARanges = split_by_fiscal_boundary(objStartMonth, objEndMonth, 3)
    objFiscalBRanges = split_by_fiscal_boundary(objStartMonth, objEndMonth, 8)
    objAllRanges: List[Tuple[Tuple[int, int], Tuple[int, int]]] = []

    def append_unique_range(objTargetRange: Tuple[Tuple[int, int], Tuple[int, int]]) -> None:
        if objTargetRange not in objAllRanges:
            objAllRanges.append(objTargetRange)

    if objFiscalARanges:
        if len(objFiscalARanges) >= 2:
            append_unique_range(objFiscalARanges[-2])
        append_unique_range(objFiscalARanges[-1])

    if objFiscalBRanges:
        if len(objFiscalBRanges) >= 2:
            append_unique_range(objFiscalBRanges[-2])
        append_unique_range(objFiscalBRanges[-1])

    return objAllRanges


def build_output_file_full_path(pszInputFileFullPath: str, pszOutputSuffix: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszStem: str = os.path.splitext(pszBaseName)[0]
    pszOutputFileName: str = pszStem + pszOutputSuffix
    return os.path.join(pszDirectory, pszOutputFileName)


def normalize_time_h_mm_to_h_mm_ss(pszTimeText: str) -> str:
    pszText: str = (pszTimeText or "").strip()
    if pszText == "":
        return ""
    if pszText.count(":") == 2:
        return pszText
    if pszText.count(":") == 1:
        return pszText + ":00"
    return pszText


def normalize_cell_text(pszCellText: str) -> str:
    pszNormalized: str = pszCellText or ""
    if "\t" in pszNormalized or '"' in pszNormalized:
        pszNormalized = pszNormalized.replace("\t", "_").replace('"', "")
    return pszNormalized


def convert_csv_to_tsv_file(pszInputCsvPath: str) -> str:
    if not os.path.exists(pszInputCsvPath):
        raise FileNotFoundError(f"Input CSV not found: {pszInputCsvPath}")

    pszOutputTsvPath: str = build_output_file_full_path(pszInputCsvPath, ".tsv")

    objRows: List[List[str]] = []
    arrEncodings: List[str] = ["utf-8-sig", "cp932"]
    objLastDecodeError: Exception | None = None

    for pszEncoding in arrEncodings:
        try:
            with open(
                pszInputCsvPath,
                mode="r",
                encoding=pszEncoding,
                newline="",
            ) as objInputFile:
                objReader: csv.reader = csv.reader(
                    objInputFile,
                    quoting=csv.QUOTE_NONE,
                )
                for objRow in objReader:
                    objRows.append(list(objRow))
            objLastDecodeError = None
            break
        except UnicodeDecodeError as objError:
            objLastDecodeError = objError
            objRows = []

    if objLastDecodeError is not None:
        raise objLastDecodeError

    if len(objRows) <= 1:
        with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
            objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
            for objRow in objRows:
                objWriter.writerow(objRow)
        return pszOutputTsvPath

    iTimeColumnIndexF: int = 5
    iTimeColumnIndexK: int = 10

    for iRowIndex in range(1, len(objRows)):
        objRow: List[str] = objRows[iRowIndex]
        if iTimeColumnIndexF < len(objRow):
            objRow[iTimeColumnIndexF] = normalize_time_h_mm_to_h_mm_ss(objRow[iTimeColumnIndexF])
        if iTimeColumnIndexK < len(objRow):
            objRow[iTimeColumnIndexK] = normalize_time_h_mm_to_h_mm_ss(objRow[iTimeColumnIndexK])
        objRow = [normalize_cell_text(objCell) for objCell in objRow]
        objRows[iRowIndex] = objRow

    if len(objRows) >= 1 and len(objRows[0]) >= 1:
        pszHeaderFirstCell: str = objRows[0][0]
        if pszHeaderFirstCell.startswith("\ufeff"):
            pszHeaderFirstCell = pszHeaderFirstCell.lstrip("\ufeff")
        if (
            len(pszHeaderFirstCell) >= 2
            and pszHeaderFirstCell.startswith('"')
            and pszHeaderFirstCell.endswith('"')
        ):
            pszHeaderFirstCell = pszHeaderFirstCell[1:-1]
            pszHeaderFirstCell = pszHeaderFirstCell.replace('""', '"')
            pszHeaderFirstCell = pszHeaderFirstCell.replace("\t", "_").replace('"', "")
        if (
            len(pszHeaderFirstCell) >= 2
            and pszHeaderFirstCell.startswith('"')
            and pszHeaderFirstCell.endswith('"')
        ):
            pszHeaderFirstCell = pszHeaderFirstCell[1:-1]
        objRows[0][0] = pszHeaderFirstCell
        objRows[0] = [normalize_cell_text(objCell) for objCell in objRows[0]]
        if len(objRows[0]) >= 4 and objRows[0][3] == "所属グループ名":
            objRows[0][3] = "所属カンパニー名"

    with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        for objRow in objRows:
            objWriter.writerow(objRow)

    return pszOutputTsvPath


def write_error_tsv(pszOutputFileFullPath: str, pszErrorMessage: str) -> None:
    pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
    if len(pszDirectory) > 0:
        os.makedirs(pszDirectory, exist_ok=True)

    with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage)
        if not pszErrorMessage.endswith("\n"):
            objFile.write("\n")


def build_removed_uninput_output_path(pszInputFileFullPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseName)

    pszOutputBaseName: str = pszRootName + "_step0001_removed_uninput.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


def make_removed_uninput_tsv_from_manhour_tsv(pszInputFileFullPath: str) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        pszDirectory: str = os.path.dirname(pszInputFileFullPath)
        pszBaseName: str = os.path.basename(pszInputFileFullPath)
        pszRootName: str
        pszExt: str
        pszRootName, pszExt = os.path.splitext(pszBaseName)
        pszErrorFileFullPath: str = os.path.join(
            pszDirectory,
            pszRootName + "_error.tsv",
        )

        write_error_tsv(
            pszErrorFileFullPath,
            "Error: input TSV file not found. Path = {0}".format(
                pszInputFileFullPath
            ),
        )
        return

    pszOutputFileFullPath: str = build_removed_uninput_output_path(pszInputFileFullPath)

    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            encoding="utf-8",
            dtype=str,
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for removing '未入力'. "
            "Detail = {0}".format(objException),
        )
        return

    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 10:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required columns G-J do not exist (need at least 10 columns). "
            "ColumnCount = {0}".format(iColumnCount),
        )
        return

    objColumnNameList: List[str] = list(objDataFrame.columns)
    pszColumnG: str = objColumnNameList[6]
    pszColumnH: str = objColumnNameList[7]
    pszColumnI: str = objColumnNameList[8]
    pszColumnJ: str = objColumnNameList[9]

    try:
        objSeriesHasUninputG = (
            objDataFrame[pszColumnG].fillna("").astype(str).str.strip() == "未入力"
        )
        objSeriesHasUninputH = (
            objDataFrame[pszColumnH].fillna("").astype(str).str.strip() == "未入力"
        )
        objSeriesHasUninputI = (
            objDataFrame[pszColumnI].fillna("").astype(str).str.strip() == "未入力"
        )
        objSeriesHasUninputJ = (
            objDataFrame[pszColumnJ].fillna("").astype(str).str.strip() == "未入力"
        )

        objSeriesHasUninputAny = (
            objSeriesHasUninputG
            | objSeriesHasUninputH
            | objSeriesHasUninputI
            | objSeriesHasUninputJ
        )

        objDataFrameFiltered: DataFrame = objDataFrame.loc[~objSeriesHasUninputAny].copy()
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while filtering rows with '未入力'. "
            "Detail = {0}".format(objException),
        )
        return

    try:
        objDataFrameFiltered.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing TSV without '未入力' rows. "
            "Detail = {0}".format(objException),
        )
        return


def build_sorted_staff_code_output_path(pszInputFileFullPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseName)

    pszStep0001Suffix: str = "_step0001_removed_uninput"
    if pszRootName.endswith(pszStep0001Suffix):
        pszRootName = pszRootName[: -len(pszStep0001Suffix)]
    pszOutputBaseName: str = (
        pszRootName + "_step0002_removed_uninput_sorted_staff_code.tsv"
    )
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


def make_sorted_staff_code_tsv_from_manhour_tsv(pszInputFileFullPath: str) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputFileFullPath}")

    pszOutputFileFullPath: str = build_sorted_staff_code_output_path(pszInputFileFullPath)

    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading manhour TSV for staff code sort. "
            "Detail = {0}".format(objException),
        )
        return

    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 2:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: staff code column (2nd column) does not exist. ColumnCount = {0}".format(
                iColumnCount
            ),
        )
        return

    objColumnNameList: List[str] = list(objDataFrame.columns)
    pszSortColumnName: str = objColumnNameList[1]

    try:
        objSorted: DataFrame = objDataFrame.copy()
        objSorted["__sort_staff_code__"] = pd.to_numeric(
            objSorted[pszSortColumnName],
            errors="coerce",
        )
        objSorted = objSorted.sort_values(
            by="__sort_staff_code__",
            ascending=True,
            kind="mergesort",
        ).drop(columns=["__sort_staff_code__"])
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while sorting by staff code. Detail = {0}".format(
                objException
            ),
        )
        return

    try:
        objSorted.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing sorted staff-code TSV. Detail = {0}".format(
                objException
            ),
        )
        return


def step0003_normalize_company_name(pszCompanyName: str) -> str:
    pszNormalized: str = pszCompanyName or ""
    pszNormalized = re.sub(
        r'"([^"]*)"',
        lambda objMatch: objMatch.group(1).replace("\t", "_"),
        pszNormalized,
    )
    pszNormalized = pszNormalized.replace('"', "")
    objReplaceTargets: List[Tuple[str, str]] = [
        ("本部", "本部"),
        ("事業開発", "事業開発"),
        ("子会社", "子会社"),
        ("投資先", "投資先"),
        ("第１インキュ", "第一インキュ"),
        ("第２インキュ", "第二インキュ"),
        ("第３インキュ", "第三インキュ"),
        ("第４インキュ", "第四インキュ"),
        ("第1インキュ", "第一インキュ"),
        ("第2インキュ", "第二インキュ"),
        ("第3インキュ", "第三インキュ"),
        ("第4インキュ", "第四インキュ"),
    ]
    for pszPrefix, pszReplacement in objReplaceTargets:
        if pszNormalized.startswith(pszPrefix):
            return pszReplacement
    return pszNormalized


def build_step0003_company_normalized_output_path(pszInputFileFullPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseName)

    pszStep0002Suffix: str = "_step0002_removed_uninput_sorted_staff_code"
    if pszRootName.endswith(pszStep0002Suffix):
        pszRootName = pszRootName[: -len(pszStep0002Suffix)]
    pszOutputBaseName: str = pszRootName + "_step0003_normalized_company_name.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


def build_step0004_company_normalized_output_path(pszInputFileFullPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseName)

    pszStep0003Suffix: str = "_step0003_normalized_company_name"
    if pszRootName.endswith(pszStep0003Suffix):
        pszRootName = pszRootName[: -len(pszStep0003Suffix)]
    pszOutputBaseName: str = pszRootName + "_step0004_normalized_project_name.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


def write_company_normalized_tsv(pszInputFileFullPath: str, pszOutputFileFullPath: str) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputFileFullPath}")

    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for company name normalization. "
            "Detail = {0}".format(objException),
        )
        return

    objColumnNameList: List[str] = list(objDataFrame.columns)
    objCandidateColumns: List[str] = [
        "計上カンパニー名",
        "計上カンパニー",
        "所属カンパニー",
        "所属カンパニー名",
    ]
    pszCompanyColumn: str | None = None
    for pszCandidate in objCandidateColumns:
        if pszCandidate in objColumnNameList:
            pszCompanyColumn = pszCandidate
            break

    if pszCompanyColumn is None:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: company name column not found. "
            "Expected one of {0}.".format(", ".join(objCandidateColumns)),
        )
        return

    try:
        objDataFrame[pszCompanyColumn] = (
            objDataFrame[pszCompanyColumn]
            .fillna("")
            .astype(str)
            .apply(step0003_normalize_company_name)
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while normalizing company name column. "
            "Detail = {0}".format(objException),
        )
        return

    try:
        objDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing normalized TSV. "
            "Detail = {0}".format(objException),
        )
        return


def make_company_normalized_tsv_from_step0002(pszInputFileFullPath: str) -> None:
    pszOutputFileFullPath: str = build_step0003_company_normalized_output_path(
        pszInputFileFullPath
    )
    write_company_normalized_tsv(pszInputFileFullPath, pszOutputFileFullPath)


def make_company_normalized_tsv_from_step0003(pszInputFileFullPath: str) -> None:
    pszOutputFileFullPath: str = build_step0004_company_normalized_output_path(
        pszInputFileFullPath
    )
    write_project_normalized_tsv(pszInputFileFullPath, pszOutputFileFullPath)


def step0004_normalize_project_code(pszProjectCode: str) -> str:
    return re.sub(r"[\s\u3000]+", "", pszProjectCode or "")


def step0004_normalize_project_name(pszProjectName: str) -> str:
    pszNormalized: str = pszProjectName or ""
    pszNormalized = re.sub(
        r'"([^"]*)"',
        lambda objMatch: objMatch.group(1).replace("\t", "_"),
        pszNormalized,
    )
    pszNormalized = pszNormalized.replace('"', "")
    pszNormalized = re.sub(
        r"((?:P\d{5}|[A-OQ-Z]\d{3}))[\u0020\u3000]+",
        r"\1_",
        pszNormalized,
    )
    if pszNormalized.startswith("【"):
        objMatchBracket: re.Match[str] | None = re.search(
            r"(P\d{5}|[A-OQ-Z]\d{3})",
            pszNormalized,
        )
        if objMatchBracket is not None:
            pszCodeBracket: str = objMatchBracket.group(1)
            pszRestBracket: str = (
                pszNormalized[: objMatchBracket.start()]
                + pszNormalized[objMatchBracket.end() :]
            )
            return pszCodeBracket + "_" + pszRestBracket
    objMatchP: re.Match[str] | None = re.match(r"^(P\d{5})(.*)$", pszNormalized)
    if objMatchP is not None:
        pszCode: str = objMatchP.group(1)
        pszRest: str = objMatchP.group(2)
        if pszRest.startswith("【"):
            pszNormalized = pszCode + "_" + pszRest
    else:
        objMatchOther: re.Match[str] | None = re.match(r"^([A-OQ-Z]\d{3})(.*)$", pszNormalized)
        if objMatchOther is not None:
            pszCodeOther: str = objMatchOther.group(1)
            pszRestOther: str = objMatchOther.group(2)
            if pszRestOther.startswith("【"):
                pszNormalized = pszCodeOther + "_" + pszRestOther
    return pszNormalized


def write_project_normalized_tsv(pszInputFileFullPath: str, pszOutputFileFullPath: str) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputFileFullPath}")

    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for project normalization. "
            "Detail = {0}".format(objException),
        )
        return

    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 8:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required columns G-H do not exist (need at least 8 columns). "
            "ColumnCount = {0}".format(iColumnCount),
        )
        return

    objColumnNameList: List[str] = list(objDataFrame.columns)
    pszColumnG: str = objColumnNameList[6]
    pszColumnH: str = objColumnNameList[7]

    try:
        objDataFrame[pszColumnG] = (
            objDataFrame[pszColumnG]
            .fillna("")
            .astype(str)
            .apply(step0004_normalize_project_code)
        )
        objDataFrame[pszColumnH] = (
            objDataFrame[pszColumnH]
            .fillna("")
            .astype(str)
            .apply(step0004_normalize_project_name)
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while normalizing project code/name columns. "
            "Detail = {0}".format(objException),
        )
        return

    try:
        objDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing normalized project TSV. "
            "Detail = {0}".format(objException),
        )
        return


def normalize_org_table_project_code_step0004(pszProjectCode: str) -> str:
    pszNormalized: str = step0004_normalize_project_name(pszProjectCode or "")
    return re.sub(r"[ \u3000]+", "_", pszNormalized)


def build_step0005_remove_ah_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0005_remove_A_or_H_project.tsv"
    )


def build_step0006_company_replaced_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0006_projects_replaced_by_管轄PJ表.tsv"
    )


def build_step0006_missing_project_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0006_projects_missing_in_管轄PJ表.tsv"
    )


def build_step0006_unique_missing_project_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0006_unique_projects_missing_in_管轄PJ表.tsv"
    )


def build_step0006_sort_asc_missing_project_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0006_sort_asc_projects_missing_in_管轄PJ表.tsv"
    )


def build_step0007_yyyy_mm_dd_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0007_yyyy_mm_dd.tsv"
    )


def build_step0007_unique_staff_code_output_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectoryFullPath: str = os.path.dirname(pszInputFileFullPath)
    pszBaseFileName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseFileName)

    pszOutputFileName: str = pszRootName + "_unique_staff_code.tsv"

    if len(pszDirectoryFullPath) == 0:
        return pszOutputFileName

    return os.path.join(pszDirectoryFullPath, pszOutputFileName)


def build_step0007_staff_code_range_output_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectoryFullPath: str = os.path.dirname(pszInputFileFullPath)
    pszBaseFileName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseFileName)

    pszOutputFileName: str = pszRootName + "_staff_code_range.tsv"

    if len(pszDirectoryFullPath) == 0:
        return pszOutputFileName

    return os.path.join(pszDirectoryFullPath, pszOutputFileName)


def build_step0008_staff_project_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0008_スタッフ別担当プロジェクト.tsv"
    )


def build_step0009_project_task_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0009_プロジェクト_タスク_工数.tsv"
    )


def build_step0009_project_staff_company_task_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0009_プロジェクト_スタッフ計上カンパニー名_タスク_工数.tsv"
    )


def build_step0009_project_company_task_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0009_プロジェクト_計上カンパニー名_タスク_工数.tsv"
    )


def build_step0010_project_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0010_計算前_プロジェクト_工数.tsv"
    )


def build_step0010_project_company_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0010_計算前_プロジェクト_計上カンパニー名_工数.tsv"
    )


def build_step0011_project_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0011_合計_プロジェクト_工数.tsv"
    )


def build_step0011_project_company_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0011_合計_プロジェクト_計上カンパニー名_工数.tsv"
    )


def build_step0012_project_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0012_昇順_合計_プロジェクト_工数.tsv"
    )


def build_step0012_project_company_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0012_昇順_合計_プロジェクト_計上カンパニー名_工数.tsv"
    )


def build_step0012_project_company_group_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0012_昇順_合計_プロジェクト_計上カンパニー名_計上グループ_工数.tsv"
    )


def build_step0013_project_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0013_各プロジェクトの工数.tsv"
    )


def build_step0013_project_company_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0013_各プロジェクトの計上カンパニー名_工数.tsv"
    )


def build_step0013_project_company_group_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0013_各プロジェクトの計上カンパニー名_計上グループ_工数.tsv"
    )


def build_step14_project_company_manhour_output_path(
    objBaseDirectoryPath: Path,
    iYear: int,
    iMonth: int,
) -> Path:
    return (
        objBaseDirectoryPath
        / f"工数_{iYear}年{iMonth:02d}月_step0014_各プロジェクトの計上カンパニー名_工数_カンパニーの工数.tsv"
    )


def normalize_step0007_yyyy_mm_dd_in_value(
    objValue: object,
    objPattern: re.Pattern[str],
) -> object:
    if not isinstance(objValue, str):
        return objValue

    pszText: str = objValue

    objMatch = objPattern.match(pszText)
    if objMatch is None:
        return pszText

    pszYear: str = objMatch.group(1)
    pszMonthRaw: str = objMatch.group(2)
    pszDayRaw: str = objMatch.group(3)

    try:
        iMonth: int = int(pszMonthRaw)
        iDay: int = int(pszDayRaw)
    except Exception:
        return pszText

    if iMonth < 1 or iMonth > 12:
        return pszText
    if iDay < 1 or iDay > 31:
        return pszText

    pszMonth: str = str(iMonth).zfill(2)
    pszDay: str = str(iDay).zfill(2)

    return pszYear + "/" + pszMonth + "/" + pszDay


def normalize_step0007_yyyy_mm_dd_in_dataframe(
    objDataFrameInput: DataFrame,
) -> DataFrame:
    objPattern: re.Pattern[str] = re.compile(r"^\s*(\d{4})/(\d{1,2})/(\d{1,2})\s*$")

    def _normalize_wrapper(objValue: object) -> object:
        return normalize_step0007_yyyy_mm_dd_in_value(objValue, objPattern)

    try:
        objDataFrameOutput: DataFrame = objDataFrameInput.map(_normalize_wrapper)
    except AttributeError:
        objDataFrameOutput = objDataFrameInput.applymap(_normalize_wrapper)

    return objDataFrameOutput


def make_step0007_yyyy_mm_dd_tsv(
    pszInputTsvPath: str,
    pszOutputTsvPath: str,
) -> None:
    if not os.path.exists(pszInputTsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputTsvPath}")

    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputTsvPath,
            sep="\t",
            header=0,
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputTsvPath,
            "Error: unexpected exception while reading TSV. Detail = {0}".format(
                objException
            ),
        )
        return

    try:
        objDataFrameOutput: DataFrame = normalize_step0007_yyyy_mm_dd_in_dataframe(
            objDataFrameInput
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputTsvPath,
            "Error: unexpected exception while converting date format. Detail = {0}".format(
                objException
            ),
        )
        return

    try:
        objDataFrameOutput.to_csv(
            pszOutputTsvPath,
            sep="\t",
            index=False,
            encoding="utf-8",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputTsvPath,
            "Error: unexpected exception while writing normalized TSV. Detail = {0}".format(
                objException
            ),
        )
        return


def make_step0007_unique_staff_code_tsv(
    pszInputFileFullPath: str,
) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        pszDirectory: str = os.path.dirname(pszInputFileFullPath)
        pszBaseName: str = os.path.basename(pszInputFileFullPath)
        pszRootName: str
        pszExt: str
        pszRootName, pszExt = os.path.splitext(pszBaseName)
        pszErrorFileFullPath: str = os.path.join(
            pszDirectory,
            pszRootName + "_error.tsv",
        )

        write_error_tsv(
            pszErrorFileFullPath,
            "Error: input TSV file not found. Path = {0}".format(
                pszInputFileFullPath
            ),
        )
        return

    pszOutputFileFullPath: str = build_step0007_unique_staff_code_output_path(
        pszInputFileFullPath
    )

    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            encoding="utf-8",
            dtype=str,
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for unique staff code list. "
            "Detail = {0}".format(objException),
        )
        return

    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 2:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required column B does not exist (need at least 2 columns). "
            "ColumnCount = {0}".format(iColumnCount),
        )
        return

    objColumnNameList: List[str] = list(objDataFrame.columns)
    pszStaffCodeColumnName: str = objColumnNameList[1]
    objSeriesStaffCode = objDataFrame.iloc[:, 1]

    try:
        objListUniqueStaffCode: List[str] = []
        objSetSeen: set[str] = set()

        for pszValueRaw in objSeriesStaffCode.tolist():
            pszValue: str = "" if pszValueRaw is None else str(pszValueRaw)
            pszValueStripped: str = pszValue.strip()

            if pszValueStripped == "":
                continue

            if pszValueStripped in objSetSeen:
                continue

            objSetSeen.add(pszValueStripped)
            objListUniqueStaffCode.append(pszValueStripped)

        objOutputDataFrame: DataFrame = DataFrame(
            {pszStaffCodeColumnName: objListUniqueStaffCode}
        )

    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while creating unique staff code list. "
            "Detail = {0}".format(objException),
        )
        return

    try:
        objOutputDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing unique staff code TSV. "
            "Detail = {0}".format(objException),
        )
        return


def analyze_step0007_staff_code_column(
    objSeriesStaffCode: pd.Series,
) -> Tuple[List[str], dict[str, Tuple[int, int]]]:
    objListUniqueStaffCode: List[str] = []
    objDictCodeToRange: dict[str, Tuple[int, int]] = {}

    iRowCount: int = objSeriesStaffCode.shape[0]
    for iRowIndex in range(iRowCount):
        objValue: object = objSeriesStaffCode.iat[iRowIndex]
        pszRaw: str = "" if objValue is None else str(objValue)
        pszCode: str = pszRaw.strip()

        if pszCode == "":
            continue

        if pszCode not in objDictCodeToRange:
            objListUniqueStaffCode.append(pszCode)
            objDictCodeToRange[pszCode] = (iRowIndex, iRowIndex)
        else:
            iFirstIndex: int
            iLastIndex: int
            iFirstIndex, iLastIndex = objDictCodeToRange[pszCode]
            objDictCodeToRange[pszCode] = (iFirstIndex, iRowIndex)

    return objListUniqueStaffCode, objDictCodeToRange


def make_step0007_staff_code_range_tsv(
    pszInputFileFullPath: str,
) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        pszDirectoryFullPath: str = os.path.dirname(pszInputFileFullPath)
        pszBaseFileName: str = os.path.basename(pszInputFileFullPath)
        pszBase: str
        pszExt: str
        pszBase, pszExt = os.path.splitext(pszBaseFileName)

        pszOutputFileFullPath: str = os.path.join(
            pszDirectoryFullPath,
            pszBase + "_error.tsv",
        )
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: input TSV file not found. Path = {0}".format(pszInputFileFullPath),
        )
        return

    pszOutputFileFullPath: str = build_step0007_staff_code_range_output_path(
        pszInputFileFullPath
    )

    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            encoding="utf-8",
            dtype=str,
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for staff code range. "
            "Detail = {0}".format(objException),
        )
        return

    iColumnCount: int = objDataFrameInput.shape[1]
    if iColumnCount < 2:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required column B does not exist (need at least 2 columns). "
            "ColumnCount = {0}".format(iColumnCount),
        )
        return

    objColumnNameList: List[str] = list(objDataFrameInput.columns)
    pszStaffCodeColumnName: str = objColumnNameList[1]
    objSeriesStaffCode = objDataFrameInput.iloc[:, 1]

    try:
        objListUniqueStaffCode: List[str]
        objDictCodeToRange: dict[str, Tuple[int, int]]
        objListUniqueStaffCode, objDictCodeToRange = analyze_step0007_staff_code_column(
            objSeriesStaffCode
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while analyzing staff code column. "
            "Detail = {0}".format(objException),
        )
        return

    if len(objListUniqueStaffCode) == 0:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: no valid staff code found in column B.",
        )
        return

    objListOutputCode: List[str] = []
    objListOutputStartRow: List[int] = []
    objListOutputEndRow: List[int] = []

    for pszCode in objListUniqueStaffCode:
        if pszCode not in objDictCodeToRange:
            write_error_tsv(
                pszOutputFileFullPath,
                "Error: internal inconsistency. Code not found in range map. "
                "Code = {0}".format(pszCode),
            )
            return

        iFirstIndex: int
        iLastIndex: int
        iFirstIndex, iLastIndex = objDictCodeToRange[pszCode]

        iStartRow: int = iFirstIndex + 2
        iEndRow: int = iLastIndex + 2

        objListOutputCode.append(pszCode)
        objListOutputStartRow.append(iStartRow)
        objListOutputEndRow.append(iEndRow)

    pszStartColumnName: str = "開始行"
    pszEndColumnName: str = "終了行"

    objDataFrameOutput: DataFrame = DataFrame(
        {
            pszStaffCodeColumnName: objListOutputCode,
            pszStartColumnName: objListOutputStartRow,
            pszEndColumnName: objListOutputEndRow,
        }
    )

    try:
        objDataFrameOutput.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing staff code range TSV. "
            "Detail = {0}".format(objException),
        )
        return


def make_step0008_staff_project_tsv(
    pszStep0007FileFullPath: str,
    pszRangeFileFullPath: str,
    pszOutputFileFullPath: str,
) -> None:
    pszErrorFileFullPath: str = pszOutputFileFullPath.replace(".tsv", "_error.tsv")

    if not os.path.isfile(pszStep0007FileFullPath):
        write_error_tsv(
            pszStep0007FileFullPath.replace(".tsv", "_error.tsv"),
            "Error: step0007 TSV file not found. Path = {0}".format(
                pszStep0007FileFullPath
            ),
        )
        return

    if not os.path.isfile(pszRangeFileFullPath):
        write_error_tsv(
            pszRangeFileFullPath.replace(".tsv", "_error.tsv"),
            "Error: step0007 staff_code_range TSV file not found. Path = {0}".format(
                pszRangeFileFullPath
            ),
        )
        return

    try:
        objDataFrameSheet4: DataFrame = pd.read_csv(
            pszStep0007FileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while reading step0007 TSV. Detail = {0}".format(
                objException
            ),
        )
        return

    objSheet4Columns: List[str] = list(objDataFrameSheet4.columns)
    if ("スタッフコード" not in objSheet4Columns) or ("プロジェクト名" not in objSheet4Columns):
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: required columns not found in step0007 TSV. "
            "Required columns: スタッフコード, プロジェクト名. "
            "Columns = {0}".format(", ".join(objSheet4Columns)),
        )
        return

    try:
        objDataFrameRange: DataFrame = pd.read_csv(
            pszRangeFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while reading step0007 staff_code_range TSV. "
            "Detail = {0}".format(objException),
        )
        return

    iRangeColumnCount: int = objDataFrameRange.shape[1]
    if iRangeColumnCount < 3:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: step0007 staff_code_range TSV must have at least 3 columns "
            "(staff_code, start_row, end_row). ColumnCount = {0}".format(
                iRangeColumnCount
            ),
        )
        return

    try:
        objDataFrameRange = objDataFrameRange.copy()
        objDataFrameRange["__start_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 1],
            errors="coerce",
        )
        objDataFrameRange["__end_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 2],
            errors="coerce",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while converting start/end rows to numeric. Detail = {0}".format(
                objException
            ),
        )
        return

    if objDataFrameRange["__start_row_excel__"].isna().any() or objDataFrameRange[
        "__end_row_excel__"
    ].isna().any():
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: start_row or end_row contains non-numeric value in step0007 staff_code_range TSV.",
        )
        return

    objListProjectListPerStaff: List[List[str]] = []
    objListStaffCode: List[str] = []

    iSheet4RowCount: int = objDataFrameSheet4.shape[0]

    for _, objRow in objDataFrameRange.iterrows():
        pszStaffCode: str = str(objRow.iloc[0])
        objListStaffCode.append(pszStaffCode)

        iStartRowExcel: int = int(objRow["__start_row_excel__"])
        iEndRowExcel: int = int(objRow["__end_row_excel__"])

        iStartIndex: int = iStartRowExcel - 2
        iEndIndex: int = iEndRowExcel - 2

        if (
            (iStartIndex < 0)
            or (iEndIndex < 0)
            or (iStartIndex > iEndIndex)
            or (iEndIndex >= iSheet4RowCount)
        ):
            write_error_tsv(
                pszErrorFileFullPath,
                "Error: invalid row range for staff code {0}. "
                "StartExcel={1}, EndExcel={2}, StartIndex={3}, EndIndex={4}, Sheet4RowCount={5}".format(
                    pszStaffCode,
                    iStartRowExcel,
                    iEndRowExcel,
                    iStartIndex,
                    iEndIndex,
                    iSheet4RowCount,
                ),
            )
            return

        objDataFrameSub: DataFrame = objDataFrameSheet4.iloc[iStartIndex : iEndIndex + 1]
        objDataFrameSub = objDataFrameSub[objDataFrameSub["スタッフコード"] == pszStaffCode]

        objSeriesPj: pd.Series = objDataFrameSub["プロジェクト名"].dropna()
        objSeriesPj = objSeriesPj.astype(str)
        objSeriesPj = objSeriesPj[objSeriesPj.str.strip() != ""]

        objSeriesPjSorted: pd.Series = objSeriesPj.sort_values()
        objSeriesPjUnique: pd.Series = objSeriesPjSorted.drop_duplicates()

        objProjectList: List[str] = objSeriesPjUnique.tolist()
        objListProjectListPerStaff.append(objProjectList)

    iStaffCount: int = len(objListProjectListPerStaff)

    if iStaffCount == 0:
        try:
            objEmpty: DataFrame = DataFrame([])
            objEmpty.to_csv(
                pszOutputFileFullPath,
                sep="\t",
                index=False,
                header=False,
                encoding="utf-8",
                lineterminator="\n",
            )
        except Exception as objException:
            write_error_tsv(
                pszErrorFileFullPath,
                "Error: unexpected exception while writing empty step0008 TSV. Detail = {0}".format(
                    objException
                ),
            )
        return

    iMaxProjectCount: int = max(len(objList) for objList in objListProjectListPerStaff)

    objRows: List[List[str]] = []
    objRow1: List[str] = [str(iIndex + 1) for iIndex in range(iStaffCount)]
    objRows.append(objRow1)

    objRow2: List[str] = [str(pszCode) for pszCode in objListStaffCode]
    objRows.append(objRow2)

    for iPjIndex in range(iMaxProjectCount):
        objRow: List[str] = []
        for iStaffIndex in range(iStaffCount):
            objProjectList: List[str] = objListProjectListPerStaff[iStaffIndex]
            if iPjIndex < len(objProjectList):
                objRow.append(objProjectList[iPjIndex])
            else:
                objRow.append("")
        objRows.append(objRow)

    try:
        objDataFrameOutput: DataFrame = DataFrame(objRows)
        objDataFrameOutput.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            header=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while writing step0008 TSV. Detail = {0}".format(
                objException
            ),
        )
        return


def read_step0009_tsv_with_encoding_candidates(
    pszInputFileFullPath: str,
    bHasHeader: bool,
) -> DataFrame:
    objEncodingCandidateList: List[str] = ["utf-8-sig", "cp932"]
    objLastDecodeError: Exception | None = None
    objDataFrameResult: DataFrame | None = None

    for pszEncoding in objEncodingCandidateList:
        try:
            if bHasHeader:
                objDataFrameResult = pd.read_csv(
                    pszInputFileFullPath,
                    sep="\t",
                    dtype=str,
                    encoding=pszEncoding,
                    engine="python",
                )
            else:
                objDataFrameResult = pd.read_csv(
                    pszInputFileFullPath,
                    sep="\t",
                    dtype=str,
                    header=None,
                    encoding=pszEncoding,
                    engine="python",
                )
            break
        except UnicodeDecodeError as objDecodeError:
            objLastDecodeError = objDecodeError
            continue

    if objDataFrameResult is None:
        if objLastDecodeError is not None:
            raise objLastDecodeError
        raise UnicodeDecodeError(
            "utf-8-sig",
            b"",
            0,
            1,
            "cannot decode TSV with utf-8-sig nor cp932",
        )

    return objDataFrameResult


def convert_step0009_time_string_to_seconds(
    pszTimeText: str,
) -> int:
    if pszTimeText is None:
        return 0

    pszWork: str = str(pszTimeText).strip()
    if len(pszWork) == 0:
        return 0

    objParts: List[str] = pszWork.split(":")
    try:
        if len(objParts) == 3:
            iHour: int = int(objParts[0])
            iMinute: int = int(objParts[1])
            iSecond: int = int(objParts[2])
        elif len(objParts) == 2:
            iHour = int(objParts[0])
            iMinute = int(objParts[1])
            iSecond = 0
        else:
            return 0
    except ValueError:
        return 0

    return iHour * 3600 + iMinute * 60 + iSecond


def convert_step0009_seconds_to_time_string(
    iTotalSeconds: int,
) -> str:
    if iTotalSeconds <= 0:
        return "0:00:00"

    iHour: int = iTotalSeconds // 3600
    iRemain: int = iTotalSeconds % 3600
    iMinute: int = iRemain // 60
    iSecond: int = iRemain % 60

    return "{0}:{1:02d}:{2:02d}".format(iHour, iMinute, iSecond)


def make_step0009_project_task_tsv(
    pszStep0007FileFullPath: str,
    pszRangeFileFullPath: str,
    pszStep0008FileFullPath: str,
    pszProjectTaskOutputPath: str,
    pszProjectStaffCompanyOutputPath: str,
) -> None:
    pszErrorFileFullPath: str = pszProjectTaskOutputPath.replace(".tsv", "_error.tsv")

    if not os.path.isfile(pszStep0007FileFullPath):
        write_error_tsv(
            pszStep0007FileFullPath.replace(".tsv", "_error.tsv"),
            "Error: step0007 TSV file not found. Path = {0}".format(
                pszStep0007FileFullPath
            ),
        )
        return

    if not os.path.isfile(pszRangeFileFullPath):
        write_error_tsv(
            pszRangeFileFullPath.replace(".tsv", "_error.tsv"),
            "Error: step0007 staff_code_range TSV file not found. Path = {0}".format(
                pszRangeFileFullPath
            ),
        )
        return

    if not os.path.isfile(pszStep0008FileFullPath):
        write_error_tsv(
            pszStep0008FileFullPath.replace(".tsv", "_error.tsv"),
            "Error: step0008 TSV file not found. Path = {0}".format(
                pszStep0008FileFullPath
            ),
        )
        return

    try:
        objDataFrameSheet4: DataFrame = read_step0009_tsv_with_encoding_candidates(
            pszStep0007FileFullPath,
            True,
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while reading step0007 TSV. Detail = {0}".format(
                objException
            ),
        )
        return

    objSheet4Columns: List[str] = list(objDataFrameSheet4.columns)
    if (
        ("スタッフコード" not in objSheet4Columns)
        or ("プロジェクト名" not in objSheet4Columns)
        or ("工数" not in objSheet4Columns)
    ):
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: required columns not found in step0007 TSV. "
            "Required columns: スタッフコード, プロジェクト名, 工数. "
            "Columns = {0}".format(", ".join(objSheet4Columns)),
        )
        return

    pszCompanyColumn: str = ""
    if "計上カンパニー名" in objSheet4Columns:
        pszCompanyColumn = "計上カンパニー名"
    elif "計上カンパニー" in objSheet4Columns:
        pszCompanyColumn = "計上カンパニー"
    elif "所属カンパニー名" in objSheet4Columns:
        pszCompanyColumn = "所属カンパニー名"
    elif "所属カンパニー" in objSheet4Columns:
        pszCompanyColumn = "所属カンパニー"
    elif "所属グループ名" in objSheet4Columns:
        pszCompanyColumn = "所属グループ名"
    elif "所属グループ" in objSheet4Columns:
        pszCompanyColumn = "所属グループ"

    try:
        objDataFrameSheet4 = objDataFrameSheet4.copy()
        objDataFrameSheet4["__time_seconds__"] = objDataFrameSheet4["工数"].apply(
            lambda pszValue: convert_step0009_time_string_to_seconds(pszValue),
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while converting time text to seconds. Detail = {0}".format(
                objException
            ),
        )
        return

    iSheet4RowCount: int = objDataFrameSheet4.shape[0]

    try:
        objDataFrameRange: DataFrame = read_step0009_tsv_with_encoding_candidates(
            pszRangeFileFullPath,
            True,
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while reading step0007 staff_code_range TSV. Detail = {0}".format(
                objException
            ),
        )
        return

    if objDataFrameRange.shape[1] < 3:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: step0007 staff_code_range TSV must have at least 3 columns "
            "(staff_code, start_row, end_row).",
        )
        return

    try:
        objDataFrameRange = objDataFrameRange.copy()
        objDataFrameRange["__start_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 1],
            errors="coerce",
        )
        objDataFrameRange["__end_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 2],
            errors="coerce",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while converting start/end rows to numeric. Detail = {0}".format(
                objException
            ),
        )
        return

    if objDataFrameRange["__start_row_excel__"].isna().any() or objDataFrameRange[
        "__end_row_excel__"
    ].isna().any():
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: start_row or end_row contains non-numeric value in step0007 staff_code_range TSV.",
        )
        return

    objDictStaffCodeToRange: dict[str, Tuple[int, int]] = {}
    for _, objRow in objDataFrameRange.iterrows():
        pszStaffCodeRange: str = str(objRow.iloc[0]).strip()
        if len(pszStaffCodeRange) == 0:
            continue

        iStartRowExcel: int = int(objRow["__start_row_excel__"])
        iEndRowExcel: int = int(objRow["__end_row_excel__"])

        iStartIndex: int = iStartRowExcel - 2
        iEndIndex: int = iEndRowExcel - 2

        if (
            (iStartIndex < 0)
            or (iEndIndex < 0)
            or (iStartIndex > iEndIndex)
            or (iEndIndex >= iSheet4RowCount)
        ):
            write_error_tsv(
                pszErrorFileFullPath,
                "Error: invalid row range for staff code {0}. "
                "StartExcel={1}, EndExcel={2}, StartIndex={3}, EndIndex={4}, Sheet4RowCount={5}".format(
                    pszStaffCodeRange,
                    iStartRowExcel,
                    iEndRowExcel,
                    iStartIndex,
                    iEndIndex,
                    iSheet4RowCount,
                ),
            )
            return

        if pszStaffCodeRange not in objDictStaffCodeToRange:
            objDictStaffCodeToRange[pszStaffCodeRange] = (iStartIndex, iEndIndex)

    try:
        objDataFrameSheet6: DataFrame = read_step0009_tsv_with_encoding_candidates(
            pszStep0008FileFullPath,
            False,
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while reading step0008 TSV. Detail = {0}".format(
                objException
            ),
        )
        return

    if objDataFrameSheet6.shape[0] < 2:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: step0008 TSV must have at least 2 rows (1st: index, 2nd: staff_code).",
        )
        return

    iSheet6RowCount: int = objDataFrameSheet6.shape[0]
    iSheet6ColumnCount: int = objDataFrameSheet6.shape[1]

    objListStaffCodeFromSheet6: List[str] = []
    for iColumnIndex in range(iSheet6ColumnCount):
        pszStaffCodeFromSheet6: str = ""
        objValue = objDataFrameSheet6.iat[1, iColumnIndex]
        if objValue is not None:
            pszStaffCodeFromSheet6 = str(objValue).strip()
        if len(pszStaffCodeFromSheet6) == 0:
            continue
        objListStaffCodeFromSheet6.append(pszStaffCodeFromSheet6)

    objListOutputRowsProjectTask: List[List[str]] = []
    objListOutputRowsProjectStaffCompany: List[List[str]] = []

    for pszStaffCode in objListStaffCodeFromSheet6:
        if pszStaffCode in objDictStaffCodeToRange:
            iStartIndex, iEndIndex = objDictStaffCodeToRange[pszStaffCode]
            objDataFrameSubStaff: DataFrame = objDataFrameSheet4.iloc[
                iStartIndex : iEndIndex + 1
            ]
            objDataFrameSubStaff = objDataFrameSubStaff[
                objDataFrameSubStaff["スタッフコード"] == pszStaffCode
            ]
        else:
            objDataFrameSubStaff = objDataFrameSheet4[
                objDataFrameSheet4["スタッフコード"] == pszStaffCode
            ]

        if objDataFrameSubStaff.empty:
            continue

        for iColumnIndex in range(iSheet6ColumnCount):
            objValueStaffCodeAtColumn = objDataFrameSheet6.iat[1, iColumnIndex]
            pszStaffCodeAtColumn: str = ""
            if objValueStaffCodeAtColumn is not None:
                pszStaffCodeAtColumn = str(objValueStaffCodeAtColumn).strip()
            if pszStaffCodeAtColumn != pszStaffCode:
                continue

            for iRowIndex in range(2, iSheet6RowCount):
                objValueProject = objDataFrameSheet6.iat[iRowIndex, iColumnIndex]
                pszProjectNameFromSheet6: str = ""
                if objValueProject is not None:
                    pszProjectNameFromSheet6 = str(objValueProject).strip()
                if len(pszProjectNameFromSheet6) == 0:
                    continue

                objDataFrameSubProject: DataFrame = objDataFrameSubStaff[
                    objDataFrameSubStaff["プロジェクト名"] == pszProjectNameFromSheet6
                ]

                if objDataFrameSubProject.empty:
                    continue

                try:
                    iTotalSeconds: int = int(
                        objDataFrameSubProject["__time_seconds__"].sum()
                    )
                except Exception:
                    iTotalSeconds = 0

                pszTimeTotal: str = convert_step0009_seconds_to_time_string(
                    iTotalSeconds
                )
                pszCompanyName: str = ""
                if pszCompanyColumn != "":
                    try:
                        objCompanySeries = objDataFrameSubProject[pszCompanyColumn].dropna()
                        if not objCompanySeries.empty:
                            pszCompanyName = str(objCompanySeries.iloc[0])
                    except Exception:
                        pszCompanyName = ""

                objListOutputRowsProjectTask.append(
                    [pszProjectNameFromSheet6, pszStaffCode, pszTimeTotal],
                )
                objListOutputRowsProjectStaffCompany.append(
                    [pszProjectNameFromSheet6, pszCompanyName, pszStaffCode, pszTimeTotal],
                )

            break

    try:
        objDataFrameOutputProjectTask: DataFrame = DataFrame(
            objListOutputRowsProjectTask
        )
        objDataFrameOutputProjectTask.to_csv(
            pszProjectTaskOutputPath,
            sep="\t",
            index=False,
            header=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while writing step0009 project task TSV. Detail = {0}".format(
                objException
            ),
        )
        return

    try:
        objDataFrameOutputProjectStaffCompany: DataFrame = DataFrame(
            objListOutputRowsProjectStaffCompany
        )
        objDataFrameOutputProjectStaffCompany.to_csv(
            pszProjectStaffCompanyOutputPath,
            sep="\t",
            index=False,
            header=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while writing step0009 project staff-company TSV. Detail = {0}".format(
                objException
            ),
        )
        return


def normalize_step0009_company_name(pszCompanyName: str) -> str:
    objReplaceTargets: List[Tuple[str, str]] = [
        ("本部", "本部"),
        ("事業開発", "事業開発"),
        ("子会社", "子会社"),
        ("投資先", "投資先"),
        ("第１インキュ", "第一インキュ"),
        ("第２インキュ", "第二インキュ"),
        ("第３インキュ", "第三インキュ"),
        ("第４インキュ", "第四インキュ"),
        ("第1インキュ", "第一インキュ"),
        ("第2インキュ", "第二インキュ"),
        ("第3インキュ", "第三インキュ"),
        ("第4インキュ", "第四インキュ"),
    ]
    for pszPrefix, pszReplacement in objReplaceTargets:
        if pszCompanyName.startswith(pszPrefix):
            return pszReplacement
    return pszCompanyName


def make_step0009_project_company_task_tsv(
    pszProjectStaffCompanyPath: str,
    pszProjectCompanyPath: str,
) -> None:
    try:
        with open(pszProjectStaffCompanyPath, "r", encoding="utf-8") as objInputFile:
            with open(pszProjectCompanyPath, "w", encoding="utf-8") as objOutputFile:
                for pszLine in objInputFile:
                    pszLineContent = pszLine.rstrip("\n")
                    if pszLineContent == "":
                        objOutputFile.write("\n")
                        continue
                    objColumns = pszLineContent.split("\t")
                    if len(objColumns) > 1:
                        objColumns[1] = normalize_step0009_company_name(objColumns[1])
                    objOutputFile.write("\t".join(objColumns) + "\n")
    except Exception as objException:
        write_error_tsv(
            pszProjectCompanyPath,
            "Error: unexpected exception while writing step0009 company task TSV. Detail = {0}".format(
                objException
            ),
        )
        return


def make_step0010_project_manhour_tsv(
    pszProjectTaskPath: str,
    pszProjectCompanyTaskPath: str,
    pszProjectManhourPath: str,
    pszProjectCompanyManhourPath: str,
) -> None:
    if not os.path.isfile(pszProjectTaskPath):
        write_error_tsv(
            pszProjectManhourPath,
            "Error: step0009 project task TSV file not found. Path = {0}".format(
                pszProjectTaskPath
            ),
        )
        return

    if not os.path.isfile(pszProjectCompanyTaskPath):
        write_error_tsv(
            pszProjectCompanyManhourPath,
            "Error: step0009 project company task TSV file not found. Path = {0}".format(
                pszProjectCompanyTaskPath
            ),
        )
        return

    try:
        with open(pszProjectTaskPath, "r", encoding="utf-8") as objInputFile:
            with open(pszProjectManhourPath, "w", encoding="utf-8") as objOutputFile:
                for pszLine in objInputFile:
                    pszLineContent: str = pszLine.rstrip("\n")
                    if pszLineContent == "":
                        objOutputFile.write("\t\n")
                        continue
                    objColumns: List[str] = pszLineContent.split("\t")
                    pszProjectName: str = objColumns[0] if len(objColumns) > 0 else ""
                    if len(objColumns) > 2:
                        pszManhour = objColumns[2]
                    elif len(objColumns) > 1:
                        pszManhour = objColumns[1]
                    else:
                        pszManhour = ""
                    objOutputFile.write(pszProjectName + "\t" + pszManhour + "\n")
    except Exception as objException:
        write_error_tsv(
            pszProjectManhourPath,
            "Error: unexpected exception while writing step0010 project manhour TSV. Detail = {0}".format(
                objException
            ),
        )
        return

    try:
        with open(pszProjectCompanyTaskPath, "r", encoding="utf-8") as objInputFile:
            with open(
                pszProjectCompanyManhourPath, "w", encoding="utf-8"
            ) as objOutputFile:
                for pszLine in objInputFile:
                    pszLineContent = pszLine.rstrip("\n")
                    if pszLineContent == "":
                        objOutputFile.write("\t\t\n")
                        continue
                    objColumns = pszLineContent.split("\t")
                    pszProjectName = objColumns[0] if len(objColumns) > 0 else ""
                    pszCompanyName = objColumns[1] if len(objColumns) > 1 else ""
                    if len(objColumns) > 3:
                        pszManhour = objColumns[3]
                    elif len(objColumns) > 2:
                        pszManhour = objColumns[2]
                    elif len(objColumns) > 1:
                        pszManhour = objColumns[1]
                    else:
                        pszManhour = ""
                    objOutputFile.write(
                        pszProjectName
                        + "\t"
                        + pszCompanyName
                        + "\t"
                        + pszManhour
                        + "\n"
                    )
    except Exception as objException:
        write_error_tsv(
            pszProjectCompanyManhourPath,
            "Error: unexpected exception while writing step0010 project company manhour TSV. Detail = {0}".format(
                objException
            ),
        )
        return


def read_org_table_company_mappings(pszOrgTableTsvPath: str) -> List[Tuple[str, str]]:
    if not os.path.isfile(pszOrgTableTsvPath):
        raise FileNotFoundError(f"Org table TSV not found: {pszOrgTableTsvPath}")

    objMappings: List[Tuple[str, str]] = []
    try:
        objOrgDataFrame: DataFrame = pd.read_csv(
            pszOrgTableTsvPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        raise RuntimeError(
            "Error: unexpected exception while reading 管轄PJ表.tsv. Detail = {0}".format(
                objException
            )
        ) from objException

    if objOrgDataFrame.shape[1] < 3:
        raise ValueError("Error: 管轄PJ表.tsv must have at least 3 columns.")

    objColumnNames: List[str] = list(objOrgDataFrame.columns)
    pszProjectColumn: str = objColumnNames[1]
    pszCompanyColumn: str = objColumnNames[2]

    for _, objRow in objOrgDataFrame.iterrows():
        pszProjectCode: str = str(objRow[pszProjectColumn] or "")
        pszCompanyName: str = str(objRow[pszCompanyColumn] or "")
        objMappings.append((pszProjectCode, pszCompanyName))

    return objMappings


def read_org_table_billing_company_map(pszOrgTableTsvPath: str) -> Dict[str, str]:
    if not os.path.isfile(pszOrgTableTsvPath):
        return {}

    objBillingMap: Dict[str, str] = {}
    objEncodingCandidateList: List[str] = ["utf-8-sig", "cp932"]
    objLastDecodeError: Exception | None = None
    objRows: List[List[str]] = []

    for pszEncoding in objEncodingCandidateList:
        try:
            with open(
                pszOrgTableTsvPath,
                mode="r",
                encoding=pszEncoding,
                newline="",
            ) as objOrgTableFile:
                objOrgTableReader = csv.reader(objOrgTableFile, delimiter="\t")
                for objRow in objOrgTableReader:
                    objRows.append(list(objRow))
            objLastDecodeError = None
            break
        except UnicodeDecodeError as objError:
            objLastDecodeError = objError
            objRows = []

    if objLastDecodeError is not None:
        raise objLastDecodeError

    for objRow in objRows:
        if len(objRow) < 3:
            continue
        pszProjectCodeOrg: str = str(objRow[1]).strip()
        pszBillingCompany: str = str(objRow[2]).strip()
        if not pszProjectCodeOrg or not pszBillingCompany:
            continue
        objMatch: re.Match[str] | None = re.match(r"^(P\d{5}|[A-OQ-Z]\d{3})", pszProjectCodeOrg)
        if objMatch is None:
            continue
        pszProjectCodePrefix: str = objMatch.group(1)
        if pszProjectCodePrefix not in objBillingMap:
            objBillingMap[pszProjectCodePrefix] = pszBillingCompany

    return objBillingMap


def read_org_table_billing_group_map(pszOrgTableTsvPath: str) -> Dict[str, str]:
    if not os.path.isfile(pszOrgTableTsvPath):
        return {}

    objGroupMap: Dict[str, str] = {}
    objEncodingCandidateList: List[str] = ["utf-8-sig", "cp932"]
    objLastDecodeError: Exception | None = None
    objRows: List[List[str]] = []

    for pszEncoding in objEncodingCandidateList:
        try:
            with open(
                pszOrgTableTsvPath,
                mode="r",
                encoding=pszEncoding,
                newline="",
            ) as objOrgTableFile:
                objOrgTableReader = csv.reader(objOrgTableFile, delimiter="\t")
                for objRow in objOrgTableReader:
                    objRows.append(list(objRow))
            objLastDecodeError = None
            break
        except UnicodeDecodeError as objError:
            objLastDecodeError = objError
            objRows = []

    if objLastDecodeError is not None:
        raise objLastDecodeError

    for objRow in objRows:
        if len(objRow) < 4:
            continue
        pszProjectCodeOrg: str = str(objRow[1]).strip()
        pszBillingGroup: str = str(objRow[3]).strip()
        if not pszProjectCodeOrg or not pszBillingGroup:
            continue
        objMatch: re.Match[str] | None = re.match(
            r"^(P\d{5}_|[A-OQ-Z]\d{3}_)", pszProjectCodeOrg
        )
        if objMatch is None:
            continue
        pszProjectCodePrefix: str = objMatch.group(1)
        if pszProjectCodePrefix not in objGroupMap:
            objGroupMap[pszProjectCodePrefix] = pszBillingGroup

    return objGroupMap


def extract_project_code_prefix_step0012(pszProjectName: str) -> str:
    iUnderscoreIndex: int = pszProjectName.find("_")
    if iUnderscoreIndex == -1:
        return pszProjectName
    return pszProjectName[:iUnderscoreIndex]


def sort_rows_by_project_prefix_step0012(
    objRows: List[Tuple[str, ...]],
) -> List[Tuple[str, ...]]:
    objIndexedRows: List[Tuple[int, Tuple[str, ...]]] = list(enumerate(objRows))
    objIndexedRows.sort(
        key=lambda objItem: (
            extract_project_code_prefix_step0012(objItem[1][0]),
            objItem[0],
        ),
    )
    return [objRow for _, objRow in objIndexedRows]


def make_step0012_project_manhour_tsv(
    pszStep0011ProjectManhourPath: str,
    pszStep0011ProjectCompanyManhourPath: str,
    pszStep0012ProjectManhourPath: str,
    pszStep0012ProjectCompanyManhourPath: str,
    pszStep0012ProjectCompanyGroupManhourPath: str,
    pszOrgTableTsvPath: str,
) -> None:
    if not os.path.isfile(pszStep0011ProjectManhourPath):
        write_error_tsv(
            pszStep0012ProjectManhourPath,
            "Error: step0011 project manhour TSV file not found. Path = {0}".format(
                pszStep0011ProjectManhourPath
            ),
        )
        return

    if not os.path.isfile(pszStep0011ProjectCompanyManhourPath):
        write_error_tsv(
            pszStep0012ProjectCompanyManhourPath,
            "Error: step0011 project company manhour TSV file not found. Path = {0}".format(
                pszStep0011ProjectCompanyManhourPath
            ),
        )
        return

    objStep0011Rows: List[Tuple[str, str]] = []
    with open(pszStep0011ProjectManhourPath, "r", encoding="utf-8") as objInputFile:
        for pszLine in objInputFile:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objStep0011Rows.append(("", ""))
                continue
            objColumns: List[str] = pszLineContent.split("\t")
            pszProjectName: str = objColumns[0] if len(objColumns) > 0 else ""
            pszManhour: str = objColumns[1] if len(objColumns) > 1 else ""
            objStep0011Rows.append((pszProjectName, pszManhour))

    objStep0011CompanyRows: List[Tuple[str, str, str]] = []
    with open(
        pszStep0011ProjectCompanyManhourPath, "r", encoding="utf-8"
    ) as objInputFile:
        for pszLine in objInputFile:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objStep0011CompanyRows.append(("", "", ""))
                continue
            objColumns: List[str] = pszLineContent.split("\t")
            pszProjectName: str = objColumns[0] if len(objColumns) > 0 else ""
            pszCompanyName: str = objColumns[1] if len(objColumns) > 1 else ""
            pszManhour: str = objColumns[2] if len(objColumns) > 2 else ""
            objStep0011CompanyRows.append(
                (pszProjectName, pszCompanyName, pszManhour)
            )

    objSortedStep0011Rows: List[Tuple[str, str]] = sort_rows_by_project_prefix_step0012(
        objStep0011Rows
    )
    objSortedStep0011CompanyRows: List[Tuple[str, str, str]] = (
        sort_rows_by_project_prefix_step0012(objStep0011CompanyRows)
    )

    with open(pszStep0012ProjectManhourPath, "w", encoding="utf-8") as objOutputFile:
        for pszProjectName, pszManhour in objSortedStep0011Rows:
            objOutputFile.write(pszProjectName + "\t" + pszManhour + "\n")

    with open(
        pszStep0012ProjectCompanyManhourPath, "w", encoding="utf-8"
    ) as objOutputFile:
        for pszProjectName, pszCompanyName, pszManhour in objSortedStep0011CompanyRows:
            objOutputFile.write(
                pszProjectName
                + "\t"
                + pszCompanyName
                + "\t"
                + pszManhour
                + "\n"
            )

    try:
        objOrgTableGroupMap: Dict[str, str] = read_org_table_billing_group_map(
            pszOrgTableTsvPath
        )
    except Exception as objException:
        write_error_tsv(
            pszStep0012ProjectCompanyGroupManhourPath,
            "Error: unexpected exception while reading 管轄PJ表.tsv. Detail = {0}".format(
                objException
            ),
        )
        return

    with open(
        pszStep0012ProjectCompanyGroupManhourPath, "w", encoding="utf-8"
    ) as objOutputFile:
        for pszProjectName, pszCompanyName, pszManhour in objSortedStep0011CompanyRows:
            if "_" in pszProjectName:
                pszProjectCodePrefix: str = pszProjectName.split("_", 1)[0] + "_"
            else:
                pszProjectCodePrefix = pszProjectName
            pszBillingGroup: str = objOrgTableGroupMap.get(pszProjectCodePrefix, "")
            objOutputFile.write(
                pszProjectName
                + "\t"
                + pszCompanyName
                + "\t"
                + pszBillingGroup
                + "\t"
                + pszManhour
                + "\n"
            )


def make_step0013_project_manhour_tsv(
    pszStep0012ProjectManhourPath: str,
    pszStep0012ProjectCompanyManhourPath: str,
    pszStep0012ProjectCompanyGroupManhourPath: str,
    pszStep0013ProjectManhourPath: str,
    pszStep0013ProjectCompanyManhourPath: str,
    pszStep0013ProjectCompanyGroupManhourPath: str,
) -> None:
    if not os.path.isfile(pszStep0012ProjectManhourPath):
        write_error_tsv(
            pszStep0013ProjectManhourPath,
            "Error: step0012 project manhour TSV file not found. Path = {0}".format(
                pszStep0012ProjectManhourPath
            ),
        )
        return

    if not os.path.isfile(pszStep0012ProjectCompanyManhourPath):
        write_error_tsv(
            pszStep0013ProjectCompanyManhourPath,
            "Error: step0012 project company manhour TSV file not found. Path = {0}".format(
                pszStep0012ProjectCompanyManhourPath
            ),
        )
        return

    if not os.path.isfile(pszStep0012ProjectCompanyGroupManhourPath):
        write_error_tsv(
            pszStep0013ProjectCompanyGroupManhourPath,
            "Error: step0012 project company group manhour TSV file not found. Path = {0}".format(
                pszStep0012ProjectCompanyGroupManhourPath
            ),
        )
        return

    objStep0012Rows: List[Tuple[str, str]] = []
    with open(pszStep0012ProjectManhourPath, "r", encoding="utf-8") as objInputFile:
        for pszLine in objInputFile:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objStep0012Rows.append(("", ""))
                continue
            objColumns: List[str] = pszLineContent.split("\t")
            pszProjectName: str = objColumns[0] if len(objColumns) > 0 else ""
            pszManhour: str = objColumns[1] if len(objColumns) > 1 else ""
            objStep0012Rows.append((pszProjectName, pszManhour))

    objStep0012CompanyRows: List[Tuple[str, str, str]] = []
    with open(
        pszStep0012ProjectCompanyManhourPath, "r", encoding="utf-8"
    ) as objInputFile:
        for pszLine in objInputFile:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objStep0012CompanyRows.append(("", "", ""))
                continue
            objColumns: List[str] = pszLineContent.split("\t")
            pszProjectName: str = objColumns[0] if len(objColumns) > 0 else ""
            pszCompanyName: str = objColumns[1] if len(objColumns) > 1 else ""
            pszManhour: str = objColumns[2] if len(objColumns) > 2 else ""
            objStep0012CompanyRows.append(
                (pszProjectName, pszCompanyName, pszManhour)
            )

    objStep0012CompanyGroupRows: List[Tuple[str, str, str, str]] = []
    with open(
        pszStep0012ProjectCompanyGroupManhourPath, "r", encoding="utf-8"
    ) as objInputFile:
        for pszLine in objInputFile:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objStep0012CompanyGroupRows.append(("", "", "", ""))
                continue
            objColumns: List[str] = pszLineContent.split("\t")
            pszProjectName: str = objColumns[0] if len(objColumns) > 0 else ""
            pszCompanyName: str = objColumns[1] if len(objColumns) > 1 else ""
            pszBillingGroup: str = objColumns[2] if len(objColumns) > 2 else ""
            pszManhour: str = objColumns[3] if len(objColumns) > 3 else ""
            objStep0012CompanyGroupRows.append(
                (pszProjectName, pszCompanyName, pszBillingGroup, pszManhour)
            )

    with open(pszStep0013ProjectManhourPath, "w", encoding="utf-8") as objOutputFile:
        for pszProjectName, pszManhour in objStep0012Rows:
            if str(pszProjectName).startswith(("A", "H")):
                continue
            objOutputFile.write(pszProjectName + "\t" + pszManhour + "\n")

    with open(
        pszStep0013ProjectCompanyManhourPath, "w", encoding="utf-8"
    ) as objOutputFile:
        for pszProjectName, pszCompanyName, pszManhour in objStep0012CompanyRows:
            if str(pszProjectName).startswith(("A", "H")):
                continue
            objOutputFile.write(
                pszProjectName + "\t" + pszCompanyName + "\t" + pszManhour + "\n"
            )

    with open(
        pszStep0013ProjectCompanyGroupManhourPath, "w", encoding="utf-8"
    ) as objOutputFile:
        for (
            pszProjectName,
            pszCompanyName,
            pszBillingGroup,
            pszManhour,
        ) in objStep0012CompanyGroupRows:
            if str(pszProjectName).startswith(("A", "H")):
                continue
            objOutputFile.write(
                pszProjectName
                + "\t"
                + pszCompanyName
                + "\t"
                + pszBillingGroup
                + "\t"
                + pszManhour
                + "\n"
            )


def make_step14_project_company_manhour_tsv(
    pszStep0012ProjectManhourPath: str,
    pszStep14ProjectCompanyManhourPath: str,
    pszOrgTableTsvPath: str,
) -> None:
    if not os.path.isfile(pszStep0012ProjectManhourPath):
        write_error_tsv(
            pszStep14ProjectCompanyManhourPath,
            "Error: step0012 project manhour TSV file not found. Path = {0}".format(
                pszStep0012ProjectManhourPath
            ),
        )
        return

    try:
        objOrgTableBillingMap: Dict[str, str] = read_org_table_billing_company_map(
            pszOrgTableTsvPath
        )
    except Exception as objException:
        write_error_tsv(
            pszStep14ProjectCompanyManhourPath,
            "Error: unexpected exception while reading 管轄PJ表.tsv. Detail = {0}".format(
                objException
            ),
        )
        return

    with open(pszStep0012ProjectManhourPath, "r", encoding="utf-8") as objInputFile, open(
        pszStep14ProjectCompanyManhourPath,
        "w",
        encoding="utf-8",
    ) as objOutputFile:
        pszZeroManhour: str = "0:00:00"
        for pszLine in objInputFile:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objOutputFile.write("\n")
                continue
            objColumns: List[str] = pszLineContent.split("\t")
            if len(objColumns) < 2:
                print(f"Warning: 不正な行をスキップしました: {pszLineContent}")
                continue
            pszProjectName: str = objColumns[0]
            pszManhour: str = objColumns[1]
            pszProjectCodePrefix: str = pszProjectName.split("_", 1)[0]
            pszBillingCompany: str = objOrgTableBillingMap.get(pszProjectCodePrefix, "")

            pszFirstIncubation: str = pszZeroManhour
            pszSecondIncubation: str = pszZeroManhour
            pszThirdIncubation: str = pszZeroManhour
            pszFourthIncubation: str = pszZeroManhour
            pszBusinessDevelopment: str = pszZeroManhour
            bIsCompanyProject: bool = re.match(r"^C\d{3}_", str(pszProjectName)) is not None
            if not bIsCompanyProject:
                if pszBillingCompany == "第一インキュ":
                    pszFirstIncubation = pszManhour
                elif pszBillingCompany == "第二インキュ":
                    pszSecondIncubation = pszManhour
                elif pszBillingCompany == "第三インキュ":
                    pszThirdIncubation = pszManhour
                elif pszBillingCompany == "第四インキュ":
                    pszFourthIncubation = pszManhour
                elif pszBillingCompany == "事業開発":
                    pszBusinessDevelopment = pszManhour
            objOutputFile.write(
                pszProjectName
                + "\t"
                + pszBillingCompany
                + "\t"
                + pszManhour
                + "\t"
                + pszFirstIncubation
                + "\t"
                + pszSecondIncubation
                + "\t"
                + pszThirdIncubation
                + "\t"
                + pszFourthIncubation
                + "\t"
                + pszBusinessDevelopment
                + "\n"
            )

def make_step0011_project_manhour_tsv(
    pszProjectManhourPath: str,
    pszProjectCompanyManhourPath: str,
    pszProjectManhourOutputPath: str,
    pszProjectCompanyManhourOutputPath: str,
    pszOrgTableTsvPath: str,
    objBaseDirectoryPath: Path,
) -> None:
    if not os.path.isfile(pszProjectManhourPath):
        write_error_tsv(
            pszProjectManhourOutputPath,
            "Error: step0010 project manhour TSV file not found. Path = {0}".format(
                pszProjectManhourPath
            ),
        )
        return

    if not os.path.isfile(pszProjectCompanyManhourPath):
        write_error_tsv(
            pszProjectCompanyManhourOutputPath,
            "Error: step0010 project company manhour TSV file not found. Path = {0}".format(
                pszProjectCompanyManhourPath
            ),
        )
        return

    objSheet0010Rows: List[Tuple[str, str]] = []
    with open(pszProjectManhourPath, "r", encoding="utf-8") as objInputFile:
        for pszLine in objInputFile:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objSheet0010Rows.append(("", ""))
                continue
            objColumns: List[str] = pszLineContent.split("\t")
            pszProjectName: str = objColumns[0] if len(objColumns) > 0 else ""
            pszManhour: str = objColumns[1] if len(objColumns) > 1 else ""
            objSheet0010Rows.append((pszProjectName, pszManhour))

    objAggregatedSeconds: Dict[str, int] = {}
    objAggregatedOrder: List[str] = []
    for pszProjectName, pszManhour in objSheet0010Rows:
        if pszProjectName == "" and pszManhour == "":
            continue
        iSeconds = convert_step0009_time_string_to_seconds(pszManhour)
        if pszProjectName not in objAggregatedSeconds:
            objAggregatedSeconds[pszProjectName] = 0
            objAggregatedOrder.append(pszProjectName)
        objAggregatedSeconds[pszProjectName] += iSeconds

    with open(pszProjectManhourOutputPath, "w", encoding="utf-8") as objOutputFile:
        for pszProjectName in objAggregatedOrder:
            pszTotalManhour: str = convert_step0009_seconds_to_time_string(
                objAggregatedSeconds[pszProjectName],
            )
            objOutputFile.write(pszProjectName + "\t" + pszTotalManhour + "\n")

    objSheet0010CompanyRows: List[Tuple[str, str, str]] = []
    with open(pszProjectCompanyManhourPath, "r", encoding="utf-8") as objInputFile:
        for pszLine in objInputFile:
            pszLineContent = pszLine.rstrip("\n")
            if pszLineContent == "":
                continue
            objColumns = pszLineContent.split("\t")
            pszProjectName = objColumns[0] if len(objColumns) > 0 else ""
            pszCompanyName = objColumns[1] if len(objColumns) > 1 else ""
            pszManhour = objColumns[2] if len(objColumns) > 2 else ""
            objSheet0010CompanyRows.append((pszProjectName, pszCompanyName, pszManhour))

    objAggregatedCompanySeconds: Dict[str, int] = {}
    objAggregatedCompanyOrder: List[str] = []
    objAggregatedCompanyNames: Dict[str, List[str]] = {}
    for pszProjectName, pszCompanyName, pszManhour in objSheet0010CompanyRows:
        if pszProjectName == "" and pszCompanyName == "" and pszManhour == "":
            continue
        iSeconds = convert_step0009_time_string_to_seconds(pszManhour)
        if pszProjectName not in objAggregatedCompanySeconds:
            objAggregatedCompanySeconds[pszProjectName] = 0
            objAggregatedCompanyOrder.append(pszProjectName)
            objAggregatedCompanyNames[pszProjectName] = []
        if pszCompanyName not in objAggregatedCompanyNames[pszProjectName]:
            objAggregatedCompanyNames[pszProjectName].append(pszCompanyName)
        objAggregatedCompanySeconds[pszProjectName] += iSeconds

    objIncubationPriority: List[str] = [
        "第一インキュ",
        "第二インキュ",
        "第三インキュ",
        "第四インキュ",
    ]
    objIncubationPrioritySet: set[str] = set(objIncubationPriority)

    try:
        objOrgTableBillingMap: Dict[str, str] = read_org_table_billing_company_map(
            pszOrgTableTsvPath
        )
    except Exception as objException:
        write_error_tsv(
            pszProjectCompanyManhourOutputPath,
            "Error: unexpected exception while reading 管轄PJ表.tsv. Detail = {0}".format(
                objException
            ),
        )
        return

    objHoldProjectLines: List[str] = []
    objMismatchProjectLines: List[str] = []

    def extract_project_code_prefix(pszProjectName: str) -> str:
        iUnderscoreIndex: int = pszProjectName.find("_")
        if iUnderscoreIndex == -1:
            return pszProjectName
        return pszProjectName[:iUnderscoreIndex]

    def select_company_name_step0011(
        pszProjectName: str,
        objCompanyNames: List[str],
    ) -> str:
        if not objCompanyNames:
            return ""
        pszProjectCodePrefix: str = extract_project_code_prefix(pszProjectName)
        pszProjectPrefix: str = pszProjectCodePrefix[:1]
        if pszProjectCodePrefix in objOrgTableBillingMap:
            pszBillingCompany: str = objOrgTableBillingMap[pszProjectCodePrefix]
            if pszProjectPrefix in ["J", "P"] and pszBillingCompany not in objCompanyNames:
                objMismatchProjectLines.append(
                    f"{pszProjectName} → {' / '.join(objCompanyNames)} (管轄PJ表={pszBillingCompany})",
                )
            return pszBillingCompany
        if pszProjectPrefix in ["A", "H"]:
            return "本部"
        return objCompanyNames[0]

    objSheet0011CompanyRows: List[Tuple[str, str, str]] = []
    with open(pszProjectCompanyManhourOutputPath, "w", encoding="utf-8") as objOutputFile:
        for pszProjectName in objAggregatedCompanyOrder:
            pszTotalManhour = convert_step0009_seconds_to_time_string(
                objAggregatedCompanySeconds[pszProjectName],
            )
            objCompanyNames = objAggregatedCompanyNames.get(pszProjectName, [])
            objIncubations = [
                name for name in objCompanyNames if name in objIncubationPrioritySet
            ]
            if len(objIncubations) > 1:
                objHoldProjectLines.append(
                    f"{pszProjectName} → {' / '.join(objCompanyNames)}",
                )
            pszCompanyName = select_company_name_step0011(
                pszProjectName,
                objCompanyNames,
            )
            objOutputFile.write(
                pszProjectName + "\t" + pszCompanyName + "\t" + pszTotalManhour + "\n",
            )
            objSheet0011CompanyRows.append((pszProjectName, pszCompanyName, pszTotalManhour))

    if objHoldProjectLines or objMismatchProjectLines:
        pszCompanyTsvLine: str = f"対象TSV: {pszProjectCompanyManhourPath}"
        print(pszCompanyTsvLine)
        write_debug_error(pszCompanyTsvLine, objBaseDirectoryPath)
        for pszLine in objHoldProjectLines:
            print(pszLine)
            write_debug_error(pszLine, objBaseDirectoryPath)
        for pszLine in objMismatchProjectLines:
            print(pszLine)
            write_debug_error(pszLine, objBaseDirectoryPath)
        objMessageParts: List[str] = []
        if objHoldProjectLines:
            objMessageParts.append("インキュがかぶっているプロジェクトがあります。")
            objMessageParts.extend(objHoldProjectLines)
        if objMismatchProjectLines:
            objMessageParts.append(
                "管轄PJ表のカンパニー名と一致しないプロジェクトがあります。"
            )
            objMessageParts.extend(objMismatchProjectLines)
        objMessage = pszCompanyTsvLine + "\n" + "\n".join(objMessageParts)
        objRoot = tk.Tk()
        objRoot.withdraw()
        messagebox.showwarning("警告", objMessage)
        objRoot.destroy()

def make_step0005_remove_ah_project_tsv(
    pszInputFileFullPath: str,
    pszOutputFileFullPath: str,
) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: input TSV file not found for A/H removal. "
            "Path = {0}".format(pszInputFileFullPath),
        )
        return

    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for A/H removal. "
            "Detail = {0}".format(objException),
        )
        return

    if objDataFrameInput.shape[1] < 7:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required column G does not exist (need at least 7 columns). "
            "ColumnCount = {0}".format(objDataFrameInput.shape[1]),
        )
        return

    objColumnNames: List[str] = list(objDataFrameInput.columns)
    pszProjectColumn: str = objColumnNames[6]

    objProjectSeries = objDataFrameInput[pszProjectColumn].fillna("").astype(str)
    objKeepMask = ~objProjectSeries.str.startswith(("A", "H"))
    objDataFrameOutput: DataFrame = objDataFrameInput.loc[objKeepMask].copy()

    try:
        objDataFrameOutput.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing A/H removed TSV. "
            "Detail = {0}".format(objException),
        )
        return


def make_step0006_company_replaced_tsv_from_step0005(
    pszInputFileFullPath: str,
    pszOrgTableTsvPath: str,
    pszOutputFileFullPath: str,
    pszMissingOutputFileFullPath: str,
) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: input TSV file not found for company replacement. "
            "Path = {0}".format(pszInputFileFullPath),
        )
        return

    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for company replacement. "
            "Detail = {0}".format(objException),
        )
        return

    if objDataFrameInput.shape[1] < 7:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required columns D/G do not exist (need at least 7 columns). "
            "ColumnCount = {0}".format(objDataFrameInput.shape[1]),
        )
        return

    try:
        objMappings: List[Tuple[str, str]] = read_org_table_company_mappings(
            pszOrgTableTsvPath
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading org table TSV. "
            "Detail = {0}".format(objException),
        )
        return

    objColumnNames: List[str] = list(objDataFrameInput.columns)
    pszCompanyColumn: str = objColumnNames[3]
    pszProjectColumn: str = objColumnNames[6]

    objCompanyValues: List[str] = []
    objMissingMask: List[bool] = []

    for _, objRow in objDataFrameInput.iterrows():
        pszProjectCode: str = str(objRow[pszProjectColumn] or "")
        pszNewCompany: str | None = None
        for pszOrgProjectCode, pszOrgCompanyName in objMappings:
            if pszOrgProjectCode != "" and pszOrgProjectCode.startswith(pszProjectCode):
                pszNewCompany = pszOrgCompanyName
                break
        if pszNewCompany is None:
            objCompanyValues.append(str(objRow[pszCompanyColumn] or ""))
            objMissingMask.append(True)
        else:
            objCompanyValues.append(pszNewCompany)
            objMissingMask.append(False)

    objDataFrameOutput: DataFrame = objDataFrameInput.copy()
    objDataFrameOutput[pszCompanyColumn] = objCompanyValues
    objMatchedDataFrame: DataFrame = objDataFrameOutput[[not is_missing for is_missing in objMissingMask]]
    objMissingDataFrame: DataFrame = objDataFrameInput[objMissingMask]

    try:
        objMatchedDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing company replaced TSV. "
            "Detail = {0}".format(objException),
        )
        return

    try:
        objMissingDataFrame.to_csv(
            pszMissingOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszMissingOutputFileFullPath,
            "Error: unexpected exception while writing missing project list TSV. "
            "Detail = {0}".format(objException),
        )
        return


def make_step0006_unique_missing_project_tsv(
    pszInputFileFullPath: str,
    pszOutputFileFullPath: str,
) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: input TSV file not found for unique missing projects. "
            "Path = {0}".format(pszInputFileFullPath),
        )
        return

    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading missing project TSV. "
            "Detail = {0}".format(objException),
        )
        return

    if objDataFrameInput.shape[1] < 8:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required columns G-H do not exist (need at least 8 columns). "
            "ColumnCount = {0}".format(objDataFrameInput.shape[1]),
        )
        return

    objColumnNames: List[str] = list(objDataFrameInput.columns)
    pszProjectColumn: str = objColumnNames[6]
    pszProjectNameColumn: str = objColumnNames[7]

    try:
        objUniqueDataFrame: DataFrame = objDataFrameInput.drop_duplicates(
            subset=[pszProjectColumn]
        )
        objOutputDataFrame: DataFrame = objUniqueDataFrame[
            [pszProjectColumn, pszProjectNameColumn]
        ].copy()
        objOutputDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing unique missing project TSV. "
            "Detail = {0}".format(objException),
        )
        return


def make_step0006_sort_asc_missing_project_tsv(
    pszInputFileFullPath: str,
    pszOutputFileFullPath: str,
) -> None:
    if not os.path.isfile(pszInputFileFullPath):
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: input TSV file not found for sorted missing projects. "
            "Path = {0}".format(pszInputFileFullPath),
        )
        return

    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading unique missing project TSV. "
            "Detail = {0}".format(objException),
        )
        return

    if objDataFrameInput.shape[1] < 1:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required project code column does not exist. "
            "ColumnCount = {0}".format(objDataFrameInput.shape[1]),
        )
        return

    try:
        objColumnNames: List[str] = list(objDataFrameInput.columns)
        pszProjectColumn: str = objColumnNames[0]
        objSortedDataFrame: DataFrame = objDataFrameInput.sort_values(
            by=pszProjectColumn,
            ascending=True,
            kind="mergesort",
        )
        objSortedDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing sorted missing project TSV. "
            "Detail = {0}".format(objException),
        )
        return
def write_org_table_tsv_from_csv(objBaseDirectoryPath: Path) -> None:
    objScriptDirectoryPath: Path = Path(__file__).resolve().parent
    objOrgTableCsvPath: Path = objScriptDirectoryPath / "管轄PJ表.csv"
    if not objOrgTableCsvPath.exists():
        objOrgTableCsvPath = objBaseDirectoryPath / "管轄PJ表.csv"

    objOrgTableTsvPath: Path = objBaseDirectoryPath / "管轄PJ表.tsv"

    if not objOrgTableCsvPath.exists():
        write_error_tsv(
            str(objOrgTableTsvPath),
            "Error: 管轄PJ表.csv が見つかりません。Path = {0}".format(objOrgTableCsvPath),
        )
        return

    objRows: List[List[str]] = []
    arrEncodings: List[str] = ["utf-8-sig", "cp932"]
    objLastDecodeError: Exception | None = None

    for pszEncoding in arrEncodings:
        try:
            with open(
                objOrgTableCsvPath,
                mode="r",
                encoding=pszEncoding,
                newline="",
            ) as objInputFile:
                objReader: csv.reader = csv.reader(objInputFile)
                for objRow in objReader:
                    objRows.append(list(objRow))
            objLastDecodeError = None
            break
        except UnicodeDecodeError as objError:
            objLastDecodeError = objError
            objRows = []

    if objLastDecodeError is not None:
        write_error_tsv(
            str(objOrgTableTsvPath),
            "Error: unexpected exception while reading 管轄PJ表.csv. Detail = {0}".format(
                objLastDecodeError
            ),
        )
        return

    for iRowIndex, objRow in enumerate(objRows):
        if len(objRow) > 1:
            objRow[1] = normalize_org_table_project_code_step0004(objRow[1])
        objRows[iRowIndex] = objRow

    objOrgTableTsvPath.parent.mkdir(parents=True, exist_ok=True)
    with open(objOrgTableTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        for objRow in objRows:
            objWriter.writerow(objRow)


def process_single_input(
    pszInputManhourCsvPath: str,
) -> tuple[int, Path | None, int | None, int | None, str | None]:
    objInputPath: Path = Path(pszInputManhourCsvPath)
    objCandidatePaths: List[Path] = [objInputPath]

    objScriptDirectoryPath: Path = Path(__file__).resolve().parent
    objCandidatePaths.append(objScriptDirectoryPath / pszInputManhourCsvPath)

    objInputDirectoryPath: Path = Path.cwd() / "input"
    objCandidatePaths.append(objInputDirectoryPath / pszInputManhourCsvPath)

    if objInputPath.suffix.lower() == ".tsv":
        pszCsvFileName: str = objInputPath.with_suffix(".csv").name
        objCandidatePaths.append(objInputPath.with_suffix(".csv"))
        objCandidatePaths.append(objScriptDirectoryPath / pszCsvFileName)
        objCandidatePaths.append(objInputDirectoryPath / pszCsvFileName)

    objExistingPaths: List[Path] = [objPath for objPath in objCandidatePaths if objPath.exists()]
    if len(objExistingPaths) > 0:
        objInputPath = objExistingPaths[0]

    if not objInputPath.exists():
        pszErrorTextFilePath: str = str(Path.cwd() / "make_manhour_to_sheet8_01_0002_error.txt")
        write_error_text_utf8(
            pszErrorTextFilePath,
            f"Error: input file not found: {pszInputManhourCsvPath}\n"
            f"CurrentDirectory: {str(Path.cwd())}\n",
        )
        raise FileNotFoundError(f"Input file not found: {pszInputManhourCsvPath}")

    objBaseDirectoryPath: Path = objInputPath.resolve().parent

    pszStep1DefaultTsvPath: str = convert_csv_to_tsv_file(str(objInputPath))
    iFileYear, iFileMonth = get_target_year_month_from_filename(str(objInputPath))
    pszStep1TsvPath: str = str(
        objBaseDirectoryPath / f"工数_{iFileYear}年{iFileMonth:02d}月.tsv"
    )
    if pszStep1DefaultTsvPath != pszStep1TsvPath:
        os.replace(pszStep1DefaultTsvPath, pszStep1TsvPath)

    make_removed_uninput_tsv_from_manhour_tsv(pszStep1TsvPath)
    pszStep0001TsvPath: str = build_removed_uninput_output_path(pszStep1TsvPath)
    make_sorted_staff_code_tsv_from_manhour_tsv(pszStep0001TsvPath)
    pszStep0002TsvPath: str = build_sorted_staff_code_output_path(pszStep0001TsvPath)
    make_company_normalized_tsv_from_step0002(pszStep0002TsvPath)
    pszStep0003TsvPath: str = build_step0003_company_normalized_output_path(
        pszStep0002TsvPath
    )
    make_company_normalized_tsv_from_step0003(pszStep0003TsvPath)
    pszStep0004TsvPath: str = build_step0004_company_normalized_output_path(
        pszStep0003TsvPath
    )

    return 0, objBaseDirectoryPath, iFileYear, iFileMonth, pszStep0004TsvPath


def main() -> int:
    objParser: argparse.ArgumentParser = argparse.ArgumentParser()
    objParser.add_argument(
        "pszInputManhourCsvPaths",
        nargs="+",
        help="Input Jobcan manhour CSV file paths",
    )
    objArgs: argparse.Namespace = objParser.parse_args()

    objScriptDirectoryPath: Path = Path(__file__).resolve().parent

    iExitCode: int = 0
    for pszInputManhourCsvPath in objArgs.pszInputManhourCsvPaths:
        try:
            iResult, objBaseDirectoryPath, iYear, iMonth, pszStep0004TsvPath = (
                process_single_input(pszInputManhourCsvPath)
            )
        except Exception as objException:
            print(
                "Error: failed to process input file: {0}. Detail = {1}".format(
                    pszInputManhourCsvPath,
                    objException,
                )
            )
            iExitCode = 1
            continue
        if iResult != 0:
            iExitCode = 1
        elif (
            objBaseDirectoryPath is not None
            and iYear is not None
            and iMonth is not None
            and pszStep0004TsvPath is not None
        ):
            write_org_table_tsv_from_csv(objBaseDirectoryPath)
            objOrgTableTsvPath: Path = objBaseDirectoryPath / "管轄PJ表.tsv"
            objStep0005Path: Path = build_step0005_remove_ah_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            make_step0005_remove_ah_project_tsv(
                pszStep0004TsvPath,
                str(objStep0005Path),
            )
            objStep0006Path: Path = build_step0006_company_replaced_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            objStep0006MissingPath: Path = build_step0006_missing_project_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            objStep0006UniqueMissingPath: Path = (
                build_step0006_unique_missing_project_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            objStep0006SortAscMissingPath: Path = (
                build_step0006_sort_asc_missing_project_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            make_step0006_company_replaced_tsv_from_step0005(
                str(objStep0005Path),
                str(objOrgTableTsvPath),
                str(objStep0006Path),
                str(objStep0006MissingPath),
            )
            objStep0007Path: Path = build_step0007_yyyy_mm_dd_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            make_step0007_yyyy_mm_dd_tsv(
                str(objStep0006Path),
                str(objStep0007Path),
            )
            make_step0007_unique_staff_code_tsv(str(objStep0007Path))
            make_step0007_staff_code_range_tsv(str(objStep0007Path))
            objStep0007StaffRangePath: str = build_step0007_staff_code_range_output_path(
                str(objStep0007Path)
            )
            objStep0008Path: Path = build_step0008_staff_project_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            make_step0008_staff_project_tsv(
                str(objStep0007Path),
                str(objStep0007StaffRangePath),
                str(objStep0008Path),
            )
            objStep0009ProjectTaskPath: Path = build_step0009_project_task_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            objStep0009ProjectStaffCompanyPath: Path = (
                build_step0009_project_staff_company_task_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            objStep0009ProjectCompanyPath: Path = (
                build_step0009_project_company_task_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            make_step0009_project_task_tsv(
                str(objStep0007Path),
                str(objStep0007StaffRangePath),
                str(objStep0008Path),
                str(objStep0009ProjectTaskPath),
                str(objStep0009ProjectStaffCompanyPath),
            )
            make_step0009_project_company_task_tsv(
                str(objStep0009ProjectStaffCompanyPath),
                str(objStep0009ProjectCompanyPath),
            )
            objStep0010ProjectManhourPath: Path = build_step0010_project_manhour_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            objStep0010ProjectCompanyManhourPath: Path = (
                build_step0010_project_company_manhour_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            make_step0010_project_manhour_tsv(
                str(objStep0009ProjectTaskPath),
                str(objStep0009ProjectCompanyPath),
                str(objStep0010ProjectManhourPath),
                str(objStep0010ProjectCompanyManhourPath),
            )
            objStep0011ProjectManhourPath: Path = build_step0011_project_manhour_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            objStep0011ProjectCompanyManhourPath: Path = (
                build_step0011_project_company_manhour_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            make_step0011_project_manhour_tsv(
                str(objStep0010ProjectManhourPath),
                str(objStep0010ProjectCompanyManhourPath),
                str(objStep0011ProjectManhourPath),
                str(objStep0011ProjectCompanyManhourPath),
                str(objOrgTableTsvPath),
                objBaseDirectoryPath,
            )
            objStep0012ProjectManhourPath: Path = build_step0012_project_manhour_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            objStep0012ProjectCompanyManhourPath: Path = (
                build_step0012_project_company_manhour_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            objStep0012ProjectCompanyGroupManhourPath: Path = (
                build_step0012_project_company_group_manhour_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            make_step0012_project_manhour_tsv(
                str(objStep0011ProjectManhourPath),
                str(objStep0011ProjectCompanyManhourPath),
                str(objStep0012ProjectManhourPath),
                str(objStep0012ProjectCompanyManhourPath),
                str(objStep0012ProjectCompanyGroupManhourPath),
                str(objOrgTableTsvPath),
            )
            objStep0013ProjectManhourPath: Path = build_step0013_project_manhour_output_path(
                objBaseDirectoryPath,
                iYear,
                iMonth,
            )
            objStep0013ProjectCompanyManhourPath: Path = (
                build_step0013_project_company_manhour_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            objStep0013ProjectCompanyGroupManhourPath: Path = (
                build_step0013_project_company_group_manhour_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            make_step0013_project_manhour_tsv(
                str(objStep0012ProjectManhourPath),
                str(objStep0012ProjectCompanyManhourPath),
                str(objStep0012ProjectCompanyGroupManhourPath),
                str(objStep0013ProjectManhourPath),
                str(objStep0013ProjectCompanyManhourPath),
                str(objStep0013ProjectCompanyGroupManhourPath),
            )
            objStep14ProjectCompanyManhourPath: Path = (
                build_step14_project_company_manhour_output_path(
                    objBaseDirectoryPath,
                    iYear,
                    iMonth,
                )
            )
            objStep14OrgTableTsvPath: Path = objScriptDirectoryPath / "管轄PJ表.tsv"
            make_step14_project_company_manhour_tsv(
                str(objStep0013ProjectManhourPath),
                str(objStep14ProjectCompanyManhourPath),
                str(objStep14OrgTableTsvPath),
            )
            if os.path.isfile(objStep14ProjectCompanyManhourPath):
                objStep0014ProjectManhourPath: Path = (
                    objBaseDirectoryPath
                    / f"工数_{iYear}年{iMonth:02d}月_step0014_各プロジェクトの工数.tsv"
                )
                objStep0014ProjectCompanyManhourPath: Path = (
                    objBaseDirectoryPath
                    / f"工数_{iYear}年{iMonth:02d}月_step0014_各プロジェクトの計上カンパニー名_工数.tsv"
                )
                objStep0014ProjectCompanyGroupManhourPath: Path = (
                    objBaseDirectoryPath
                    / f"工数_{iYear}年{iMonth:02d}月_step0014_各プロジェクトの計上カンパニー名_計上グループ_工数.tsv"
                )
                shutil.copyfile(
                    objStep0013ProjectManhourPath,
                    objStep0014ProjectManhourPath,
                )
                shutil.copyfile(
                    objStep0013ProjectCompanyManhourPath,
                    objStep0014ProjectCompanyManhourPath,
                )
                shutil.copyfile(
                    objStep0013ProjectCompanyGroupManhourPath,
                    objStep0014ProjectCompanyGroupManhourPath,
                )
            make_step0006_unique_missing_project_tsv(
                str(objStep0006MissingPath),
                str(objStep0006UniqueMissingPath),
            )
            make_step0006_sort_asc_missing_project_tsv(
                str(objStep0006UniqueMissingPath),
                str(objStep0006SortAscMissingPath),
            )

    return iExitCode


if __name__ == "__main__":
    raise SystemExit(main())
