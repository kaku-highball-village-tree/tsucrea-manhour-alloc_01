from __future__ import annotations

import argparse
import csv
import re
from datetime import timedelta
from pathlib import Path
from typing import List


INVALID_FILE_CHARS_PATTERN: re.Pattern[str] = re.compile(r'[\\/:*?"<>|]')
YEAR_MONTH_PATTERN: re.Pattern[str] = re.compile(r"(\d{2})\.(\d{1,2})月")
DURATION_TEXT_PATTERN: re.Pattern[str] = re.compile(r"^\s*(\d+)\s+day(?:s)?,\s*(\d+):(\d{2}):(\d{2})\s*$")
TIME_TEXT_PATTERN: re.Pattern[str] = re.compile(r"^\d+:\d{2}:\d{2}$")
SALARY_STEP0001_FILE_PATTERN: re.Pattern[str] = re.compile(
    r"^給与配賦アルバイト_step0001_(\d{4}年\d{2}月)\.tsv$"
)
STAFF_MANHOUR_STEP0001_FILE_PATTERN: re.Pattern[str] = re.compile(
    r"^スタッフ別工数_step0001_(\d{4}年\d{2}月)\.tsv$"
)


def build_candidate_paths(pszInputPath: str) -> List[Path]:
    objInputPath: Path = Path(pszInputPath)
    objScriptDirectoryPath: Path = Path(__file__).resolve().parent
    objInputDirectoryPath: Path = Path.cwd() / "input"
    return [
        objInputPath,
        objScriptDirectoryPath / pszInputPath,
        objInputDirectoryPath / pszInputPath,
    ]


def resolve_existing_input_path(pszInputPath: str) -> Path:
    for objCandidatePath in build_candidate_paths(pszInputPath):
        if objCandidatePath.exists():
            return objCandidatePath
    raise FileNotFoundError(f"Input file not found: {pszInputPath}")


def sanitize_sheet_name_for_file_name(pszSheetName: str) -> str:
    pszSanitized: str = INVALID_FILE_CHARS_PATTERN.sub("_", pszSheetName).strip()
    if pszSanitized == "":
        return "Sheet"
    return pszSanitized


def build_unique_output_path(
    objBaseDirectoryPath: Path,
    pszExcelStem: str,
    pszSanitizedSheetName: str,
    objUsedPaths: set[Path],
) -> Path:
    objOutputPath: Path = objBaseDirectoryPath / f"{pszExcelStem}_{pszSanitizedSheetName}.tsv"
    objUsedPaths.add(objOutputPath)
    return objOutputPath


def format_timedelta_as_h_mm_ss(objDuration: timedelta) -> str:
    iTotalSeconds: int = int(objDuration.total_seconds())
    iSign: int = -1 if iTotalSeconds < 0 else 1
    iAbsTotalSeconds: int = abs(iTotalSeconds)
    iHours: int = iAbsTotalSeconds // 3600
    iMinutes: int = (iAbsTotalSeconds % 3600) // 60
    iSeconds: int = iAbsTotalSeconds % 60
    pszPrefix: str = "-" if iSign < 0 else ""
    return f"{pszPrefix}{iHours}:{iMinutes:02d}:{iSeconds:02d}"


def normalize_duration_text_if_needed(pszText: str) -> str:
    objMatch = DURATION_TEXT_PATTERN.match(pszText)
    if objMatch is None:
        return pszText
    iDays: int = int(objMatch.group(1))
    iHours: int = int(objMatch.group(2))
    iMinutes: int = int(objMatch.group(3))
    iSeconds: int = int(objMatch.group(4))
    iTotalHours: int = iDays * 24 + iHours
    return f"{iTotalHours}:{iMinutes:02d}:{iSeconds:02d}"


def normalize_cell_value(objValue: object) -> str:
    if objValue is None:
        return ""
    if isinstance(objValue, timedelta):
        return format_timedelta_as_h_mm_ss(objValue)
    pszText: str = str(objValue)
    pszText = normalize_duration_text_if_needed(pszText)
    return pszText.replace("\t", "_")


def write_sheet_to_tsv(objOutputPath: Path, objRows: List[List[object]]) -> None:
    with open(objOutputPath, mode="w", encoding="utf-8", newline="") as objFile:
        objWriter: csv.writer = csv.writer(objFile, delimiter="\t", lineterminator="\n")
        for objRow in objRows:
            objWriter.writerow([normalize_cell_value(objValue) for objValue in objRow])


def read_tsv_rows(objInputPath: Path) -> List[List[str]]:
    objRows: List[List[str]] = []
    with open(objInputPath, mode="r", encoding="utf-8-sig", newline="") as objFile:
        objReader = csv.reader(objFile, delimiter="\t")
        for objRow in objReader:
            objRows.append(list(objRow))
    return objRows


def is_blank_text(pszText: str) -> bool:
    return (pszText or "").strip().replace("\u3000", "") == ""


def extract_year_month_text_from_path(objInputPath: Path) -> str:
    objMatch = YEAR_MONTH_PATTERN.search(str(objInputPath))
    if objMatch is None:
        raise ValueError(f"Could not extract YY.MM月 from input path: {objInputPath}")
    iYear: int = 2000 + int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    return f"{iYear}年{iMonth:02d}月"


def extract_year_month_text_from_step0001_file_name(pszFileName: str) -> str | None:
    objSalaryMatch = SALARY_STEP0001_FILE_PATTERN.match(pszFileName)
    if objSalaryMatch is not None:
        return objSalaryMatch.group(1)

    objStaffManhourMatch = STAFF_MANHOUR_STEP0001_FILE_PATTERN.match(pszFileName)
    if objStaffManhourMatch is not None:
        return objStaffManhourMatch.group(1)

    return None


def get_effective_column_count(objRow: List[str]) -> int:
    for iIndex in range(len(objRow) - 1, -1, -1):
        if not is_blank_text(objRow[iIndex]):
            return iIndex + 1
    return 0


def is_jobcan_long_format_tsv(objRows: List[List[str]]) -> bool:
    objNonEmptyRows: List[List[str]] = [
        objRow for objRow in objRows if any(not is_blank_text(pszCell) for pszCell in objRow)
    ]
    if not objNonEmptyRows:
        return False

    iTotal: int = len(objNonEmptyRows)
    iFourColumnsLike: int = 0
    iTimeTextRows: int = 0
    iProjectCodeRows: int = 0
    for objRow in objNonEmptyRows:
        iEffectiveColumns: int = get_effective_column_count(objRow)
        if 3 <= iEffectiveColumns <= 5:
            iFourColumnsLike += 1

        if len(objRow) >= 4:
            pszTimeText: str = (objRow[3] or "").strip()
            if TIME_TEXT_PATTERN.match(pszTimeText) is not None or DURATION_TEXT_PATTERN.match(pszTimeText) is not None:
                iTimeTextRows += 1

        if len(objRow) >= 2:
            pszProjectText: str = (objRow[1] or "").strip()
            if re.match(r"^(P\d{5}|[A-OQ-Z]\d{3})", pszProjectText) is not None:
                iProjectCodeRows += 1

    return (
        iFourColumnsLike / iTotal >= 0.7
        and iTimeTextRows / iTotal >= 0.5
        and iProjectCodeRows / iTotal >= 0.5
    )


def normalize_project_name_for_jobcan_long_tsv(pszProjectName: str) -> str:
    pszNormalized: str = pszProjectName or ""
    pszNormalized = pszNormalized.replace("\t", "_")
    pszNormalized = re.sub(r"(P\d{5})(?![ _\t　【])", r"\1_", pszNormalized)
    pszNormalized = re.sub(r"([A-OQ-Z]\d{3})(?![ _\t　【])", r"\1_", pszNormalized)
    pszNormalized = re.sub(r"^([A-OQ-Z]\d{3}) +", r"\1_", pszNormalized)
    pszNormalized = re.sub(r"([A-OQ-Z]\d{3})[ 　]+", r"\1_", pszNormalized)
    pszNormalized = re.sub(r"(P\d{5})[ 　]+", r"\1_", pszNormalized)
    return pszNormalized


def process_jobcan_long_tsv_input(objResolvedInputPath: Path, objRows: List[List[str]]) -> int:
    pszYearMonthText: str = extract_year_month_text_from_path(objResolvedInputPath)

    objOutputRows: List[List[str]] = []
    pszCurrentStaffName: str = ""
    pszLastOutputStaffName: str = ""
    for objRow in objRows:
        if not any(not is_blank_text(pszCell) for pszCell in objRow):
            continue
        if len(objRow) < 4:
            continue

        pszStaffName: str = (objRow[0] or "").strip()
        if pszStaffName != "":
            pszCurrentStaffName = pszStaffName
        if pszCurrentStaffName == "":
            continue

        pszProjectName: str = normalize_project_name_for_jobcan_long_tsv((objRow[1] or "").strip())
        pszManhour: str = (objRow[3] or "").strip()
        if pszProjectName == "" and pszManhour == "":
            continue

        pszOutputStaffName: str = pszCurrentStaffName
        if pszCurrentStaffName == pszLastOutputStaffName:
            pszOutputStaffName = ""
        else:
            pszLastOutputStaffName = pszCurrentStaffName

        objOutputRows.append([pszOutputStaffName, pszProjectName, pszManhour])

    if not objOutputRows:
        raise ValueError("No output rows generated for Jobcan long-format TSV")

    objOutputPath: Path = (
        objResolvedInputPath.resolve().parent
        / f"スタッフ別工数_step0001_{pszYearMonthText}.tsv"
    )
    write_sheet_to_tsv(objOutputPath, objOutputRows)
    return 0




def is_salary_allocation_parttime_tsv(objRows: List[List[str]]) -> bool:
    if len(objRows) < 2:
        return False

    if is_jobcan_long_format_tsv(objRows):
        return False

    objHeaderRowCandidates: List[List[str]] = [objRows[0]]
    if len(objRows) >= 2:
        objHeaderRowCandidates.append(objRows[1])
    iHeaderNonBlankCount: int = 0
    for objHeaderRow in objHeaderRowCandidates:
        iHeaderNonBlankCount = max(
            iHeaderNonBlankCount,
            sum(1 for pszCell in objHeaderRow if (pszCell or "").strip() != ""),
        )
    if iHeaderNonBlankCount < 3:
        return False

    objTotalRow: List[str] | None = None
    for objRow in objRows:
        if objRow and (objRow[0] or "").strip() == "合計":
            objTotalRow = objRow
            break
    if objTotalRow is None:
        return False

    iNumericCount: int = 0
    for pszCell in objTotalRow[1:]:
        pszValue: str = (pszCell or "").strip()
        if re.match(r"^-?\d+(?:\.\d+)?$", pszValue) is not None:
            iNumericCount += 1
    return iNumericCount >= 3


def build_error_copy_output_path(objInputPath: Path) -> Path:
    return objInputPath.resolve().with_name(f"{objInputPath.stem}_error{objInputPath.suffix}")


def is_specific_fallback_case_for_staff_step0001_error(objInputPath: Path) -> bool:
    if objInputPath.name != "作成用データ：工数25.12月_Sheet1.tsv":
        return False
    objSalaryPath: Path = objInputPath.resolve().parent / "給与配賦アルバイト_step0001_2025年12月.tsv"
    return objSalaryPath.exists()


def write_specific_staff_step0001_error_file(objInputPath: Path) -> int:
    pszYearMonthText: str = extract_year_month_text_from_path(objInputPath)
    objOutputPath: Path = (
        objInputPath.resolve().parent
        / f"スタッフ別工数_step0001_{pszYearMonthText}_error.tsv"
    )
    objOutputRows: List[List[str]] = [
        ["エラー種別", "例外暫定未定義処理"],
        ["判定", "is_salary_allocation_parttime_tsv == False"],
        ["対象入力", str(objInputPath)],
        ["説明", "給与配賦TSV本処理の対象形式として判定できませんでした。"],
    ]
    write_sheet_to_tsv(objOutputPath, objOutputRows)
    return 0


def process_tsv_input(objResolvedInputPath: Path) -> int:
    objRows: List[List[str]] = read_tsv_rows(objResolvedInputPath)
    if len(objRows) < 2:
        raise ValueError(f"Input TSV has too few rows: {objResolvedInputPath}")

    if is_jobcan_long_format_tsv(objRows):
        return process_jobcan_long_tsv_input(objResolvedInputPath, objRows)

    if not is_salary_allocation_parttime_tsv(objRows):
        if is_specific_fallback_case_for_staff_step0001_error(objResolvedInputPath):
            return write_specific_staff_step0001_error_file(objResolvedInputPath)

        objErrorOutputPath: Path = build_error_copy_output_path(objResolvedInputPath)
        write_sheet_to_tsv(objErrorOutputPath, objRows)
        return 0

    objRowsWithoutA: List[List[str]] = [objRow[1:] if len(objRow) >= 1 else [] for objRow in objRows]

    if len(objRowsWithoutA[0]) < 1:
        raise ValueError("B1 text could not be found after removing column A")
    pszTitle: str = (objRowsWithoutA[0][0] or "").strip()
    if pszTitle == "":
        raise ValueError("B1 text is empty")

    pszYearMonthText: str = extract_year_month_text_from_path(objResolvedInputPath)

    objOutputRows: List[List[str]] = []
    if len(objRowsWithoutA) >= 2:
        objOutputRows.append(objRowsWithoutA[1])
    objOutputRows.extend(objRowsWithoutA[3:])

    objFilteredOutputRows: List[List[str]] = []
    for objRow in objOutputRows:
        if any(not is_blank_text(pszCell) for pszCell in objRow):
            objFilteredOutputRows.append(objRow)

    if not objFilteredOutputRows:
        raise ValueError("No output rows after applying TSV row rules")

    objOutputPath: Path = (
        objResolvedInputPath.resolve().parent
        / f"{pszTitle}_step0001_{pszYearMonthText}.tsv"
    )
    write_sheet_to_tsv(objOutputPath, objFilteredOutputRows)
    return 0


def process_staff_manhour_step0002_from_step0001_pair(
    objSalaryStep0001Path: Path,
    objStaffManhourStep0001Path: Path,
) -> int:
    pszSalaryYearMonthText: str | None = extract_year_month_text_from_step0001_file_name(
        objSalaryStep0001Path.name
    )
    pszStaffManhourYearMonthText: str | None = extract_year_month_text_from_step0001_file_name(
        objStaffManhourStep0001Path.name
    )
    if pszSalaryYearMonthText is None or pszStaffManhourYearMonthText is None:
        raise ValueError("Could not extract year-month from step0001 file names")
    if pszSalaryYearMonthText != pszStaffManhourYearMonthText:
        raise ValueError("Year-month mismatch between salary and staff-manhour step0001 files")

    objSalaryRows: List[List[str]] = read_tsv_rows(objSalaryStep0001Path)
    if not objSalaryRows:
        raise ValueError(f"Input TSV has no rows: {objSalaryStep0001Path}")

    objAllowedStaffNames: set[str] = {
        (pszCell or "").strip()
        for pszCell in objSalaryRows[0]
        if (pszCell or "").strip() != ""
    }
    if not objAllowedStaffNames:
        raise ValueError(f"No staff names found in first row: {objSalaryStep0001Path}")

    objStaffManhourRows: List[List[str]] = read_tsv_rows(objStaffManhourStep0001Path)
    if not objStaffManhourRows:
        raise ValueError(f"Input TSV has no rows: {objStaffManhourStep0001Path}")

    objOutputRows: List[List[str]] = []
    pszCurrentStaffName: str = ""
    for objRow in objStaffManhourRows:
        if not objRow:
            continue

        pszStaffName: str = (objRow[0] if objRow else "").strip()
        if pszStaffName != "":
            pszCurrentStaffName = pszStaffName
        if pszCurrentStaffName == "":
            continue

        if pszCurrentStaffName in objAllowedStaffNames:
            objOutputRows.append(objRow)

    if not objOutputRows:
        raise ValueError("No rows remained after filtering by salary step0001 first-row staff names")

    objOutputPath: Path = (
        objStaffManhourStep0001Path.resolve().parent
        / f"スタッフ別工数_step0002_{pszStaffManhourYearMonthText}.tsv"
    )
    write_sheet_to_tsv(objOutputPath, objOutputRows)

    process_staff_manhour_step0003_from_step0002(
        objSalaryStep0001Path,
        objOutputPath,
        pszStaffManhourYearMonthText,
    )
    return 0






def find_first_non_blank_index(objRow: List[str]) -> int | None:
    for iIndex, pszValue in enumerate(objRow):
        if (pszValue or "").strip() != "":
            return iIndex
    return None


def find_first_numeric_like_index(objRow: List[str], iStartIndex: int = 0) -> int | None:
    for iIndex in range(max(iStartIndex, 0), len(objRow)):
        pszValue: str = (objRow[iIndex] or "").strip()
        if re.match(r"^-?\d+(?:\.\d+)?$", pszValue) is not None:
            return iIndex
    return None

def build_salary_total_value_by_staff_from_step0001(objSalaryStep0001Path: Path) -> dict[str, str]:
    objSalaryRows: List[List[str]] = read_tsv_rows(objSalaryStep0001Path)
    if not objSalaryRows:
        raise ValueError(f"Input TSV has no rows: {objSalaryStep0001Path}")

    objHeaderRow: List[str] = objSalaryRows[0]
    iTotalRowIndex: int | None = None
    for iRowIndex, objRow in enumerate(objSalaryRows):
        if not objRow:
            continue
        if (objRow[0] or "").strip() == "合計":
            iTotalRowIndex = iRowIndex
            break
    if iTotalRowIndex is None:
        raise ValueError(f"No total row found in salary step0001 TSV: {objSalaryStep0001Path}")

    objTotalRow: List[str] = objSalaryRows[iTotalRowIndex]
    bTotalRowStartsWithLabel: bool = len(objTotalRow) >= 1 and (objTotalRow[0] or "").strip() == "合計"

    iFirstStaffColumnIndex: int | None = find_first_non_blank_index(objHeaderRow)
    iFirstValueColumnIndex: int | None = find_first_numeric_like_index(
        objTotalRow,
        iStartIndex=1 if bTotalRowStartsWithLabel else 0,
    )
    iColumnOffset: int = 0
    if iFirstStaffColumnIndex is not None and iFirstValueColumnIndex is not None:
        iColumnOffset = iFirstValueColumnIndex - iFirstStaffColumnIndex

    objTotalValueByStaff: dict[str, str] = {}
    for iColumnIndex, pszStaffNameRaw in enumerate(objHeaderRow):
        pszStaffName: str = (pszStaffNameRaw or "").strip()
        if pszStaffName == "":
            continue
        iValueColumnIndex: int = iColumnIndex + iColumnOffset
        pszValue: str = objTotalRow[iValueColumnIndex] if 0 <= iValueColumnIndex < len(objTotalRow) else ""
        objTotalValueByStaff[pszStaffName] = pszValue

    return objTotalValueByStaff


def process_staff_manhour_step0003_from_step0002(
    objSalaryStep0001Path: Path,
    objStaffManhourStep0002Path: Path,
    pszYearMonthText: str,
) -> int:
    objTotalValueByStaff: dict[str, str] = build_salary_total_value_by_staff_from_step0001(
        objSalaryStep0001Path
    )

    objStep0002Rows: List[List[str]] = read_tsv_rows(objStaffManhourStep0002Path)
    if not objStep0002Rows:
        raise ValueError(f"Input TSV has no rows: {objStaffManhourStep0002Path}")

    objOutputRows: List[List[str]] = []
    pszCurrentStaffName: str = ""
    pszPreviousStaffName: str = ""
    for objRow in objStep0002Rows:
        objNewRow: List[str] = list(objRow)
        pszStaffNameCell: str = (objNewRow[0] if objNewRow else "").strip()
        if pszStaffNameCell != "":
            pszCurrentStaffName = pszStaffNameCell
        if pszCurrentStaffName == "":
            objOutputRows.append(objNewRow)
            continue

        if pszCurrentStaffName != pszPreviousStaffName:
            while len(objNewRow) < 4:
                objNewRow.append("")
            objNewRow[3] = objTotalValueByStaff.get(pszCurrentStaffName, "")
            pszPreviousStaffName = pszCurrentStaffName

        objOutputRows.append(objNewRow)

    objOutputPath: Path = (
        objStaffManhourStep0002Path.resolve().parent
        / f"スタッフ別工数_step0003_{pszYearMonthText}.tsv"
    )
    write_sheet_to_tsv(objOutputPath, objOutputRows)
    objStep0004Path: Path = process_staff_manhour_step0004_from_step0003(objOutputPath, pszYearMonthText)
    process_staff_manhour_step0005_from_step0004_and_salary_step0001(
        objStep0004Path,
        objSalaryStep0001Path,
        pszYearMonthText,
    )
    return 0


def parse_h_mm_ss_to_seconds(pszManhourText: str) -> int:
    pszText: str = (pszManhourText or "").strip()
    objMatch = re.match(r"^(\d+):(\d{2}):(\d{2})$", pszText)
    if objMatch is None:
        raise ValueError(f"Invalid manhour text format (expected H:MM:SS): {pszManhourText}")

    iHours: int = int(objMatch.group(1))
    iMinutes: int = int(objMatch.group(2))
    iSeconds: int = int(objMatch.group(3))
    if iMinutes >= 60 or iSeconds >= 60:
        raise ValueError(f"Invalid manhour text value (minutes/seconds out of range): {pszManhourText}")
    return iHours * 3600 + iMinutes * 60 + iSeconds


def parse_integer_text(pszValueText: str) -> int:
    pszText: str = (pszValueText or "").strip()
    if re.match(r"^-?\d+$", pszText) is None:
        raise ValueError(f"Invalid integer text: {pszValueText}")
    return int(pszText)


def allocate_integer_values_by_ratio(iTotalValue: int, objDurationsInSeconds: List[int]) -> List[int]:
    if iTotalValue < 0:
        raise ValueError(f"Total value must be non-negative: {iTotalValue}")
    if not objDurationsInSeconds:
        raise ValueError("No duration rows found for allocation")

    iTotalSeconds: int = sum(objDurationsInSeconds)
    if iTotalSeconds <= 0:
        raise ValueError("Total duration must be greater than zero for allocation")

    objBaseValues: List[int] = []
    objRemainders: List[tuple[int, int]] = []
    iAllocatedBaseSum: int = 0
    for iIndex, iSeconds in enumerate(objDurationsInSeconds):
        if iSeconds < 0:
            raise ValueError("Duration seconds must be non-negative")
        iNumerator: int = iTotalValue * iSeconds
        iBaseValue: int = iNumerator // iTotalSeconds
        iRemainder: int = iNumerator % iTotalSeconds
        objBaseValues.append(iBaseValue)
        objRemainders.append((iRemainder, iIndex))
        iAllocatedBaseSum += iBaseValue

    iDifference: int = iTotalValue - iAllocatedBaseSum
    if iDifference < 0:
        raise ValueError("Allocated base sum exceeded total value")

    objRemainders.sort(key=lambda objItem: (-objItem[0], objItem[1]))
    for iOffset in range(iDifference):
        _, iTargetIndex = objRemainders[iOffset]
        objBaseValues[iTargetIndex] += 1

    if sum(objBaseValues) != iTotalValue:
        raise ValueError("Allocated integer sum does not match total value")
    return objBaseValues


def process_staff_manhour_step0004_from_step0003(
    objStaffManhourStep0003Path: Path,
    pszYearMonthText: str,
) -> Path:
    objStep0003Rows: List[List[str]] = read_tsv_rows(objStaffManhourStep0003Path)
    if not objStep0003Rows:
        raise ValueError(f"Input TSV has no rows: {objStaffManhourStep0003Path}")

    objOutputRows: List[List[str]] = [list(objRow) for objRow in objStep0003Rows]

    objCurrentStaffRows: List[int] = []
    objCurrentStaffDurations: List[int] = []
    pszCurrentStaffName: str = ""
    iCurrentStaffTotalValue: int | None = None

    def flush_current_staff_rows() -> None:
        if not objCurrentStaffRows:
            return
        if pszCurrentStaffName == "":
            raise ValueError("Staff name could not be resolved for step0003 block")
        if iCurrentStaffTotalValue is None:
            raise ValueError(f"Total value column is missing for staff: {pszCurrentStaffName}")
        objAllocatedValues: List[int] = allocate_integer_values_by_ratio(
            iCurrentStaffTotalValue,
            objCurrentStaffDurations,
        )
        for iRowIndex, iAllocatedValue in zip(objCurrentStaffRows, objAllocatedValues):
            while len(objOutputRows[iRowIndex]) < 5:
                objOutputRows[iRowIndex].append("")
            objOutputRows[iRowIndex][4] = str(iAllocatedValue)

    for iRowIndex, objRow in enumerate(objStep0003Rows):
        objNewRow: List[str] = list(objRow)
        pszStaffNameCell: str = (objNewRow[0] if len(objNewRow) >= 1 else "").strip()
        pszManhourText: str = (objNewRow[2] if len(objNewRow) >= 3 else "").strip()

        if pszStaffNameCell != "":
            flush_current_staff_rows()
            objCurrentStaffRows = []
            objCurrentStaffDurations = []
            pszCurrentStaffName = pszStaffNameCell
            pszTotalValueText: str = (objNewRow[3] if len(objNewRow) >= 4 else "").strip()
            if pszTotalValueText == "":
                raise ValueError(f"Total value is blank for staff: {pszCurrentStaffName}")
            iCurrentStaffTotalValue = parse_integer_text(pszTotalValueText)
        elif pszCurrentStaffName == "":
            raise ValueError(f"Staff name could not be resolved at row index: {iRowIndex}")

        iDurationInSeconds: int = parse_h_mm_ss_to_seconds(pszManhourText)
        objCurrentStaffRows.append(iRowIndex)
        objCurrentStaffDurations.append(iDurationInSeconds)

    flush_current_staff_rows()

    objOutputPath: Path = (
        objStaffManhourStep0003Path.resolve().parent
        / f"スタッフ別工数_step0004_{pszYearMonthText}.tsv"
    )
    write_sheet_to_tsv(objOutputPath, objOutputRows)
    return objOutputPath


def collect_non_blank_staff_names_from_step0004_rows(objRows: List[List[str]]) -> set[str]:
    objStaffNames: set[str] = set()
    for objRow in objRows:
        pszStaffName: str = (objRow[0] if len(objRow) >= 1 else "").strip()
        if pszStaffName != "":
            objStaffNames.add(pszStaffName)
    return objStaffNames


def process_staff_manhour_step0005_from_step0004_and_salary_step0001(
    objStaffManhourStep0004Path: Path,
    objSalaryStep0001Path: Path,
    pszYearMonthText: str,
) -> Path:
    objStep0004Rows: List[List[str]] = read_tsv_rows(objStaffManhourStep0004Path)
    if not objStep0004Rows:
        raise ValueError(f"Input TSV has no rows: {objStaffManhourStep0004Path}")

    objOutputRows: List[List[str]] = [list(objRow) for objRow in objStep0004Rows]
    objStaffNamesInStep0004: set[str] = collect_non_blank_staff_names_from_step0004_rows(objStep0004Rows)

    objSalaryRows: List[List[str]] = read_tsv_rows(objSalaryStep0001Path)
    if not objSalaryRows:
        raise ValueError(f"Input TSV has no rows: {objSalaryStep0001Path}")

    objHeaderRow: List[str] = objSalaryRows[0]
    objTotalValueByStaff: dict[str, str] = build_salary_total_value_by_staff_from_step0001(objSalaryStep0001Path)

    iOutputColumnCount: int = 0
    for objRow in objOutputRows:
        if len(objRow) > iOutputColumnCount:
            iOutputColumnCount = len(objRow)
    iOutputColumnCount = max(iOutputColumnCount, 5)

    for pszStaffNameRaw in objHeaderRow:
        pszStaffName: str = (pszStaffNameRaw or "").strip()
        if pszStaffName == "":
            continue
        if pszStaffName in objStaffNamesInStep0004:
            continue

        objAppendRow: List[str] = [""] * iOutputColumnCount
        objAppendRow[0] = pszStaffName
        objAppendRow[3] = objTotalValueByStaff.get(pszStaffName, "")
        objOutputRows.append(objAppendRow)

    objOutputPath: Path = (
        objStaffManhourStep0004Path.resolve().parent
        / f"スタッフ別工数_step0005_{pszYearMonthText}.tsv"
    )
    write_sheet_to_tsv(objOutputPath, objOutputRows)
    return objOutputPath

def process_single_input(pszInputXlsxPath: str) -> int:
    objResolvedInputPath: Path = resolve_existing_input_path(pszInputXlsxPath)
    pszSuffix: str = objResolvedInputPath.suffix.lower()

    if pszSuffix == ".tsv":
        return process_tsv_input(objResolvedInputPath)

    if pszSuffix != ".xlsx":
        raise ValueError(f"Unsupported extension (only .xlsx/.tsv): {objResolvedInputPath}")

    objBaseDirectoryPath: Path = objResolvedInputPath.resolve().parent
    pszExcelStem: str = objResolvedInputPath.stem

    try:
        import openpyxl
    except Exception as objException:
        raise RuntimeError(f"Failed to import openpyxl: {objException}") from objException

    try:
        objWorkbook = openpyxl.load_workbook(
            filename=objResolvedInputPath,
            read_only=True,
            data_only=True,
        )
    except Exception as objException:
        raise RuntimeError(f"Failed to read workbook: {objResolvedInputPath}. Detail = {objException}") from objException

    objUsedPaths: set[Path] = set()
    try:
        for objWorksheet in objWorkbook.worksheets:
            pszSanitizedSheetName: str = sanitize_sheet_name_for_file_name(objWorksheet.title)
            objOutputPath: Path = build_unique_output_path(
                objBaseDirectoryPath,
                pszExcelStem,
                pszSanitizedSheetName,
                objUsedPaths,
            )
            objRows: List[List[object]] = [list(objRow) for objRow in objWorksheet.iter_rows(values_only=True)]
            write_sheet_to_tsv(objOutputPath, objRows)
    finally:
        objWorkbook.close()

    return 0


def main() -> int:
    objParser: argparse.ArgumentParser = argparse.ArgumentParser()
    objParser.add_argument(
        "pszInputXlsxPaths",
        nargs="+",
        help="Input Excel (.xlsx) file paths",
    )
    objArgs: argparse.Namespace = objParser.parse_args()

    iExitCode: int = 0
    objHandledInputPaths: set[Path] = set()

    objSalaryStep0001ByYearMonth: dict[str, Path] = {}
    objStaffManhourStep0001ByYearMonth: dict[str, Path] = {}
    for pszInputXlsxPath in objArgs.pszInputXlsxPaths:
        try:
            objResolvedInputPath: Path = resolve_existing_input_path(pszInputXlsxPath)
        except Exception:
            continue

        pszYearMonthText: str | None = extract_year_month_text_from_step0001_file_name(
            objResolvedInputPath.name
        )
        if pszYearMonthText is None:
            continue

        objSalaryMatch = SALARY_STEP0001_FILE_PATTERN.match(objResolvedInputPath.name)
        if objSalaryMatch is not None:
            objSalaryStep0001ByYearMonth[pszYearMonthText] = objResolvedInputPath

        objStaffManhourMatch = STAFF_MANHOUR_STEP0001_FILE_PATTERN.match(objResolvedInputPath.name)
        if objStaffManhourMatch is not None:
            objStaffManhourStep0001ByYearMonth[pszYearMonthText] = objResolvedInputPath

    for pszYearMonthText, objSalaryStep0001Path in objSalaryStep0001ByYearMonth.items():
        objStaffManhourStep0001Path: Path | None = objStaffManhourStep0001ByYearMonth.get(
            pszYearMonthText
        )
        if objStaffManhourStep0001Path is None:
            continue
        try:
            process_staff_manhour_step0002_from_step0001_pair(
                objSalaryStep0001Path,
                objStaffManhourStep0001Path,
            )
            objHandledInputPaths.add(objSalaryStep0001Path.resolve())
            objHandledInputPaths.add(objStaffManhourStep0001Path.resolve())
        except Exception as objException:
            print(
                "Error: failed to process step0002 pair: {0} / {1}. Detail = {2}".format(
                    objSalaryStep0001Path,
                    objStaffManhourStep0001Path,
                    objException,
                )
            )
            iExitCode = 1

    for pszInputXlsxPath in objArgs.pszInputXlsxPaths:
        try:
            objResolvedInputPath = resolve_existing_input_path(pszInputXlsxPath)
            if objResolvedInputPath.resolve() in objHandledInputPaths:
                continue
            process_single_input(pszInputXlsxPath)
        except Exception as objException:
            print(
                "Error: failed to process input file: {0}. Detail = {1}".format(
                    pszInputXlsxPath,
                    objException,
                )
            )
            iExitCode = 1
            continue

    return iExitCode


if __name__ == "__main__":
    raise SystemExit(main())
