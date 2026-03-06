import csv
import os
import re
import shutil
import sys
from typing import Dict, List, Optional, Tuple


def get_target_year_month_from_filename(pszInputFilePath: str) -> Tuple[int, int]:
    pszBaseName: str = os.path.basename(pszInputFilePath)
    objMatch: re.Match[str] | None = re.search(r"(\d{2})\.(\d{1,2})\.csv$", pszBaseName)
    if objMatch is None:
        raise ValueError("入力ファイル名から対象年月を取得できません。")
    iYearTwoDigits: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    iYear: int = 2000 + iYearTwoDigits
    return iYear, iMonth


def get_target_year_month_from_period_row(pszRowA: str) -> Tuple[int, int]:
    pszNormalized: str = re.sub(r"[ \u3000]", "", pszRowA)
    pszNormalized = pszNormalized.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    objMatch: re.Match[str] | None = re.search(r"(?:自)?(\d{4})年(\d{1,2})月(?:度)?", pszNormalized)
    if objMatch is None:
        objMatch = re.search(r"(\d{4})[./-](\d{1,2})", pszNormalized)
    if objMatch is None:
        raise ValueError("集計期間から対象年月を取得できません。")
    iYear: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    return iYear, iMonth


def read_csv_rows(pszInputFilePath: str) -> List[List[str]]:
    objRows: List[List[str]] = []
    try:
        with open(pszInputFilePath, mode="r", encoding="utf-8-sig", errors="strict", newline="") as objFile:
            objReader: csv.reader = csv.reader(objFile)
            for objRow in objReader:
                objRows.append(objRow)
        append_debug_log("input decoded as utf-8-sig")
        return objRows
    except UnicodeDecodeError:
        append_debug_log("utf-8-sig decode failed; retrying with cp932")

    with open(pszInputFilePath, mode="r", encoding="cp932", errors="strict", newline="") as objFile:
        objReader = csv.reader(objFile)
        for objRow in objReader:
            objRows.append(objRow)
    append_debug_log("input decoded as cp932")
    return objRows


def write_tsv_rows(pszOutputFilePath: str, objRows: List[List[str]]) -> None:
    with open(pszOutputFilePath, mode="w", encoding="utf-8", newline="") as objFile:
        objWriter: csv.writer = csv.writer(objFile, delimiter="\t", lineterminator="\n")
        for objRow in objRows:
            objWriter.writerow(objRow)


def read_tsv_rows(pszInputFilePath: str) -> List[List[str]]:
    objRows: List[List[str]] = []
    with open(pszInputFilePath, mode="r", encoding="utf-8", newline="") as objFile:
        objReader: csv.reader = csv.reader(objFile, delimiter="\t")
        for objRow in objReader:
            objRows.append(objRow)
    return objRows


def build_first_column_rows(objRows: List[List[str]]) -> List[List[str]]:
    return [[objRow[0] if objRow else ""] for objRow in objRows]



def build_unique_subjects(objSubjectRows: List[List[str]]) -> List[str]:
    objSubjects: List[str] = []
    objSeen: set[str] = set()
    for objRow in objSubjectRows:
        pszValue: str = objRow[0] if objRow else ""
        if pszValue == "" or pszValue in objSeen:
            continue
        objSeen.add(pszValue)
        objSubjects.append(pszValue)
    return objSubjects



def build_union_subject_order(objSubjectLists: List[List[str]]) -> List[str]:
    objAppearanceOrder: dict[str, int] = {}
    iCounter: int = 0
    for objSubjectList in objSubjectLists:
        for pszSubject in objSubjectList:
            if pszSubject not in objAppearanceOrder:
                objAppearanceOrder[pszSubject] = iCounter
                iCounter += 1

    objAdjacency: dict[str, set[str]] = {psz: set() for psz in objAppearanceOrder}
    objIndegree: dict[str, int] = {psz: 0 for psz in objAppearanceOrder}
    for objSubjectList in objSubjectLists:
        for iIndex in range(len(objSubjectList) - 1):
            pszBefore: str = objSubjectList[iIndex]
            pszAfter: str = objSubjectList[iIndex + 1]
            if pszAfter not in objAdjacency[pszBefore]:
                objAdjacency[pszBefore].add(pszAfter)
                objIndegree[pszAfter] += 1

    objOrderedSubjects: List[str] = []
    objReady: List[str] = [
        pszSubject for pszSubject, iDegree in objIndegree.items() if iDegree == 0
    ]
    objReady.sort(key=lambda pszSubject: objAppearanceOrder[pszSubject])

    while objReady:
        pszSubject = objReady.pop(0)
        objOrderedSubjects.append(pszSubject)
        for pszNext in sorted(objAdjacency[pszSubject], key=lambda psz: objAppearanceOrder[psz]):
            objIndegree[pszNext] -= 1
            if objIndegree[pszNext] == 0:
                objReady.append(pszNext)
        objReady.sort(key=lambda pszSubject: objAppearanceOrder[pszSubject])

    if len(objOrderedSubjects) != len(objAppearanceOrder):
        return list(objAppearanceOrder.keys())

    return objOrderedSubjects


def build_cumulative_subject_order(objSubjectLists: List[List[str]]) -> List[str]:
    objOrderedSubjects: List[str] = []
    for objSubjectList in objSubjectLists:
        iInsertAfterIndex: int = -1
        for pszSubject in objSubjectList:
            if pszSubject in objOrderedSubjects:
                iInsertAfterIndex = objOrderedSubjects.index(pszSubject)
                continue
            objOrderedSubjects.insert(iInsertAfterIndex + 1, pszSubject)
            iInsertAfterIndex += 1
    return objOrderedSubjects


def sort_vertical_file_paths(objFilePaths: List[str]) -> List[str]:
    def get_sort_key(pszFilePath: str) -> tuple[int, int, str]:
        objMatch = re.search(r"_(\d{4})年(\d{2})月", pszFilePath)
        if objMatch is None:
            return (9999, 99, pszFilePath)
        return (int(objMatch.group(1)), int(objMatch.group(2)), pszFilePath)

    return sorted(objFilePaths, key=get_sort_key)



def build_subject_vertical_rows(objSubjects: List[str]) -> List[List[str]]:
    return [[pszSubject] for pszSubject in objSubjects]


def transpose_rows(objRows: List[List[str]]) -> List[List[str]]:
    if not objRows:
        return []
    iMaxColumns: int = max(len(objRow) for objRow in objRows)
    objPaddedRows: List[List[str]] = [
        objRow + [""] * (iMaxColumns - len(objRow)) for objRow in objRows
    ]
    return [
        [objPaddedRows[iRowIndex][iColumnIndex] for iRowIndex in range(len(objPaddedRows))]
        for iColumnIndex in range(iMaxColumns)
    ]



def normalize_project_name(pszProjectName: str) -> str:
    if pszProjectName == "":
        return pszProjectName
    normalized = pszProjectName.replace("\t", "_")
    normalized = re.sub(r"(P\d{5})(?![ _\t　【])", r"\1_", normalized)
    normalized = re.sub(r"([A-OQ-Z]\d{3})(?![ _\t　【])", r"\1_", normalized)
    normalized = re.sub(r"^(J\d{3}) +", r"\1_", normalized)
    normalized = re.sub(r"([A-OQ-Z]\d{3})[ 　]+", r"\1_", normalized)
    normalized = re.sub(r"(P\d{5})[ 　]+", r"\1_", normalized)

    if normalized.startswith("【廃番】"):
        try:
            pszCode: str | None = None
            iCodeIndex: int = -1
            for pszPrefix in ["J", "A", "C", "H", "M", "P"]:
                iSearchFrom: int = 0
                iCodeLength: int = 4
                if pszPrefix == "P":
                    iCodeLength = 6
                while True:
                    iFoundIndex: int = normalized.find(pszPrefix, iSearchFrom)
                    if iFoundIndex == -1:
                        break
                    if iFoundIndex + iCodeLength <= len(normalized):
                        pszCode = normalized[iFoundIndex:iFoundIndex + iCodeLength]
                        iCodeIndex = iFoundIndex
                        break
                    iSearchFrom = iFoundIndex + 1
                if pszCode is not None:
                    break
            if pszCode is None or iCodeIndex == -1:
                return normalized
            pszHead: str = normalized[:iCodeIndex]
            pszTail: str = normalized[iCodeIndex + len(pszCode):]
            return pszCode + "_" + pszHead + pszTail
        except Exception:
            return normalized

    if normalized.startswith("【"):
        iBracketEndIndex: int = normalized.find("】")
        if iBracketEndIndex != -1:
            pszAfterBracket: str = normalized[iBracketEndIndex + 1:]
            objMatch = re.search(r"(P\d{5}|[A-OQ-Z]\d{3})", pszAfterBracket)
            if objMatch is not None:
                pszCode = objMatch.group(1)
                pszBeforeCode: str = pszAfterBracket[:objMatch.start()]
                pszAfterCode: str = pszAfterBracket[objMatch.end():]
                if pszAfterCode.startswith(" ") or pszAfterCode.startswith("　"):
                    pszAfterCode = pszAfterCode[1:]
                pszRest: str = normalized[: iBracketEndIndex + 1] + pszBeforeCode + pszAfterCode
                return pszCode + "_" + pszRest

    if len(normalized) >= 1 and normalized[0] in ["J", "A", "C", "H", "M"]:
        if len(normalized) >= 5:
            pszNextChar: str = normalized[4]
            if pszNextChar == "【":
                return normalized[:4] + "_" + normalized[4:]
            if pszNextChar == " " or pszNextChar == "　":
                return normalized[:4] + "_" + normalized[5:]
        return normalized

    if len(normalized) >= 1 and normalized[0] == "P":
        if len(normalized) >= 7:
            pszNextCharP: str = normalized[6]
            if pszNextCharP == "【":
                return normalized[:6] + "_" + normalized[6:]
            if pszNextCharP == " " or pszNextCharP == "　":
                return normalized[:6] + "_" + normalized[7:]
        return normalized

    return normalized


def normalize_project_names_in_row(objRows: List[List[str]], iRowIndex: int) -> None:
    if iRowIndex < 0 or iRowIndex >= len(objRows):
        return
    objTargetRow = objRows[iRowIndex]
    for iIndex, pszProjectName in enumerate(objTargetRow):
        objTargetRow[iIndex] = normalize_project_name(pszProjectName)


def is_valid_project_subject_name(pszName: str) -> bool:
    pszText: str = (pszName or "").strip()
    if pszText == "":
        return True

    objAllowedNames: set[str] = {
        "科目名",
        "合計",
        "本部",
        "1Cカンパニー販管費",
        "2Cカンパニー販管費",
        "3Cカンパニー販管費",
        "4Cカンパニー販管費",
        "事業開発カンパニー販管費",
        "社長室カンパニー販管費",
        "本部カンパニー販管費",
        "その他",
    }
    if pszText in objAllowedNames:
        return True

    return re.match(r"^(P\d{5}|[A-OQ-Z]\d{3})[ _\t　].+", pszText) is not None


def collect_invalid_project_subject_cells(
    objRows: List[List[str]],
    objTargetRowIndices: List[int],
) -> List[tuple[int, int, str]]:
    objInvalidCells: List[tuple[int, int, str]] = []
    for iRowIndex in objTargetRowIndices:
        if iRowIndex < 0 or iRowIndex >= len(objRows):
            continue
        objRow = objRows[iRowIndex]
        for iColumnIndex, pszValue in enumerate(objRow):
            if is_valid_project_subject_name(pszValue):
                continue
            objInvalidCells.append((iRowIndex + 1, iColumnIndex + 1, pszValue))
    return objInvalidCells


def write_project_name_validation_error_file(
    pszErrorFilePath: str,
    pszInputFilePath: str,
    objInvalidCells: List[tuple[int, int, str]],
) -> None:
    with open(pszErrorFilePath, mode="w", encoding="utf-8", newline="") as objErrorFile:
        objErrorFile.write("プロジェクト名形式エラー\n")
        objErrorFile.write(f"入力ファイル: {pszInputFilePath}\n")
        objErrorFile.write(f"エラー件数: {len(objInvalidCells)}\n")
        objErrorFile.write("\n")
        for iRowNumber, iColumnNumber, pszValue in objInvalidCells:
            objErrorFile.write(
                f"行:{iRowNumber} 列:{iColumnNumber} 値:{pszValue}\n"
            )


def find_row_index_with_subject_tab(objRows: List[List[str]], iStartIndex: int) -> int | None:
    for iRowIndex in range(iStartIndex, len(objRows)):
        objRow = objRows[iRowIndex]
        if any(
            "科目名\t" in pszValue or pszValue.strip() == "科目名"
            for pszValue in objRow
        ):
            return iRowIndex
    return None


def build_pj_name_vertical_rows(objRows: List[List[str]]) -> List[List[str]]:
    if not objRows:
        return []

    objHeaderRow: List[str] = objRows[0]
    objItemRows: List[List[str]] = objRows[1:]

    objVerticalRows: List[List[str]] = []
    objVerticalHeader: List[str] = ["PJ名称"]
    for objItemRow in objItemRows:
        pszItemName: str = objItemRow[0] if len(objItemRow) > 0 else ""
        objVerticalHeader.append(pszItemName)
    objVerticalRows.append(objVerticalHeader)

    for iColumnIndex in range(1, len(objHeaderRow)):
        pszProjectName: str = objHeaderRow[iColumnIndex]
        objVerticalRow: List[str] = [pszProjectName]
        for objItemRow in objItemRows:
            pszValue: str = objItemRow[iColumnIndex] if len(objItemRow) > iColumnIndex else ""
            objVerticalRow.append(pszValue)
        objVerticalRows.append(objVerticalRow)

    return objVerticalRows


def write_first_row_tabs_to_newlines(pszInputFilePath: str, pszOutputFilePath: str) -> None:
    with open(pszInputFilePath, mode="r", encoding="utf-8", newline="") as objInputFile:
        pszFirstLine: str = objInputFile.readline()
    pszConverted: str = pszFirstLine.replace("\t", "\n")
    with open(pszOutputFilePath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objOutputFile.write(pszConverted)


def insert_company_expense_columns(objRows: List[List[str]]) -> None:
    if not objRows:
        return
    objHeaderRow: List[str] = objRows[0]
    try:
        iHeadOfficeIndex: int = objHeaderRow.index("本部")
    except ValueError:
        return

    objExpenseColumns: List[str] = [
        "1Cカンパニー販管費",
        "2Cカンパニー販管費",
        "3Cカンパニー販管費",
        "4Cカンパニー販管費",
        "事業開発カンパニー販管費",
        "社長室カンパニー販管費",
        "本部カンパニー販管費",
    ]
    iInsertIndex: int = iHeadOfficeIndex + 1
    objHeaderRow[iInsertIndex:iInsertIndex] = objExpenseColumns
    for objRow in objRows[1:]:
        objRow[iInsertIndex:iInsertIndex] = ["0"] * len(objExpenseColumns)


COMPANY_EXPENSE_REPLACEMENTS: dict[str, str] = {
    "1Cカンパニー販管費": "C001_1Cカンパニー販管費",
    "2Cカンパニー販管費": "C002_2Cカンパニー販管費",
    "3Cカンパニー販管費": "C003_3Cカンパニー販管費",
    "4Cカンパニー販管費": "C004_4Cカンパニー販管費",
    "事業開発カンパニー販管費": "C005_事業開発カンパニー販管費",
    "社長室カンパニー販管費": "C006_社長室カンパニー販管費",
    "本部カンパニー販管費": "C007_本部カンパニー販管費",
}


def replace_company_expense_labels(objRows: List[List[str]], objReplacementMap: dict[str, str]) -> None:
    for objRow in objRows:
        for iIndex, pszValue in enumerate(objRow):
            if pszValue in objReplacementMap:
                objRow[iIndex] = objReplacementMap[pszValue]


def append_debug_log(pszMessage: str, pszDebugFilePath: str = "debug.txt") -> None:
    with open(pszDebugFilePath, mode="a", encoding="utf-8", newline="") as objDebugFile:
        objDebugFile.write(f"{pszMessage}\n")


def insert_allocated_sga_row(objRows: List[List[str]]) -> None:
    if not objRows:
        return
    for iRowIndex, objRow in enumerate(objRows):
        if not objRow or objRow[0] != "販売費及び一般管理費計":
            continue
        iColumnCount: int = len(objRow)
        objRows.insert(iRowIndex + 1, ["配賦販管費"] + ["0"] * max(iColumnCount - 1, 0))
        return


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python src/PL_CsvToTsv_Cmd.py <csv_file> [<csv_file> ...]")
        return 1

    iExitCode: int = 0
    objCostReportVerticalFilePaths: List[str] = []
    objCostReportProjectNameVerticalFilePaths: List[str] = []
    objProfitLossProjectNameVerticalFilePaths: List[str] = []
    objProfitLossVerticalFilePaths: List[str] = []
    for pszInputFilePath in sys.argv[1:]:
        try:
            append_debug_log("start")
            iFileYear: int
            iFileMonth: int
            iFileYear, iFileMonth = get_target_year_month_from_filename(pszInputFilePath)
            append_debug_log(f"filename parsed: {iFileYear}-{iFileMonth:02d}")

            if not os.path.isfile(pszInputFilePath):
                raise FileNotFoundError(f"入力ファイルが存在しません: {pszInputFilePath}")

            objRows: List[List[str]] = read_csv_rows(pszInputFilePath)
            if len(objRows) < 2:
                raise ValueError("集計期間の取得に必要な行が存在しません。")
            append_debug_log(f"rows read: {len(objRows)}")

            normalize_project_names_in_row(objRows, 7)
            iSubjectRowIndex = find_row_index_with_subject_tab(objRows, 8)
            if iSubjectRowIndex is not None:
                normalize_project_names_in_row(objRows, iSubjectRowIndex)
            append_debug_log("project names normalized")

            objValidationTargetRowIndices: List[int] = [7]
            if iSubjectRowIndex is not None and iSubjectRowIndex != 7:
                objValidationTargetRowIndices.append(iSubjectRowIndex)
            objInvalidCells = collect_invalid_project_subject_cells(
                objRows,
                objValidationTargetRowIndices,
            )

            pszRowA: str = objRows[1][1] if len(objRows[1]) > 1 else ""
            append_debug_log(f"B2 value: {pszRowA}")
            pszRowANormalized: str = re.sub(r"[ \u3000]", "", pszRowA)
            if "期首振戻" in pszRowANormalized:
                append_debug_log("period parse skipped due to 期首振戻; using filename")
            else:
                iPeriodYear: int
                iPeriodMonth: int
                iPeriodYear, iPeriodMonth = get_target_year_month_from_period_row(pszRowA)
                append_debug_log(f"period parsed: {iPeriodYear}-{iPeriodMonth:02d}")

                if iFileYear != iPeriodYear or iFileMonth != iPeriodMonth:
                    raise ValueError("ファイル名と集計期間の対象年月が一致しません。")
                append_debug_log("period matches filename")

            pszMonth: str = f"{iFileMonth:02d}"
            if objInvalidCells:
                pszProjectNameValidationErrorPath: str = (
                    f"損益計算書_{iFileYear}年{pszMonth}月_プロジェクト名形式エラー_error.txt"
                )
                write_project_name_validation_error_file(
                    pszProjectNameValidationErrorPath,
                    pszInputFilePath,
                    objInvalidCells,
                )
                append_debug_log(
                    f"project name format error file written: {pszProjectNameValidationErrorPath}"
                )

            pszOutputFilePath: str = f"損益計算書_{iFileYear}年{pszMonth}月.tsv"
            pszCostReportFilePath: str = f"製造原価報告書_{iFileYear}年{pszMonth}月.tsv"
            objOutputRows: List[List[str]] = []
            objCostReportRows: List[List[str]] = []
            iSplitIndex: int | None = None
            for iRowIndex in range(7, len(objRows) - 1):
                objRow: List[str] = objRows[iRowIndex]
                objNextRow: List[str] = objRows[iRowIndex + 1]
                if objRow and objNextRow and objRow[0] == "当期純利益" and objNextRow[0] == "科目名":
                    iSplitIndex = iRowIndex
                    break

            if iSplitIndex is None:
                for iRowIndex in range(7, len(objRows)):
                    objRow = objRows[iRowIndex]
                    objOutputRows.append(objRow[:])
            else:
                for iRowIndex in range(7, iSplitIndex + 1):
                    objRow = objRows[iRowIndex]
                    objOutputRows.append(objRow[:])
                for iRowIndex in range(iSplitIndex + 1, len(objRows)):
                    objRow = objRows[iRowIndex]
                    objCostReportRows.append(objRow[:])
            append_debug_log(f"output rows prepared: {len(objOutputRows)}")

            insert_allocated_sga_row(objOutputRows)
            append_debug_log("allocated sga row inserted")

            if (iFileYear, iFileMonth) <= (2025, 7):
                insert_company_expense_columns(objOutputRows)
                append_debug_log("company expense columns inserted")

            replace_company_expense_labels(
                objOutputRows,
                COMPANY_EXPENSE_REPLACEMENTS,
            )
            append_debug_log("company expense labels replaced")

            write_tsv_rows(pszOutputFilePath, objOutputRows)
            append_debug_log(f"tsv written: {pszOutputFilePath}")
            objOutputTsvRows: List[List[str]] = read_tsv_rows(pszOutputFilePath)
            objOutputVerticalRows: List[List[str]] = build_first_column_rows(objOutputTsvRows)
            pszOutputVerticalFilePath: str = (
                f"損益計算書_{iFileYear}年{pszMonth}月_科目名_vertical.tsv"
            )
            write_tsv_rows(pszOutputVerticalFilePath, objOutputVerticalRows)
            append_debug_log(f"vertical tsv written: {pszOutputVerticalFilePath}")
            objProfitLossVerticalFilePaths.append(pszOutputVerticalFilePath)

            if objCostReportRows:
                write_tsv_rows(pszCostReportFilePath, objCostReportRows)
                append_debug_log(f"tsv written: {pszCostReportFilePath}")
                objCostReportTsvRows: List[List[str]] = read_tsv_rows(pszCostReportFilePath)
                objCostReportVerticalRows: List[List[str]] = build_first_column_rows(objCostReportTsvRows)
                pszCostReportVerticalFilePath: str = (
                    f"製造原価報告書_{iFileYear}年{pszMonth}月_科目名_vertical.tsv"
                )
                write_tsv_rows(pszCostReportVerticalFilePath, objCostReportVerticalRows)
                append_debug_log(f"vertical tsv written: {pszCostReportVerticalFilePath}")

                objCostReportVerticalFilePaths.append(pszCostReportVerticalFilePath)


            pszVerticalOutputFilePath: str = f"損益計算書_{iFileYear}年{pszMonth}月_PJ名称_vertical.tsv"
            write_first_row_tabs_to_newlines(pszOutputFilePath, pszVerticalOutputFilePath)
            append_debug_log(f"vertical tsv written: {pszVerticalOutputFilePath}")
        except Exception as objException:
            iExitCode = 1
            append_debug_log(f"error: {objException}")
            print(objException)
            try:
                iErrorYear: int
                iErrorMonth: int
                iErrorYear, iErrorMonth = get_target_year_month_from_filename(pszInputFilePath)
                pszErrorMonth: str = f"{iErrorMonth:02d}"
                pszErrorFilePath: str = f"損益計算書_{iErrorYear}年{pszErrorMonth}月_error.txt"
            except Exception:
                pszBaseName: str = os.path.basename(pszInputFilePath)
                pszErrorFilePath = f"{pszBaseName}_error.txt"
            with open(pszErrorFilePath, mode="w", encoding="utf-8", newline="") as objErrorFile:
                objErrorFile.write(str(objException))

    create_union_subject_vertical_tsvs(objCostReportVerticalFilePaths)
    create_union_subject_vertical_tsvs(objProfitLossVerticalFilePaths)
    create_profit_loss_union_tsvs(
        objProfitLossVerticalFilePaths,
        objProfitLossProjectNameVerticalFilePaths,
    )
    create_cost_report_union_tsvs(
        objCostReportVerticalFilePaths,
        objCostReportProjectNameVerticalFilePaths,
    )
    create_union_project_name_vertical_tsvs(
        objCostReportProjectNameVerticalFilePaths,
        bWriteHorizontal=True,
    )
    create_union_project_name_vertical_tsvs(
        objProfitLossProjectNameVerticalFilePaths,
        bWriteHorizontal=True,
    )
    create_drag_and_drop_manhour_and_pl_folder()
    return iExitCode


def create_drag_and_drop_manhour_and_pl_folder() -> None:
    pszScriptDirectory: str = os.path.dirname(os.path.abspath(__file__))

    def extract_year_month_from_path(pszPath: str) -> Optional[Tuple[int, int]]:
        pszBaseName: str = os.path.basename(pszPath)
        objMatch = re.search(r"_(\d{4})年(\d{1,2})月", pszBaseName)
        if objMatch is None:
            return None
        try:
            iYear: int = int(objMatch.group(1))
            iMonth: int = int(objMatch.group(2))
        except ValueError:
            return None
        if iMonth < 1 or iMonth > 12:
            return None
        return iYear, iMonth

    def next_year_month(iYear: int, iMonth: int) -> Tuple[int, int]:
        iMonth += 1
        if iMonth > 12:
            return iYear + 1, 1
        return iYear, iMonth

    def range_length(objStart: Tuple[int, int], objEnd: Tuple[int, int]) -> int:
        iYearStart, iMonthStart = objStart
        iYearEnd, iMonthEnd = objEnd
        return (iYearEnd * 12 + iMonthEnd) - (iYearStart * 12 + iMonthStart) + 1

    def update_best_range(
        objCurrentBest: Optional[Tuple[Tuple[int, int], Tuple[int, int]]],
        iCurrentBestLength: int,
        objCandidateStart: Tuple[int, int],
        objCandidateEnd: Tuple[int, int],
    ) -> Tuple[Optional[Tuple[Tuple[int, int], Tuple[int, int]]], int]:
        iCandidateLength: int = range_length(objCandidateStart, objCandidateEnd)
        if iCandidateLength > iCurrentBestLength:
            return (objCandidateStart, objCandidateEnd), iCandidateLength
        if iCandidateLength == iCurrentBestLength and objCurrentBest is not None:
            if objCandidateEnd > objCurrentBest[1]:
                return (objCandidateStart, objCandidateEnd), iCandidateLength
        if objCurrentBest is None:
            return (objCandidateStart, objCandidateEnd), iCandidateLength
        return objCurrentBest, iCurrentBestLength

    def find_best_continuous_range(
        objYearMonths: List[Tuple[int, int]],
    ) -> Optional[Tuple[Tuple[int, int], Tuple[int, int]]]:
        if not objYearMonths:
            return None
        objSorted: List[Tuple[int, int]] = sorted(set(objYearMonths))
        objBestRange: Optional[Tuple[Tuple[int, int], Tuple[int, int]]] = None
        iBestLength: int = 0
        objCurrentStart: Tuple[int, int] = objSorted[0]
        objCurrentEnd: Tuple[int, int] = objSorted[0]
        for objMonth in objSorted[1:]:
            if next_year_month(*objCurrentEnd) == objMonth:
                objCurrentEnd = objMonth
                continue
            objBestRange, iBestLength = update_best_range(
                objBestRange,
                iBestLength,
                objCurrentStart,
                objCurrentEnd,
            )
            objCurrentStart = objMonth
            objCurrentEnd = objMonth
        objBestRange, _ = update_best_range(
            objBestRange,
            iBestLength,
            objCurrentStart,
            objCurrentEnd,
        )
        return objBestRange

    def month_to_ordinal(objMonth: Tuple[int, int]) -> int:
        iYear, iMonth = objMonth
        return iYear * 12 + iMonth

    def is_month_in_range(
        objMonth: Tuple[int, int],
        objRange: Tuple[Tuple[int, int], Tuple[int, int]],
    ) -> bool:
        iValue: int = month_to_ordinal(objMonth)
        iStart: int = month_to_ordinal(objRange[0])
        iEnd: int = month_to_ordinal(objRange[1])
        return iStart <= iValue <= iEnd

    def write_selected_range_file(
        pszDirectory: str,
        objRange: Tuple[Tuple[int, int], Tuple[int, int]],
    ) -> None:
        iStartYear, iStartMonth = objRange[0]
        iEndYear, iEndMonth = objRange[1]
        pszOutputPath: str = os.path.join(
            pszDirectory,
            "PL_CsvToTsv_Cmd_SelectedRange.txt",
        )
        pszStartText: str = f"{iStartYear:04d}/{iStartMonth:02d}"
        pszEndText: str = f"{iEndYear:04d}/{iEndMonth:02d}"
        objLines: List[str] = [
            "採用範囲:",
            f"開始: {pszStartText}",
            f"終了: {pszEndText}",
        ]
        with open(pszOutputPath, "w", encoding="utf-8", newline="") as objFile:
            objFile.write("\n".join(objLines) + "\n")

    objMonthFiles: Dict[Tuple[int, int], Dict[str, List[str] | str]] = {}
    for pszFileName in os.listdir(pszScriptDirectory):
        pszSourcePath: str = os.path.join(pszScriptDirectory, pszFileName)
        if not os.path.isfile(pszSourcePath):
            continue
        objMonth: Optional[Tuple[int, int]] = extract_year_month_from_path(pszFileName)
        if objMonth is None:
            continue

        if (
            pszFileName.startswith("工数_")
            and (
                pszFileName.endswith("_step0014_各プロジェクトの計上カンパニー名_工数_カンパニーの工数.tsv")
                or pszFileName.endswith("_step15_各プロジェクトの工数.tsv")
            )
        ):
            objMonthEntry = objMonthFiles.setdefault(objMonth, {"manhour": [], "pl": ""})
            objMonthEntry["manhour"].append(pszSourcePath)
            continue

        if (
            pszFileName.startswith("損益計算書_")
            and pszFileName.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv")
        ):
            objMonthEntry = objMonthFiles.setdefault(objMonth, {"manhour": [], "pl": ""})
            objMonthEntry["pl"] = pszSourcePath

    objPairMonths: List[Tuple[int, int]] = [
        objMonth
        for objMonth, objEntry in objMonthFiles.items()
        if objEntry["manhour"] and objEntry["pl"]
    ]
    objSelectedRange = find_best_continuous_range(objPairMonths)
    if objSelectedRange is None:
        return

    objSelectedSourcePaths: List[str] = []
    for objMonth in sorted(objPairMonths):
        if not is_month_in_range(objMonth, objSelectedRange):
            continue
        objEntry = objMonthFiles[objMonth]
        objSelectedSourcePaths.extend(sorted(objEntry["manhour"]))
        objSelectedSourcePaths.append(objEntry["pl"])

    if not objSelectedSourcePaths:
        return

    write_selected_range_file(pszScriptDirectory, objSelectedRange)

    pszOutputDirectory: str = os.path.join(pszScriptDirectory, "DragAndDropManhourAndPl")
    os.makedirs(pszOutputDirectory, exist_ok=True)
    for pszSourcePath in objSelectedSourcePaths:
        pszDestinationPath: str = os.path.join(pszOutputDirectory, os.path.basename(pszSourcePath))
        shutil.copy2(pszSourcePath, pszDestinationPath)

    if hasattr(os, "startfile"):
        os.startfile(pszOutputDirectory)


def create_union_subject_vertical_tsvs(objCostReportVerticalFilePaths: List[str]) -> None:
    if not objCostReportVerticalFilePaths:
        return

    objSubjectLists: List[List[str]] = []
    objSubjectsByFilePath: dict[str, List[str]] = {}
    for pszFilePath in sort_vertical_file_paths(objCostReportVerticalFilePaths):
        objRows: List[List[str]] = read_tsv_rows(pszFilePath)
        objSubjects: List[str] = build_unique_subjects(objRows)
        objSubjectLists.append(objSubjects)
        objSubjectsByFilePath[pszFilePath] = objSubjects
    objUnionSubjects: List[str] = build_cumulative_subject_order(objSubjectLists)
    objUnionRows: List[List[str]] = build_subject_vertical_rows(objUnionSubjects)

    for pszFilePath in sort_vertical_file_paths(objCostReportVerticalFilePaths):
        pszUnionFilePath: str = pszFilePath.replace("_科目名_vertical.tsv", "_科目名_A∪B_vertical.tsv")
        write_tsv_rows(pszUnionFilePath, objUnionRows)
        append_debug_log(f"union vertical tsv written: {pszUnionFilePath}")
        if pszFilePath.startswith("損益計算書_") or pszFilePath.startswith("製造原価報告書_"):
            objSubjects: List[str] = objSubjectsByFilePath.get(pszFilePath, [])
            objSubjectSet: set[str] = set(objSubjects)
            objMissingSubjects: List[str] = [
                pszSubject for pszSubject in objUnionSubjects if pszSubject not in objSubjectSet
            ]
            pszMissingFilePath: str = pszFilePath.replace(
                "_科目名_vertical.tsv",
                "_科目名_vertical_subject.tsv",
            )
            write_tsv_rows(pszMissingFilePath, build_subject_vertical_rows(objMissingSubjects))
            append_debug_log(f"vertical subject tsv written: {pszMissingFilePath}")


def create_profit_loss_union_tsvs(
    objProfitLossVerticalFilePaths: List[str],
    objProfitLossProjectNameVerticalFilePaths: List[str],
) -> None:
    if not objProfitLossVerticalFilePaths:
        return

    for pszVerticalFilePath in objProfitLossVerticalFilePaths:
        pszUnionVerticalFilePath: str = pszVerticalFilePath.replace(
            "_科目名_vertical.tsv",
            "_科目名_A∪B_vertical.tsv",
        )
        if not os.path.isfile(pszUnionVerticalFilePath):
            continue
        pszProfitLossFilePath: str = pszVerticalFilePath.replace("_科目名_vertical.tsv", ".tsv")
        if not os.path.isfile(pszProfitLossFilePath):
            continue

        objUnionRows: List[List[str]] = read_tsv_rows(pszUnionVerticalFilePath)
        objProfitLossRows: List[List[str]] = read_tsv_rows(pszProfitLossFilePath)
        objSubjectRows: List[str] = [objRow[0] if objRow else "" for objRow in objUnionRows]
        objProfitLossRowMap: dict[str, List[str]] = {}
        for objRow in objProfitLossRows:
            pszSubject: str = objRow[0] if objRow else ""
            if pszSubject in objProfitLossRowMap:
                continue
            objProfitLossRowMap[pszSubject] = objRow

        iColumnCount: int = len(objProfitLossRows[0]) if objProfitLossRows else 1
        objUnionProfitLossRows: List[List[str]] = []
        for pszSubject in objSubjectRows:
            if pszSubject in objProfitLossRowMap:
                objUnionProfitLossRows.append(objProfitLossRowMap[pszSubject])
            else:
                objUnionProfitLossRows.append([pszSubject] + ["0"] * max(iColumnCount - 1, 0))

        pszUnionProfitLossFilePath: str = pszVerticalFilePath.replace("_科目名_vertical.tsv", "_A∪B.tsv")
        write_tsv_rows(pszUnionProfitLossFilePath, objUnionProfitLossRows)
        append_debug_log(f"union tsv written: {pszUnionProfitLossFilePath}")
        pszUnionProfitLossVerticalFilePath: str = pszUnionProfitLossFilePath.replace(
            "_A∪B.tsv",
            "_A∪B_vertical.tsv",
        )
        write_tsv_rows(pszUnionProfitLossVerticalFilePath, transpose_rows(objUnionProfitLossRows))
        append_debug_log(f"union vertical tsv written: {pszUnionProfitLossVerticalFilePath}")
        objUnionProfitLossVerticalRows: List[List[str]] = read_tsv_rows(pszUnionProfitLossVerticalFilePath)
        objProjectNameVerticalRows: List[List[str]] = build_first_column_rows(objUnionProfitLossVerticalRows)
        pszProjectNameVerticalFilePath: str = pszUnionProfitLossVerticalFilePath.replace(
            "_A∪B_vertical.tsv",
            "_A∪B_プロジェクト名_vertical.tsv",
        )
        write_tsv_rows(pszProjectNameVerticalFilePath, objProjectNameVerticalRows)
        append_debug_log(f"project name vertical tsv written: {pszProjectNameVerticalFilePath}")
        objProfitLossProjectNameVerticalFilePaths.append(pszProjectNameVerticalFilePath)


def create_cost_report_union_tsvs(
    objCostReportVerticalFilePaths: List[str],
    objCostReportProjectNameVerticalFilePaths: List[str],
) -> None:
    if not objCostReportVerticalFilePaths:
        return

    for pszVerticalFilePath in objCostReportVerticalFilePaths:
        pszUnionVerticalFilePath: str = pszVerticalFilePath.replace(
            "_科目名_vertical.tsv",
            "_科目名_A∪B_vertical.tsv",
        )
        if not os.path.isfile(pszUnionVerticalFilePath):
            continue
        pszCostReportFilePath: str = pszVerticalFilePath.replace("_科目名_vertical.tsv", ".tsv")
        if not os.path.isfile(pszCostReportFilePath):
            continue

        objUnionRows: List[List[str]] = read_tsv_rows(pszUnionVerticalFilePath)
        objCostReportRows: List[List[str]] = read_tsv_rows(pszCostReportFilePath)
        objSubjectRows: List[str] = [objRow[0] if objRow else "" for objRow in objUnionRows]
        objCostReportRowMap: dict[str, List[str]] = {}
        for objRow in objCostReportRows:
            pszSubject: str = objRow[0] if objRow else ""
            if pszSubject in objCostReportRowMap:
                continue
            objCostReportRowMap[pszSubject] = objRow

        iColumnCount: int = len(objCostReportRows[0]) if objCostReportRows else 1
        objUnionCostReportRows: List[List[str]] = []
        for pszSubject in objSubjectRows:
            if pszSubject in objCostReportRowMap:
                objUnionCostReportRows.append(objCostReportRowMap[pszSubject])
            else:
                objUnionCostReportRows.append([pszSubject] + ["0"] * max(iColumnCount - 1, 0))

        pszUnionCostReportFilePath: str = pszVerticalFilePath.replace("_科目名_vertical.tsv", "_A∪B.tsv")
        write_tsv_rows(pszUnionCostReportFilePath, objUnionCostReportRows)
        append_debug_log(f"union tsv written: {pszUnionCostReportFilePath}")
        pszUnionCostReportVerticalFilePath: str = pszUnionCostReportFilePath.replace(
            "_A∪B.tsv",
            "_A∪B_vertical.tsv",
        )
        write_tsv_rows(pszUnionCostReportVerticalFilePath, transpose_rows(objUnionCostReportRows))
        append_debug_log(f"union vertical tsv written: {pszUnionCostReportVerticalFilePath}")
        objUnionCostReportVerticalRows: List[List[str]] = read_tsv_rows(pszUnionCostReportVerticalFilePath)
        objProjectNameVerticalRows: List[List[str]] = build_first_column_rows(objUnionCostReportVerticalRows)
        pszProjectNameVerticalFilePath: str = pszUnionCostReportVerticalFilePath.replace(
            "_A∪B_vertical.tsv",
            "_A∪B_プロジェクト名_vertical.tsv",
        )
        write_tsv_rows(pszProjectNameVerticalFilePath, objProjectNameVerticalRows)
        append_debug_log(f"project name vertical tsv written: {pszProjectNameVerticalFilePath}")
        objCostReportProjectNameVerticalFilePaths.append(pszProjectNameVerticalFilePath)


def create_union_project_name_vertical_tsvs(
    objProjectNameVerticalFilePaths: List[str],
    bWriteHorizontal: bool = False,
) -> None:
    if not objProjectNameVerticalFilePaths:
        return

    objProjectNameLists: List[List[str]] = []
    objProjectNamesByFilePath: dict[str, List[str]] = {}
    for pszFilePath in sort_vertical_file_paths(objProjectNameVerticalFilePaths):
        objRows: List[List[str]] = read_tsv_rows(pszFilePath)
        objProjectNames: List[str] = build_unique_subjects(objRows)
        if objProjectNames and objProjectNames[0] == "科目名":
            objProjectNames = objProjectNames[1:]
        objProjectNameLists.append(objProjectNames)
        objProjectNamesByFilePath[pszFilePath] = objProjectNames

    objUnionProjectNames: List[str] = build_cumulative_subject_order(objProjectNameLists)

    for pszFilePath in sort_vertical_file_paths(objProjectNameVerticalFilePaths):
        pszBaseVerticalFilePath: str = pszFilePath.replace(
            "_A∪B_プロジェクト名_vertical.tsv",
            "_A∪B_vertical.tsv",
        )
        if not os.path.isfile(pszBaseVerticalFilePath):
            continue
        objBaseVerticalRows: List[List[str]] = read_tsv_rows(pszBaseVerticalFilePath)
        if not objBaseVerticalRows:
            continue
        objHeaderRow: List[str] = objBaseVerticalRows[0]
        objProjectRows: List[List[str]] = objBaseVerticalRows[1:]
        objProjectRowMap: dict[str, List[str]] = {}
        for objRow in objProjectRows:
            pszProjectName: str = objRow[0] if objRow else ""
            if pszProjectName in objProjectRowMap:
                continue
            objProjectRowMap[pszProjectName] = objRow

        iColumnCount: int = len(objHeaderRow)
        objUnionRows: List[List[str]] = [objHeaderRow]
        for pszProjectName in objUnionProjectNames:
            if pszProjectName in objProjectRowMap:
                objUnionRows.append(objProjectRowMap[pszProjectName])
            else:
                objUnionRows.append([pszProjectName] + ["0"] * max(iColumnCount - 1, 0))

        pszUnionFilePath: str = pszFilePath.replace(
            "_A∪B_プロジェクト名_vertical.tsv",
            "_A∪B_プロジェクト名_C∪D_vertical.tsv",
        )
        write_tsv_rows(pszUnionFilePath, objUnionRows)
        append_debug_log(f"union project name vertical tsv written: {pszUnionFilePath}")
        if bWriteHorizontal:
            pszUnionHorizontalFilePath: str = pszUnionFilePath.replace(
                "_A∪B_プロジェクト名_C∪D_vertical.tsv",
                "_A∪B_プロジェクト名_C∪D.tsv",
            )
            write_tsv_rows(pszUnionHorizontalFilePath, transpose_rows(objUnionRows))
            append_debug_log(f"union project name tsv written: {pszUnionHorizontalFilePath}")
        objProjectNames: List[str] = objProjectNamesByFilePath.get(pszFilePath, [])
        objProjectNameSet: set[str] = set(objProjectNames)
        objMissingProjectNames: List[str] = [
            pszName for pszName in objUnionProjectNames if pszName not in objProjectNameSet
        ]
        pszMissingFilePath: str = pszFilePath.replace(
            "_A∪B_プロジェクト名_vertical.tsv",
            "_プロジェクト名_vertical_subject.tsv",
        )
        write_tsv_rows(pszMissingFilePath, build_subject_vertical_rows(objMissingProjectNames))
        append_debug_log(f"project name vertical subject tsv written: {pszMissingFilePath}")


if __name__ == "__main__":
    sys.exit(main())
