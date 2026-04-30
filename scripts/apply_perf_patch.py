# -*- coding: utf-8 -*-
"""
Apply performance optimizations to the 5 dedup VBA modules.
Preserves GBK encoding. Run from project root.

Optimizations:
1. Bulk array read (replace cell-by-cell access)
2. Union batch delete (replace row-by-row delete)
3. Combined blank-check + key-build (eliminate double traversal)
"""
import os
import sys

MODULES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "VBA_Export", "Modules")
ENCODING = "gbk"


def read_file(path):
    with open(path, "r", encoding=ENCODING) as f:
        return f.read()


def write_file(path, content):
    with open(path, "w", encoding=ENCODING) as f:
        f.write(content)


def patch_replace(content, old, new, filename, label):
    if old not in content:
        print(f"  WARNING: patch '{label}' not found in {filename}")
        return content
    if content.count(old) > 1:
        print(f"  WARNING: patch '{label}' matches {content.count(old)} times in {filename}, applying first only")
    content = content.replace(old, new, 1)
    print(f"  Applied: {label}")
    return content


# ============================================================
# Module 1: 功能3dot11按批注检查重复数据.bas
# Optimization 1 only: bulk array read
# ============================================================
def patch_3dot11():
    fname = "功能3dot11按批注检查重复数据.bas"
    path = os.path.join(MODULES_DIR, fname)
    content = read_file(path)

    # 1a. Modify MarkDuplicateRowsByComment to use array
    old = (
        "Private Function MarkDuplicateRowsByComment(ByVal ws As Worksheet) As Long\r\n"
        "    Dim keyCols As Collection\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set keyCols = FindCommentMarkedColumns(ws)\r\n"
        "    If keyCols Is Nothing Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    ElseIf keyCols.Count = 0 Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    End If\r\n"
        "    If keyCols Is Nothing Then Exit Function\r\n"
        "    If keyCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        rowKey = BuildRowKeyByColumns(ws, r, keyCols)\r\n"
        "        If seen.Exists(rowKey) Then\r\n"
        "            MarkDuplicateRow ws, r, firstCol, lastCol\r\n"
        "            MarkDuplicateRowsByComment = MarkDuplicateRowsByComment + 1\r\n"
        "        Else\r\n"
        "            seen.Add rowKey, True\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "End Function"
    )
    new = (
        "Private Function MarkDuplicateRowsByComment(ByVal ws As Worksheet) As Long\r\n"
        "    Dim keyCols As Collection\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim dataArr As Variant\r\n"
        "    Dim rng As Range\r\n"
        "    Dim colOffset As Long\r\n"
        "    Dim rowOffset As Long\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set keyCols = FindCommentMarkedColumns(ws)\r\n"
        "    If keyCols Is Nothing Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    ElseIf keyCols.Count = 0 Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    End If\r\n"
        "    If keyCols Is Nothing Then Exit Function\r\n"
        "    If keyCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    Set rng = ws.Range(ws.Cells(DATA_START_ROW, firstCol), ws.Cells(lastRow, lastCol))\r\n"
        "    If rng.Cells.CountLarge = 1 Then\r\n"
        "        ReDim dataArr(1 To 1, 1 To 1)\r\n"
        "        dataArr(1, 1) = rng.Value2\r\n"
        "    Else\r\n"
        "        dataArr = rng.Value2\r\n"
        "    End If\r\n"
        "    colOffset = firstCol - 1\r\n"
        "    rowOffset = DATA_START_ROW - 1\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        rowKey = BuildRowKeyFromArray(dataArr, r - rowOffset, keyCols, colOffset)\r\n"
        "        If seen.Exists(rowKey) Then\r\n"
        "            MarkDuplicateRow ws, r, firstCol, lastCol\r\n"
        "            MarkDuplicateRowsByComment = MarkDuplicateRowsByComment + 1\r\n"
        "        Else\r\n"
        "            seen.Add rowKey, True\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "End Function"
    )
    content = patch_replace(content, old, new, fname, "MarkDuplicateRowsByComment -> array read")

    # 1b. Replace BuildRowKeyByColumns + GetCellKeyValue with BuildRowKeyFromArray
    old = (
        "Private Function BuildRowKeyByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As String\r\n"
        "    Dim idx As Variant\r\n"
        "    Dim parts As String\r\n"
        "\r\n"
        "    For Each idx In colIndexes\r\n"
        "        parts = parts & KEY_SEP & GetCellKeyValue(ws.Cells(rowIndex, CLng(idx)))\r\n"
        "    Next idx\r\n"
        "    BuildRowKeyByColumns = parts\r\n"
        "End Function\r\n"
        "\r\n"
        "Private Function GetCellKeyValue(ByVal c As Range) As String\r\n"
        "    On Error Resume Next\r\n"
        "    GetCellKeyValue = NormalizeText(c.Value2)\r\n"
        "    On Error GoTo 0\r\n"
        "End Function"
    )
    new = (
        "Private Function BuildRowKeyFromArray(ByVal dataArr As Variant, ByVal arrRow As Long, _\r\n"
        "                                     ByVal colIndexes As Collection, ByVal colOffset As Long) As String\r\n"
        "    Dim idx As Variant\r\n"
        "    Dim parts As String\r\n"
        "\r\n"
        "    For Each idx In colIndexes\r\n"
        "        parts = parts & KEY_SEP & NormalizeText(dataArr(arrRow, CLng(idx) - colOffset))\r\n"
        "    Next idx\r\n"
        "    BuildRowKeyFromArray = parts\r\n"
        "End Function"
    )
    content = patch_replace(content, old, new, fname, "BuildRowKeyByColumns -> BuildRowKeyFromArray")

    write_file(path, content)
    print(f"OK: {fname}")


# ============================================================
# Module 2: 功能3dot14按批注删除重复行.bas
# All 3 optimizations
# ============================================================
def patch_3dot14():
    fname = "功能3dot14按批注删除重复行.bas"
    path = os.path.join(MODULES_DIR, fname)
    content = read_file(path)

    # 2a. Modify DeleteDuplicateRowsByComment: array read + combined blank/key + Union delete
    old = (
        "Private Function DeleteDuplicateRowsByComment(ByVal ws As Worksheet) As Long\r\n"
        "    Dim keyCols As Collection\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim rowsToDelete As Collection\r\n"
        "    Dim i As Long\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set keyCols = FindCommentMarkedColumns(ws)\r\n"
        "    If keyCols Is Nothing Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    ElseIf keyCols.Count = 0 Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    End If\r\n"
        "    If keyCols Is Nothing Then Exit Function\r\n"
        "    If keyCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    Set rowsToDelete = New Collection\r\n"
        "\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        If Not RowIsBlankByColumns(ws, r, keyCols) Then\r\n"
        "            rowKey = BuildRowKeyByColumns(ws, r, keyCols)\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                rowsToDelete.Add r\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "\r\n"
        "    For i = rowsToDelete.Count To 1 Step -1\r\n"
        "        ws.Rows(CLng(rowsToDelete(i))).EntireRow.Delete\r\n"
        "        DeleteDuplicateRowsByComment = DeleteDuplicateRowsByComment + 1\r\n"
        "    Next i\r\n"
        "End Function"
    )
    new = (
        "Private Function DeleteDuplicateRowsByComment(ByVal ws As Worksheet) As Long\r\n"
        "    Dim keyCols As Collection\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim rowsToDelete As Collection\r\n"
        "    Dim i As Long\r\n"
        "    Dim dataArr As Variant\r\n"
        "    Dim rng As Range\r\n"
        "    Dim colOffset As Long\r\n"
        "    Dim rowOffset As Long\r\n"
        "    Dim isBlank As Boolean\r\n"
        "    Dim delRange As Range\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set keyCols = FindCommentMarkedColumns(ws)\r\n"
        "    If keyCols Is Nothing Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    ElseIf keyCols.Count = 0 Then\r\n"
        "        Set keyCols = BuildColumnCollection(firstCol, lastCol)\r\n"
        "    End If\r\n"
        "    If keyCols Is Nothing Then Exit Function\r\n"
        "    If keyCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    Set rng = ws.Range(ws.Cells(DATA_START_ROW, firstCol), ws.Cells(lastRow, lastCol))\r\n"
        "    If rng.Cells.CountLarge = 1 Then\r\n"
        "        ReDim dataArr(1 To 1, 1 To 1)\r\n"
        "        dataArr(1, 1) = rng.Value2\r\n"
        "    Else\r\n"
        "        dataArr = rng.Value2\r\n"
        "    End If\r\n"
        "    colOffset = firstCol - 1\r\n"
        "    rowOffset = DATA_START_ROW - 1\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    Set rowsToDelete = New Collection\r\n"
        "\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        rowKey = BuildRowKeyOrBlank(dataArr, r - rowOffset, keyCols, colOffset, isBlank)\r\n"
        "        If Not isBlank Then\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                rowsToDelete.Add r\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "\r\n"
        "    For i = 1 To rowsToDelete.Count\r\n"
        "        If delRange Is Nothing Then\r\n"
        "            Set delRange = ws.Rows(CLng(rowsToDelete(i))).EntireRow\r\n"
        "        Else\r\n"
        "            Set delRange = Union(delRange, ws.Rows(CLng(rowsToDelete(i))).EntireRow)\r\n"
        "        End If\r\n"
        "    Next i\r\n"
        "    If Not delRange Is Nothing Then\r\n"
        "        DeleteDuplicateRowsByComment = rowsToDelete.Count\r\n"
        "        delRange.Delete\r\n"
        "    End If\r\n"
        "End Function"
    )
    content = patch_replace(content, old, new, fname, "DeleteDuplicateRowsByComment -> array+union")

    # 2b. Replace BuildRowKeyByColumns + RowIsBlankByColumns with BuildRowKeyOrBlank
    old = (
        "Private Function BuildRowKeyByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As String\r\n"
        "    Dim idx As Variant\r\n"
        "    Dim parts As String\r\n"
        "\r\n"
        "    For Each idx In colIndexes\r\n"
        "        parts = parts & KEY_SEP & NormalizeText(ws.Cells(rowIndex, CLng(idx)).Value2)\r\n"
        "    Next idx\r\n"
        "    BuildRowKeyByColumns = parts\r\n"
        "End Function\r\n"
        "\r\n"
        "Private Function RowIsBlankByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As Boolean\r\n"
        "    Dim idx As Variant\r\n"
        "    For Each idx In colIndexes\r\n"
        "        If Len(NormalizeText(ws.Cells(rowIndex, CLng(idx)).Value2)) > 0 Then\r\n"
        "            Exit Function\r\n"
        "        End If\r\n"
        "    Next idx\r\n"
        "    RowIsBlankByColumns = True\r\n"
        "End Function"
    )
    new = (
        "Private Function BuildRowKeyOrBlank(ByVal dataArr As Variant, ByVal arrRow As Long, _\r\n"
        "                                    ByVal colIndexes As Collection, ByVal colOffset As Long, _\r\n"
        "                                    ByRef isBlank As Boolean) As String\r\n"
        "    Dim idx As Variant\r\n"
        "    Dim parts As String\r\n"
        "    Dim cellText As String\r\n"
        "\r\n"
        "    isBlank = True\r\n"
        "    For Each idx In colIndexes\r\n"
        "        cellText = NormalizeText(dataArr(arrRow, CLng(idx) - colOffset))\r\n"
        "        parts = parts & KEY_SEP & cellText\r\n"
        "        If Len(cellText) > 0 Then isBlank = False\r\n"
        "    Next idx\r\n"
        "    If Not isBlank Then BuildRowKeyOrBlank = parts\r\n"
        "End Function"
    )
    content = patch_replace(content, old, new, fname, "BuildRowKey+RowIsBlank -> BuildRowKeyOrBlank")

    write_file(path, content)
    print(f"OK: {fname}")


# ============================================================
# Generic patch for config-based modules (3dot12, 3dot13, 3dot15)
# They share identical helper function signatures
# ============================================================

# BuildRowKeyByColumns + RowIsBlankByColumns -> BuildRowKeyOrBlank
CONFIG_OLD_HELPERS = (
    "Private Function BuildRowKeyByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As String\r\n"
    "    Dim idx As Variant\r\n"
    "    Dim parts As String\r\n"
    "\r\n"
    "    For Each idx In colIndexes\r\n"
    "        parts = parts & KEY_SEP & NormalizeText(ws.Cells(rowIndex, CLng(idx)).Value2)\r\n"
    "    Next idx\r\n"
    "    BuildRowKeyByColumns = parts\r\n"
    "End Function"
)

CONFIG_OLD_BLANK = (
    "Private Function RowIsBlankByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As Boolean\r\n"
    "    Dim idx As Variant\r\n"
    "    For Each idx In colIndexes\r\n"
    "        If Len(NormalizeText(ws.Cells(rowIndex, CLng(idx)).Value2)) > 0 Then\r\n"
    "            Exit Function\r\n"
    "        End If\r\n"
    "    Next idx\r\n"
    "    RowIsBlankByColumns = True\r\n"
    "End Function"
)

CONFIG_NEW_COMBINED = (
    "Private Function BuildRowKeyOrBlank(ByVal dataArr As Variant, ByVal arrRow As Long, _\r\n"
    "                                    ByVal colIndexes As Collection, ByVal colOffset As Long, _\r\n"
    "                                    ByRef isBlank As Boolean) As String\r\n"
    "    Dim idx As Variant\r\n"
    "    Dim parts As String\r\n"
    "    Dim cellText As String\r\n"
    "\r\n"
    "    isBlank = True\r\n"
    "    For Each idx In colIndexes\r\n"
    "        cellText = NormalizeText(dataArr(arrRow, CLng(idx) - colOffset))\r\n"
    "        parts = parts & KEY_SEP & cellText\r\n"
    "        If Len(cellText) > 0 Then isBlank = False\r\n"
    "    Next idx\r\n"
    "    If Not isBlank Then BuildRowKeyOrBlank = parts\r\n"
    "End Function"
)


# ============================================================
# Module 3: 功能3dot12按配置检查重复数据.bas
# Optimization 1 + 3 (no delete, so no Union)
# ============================================================
def patch_3dot12():
    fname = "功能3dot12按配置检查重复数据.bas"
    path = os.path.join(MODULES_DIR, fname)
    content = read_file(path)

    # 3a. Modify MarkDuplicateRowsByIndexes: array read + combined blank/key
    old = (
        "Private Function MarkDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "\r\n"
        "    If dedupeCols Is Nothing Then Exit Function\r\n"
        "    If dedupeCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        If Not RowIsBlankByColumns(ws, r, dedupeCols) Then\r\n"
        "            rowKey = BuildRowKeyByColumns(ws, r, dedupeCols)\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                MarkDuplicateRow ws, r, firstCol, lastCol\r\n"
        "                MarkDuplicateRowsByIndexes = MarkDuplicateRowsByIndexes + 1\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "End Function"
    )
    new = (
        "Private Function MarkDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim dataArr As Variant\r\n"
        "    Dim rng As Range\r\n"
        "    Dim colOffset As Long\r\n"
        "    Dim rowOffset As Long\r\n"
        "    Dim isBlank As Boolean\r\n"
        "\r\n"
        "    If dedupeCols Is Nothing Then Exit Function\r\n"
        "    If dedupeCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set rng = ws.Range(ws.Cells(DATA_START_ROW, firstCol), ws.Cells(lastRow, lastCol))\r\n"
        "    If rng.Cells.CountLarge = 1 Then\r\n"
        "        ReDim dataArr(1 To 1, 1 To 1)\r\n"
        "        dataArr(1, 1) = rng.Value2\r\n"
        "    Else\r\n"
        "        dataArr = rng.Value2\r\n"
        "    End If\r\n"
        "    colOffset = firstCol - 1\r\n"
        "    rowOffset = DATA_START_ROW - 1\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        rowKey = BuildRowKeyOrBlank(dataArr, r - rowOffset, dedupeCols, colOffset, isBlank)\r\n"
        "        If Not isBlank Then\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                MarkDuplicateRow ws, r, firstCol, lastCol\r\n"
        "                MarkDuplicateRowsByIndexes = MarkDuplicateRowsByIndexes + 1\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "End Function"
    )
    content = patch_replace(content, old, new, fname, "MarkDuplicateRowsByIndexes -> array read")

    # 3b. Replace helpers
    content = patch_replace(content, CONFIG_OLD_HELPERS, "", fname, "remove old BuildRowKeyByColumns")
    content = patch_replace(content, CONFIG_OLD_BLANK, CONFIG_NEW_COMBINED, fname, "RowIsBlankByColumns -> BuildRowKeyOrBlank")

    write_file(path, content)
    print(f"OK: {fname}")


# ============================================================
# Module 4: 功能3dot13按配置删除重复行.bas
# All 3 optimizations
# ============================================================
def patch_3dot13():
    fname = "功能3dot13按配置删除重复行.bas"
    path = os.path.join(MODULES_DIR, fname)
    content = read_file(path)

    # 4a. Modify DeleteDuplicateRowsByIndexes: array read + combined blank/key + Union delete
    old = (
        "Private Function DeleteDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long\r\n"
        "    Dim seen As Object\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim rowsToDelete As Collection\r\n"
        "    Dim i As Long\r\n"
        "\r\n"
        "    If dedupeCols Is Nothing Then Exit Function\r\n"
        "    If dedupeCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    Set rowsToDelete = New Collection\r\n"
        "\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        If Not RowIsBlankByColumns(ws, r, dedupeCols) Then\r\n"
        "            rowKey = BuildRowKeyByColumns(ws, r, dedupeCols)\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                rowsToDelete.Add r\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "\r\n"
        "    For i = rowsToDelete.Count To 1 Step -1\r\n"
        "        ws.Rows(CLng(rowsToDelete(i))).EntireRow.Delete\r\n"
        "        DeleteDuplicateRowsByIndexes = DeleteDuplicateRowsByIndexes + 1\r\n"
        "    Next i\r\n"
        "End Function"
    )
    new = (
        "Private Function DeleteDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim rowsToDelete As Collection\r\n"
        "    Dim i As Long\r\n"
        "    Dim dataArr As Variant\r\n"
        "    Dim rng As Range\r\n"
        "    Dim colOffset As Long\r\n"
        "    Dim rowOffset As Long\r\n"
        "    Dim isBlank As Boolean\r\n"
        "    Dim delRange As Range\r\n"
        "\r\n"
        "    If dedupeCols Is Nothing Then Exit Function\r\n"
        "    If dedupeCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set rng = ws.Range(ws.Cells(DATA_START_ROW, firstCol), ws.Cells(lastRow, lastCol))\r\n"
        "    If rng.Cells.CountLarge = 1 Then\r\n"
        "        ReDim dataArr(1 To 1, 1 To 1)\r\n"
        "        dataArr(1, 1) = rng.Value2\r\n"
        "    Else\r\n"
        "        dataArr = rng.Value2\r\n"
        "    End If\r\n"
        "    colOffset = firstCol - 1\r\n"
        "    rowOffset = DATA_START_ROW - 1\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    Set rowsToDelete = New Collection\r\n"
        "\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        rowKey = BuildRowKeyOrBlank(dataArr, r - rowOffset, dedupeCols, colOffset, isBlank)\r\n"
        "        If Not isBlank Then\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                rowsToDelete.Add r\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "\r\n"
        "    For i = 1 To rowsToDelete.Count\r\n"
        "        If delRange Is Nothing Then\r\n"
        "            Set delRange = ws.Rows(CLng(rowsToDelete(i))).EntireRow\r\n"
        "        Else\r\n"
        "            Set delRange = Union(delRange, ws.Rows(CLng(rowsToDelete(i))).EntireRow)\r\n"
        "        End If\r\n"
        "    Next i\r\n"
        "    If Not delRange Is Nothing Then\r\n"
        "        DeleteDuplicateRowsByIndexes = rowsToDelete.Count\r\n"
        "        delRange.Delete\r\n"
        "    End If\r\n"
        "End Function"
    )
    content = patch_replace(content, old, new, fname, "DeleteDuplicateRowsByIndexes -> array+union")

    # 4b. Replace helpers
    content = patch_replace(content, CONFIG_OLD_BLANK, CONFIG_NEW_COMBINED, fname, "RowIsBlankByColumns -> BuildRowKeyOrBlank")
    content = patch_replace(content, CONFIG_OLD_HELPERS, "", fname, "remove old BuildRowKeyByColumns")

    write_file(path, content)
    print(f"OK: {fname}")


# ============================================================
# Module 5: 功能3dot15按配置去重追加到目标.bas
# All 3 optimizations (same DeleteDuplicateRowsByIndexes pattern)
# ============================================================
def patch_3dot15():
    fname = "功能3dot15按配置去重追加到目标.bas"
    path = os.path.join(MODULES_DIR, fname)
    content = read_file(path)

    # 5a. Same DeleteDuplicateRowsByIndexes patch as 3dot13
    old = (
        "Private Function DeleteDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long\r\n"
        "    Dim seen As Object\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim rowsToDelete As Collection\r\n"
        "    Dim i As Long\r\n"
        "\r\n"
        "    If dedupeCols Is Nothing Then Exit Function\r\n"
        "    If dedupeCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    Set rowsToDelete = New Collection\r\n"
        "\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        If Not RowIsBlankByColumns(ws, r, dedupeCols) Then\r\n"
        "            rowKey = BuildRowKeyByColumns(ws, r, dedupeCols)\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                rowsToDelete.Add r\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "\r\n"
        "    For i = rowsToDelete.Count To 1 Step -1\r\n"
        "        ws.Rows(CLng(rowsToDelete(i))).EntireRow.Delete\r\n"
        "        DeleteDuplicateRowsByIndexes = DeleteDuplicateRowsByIndexes + 1\r\n"
        "    Next i\r\n"
        "End Function"
    )
    new = (
        "Private Function DeleteDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long\r\n"
        "    Dim seen As Object\r\n"
        "    Dim firstCol As Long\r\n"
        "    Dim lastCol As Long\r\n"
        "    Dim lastRow As Long\r\n"
        "    Dim r As Long\r\n"
        "    Dim rowKey As String\r\n"
        "    Dim rowsToDelete As Collection\r\n"
        "    Dim i As Long\r\n"
        "    Dim dataArr As Variant\r\n"
        "    Dim rng As Range\r\n"
        "    Dim colOffset As Long\r\n"
        "    Dim rowOffset As Long\r\n"
        "    Dim isBlank As Boolean\r\n"
        "    Dim delRange As Range\r\n"
        "\r\n"
        "    If dedupeCols Is Nothing Then Exit Function\r\n"
        "    If dedupeCols.Count = 0 Then Exit Function\r\n"
        "\r\n"
        "    firstCol = GetFirstUsedColumn(ws)\r\n"
        "    lastCol = GetLastUsedColumn(ws)\r\n"
        "    lastRow = GetLastUsedRow(ws)\r\n"
        "    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function\r\n"
        "\r\n"
        "    Set rng = ws.Range(ws.Cells(DATA_START_ROW, firstCol), ws.Cells(lastRow, lastCol))\r\n"
        "    If rng.Cells.CountLarge = 1 Then\r\n"
        "        ReDim dataArr(1 To 1, 1 To 1)\r\n"
        "        dataArr(1, 1) = rng.Value2\r\n"
        "    Else\r\n"
        "        dataArr = rng.Value2\r\n"
        "    End If\r\n"
        "    colOffset = firstCol - 1\r\n"
        "    rowOffset = DATA_START_ROW - 1\r\n"
        "\r\n"
        "    Set seen = CreateObject(\"Scripting.Dictionary\")\r\n"
        "    Set rowsToDelete = New Collection\r\n"
        "\r\n"
        "    For r = DATA_START_ROW To lastRow\r\n"
        "        rowKey = BuildRowKeyOrBlank(dataArr, r - rowOffset, dedupeCols, colOffset, isBlank)\r\n"
        "        If Not isBlank Then\r\n"
        "            If seen.Exists(rowKey) Then\r\n"
        "                rowsToDelete.Add r\r\n"
        "            Else\r\n"
        "                seen.Add rowKey, True\r\n"
        "            End If\r\n"
        "        End If\r\n"
        "    Next r\r\n"
        "\r\n"
        "    For i = 1 To rowsToDelete.Count\r\n"
        "        If delRange Is Nothing Then\r\n"
        "            Set delRange = ws.Rows(CLng(rowsToDelete(i))).EntireRow\r\n"
        "        Else\r\n"
        "            Set delRange = Union(delRange, ws.Rows(CLng(rowsToDelete(i))).EntireRow)\r\n"
        "        End If\r\n"
        "    Next i\r\n"
        "    If Not delRange Is Nothing Then\r\n"
        "        DeleteDuplicateRowsByIndexes = rowsToDelete.Count\r\n"
        "        delRange.Delete\r\n"
        "    End If\r\n"
        "End Function"
    )
    content = patch_replace(content, old, new, fname, "DeleteDuplicateRowsByIndexes -> array+union")

    # 5b. Replace helpers
    content = patch_replace(content, CONFIG_OLD_BLANK, CONFIG_NEW_COMBINED, fname, "RowIsBlankByColumns -> BuildRowKeyOrBlank")
    content = patch_replace(content, CONFIG_OLD_HELPERS, "", fname, "remove old BuildRowKeyByColumns")

    write_file(path, content)
    print(f"OK: {fname}")


def main():
    print("Applying performance patches to 5 dedup modules...\n")
    patch_3dot11()
    print()
    patch_3dot14()
    print()
    patch_3dot12()
    print()
    patch_3dot13()
    print()
    patch_3dot15()
    print("\nDone. All patches applied with GBK encoding preserved.")


if __name__ == "__main__":
    main()
