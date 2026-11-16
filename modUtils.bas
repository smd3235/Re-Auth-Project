Attribute VB_Name = "modUtils"
Option Explicit

' ============================================================================
'  Module: modUtils
' ============================================================================

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        LastUsedRow = 0
        Exit Function
    End If
    LastUsedRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function

Public Function LastUsedColumn(ByVal ws As Worksheet) As Long
    On Error Resume Next
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        LastUsedColumn = 0
        Exit Function
    End If
    LastUsedColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function

Public Function FindHeaderRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long, r As Long, hits As Long
    On Error Resume Next
    lastRow = LastUsedRow(ws)
    FindHeaderRow = 0
    For r = 1 To Application.WorksheetFunction.Min(lastRow, 20)
        hits = 0
        If FindHeaderColumn(ws, r, MANAGER_HEADER) > 0 Then hits = hits + 1
        If FindHeaderColumn(ws, r, AUTH_STATUS_HEADER) > 0 Then hits = hits + 1
        If FindHeaderColumn(ws, r, PATIENT_HEADER) > 0 Then hits = hits + 1
        If FindHeaderColumn(ws, r, REAUTH_DUE_HEADER) > 0 Then hits = hits + 1
        If hits >= 3 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim c As Long, lastCol As Long, h As String
    On Error Resume Next
    FindHeaderColumn = 0
    lastCol = LastUsedColumn(ws)
    For c = 1 To lastCol
        h = LCase$(Trim$(CStr(ws.Cells(headerRow, c).Value)))
        If h = LCase$(Trim$(headerText)) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Public Function FindColumnByHeader(ByVal headers As Variant, ByVal headerText As String) As Long
    Dim c As Long
    On Error Resume Next
    FindColumnByHeader = 0
    If Not IsArray(headers) Then Exit Function
    For c = LBound(headers) To UBound(headers)
        If LCase$(Trim$(CStr(headers(c)))) = LCase$(Trim$(headerText)) Then
            FindColumnByHeader = c
            Exit Function
        End If
    Next c
End Function

Public Function TryGetExcelDate(ByVal v As Variant, ByRef dt As Date) As Boolean
    On Error Resume Next
    TryGetExcelDate = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsDate(v) Then
        dt = CDate(v)
        TryGetExcelDate = True
        Exit Function
    End If
    If IsNumeric(v) Then
        dt = CDbl(v)
        TryGetExcelDate = True
        Exit Function
    End If
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    dt = CDate(CStr(v))
    If Err.Number = 0 Then TryGetExcelDate = True
    Err.Clear
End Function

Public Sub NormalizeDateColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal col As Long)
    Dim r As Long, v As Variant, dt As Date
    On Error Resume Next
    If col <= 0 Or headerRow <= 0 Or lastRow <= headerRow Then Exit Sub
    For r = headerRow + 1 To lastRow
        v = ws.Cells(r, col).Value
        If TryGetExcelDate(v, dt) Then
            ws.Cells(r, col).Value = dt
            ws.Cells(r, col).NumberFormat = "m/d/yyyy"
        End If
    Next r
End Sub

Public Function SafeSheetName(ByVal name As String, ByVal wb As Workbook) As String
    Dim s As String, i As Long, base As String, candidate As String
    s = Trim$(name)
    If Len(s) = 0 Then s = "Unassigned"
    s = Replace(s, "\", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "?", "_")
    s = Replace(s, "*", "_")
    s = Replace(s, "[", "(")
    s = Replace(s, "]", ")")
    s = Replace(s, ":", "-")
    If Len(s) > 31 Then s = Left$(s, 31)
    base = s
    candidate = base
    i = 1
    Do While SheetExists(wb, candidate)
        candidate = Left$(base, 31 - Len(CStr(i)) - 1) & "_" & CStr(i)
        i = i + 1
    Loop
    SafeSheetName = candidate
End Function

Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Public Sub SortManagersByLastName(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As String
    If Not IsArray(arr) Then Exit Sub
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If LCase$(LastNameKey(CStr(arr(i)))) > LCase$(LastNameKey(CStr(arr(j)))) Then
                tmp = CStr(arr(i)): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function LastNameKey(ByVal fullName As String) As String
    Dim parts() As String, p As String
    p = Trim$(fullName)
    If Len(p) = 0 Then LastNameKey = "zzzz": Exit Function
    parts = Split(p, " ")
    LastNameKey = parts(UBound(parts)) & ", " & p
End Function

Public Function BuildOutputArray(ByVal data As Variant, ByVal rows As Collection, ByVal lastCol As Long) As Variant
    Dim arr() As Variant, c As Long, i As Long, r As Long, item As Variant
    ReDim arr(1 To rows.Count + 1, 1 To lastCol)
    For c = 1 To lastCol
        arr(1, c) = data(1, c)
    Next c
    i = 2
    For Each item In rows
        r = CLng(item)
        For c = 1 To lastCol
            arr(i, c) = data(r, c)
        Next c
        i = i + 1
    Next item
    BuildOutputArray = arr
End Function

Public Function SortRowsByHeader(ByVal data As Variant, ByVal rows As Collection, ByVal headerName As String) As Collection
    Dim arr() As Variant
    Dim i As Long, j As Long, n As Long
    Dim srcCol As Long, c As Long
    Dim tmpRow As Long, tmpKey As String
    Dim keys() As String
    Dim result As New Collection

    On Error Resume Next
    Set SortRowsByHeader = rows
    n = rows.Count
    If n <= 1 Then Exit Function
    If Not IsArray(data) Then Exit Function

    srcCol = 0
    For c = 1 To UBound(data, 2)
        If LCase$(Trim$(CStr(data(1, c)))) = LCase$(Trim$(headerName)) Then
            srcCol = c
            Exit For
        End If
    Next c
    If srcCol = 0 Then Exit Function

    ReDim arr(1 To n)
    ReDim keys(1 To n)
    For i = 1 To n
        arr(i) = CLng(rows(i))
        keys(i) = LCase$(Trim$(CStr(data(arr(i), srcCol))))
    Next i

    ' simple insertion sort (stable enough for typical tab sizes)
    For i = 2 To n
        tmpRow = arr(i)
        tmpKey = keys(i)
        j = i - 1
        Do While j >= 1
            If StrComp(keys(j), tmpKey, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        arr(j + 1) = tmpRow
        keys(j + 1) = tmpKey
    Next i

    For i = 1 To n
        result.Add arr(i)
    Next i
    Set SortRowsByHeader = result
End Function

Public Function BuildManagerOutputArray(ByVal data As Variant, ByVal rows As Collection, ByVal orderedHeaders As Variant) As Variant
    Dim arr() As Variant
    Dim srcCols() As Long
    Dim c As Long, outC As Long, i As Long, r As Long, srcC As Long
    Dim item As Variant, h As String

    ReDim srcCols(LBound(orderedHeaders) To UBound(orderedHeaders))
    For c = LBound(orderedHeaders) To UBound(orderedHeaders)
        h = CStr(orderedHeaders(c))
        srcC = 0
        For outC = 1 To UBound(data, 2)
            If LCase$(Trim$(CStr(data(1, outC)))) = LCase$(Trim$(h)) Then
                srcC = outC
                Exit For
            End If
        Next outC
        srcCols(c) = srcC
    Next c

    ReDim arr(1 To rows.Count + 1, 1 To UBound(orderedHeaders) - LBound(orderedHeaders) + 1)

    outC = 1
    For c = LBound(orderedHeaders) To UBound(orderedHeaders)
        arr(1, outC) = CStr(orderedHeaders(c))
        outC = outC + 1
    Next c

    i = 2
    For Each item In rows
        r = CLng(item)
        outC = 1
        For c = LBound(orderedHeaders) To UBound(orderedHeaders)
            srcC = srcCols(c)
            If srcC > 0 Then
                arr(i, outC) = data(r, srcC)
            Else
                arr(i, outC) = ""
            End If
            outC = outC + 1
        Next c
        i = i + 1
    Next item

    BuildManagerOutputArray = arr
End Function

Public Sub HideUnusedCells(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    On Error Resume Next
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    If lastRow > 0 And lastRow < ws.Rows.Count Then ws.Rows((lastRow + 1) & ":" & ws.Rows.Count).Hidden = True
    If lastCol > 0 And lastCol < ws.Columns.Count Then ws.Columns(lastCol + 1).Resize(, ws.Columns.Count - lastCol).Hidden = True
End Sub

Public Function BuildOutputPath(ByVal srcWb As Workbook) As String
    Dim reportsDir As String, stamp As String
    On Error Resume Next
    reportsDir = GetReportsDir()
    If Len(reportsDir) = 0 Then reportsDir = GetProjectRoot() & "\Reports\"
    stamp = Format(Now, "yyyy-mm-dd_HHMM")
    BuildOutputPath = reportsDir & "ReAuth_Report_" & stamp & ".xlsx"
End Function

Public Sub ApplyUniformColumnWidths(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim c As Long, w As Double, fitRow As Long
    On Error Resume Next
    fitRow = LastUsedRow(ws)
    If fitRow < 1 Then fitRow = 200
    If fitRow > 5000 Then fitRow = 5000
    For c = firstCol To lastCol
        ws.Columns(c).ColumnWidth = 18
    Next c
    ws.Range(ws.Cells(1, firstCol), ws.Cells(fitRow, lastCol)).Columns.AutoFit
    For c = firstCol To lastCol
        w = ws.Columns(c).ColumnWidth
        If w < 12 Or w = 0 Then w = 12
        If w > 42 Then w = 42
        ws.Columns(c).ColumnWidth = w
    Next c
End Sub


' ----------------------------------------------------------------------------
' Portable runtime paths (avoid hardcoded machine-specific locations)
' ----------------------------------------------------------------------------
Public Function GetProjectRoot() As String
    On Error Resume Next
    GetProjectRoot = ThisWorkbook.Path
    If Len(GetProjectRoot) = 0 Then GetProjectRoot = CurDir$
    If Len(GetProjectRoot) = 0 Then GetProjectRoot = Environ$("TEMP")
End Function

Public Function GetLogDir() As String
    Dim p As String
    On Error Resume Next
    p = GetProjectRoot() & "\log\"
    EnsureFolderExists p
    GetLogDir = p
End Function

Public Function GetReportsDir() As String
    Dim p As String
    On Error Resume Next
    p = GetProjectRoot() & "\Reports\"
    EnsureFolderExists p
    GetReportsDir = p
End Function

Public Function LOG_PATH() As String
    LOG_PATH = GetLogDir() & LOG_FILE_NAME
End Function

Public Function COMPLETION_MARKER_PATH() As String
    COMPLETION_MARKER_PATH = GetLogDir() & COMPLETION_MARKER_NAME
End Function

Public Function AUTO_SOURCE_PATH_FILE() As String
    AUTO_SOURCE_PATH_FILE = GetLogDir() & AUTO_SOURCE_FILE_NAME
End Function

Public Function AUTO_SKIP_PRIOR_FILE() As String
    AUTO_SKIP_PRIOR_FILE = GetLogDir() & AUTO_SKIP_PRIOR_FILE_NAME
End Function

Public Sub EnsureFolderExists(ByVal folderPath As String)
    On Error Resume Next
    If Len(folderPath) = 0 Then Exit Sub
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        ' try simple MkDir chain for one level deep common case
        Dim root As String
        root = Left$(folderPath, Len(folderPath) - 1)
        root = Left$(root, InStrRev(root, "\"))
        If Len(root) > 0 Then
            If Len(Dir$(root, vbDirectory)) = 0 Then MkDir root
        End If
        MkDir folderPath
    End If
End Sub

Public Function SelectExcelFile(ByVal title As String, ByVal initialDir As String) As Variant
    Dim fd As FileDialog
    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If fd Is Nothing Then
        SelectExcelFile = False
        Exit Function
    End If
    With fd
        .title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel/CSV Files", "*.xlsx;*.xlsm;*.xls;*.csv"
        If Len(initialDir) > 0 Then .InitialFileName = initialDir & "\"
        If .Show = -1 Then
            SelectExcelFile = .SelectedItems(1)
        Else
            SelectExcelFile = False
        End If
    End With
End Function
