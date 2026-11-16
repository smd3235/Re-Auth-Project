Attribute VB_Name = "modFormatting"
Option Explicit

' ============================================================================
'  Module: modFormatting - visual styling with defensive error logging
' ============================================================================

Public Sub FormatDashboard(ByVal ws As Worksheet, ByVal reportStamp As String, ByVal sourceName As String, _
    ByVal totalCount As Long, ByVal pastDue As Long, ByVal upcoming As Long, ByVal denied As Long, ByVal otherStatus As Long, _
    ByVal managers As Variant, ByVal managerCounts As Object, ByVal managerPastDue As Object, ByVal managerSheetNames As Object)

    Dim i As Long, rowPtr As Long, mgr As String
    Dim rng As Range
    Dim lastDataRow As Long
    Dim pct As Double

    On Error GoTo EH
    LogMsg "FormatDashboard start"
    ws.Cells.Clear
    ws.Name = DASHBOARD_SHEET

    ws.Columns("A:J").ColumnWidth = 12
    ws.Columns("A").ColumnWidth = 30

    Set rng = ws.Range("A1:J1")
    rng.Merge
    rng.Value = "RE-AUTHORIZATION DASHBOARD"
    With rng
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 34

    Set rng = ws.Range("A2:J2")
    rng.Merge
    rng.Value = "Generated " & reportStamp & "   |   Source: " & sourceName
    With rng
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 20
    ws.Rows(3).RowHeight = 8

    AddKpiCard ws, "A4:B4", "A5:B5", "TOTAL RE-AUTHS", CStr(totalCount), RGB(68, 114, 196), RGB(255, 255, 255)
    AddKpiCard ws, "C4:D4", "C5:D5", "PAST DUE", CStr(pastDue), RGB(192, 0, 0), RGB(255, 255, 255)
    AddKpiCard ws, "E4:F4", "E5:F5", "UPCOMING", CStr(upcoming), RGB(56, 142, 60), RGB(255, 255, 255)
    AddKpiCard ws, "G4:H4", "G5:H5", "DENIED", CStr(denied), RGB(237, 125, 49), RGB(255, 255, 255)
    AddKpiCard ws, "I4:J4", "I5:J5", "OTHER", CStr(otherStatus), RGB(128, 128, 128), RGB(255, 255, 255)

    ws.Rows(4).RowHeight = 22
    ws.Rows(5).RowHeight = 36
    ws.Rows(6).RowHeight = 10

    ' Manager summary - full width A:J aligned with KPI grid
    Set rng = ws.Range("A7:J7")
    rng.Merge
    rng.Value = "  MANAGER BREAKDOWN"
    With rng
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(7).RowHeight = 24

    ws.Range("A8:D8").Merge
    ws.Range("A8").Value = "Assigned Sales Manager"
    ws.Range("E8").Value = "Total"
    ws.Range("F8").Value = "Past Due"
    ws.Range("G8:H8").Merge
    ws.Range("G8").Value = "% Past Due"
    ws.Range("I8:J8").Merge
    ws.Range("I8").Value = "Tab"
    With ws.Range("A8:J8")
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(8).RowHeight = 24

    rowPtr = 9
    If IsArray(managers) Then
        For i = LBound(managers) To UBound(managers)
            mgr = CStr(managers(i))
            ws.Range("A" & rowPtr & ":D" & rowPtr).Merge
            ws.Cells(rowPtr, 1).Value = mgr
            ws.Cells(rowPtr, 5).Value = CLng(managerCounts(mgr))
            ws.Cells(rowPtr, 6).Value = CLng(managerPastDue(mgr))
            ws.Range("G" & rowPtr & ":H" & rowPtr).Merge
            If CLng(managerCounts(mgr)) > 0 Then
                pct = CDbl(managerPastDue(mgr)) / CDbl(managerCounts(mgr))
            Else
                pct = 0
            End If
            ws.Cells(rowPtr, 7).Value = pct
            ws.Cells(rowPtr, 7).NumberFormat = "0.0%"
            ws.Range("I" & rowPtr & ":J" & rowPtr).Merge
            ws.Cells(rowPtr, 9).Value = ""

            If Not managerSheetNames Is Nothing Then
                If managerSheetNames.Exists(mgr) Then
                    On Error Resume Next
                    ws.Hyperlinks.Add Anchor:=ws.Cells(rowPtr, 9), Address:="", SubAddress:="'" & CStr(managerSheetNames(mgr)) & "'!A1", TextToDisplay:="Open"
                    On Error GoTo EH
                    ws.Cells(rowPtr, 9).Font.Bold = True
                End If
            End If

            With ws.Range(ws.Cells(rowPtr, 1), ws.Cells(rowPtr, 10))
                .Font.Name = "Segoe UI"
                .Font.Size = 10
                .Borders.LineStyle = xlContinuous
                .VerticalAlignment = xlCenter
            End With
            ws.Cells(rowPtr, 1).HorizontalAlignment = xlLeft
            ws.Cells(rowPtr, 5).HorizontalAlignment = xlCenter
            ws.Cells(rowPtr, 6).HorizontalAlignment = xlCenter
            ws.Cells(rowPtr, 7).HorizontalAlignment = xlCenter
            ws.Cells(rowPtr, 9).HorizontalAlignment = xlCenter
            ws.Rows(rowPtr).RowHeight = 21

            If rowPtr Mod 2 = 1 Then
                ws.Range(ws.Cells(rowPtr, 1), ws.Cells(rowPtr, 10)).Interior.Color = RGB(242, 247, 255)
            Else
                ws.Range(ws.Cells(rowPtr, 1), ws.Cells(rowPtr, 10)).Interior.Color = RGB(255, 255, 255)
            End If

            If pct >= 0.5 Then
                ws.Cells(rowPtr, 7).Interior.Color = RGB(255, 199, 206)
                ws.Cells(rowPtr, 7).Font.Color = RGB(156, 0, 6)
                ws.Cells(rowPtr, 7).Font.Bold = True
            ElseIf pct >= 0.25 Then
                ws.Cells(rowPtr, 7).Interior.Color = RGB(255, 235, 156)
                ws.Cells(rowPtr, 7).Font.Color = RGB(156, 101, 0)
                ws.Cells(rowPtr, 7).Font.Bold = True
            End If
            If CLng(managerPastDue(mgr)) > 0 Then
                ws.Cells(rowPtr, 6).Font.Color = RGB(192, 0, 0)
                ws.Cells(rowPtr, 6).Font.Bold = True
            End If

            rowPtr = rowPtr + 1
        Next i
    End If

    lastDataRow = Application.WorksheetFunction.Max(8, rowPtr - 1)
    On Error Resume Next
    ws.Range("A8:J" & lastDataRow).AutoFilter
    On Error GoTo EH

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 10
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 10
    ws.Columns("J").ColumnWidth = 10

    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 8
        .FreezePanes = True
    End With
    ws.Range("A1").Select
    LogMsg "FormatDashboard done"
    Exit Sub
EH:
    LogErr "FormatDashboard"
    Resume Next
End Sub

Private Sub AddKpiCard(ByVal ws As Worksheet, ByVal labelAddr As String, ByVal valueAddr As String, _
    ByVal labelText As String, ByVal valueText As String, ByVal fillColor As Long, ByVal fontColor As Long)
    On Error Resume Next
    With ws.Range(labelAddr)
        .Merge
        .Value = labelText
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = fontColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.Color = fillColor
    End With
    With ws.Range(valueAddr)
        .Merge
        .Value = valueText
        .Font.Name = "Segoe UI"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = fontColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = fillColor
    End With
    If Err.Number <> 0 Then LogErr "AddKpiCard"
    Err.Clear
End Sub

Public Sub FormatManagerSheet(ByVal ws As Worksheet, ByVal managerName As String, ByVal totalRows As Long, _
    ByVal lastCol As Long, ByVal headerRow As Long)

    Dim dataLastRow As Long, r As Long, c As Long
    Dim statusCol As Long, authCol As Long
    Dim statusVal As String

    On Error GoTo EH
    LogMsg "FormatManagerSheet start: " & managerName & " lastCol=" & lastCol & " headerRow=" & headerRow
    dataLastRow = LastUsedRow(ws)

    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Merge
    With ws.Range("A1")
        .Value = "RE-AUTH  |  " & UCase$(managerName)
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 17
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(1).RowHeight = 28

    ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol)).Merge
    With ws.Range("A2")
        .Value = "Total re-auths: " & CStr(totalRows)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(2).RowHeight = 18

    ws.Range(ws.Cells(3, 1), ws.Cells(3, lastCol)).Merge
    With ws.Range("A3")
        .Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:nn AM/PM")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(3).RowHeight = 16
    ws.Rows(4).RowHeight = 5
    ws.Rows("5:9").RowHeight = 18

    ' Header row
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .Borders.LineStyle = xlContinuous
        .AutoFilter
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(headerRow).RowHeight = 32

    ' Data
    If dataLastRow > headerRow Then
        With ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(dataLastRow, lastCol))
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With
        For r = headerRow + 1 To dataLastRow
            If (r - headerRow) Mod 2 = 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(242, 247, 255)
            Else
                ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 255, 255)
            End If
            ws.Rows(r).RowHeight = 19
        Next r
    End If

    ' Status highlighting
    statusCol = 0: authCol = 0
    For c = 1 To lastCol
        If LCase$(Trim$(CStr(ws.Cells(headerRow, c).Value))) = LCase$(AUTH_STATUS_HEADER) Then statusCol = c
        If LCase$(Trim$(CStr(ws.Cells(headerRow, c).Value))) = LCase$(AUTHORIZATION_STATUS_HEADER) Then authCol = c
    Next c

    If statusCol > 0 And dataLastRow > headerRow Then
        For r = headerRow + 1 To dataLastRow
            statusVal = LCase$(Trim$(CStr(ws.Cells(r, statusCol).Value)))
            If InStr(statusVal, "past due") > 0 Then
                ws.Cells(r, statusCol).Interior.Color = RGB(255, 199, 206)
                ws.Cells(r, statusCol).Font.Color = RGB(156, 0, 6)
                ws.Cells(r, statusCol).Font.Bold = True
            ElseIf InStr(statusVal, "upcoming") > 0 Then
                ws.Cells(r, statusCol).Interior.Color = RGB(198, 239, 206)
                ws.Cells(r, statusCol).Font.Color = RGB(0, 97, 0)
                ws.Cells(r, statusCol).Font.Bold = True
            ElseIf InStr(statusVal, "denied") > 0 Then
                ws.Cells(r, statusCol).Interior.Color = RGB(255, 235, 156)
                ws.Cells(r, statusCol).Font.Color = RGB(156, 101, 0)
                ws.Cells(r, statusCol).Font.Bold = True
            End If
        Next r
    End If
    If authCol > 0 And dataLastRow > headerRow Then
        For r = headerRow + 1 To dataLastRow
            statusVal = LCase$(Trim$(CStr(ws.Cells(r, authCol).Value)))
            If InStr(statusVal, "denied") > 0 Then
                ws.Cells(r, authCol).Font.Color = RGB(156, 101, 0)
                ws.Cells(r, authCol).Font.Bold = True
            End If
        Next r
    End If

    LogMsg "FormatManagerSheet widths: " & managerName
    ApplyUniformColumnWidths ws, 1, lastCol

    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = headerRow
        .FreezePanes = True
    End With
    ws.Range("A1").Select
    LogMsg "FormatManagerSheet done: " & managerName
    Exit Sub
EH:
    LogErr "FormatManagerSheet:" & managerName
    Err.Clear
    Resume Next
End Sub

Public Sub BuildRepTotalsSection(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastCol As Long)
    Dim repCounts As Object
    Dim repOrder As Collection
    Dim r As Long, i As Long, repColIdx As Long, lastRow As Long
    Dim repName As String, parts As String, summaryText As String
    Dim blockRg As Range, repKey As Variant
    Dim useWidth As Long

    On Error GoTo EH
    LogMsg "BuildRepTotalsSection start"
    useWidth = lastCol
    If useWidth < 4 Then useWidth = 4

    repColIdx = 0
    For i = 1 To lastCol
        If LCase$(Trim$(CStr(ws.Cells(headerRow, i).Value))) = LCase$(WCC_HEADER) Then
            repColIdx = i
            Exit For
        End If
    Next i
    If repColIdx = 0 Then repColIdx = 1

    Set repCounts = CreateObject("Scripting.Dictionary")
    repCounts.CompareMode = vbTextCompare
    Set repOrder = New Collection

    lastRow = LastUsedRow(ws)
    If lastRow >= headerRow + 1 Then
        For r = headerRow + 1 To lastRow
            repName = Trim$(CStr(ws.Cells(r, repColIdx).Value))
            If Len(repName) = 0 Then repName = "Unassigned"
            If Not repCounts.Exists(repName) Then
                repCounts.Add repName, 1
                repOrder.Add repName
            Else
                repCounts(repName) = CLng(repCounts(repName)) + 1
            End If
        Next r
    End If

    parts = ""
    For i = 1 To repOrder.Count
        repKey = repOrder(i)
        If Len(parts) > 0 Then parts = parts & "     |     "
        parts = parts & CStr(repKey) & ": " & CStr(repCounts(repKey))
    Next i
    If Len(parts) = 0 Then parts = "No sales rep assignments"

    summaryText = "RE-AUTHS BY SALES REP" & vbCrLf & parts

    ws.Rows("5:9").RowHeight = 18
    Set blockRg = ws.Range(ws.Cells(5, 1), ws.Cells(9, useWidth))
    blockRg.UnMerge
    blockRg.Merge
    blockRg.Value = summaryText
    With blockRg
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
        .IndentLevel = 1
    End With
    LogMsg "BuildRepTotalsSection done"
    Exit Sub
EH:
    LogErr "BuildRepTotalsSection"
    Err.Clear
End Sub

Public Sub FormatRawData(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal lastRow As Long)
    Dim r As Long
    On Error GoTo EH
    LogMsg "FormatRawData start rows=" & lastRow & " cols=" & lastCol
    If lastRow < 1 Then Exit Sub
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .Borders.LineStyle = xlContinuous
        .AutoFilter
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30
    If lastRow > 1 Then
        With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = False
        End With
        For r = 2 To lastRow
            If r Mod 2 = 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(242, 247, 255)
            End If
        Next r
        ws.Range(ws.Rows(2), ws.Rows(lastRow)).RowHeight = 18
    End If
    ApplyUniformColumnWidths ws, 1, lastCol

    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    ws.Range("A1").Select
    LogMsg "FormatRawData done"
    Exit Sub
EH:
    LogErr "FormatRawData"
    Err.Clear
End Sub
