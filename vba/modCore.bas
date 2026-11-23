Attribute VB_Name = "modCore"
Option Explicit

' ============================================================================
'  Module: modCore - Build Re-Auth report (Dashboard + Manager tabs + Raw Data)
'  Realtime step logging and error capture included.
' ============================================================================

Public Function CreateReAuthReport() As Boolean
    Dim srcWb As Workbook, sourceWb As Workbook, raw As Worksheet, outWb As Workbook
    Dim ws As Worksheet, dashboard As Worksheet, rawOut As Worksheet
    Dim headerRow As Long, lastRow As Long, lastCol As Long
    Dim managerCol As Long, statusCol As Long, dueCol As Long, startCol As Long
    Dim r As Long, i As Long, c As Long
    Dim manager As String, statusNorm As String
    Dim managerRows As Object, managerNames As Object
    Dim managerCounts As Object, managerPastDue As Object
    Dim managerSheetNames As Object
    Dim currentData As Variant, arrOut As Variant, arrRaw As Variant
    Dim managers As Variant
    Dim totalCount As Long, pastDue As Long, upcoming As Long, denied As Long, otherStatus As Long
    Dim savePath As String, sourcePath As Variant, sourceName As String, reportStamp As String
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim autoSourcePath As String
    Dim fso As Object, ts As Object
    Dim rowsCol As Collection
    Dim mgrHeaders As Variant, mgrColCount As Long
    Dim failMsg As String

    CreateReAuthReport = False
    mInReportRun = True
    mReportStartTime = Now
    failMsg = ""

    On Error GoTo CleanFail

    mRunLogPath = BuildRunLogPath
    LogMsg "=== CreateReAuthReport START ==="
    LogStep "CreateReAuthReport", "init run log=" & mRunLogPath
    LogMsg "Excel=" & Application.Version & " Template=" & ThisWorkbook.FullName
    LogMsg "Context: version=" & MACRO_VERSION & " user=" & Environ$("USERNAME") & _
            " OS=" & Application.OperatingSystem & " build=" & Application.Build
    LogMsg "Context: project_root=" & GetProjectRoot()
    LogMsg "Context: log_dir=" & GetLogDir() & " reports_dir=" & GetReportsDir()
    LogMsg "Context: auto_source_file=" & AUTO_SOURCE_PATH_FILE() & _
            " exists=" & CStr(Len(Dir$(AUTO_SOURCE_PATH_FILE())) > 0)

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Selecting raw data..."

    Set srcWb = ThisWorkbook
    LogStep "CreateReAuthReport", "source selection"

    autoSourcePath = ""
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FileExists(AUTO_SOURCE_PATH_FILE) Then
            Set ts = fso.OpenTextFile(AUTO_SOURCE_PATH_FILE, 1)
            If Not ts Is Nothing Then
                autoSourcePath = Trim$(ts.ReadAll)
                ts.Close
                Set ts = Nothing
            End If
        End If
    End If
    Set fso = Nothing
    Err.Clear
    On Error GoTo CleanFail

    If Len(autoSourcePath) > 0 Then
        sourcePath = autoSourcePath
        LogMsg "AUTO source path: " & CStr(sourcePath)
    Else
        LogMsg "Prompting for source file"
        sourcePath = SelectExcelFile("Select the Re-Auth raw data file (CSV or Excel)", GetProjectRoot())
    End If

    If VarType(sourcePath) = vbBoolean Then
        LogMsg "CANCELLED: no source selected"
        MsgBox "No source file selected. Report was not created.", vbExclamation
        GoTo CleanExit
    End If

    If StrComp(CStr(sourcePath), srcWb.FullName, vbTextCompare) = 0 Then
        LogMsg "ABORT: selected template as source"
        MsgBox "Please select the raw data file, not the template workbook.", vbExclamation
        GoTo CleanExit
    End If

    LogStep "CreateReAuthReport", "opening source " & CStr(sourcePath)
    Application.StatusBar = "Opening source data..."
    If Len(Dir$(CStr(sourcePath))) = 0 Then
        Err.Raise vbObjectError + 1004, "CreateReAuthReport", "Source file not found: " & CStr(sourcePath)
    End If
    LogMsg "Source file bytes=" & CStr(FileLen(CStr(sourcePath)))
    Set sourceWb = Workbooks.Open(Filename:=CStr(sourcePath), ReadOnly:=True, UpdateLinks:=False)
    If sourceWb Is Nothing Then Err.Raise vbObjectError + 1001, "CreateReAuthReport", "Failed to open source workbook"
    sourceName = sourceWb.Name
    reportStamp = Format(Now, "yyyy-mm-dd hh:nn AM/PM")
    LogMsg "Source opened: " & sourceWb.FullName & " | sheets=" & sourceWb.Worksheets.Count & _
            " | sheet1=" & sourceWb.Worksheets(1).Name

    Set raw = sourceWb.Worksheets(1)
    headerRow = FindHeaderRow(raw)
    If headerRow = 0 Then
        LogMsg "ERROR: required headers not found"
        MsgBox "Could not find required headers in source file. Need at least: " & MANAGER_HEADER & ", " & AUTH_STATUS_HEADER & ", " & PATIENT_HEADER & ", " & REAUTH_DUE_HEADER, vbCritical
        GoTo CleanExit
    End If

    managerCol = FindHeaderColumn(raw, headerRow, MANAGER_HEADER)
    statusCol = FindHeaderColumn(raw, headerRow, AUTH_STATUS_HEADER)
    dueCol = FindHeaderColumn(raw, headerRow, REAUTH_DUE_HEADER)
    startCol = FindHeaderColumn(raw, headerRow, AUTH_START_HEADER)

    lastRow = LastUsedRow(raw)
    lastCol = LastUsedColumn(raw)
    LogMsg "Source parsed: headerRow=" & headerRow & " lastRow=" & lastRow & " lastCol=" & lastCol & _
            " managerCol=" & managerCol & " statusCol=" & statusCol & " dueCol=" & dueCol & " startCol=" & startCol
    LogMsg "Header mapping: manager=" & HeaderState(managerCol) & " status=" & HeaderState(statusCol) & _
            " due=" & HeaderState(dueCol) & " start=" & HeaderState(startCol) & _
            " denial=" & HeaderState(FindHeaderColumn(raw, headerRow, DENIAL_DATE_HEADER))
    If managerCol = 0 Then LogMsg "WARN: header not found '" & MANAGER_HEADER & "' - all rows will be Unassigned"
    If statusCol = 0 Then LogMsg "WARN: header not found '" & AUTH_STATUS_HEADER & "' - status counts will be 0"
    If dueCol = 0 Then LogMsg "WARN: header not found '" & REAUTH_DUE_HEADER & "' - due dates not normalized"
    If startCol = 0 Then LogMsg "WARN: header not found '" & AUTH_START_HEADER & "' - start dates not normalized"

    If lastRow <= headerRow Then
        LogMsg "No data rows found"
        MsgBox "No data rows found in the source file.", vbInformation
        GoTo CleanExit
    End If

    LogStep "CreateReAuthReport", "normalizing dates"
    NormalizeDateColumn raw, headerRow, lastRow, dueCol
    NormalizeDateColumn raw, headerRow, lastRow, startCol
    NormalizeDateColumn raw, headerRow, lastRow, FindHeaderColumn(raw, headerRow, "Initial Authorization Denial Date")

    LogStep "CreateReAuthReport", "loading data array"
    currentData = raw.Range(raw.Cells(headerRow, 1), raw.Cells(lastRow, lastCol)).Value
    LogMsg "Loaded rows=" & UBound(currentData, 1) & " cols=" & UBound(currentData, 2)

    Set managerRows = CreateObject("Scripting.Dictionary")
    managerRows.CompareMode = vbTextCompare
    Set managerNames = CreateObject("Scripting.Dictionary")
    managerNames.CompareMode = vbTextCompare
    Set managerCounts = CreateObject("Scripting.Dictionary")
    managerCounts.CompareMode = vbTextCompare
    Set managerPastDue = CreateObject("Scripting.Dictionary")
    managerPastDue.CompareMode = vbTextCompare
    Set managerSheetNames = CreateObject("Scripting.Dictionary")
    managerSheetNames.CompareMode = vbTextCompare

    totalCount = 0: pastDue = 0: upcoming = 0: denied = 0: otherStatus = 0

    LogStep "CreateReAuthReport", "grouping records"
    For r = 2 To UBound(currentData, 1)
        totalCount = totalCount + 1

        If managerCol > 0 Then
            manager = Trim$(CStrSafe(currentData(r, managerCol)))
        Else
            manager = ""
        End If
        If Len(manager) = 0 Then manager = "Unassigned"

        If Not managerRows.Exists(manager) Then
            Set rowsCol = New Collection
            managerRows.Add manager, rowsCol
            managerNames.Add manager, manager
            managerCounts.Add manager, 0
            managerPastDue.Add manager, 0
        End If
        managerRows(manager).Add r
        managerCounts(manager) = CLng(managerCounts(manager)) + 1

        statusNorm = ""
        If statusCol > 0 Then statusNorm = LCase$(Trim$(CStrSafe(currentData(r, statusCol))))
        Select Case True
            Case InStr(statusNorm, "past due") > 0
                pastDue = pastDue + 1
                managerPastDue(manager) = CLng(managerPastDue(manager)) + 1
            Case InStr(statusNorm, "upcoming") > 0
                upcoming = upcoming + 1
            Case InStr(statusNorm, "denied") > 0
                denied = denied + 1
            Case Else
                otherStatus = otherStatus + 1
        End Select
    Next r

    LogMsg "Counts: Total=" & totalCount & " PastDue=" & pastDue & " Upcoming=" & upcoming & _
            " Denied=" & denied & " Other=" & otherStatus & " Managers=" & managerNames.Count
    LogMsg "Phase grouping elapsed(s)=" & Format(TimerDiff(mReportStartTime, Now), "0.00")

    If totalCount = 0 Then
        LogMsg "No records to report"
        MsgBox "No records found to include in the report.", vbInformation
        GoTo CleanExit
    End If

    LogStep "CreateReAuthReport", "creating output workbook"
    Application.StatusBar = "Creating report..."
    Set outWb = Workbooks.Add(xlWBATWorksheet)
    If outWb Is Nothing Then Err.Raise vbObjectError + 1002, "CreateReAuthReport", "Failed to create output workbook"
    Set dashboard = outWb.Worksheets(1)
    dashboard.Name = DASHBOARD_SHEET

    managers = managerNames.Keys
    If managerNames.Count > 1 Then SortManagersByLastName managers

    mgrHeaders = Array( _
        WCC_HEADER, _
        ASSOC_WCC_HEADER, _
        AUTH_STATUS_HEADER, _
        PATIENT_HEADER, _
        COMPANY_HEADER, _
        REAUTH_DUE_HEADER, _
        AUTH_START_HEADER, _
        AUTHORIZATION_STATUS_HEADER, _
        TASK_HEADER, _
        PACKET_STATUS_HEADER, _
        TREATMENT_PHYSICIAN_HEADER _
    )
    mgrColCount = UBound(mgrHeaders) - LBound(mgrHeaders) + 1
    LogMsg "Manager tab columns=" & mgrColCount & " headerRow=" & MANAGER_HEADER_ROW

    For i = LBound(managers) To UBound(managers)
        manager = CStr(managers(i))
        LogStep "CreateReAuthReport", "manager tab " & manager & " rows=" & managerRows(manager).Count
        Application.StatusBar = "Creating tab: " & manager
        Set ws = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
        ws.Name = SafeSheetName(manager, outWb)
        managerSheetNames(manager) = ws.Name

        arrOut = BuildManagerOutputArray(currentData, SortRowsByHeader(currentData, managerRows(manager), WCC_HEADER), mgrHeaders)
        ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(MANAGER_HEADER_ROW + UBound(arrOut, 1) - 1, mgrColCount)).Value = arrOut
        For c = LBound(mgrHeaders) To UBound(mgrHeaders)
            ws.Cells(MANAGER_HEADER_ROW, c - LBound(mgrHeaders) + 1).Value = CStr(mgrHeaders(c))
        Next c
        FormatManagerSheet ws, manager, managerRows(manager).Count, mgrColCount, MANAGER_HEADER_ROW
        BuildRepTotalsSection ws, MANAGER_HEADER_ROW, mgrColCount
        ApplyDateFormatsByHeader ws, MANAGER_HEADER_ROW, LastUsedRow(ws), REAUTH_DUE_HEADER
        ApplyDateFormatsByHeader ws, MANAGER_HEADER_ROW, LastUsedRow(ws), AUTH_START_HEADER
        HideUnusedCells ws
        LogMsg "Tab done: " & ws.Name
    Next i
    LogMsg "Phase manager tabs elapsed(s)=" & Format(TimerDiff(mReportStartTime, Now), "0.00")

    LogStep "CreateReAuthReport", "building Raw Data"
    Application.StatusBar = "Creating Raw Data..."
    Set rawOut = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
    rawOut.Name = RAW_SHEET
    Dim allRows As New Collection
    For r = 2 To UBound(currentData, 1)
        allRows.Add r
    Next r
    arrRaw = BuildOutputArray(currentData, allRows, lastCol)
    rawOut.Range(rawOut.Cells(1, 1), rawOut.Cells(UBound(arrRaw, 1), lastCol)).Value = arrRaw
    FormatRawData rawOut, lastCol, UBound(arrRaw, 1)
    ApplyDateFormatsByHeader rawOut, 1, UBound(arrRaw, 1), REAUTH_DUE_HEADER
    ApplyDateFormatsByHeader rawOut, 1, UBound(arrRaw, 1), AUTH_START_HEADER
    ApplyDateFormatsByHeader rawOut, 1, UBound(arrRaw, 1), "Initial Authorization Denial Date"
    HideUnusedCells rawOut
    LogMsg "Raw Data done rows=" & (UBound(arrRaw, 1) - 1) & " data (+header)"
    LogMsg "Phase raw data elapsed(s)=" & Format(TimerDiff(mReportStartTime, Now), "0.00")

    LogStep "CreateReAuthReport", "building Dashboard"
    Application.StatusBar = "Building dashboard..."
    FormatDashboard dashboard, reportStamp, sourceName, totalCount, pastDue, upcoming, denied, otherStatus, _
        managers, managerCounts, managerPastDue, managerSheetNames
    HideUnusedCells dashboard
    dashboard.Activate
    LogMsg "Phase dashboard elapsed(s)=" & Format(TimerDiff(mReportStartTime, Now), "0.00")

    savePath = BuildOutputPath(srcWb)
    If Len(savePath) = 0 Then Err.Raise vbObjectError + 1006, "CreateReAuthReport", "BuildOutputPath returned an empty path"
    LogStep "CreateReAuthReport", "saving " & savePath
    Application.StatusBar = "Saving report..."
    outWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    If Len(Dir$(savePath)) = 0 Then
        Err.Raise vbObjectError + 1005, "CreateReAuthReport", "SaveAs reported success but output file not found: " & savePath
    End If
    LogMsg "Saved: " & savePath & " | bytes=" & CStr(FileLen(savePath))
    LogMsg "Phase save elapsed(s)=" & Format(TimerDiff(mReportStartTime, Now), "0.00")

    WriteCompletionMarker savePath
    LogMsg "Run duration(s): " & Format(TimerDiff(mReportStartTime, Now), "0.00")
    LogMsg "=== CreateReAuthReport COMPLETE ==="
    CreateReAuthReport = True
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    mInReportRun = False
    LogMsg "CleanExit: success=" & CStr(CreateReAuthReport) & " step=" & mCurrentStep
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    Exit Function

CleanFail:
    Dim errNum As Long, errDesc As String, errLine As Long
    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl
    failMsg = "CreateReAuthReport failed | Err=" & errNum & " | " & errDesc & " | Erl=" & errLine & _
              " | Step=" & mCurrentStep
    On Error Resume Next
    LogMsg failMsg
    LogErr "CreateReAuthReport", errNum, errDesc, errLine
    mInReportRun = False
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    WriteCompletionMarker "(failed) " & failMsg
    MsgBox "Report creation failed." & vbCrLf & vbCrLf & _
           "Error: " & errNum & " - " & errDesc & vbCrLf & _
           "Line: " & errLine & vbCrLf & _
           "Step: " & mCurrentStep & vbCrLf & vbCrLf & _
           "See log\macro-debug.log and log\reauth-complete.txt for details.", vbCritical
    CreateReAuthReport = False
End Function

Private Function HeaderState(ByVal colIdx As Long) As String
    If colIdx > 0 Then
        HeaderState = "col" & CStr(colIdx)
    Else
        HeaderState = "MISSING"
    End If
End Function

Private Function CStrSafe(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        CStrSafe = ""
    Else
        CStrSafe = CStr(v)
    End If
End Function

Private Sub ApplyDateFormatsByHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal headerName As String)
    Dim colIdx As Long, r As Long
    On Error Resume Next
    colIdx = FindHeaderColumn(ws, headerRow, headerName)
    If colIdx <= 0 Or lastRow <= headerRow Then Exit Sub
    For r = headerRow + 1 To lastRow
        If IsDate(ws.Cells(r, colIdx).Value) Then ws.Cells(r, colIdx).NumberFormat = "m/d/yyyy"
    Next r
End Sub

Private Function TimerDiff(ByVal tStart As Date, ByVal tEnd As Date) As Double
    On Error Resume Next
    TimerDiff = (tEnd - tStart) * 86400#
End Function
