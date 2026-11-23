Attribute VB_Name = "modLogging"
Option Explicit

' ============================================================================
'  Module: modLogging - centralized realtime debug/error logging
'  Writes to LOG_PATH (always) and to per-run mRunLogPath during report runs.
' ============================================================================

Public Sub LogMsg(ByVal msg As String)
    Dim lineOut As String
    On Error Resume Next
    lineOut = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
    AppendLogLine LOG_PATH, lineOut
    If mInReportRun And Len(mRunLogPath) > 0 Then AppendLogLine mRunLogPath, lineOut
End Sub

Private Sub AppendLogLine(ByVal filePath As String, ByVal lineOut As String)
    Dim f As Integer
    On Error Resume Next
    EnsureDirForFile filePath
    f = FreeFile
    Open filePath For Append As #f
    If Err.Number = 0 Then
        Print #f, lineOut
        Close #f
    Else
        Debug.Print "LOG WRITE FAILED [" & filePath & "] Err=" & Err.Number & " " & Err.Description & " :: " & lineOut
        Err.Clear
        AppendFallbackLine lineOut
    End If
    Err.Clear
End Sub

Private Sub AppendFallbackLine(ByVal lineOut As String)
    Dim f As Integer, tmpPath As String
    On Error Resume Next
    tmpPath = Environ$("TEMP")
    If Len(tmpPath) = 0 Then tmpPath = Environ$("TMP")
    If Len(tmpPath) = 0 Then Exit Sub
    f = FreeFile
    Open tmpPath & "\reauth-log-fallback.log" For Append As #f
    If Err.Number = 0 Then
        Print #f, lineOut
        Close #f
    End If
    Err.Clear
End Sub

Public Sub LogErr(ByVal procName As String, Optional ByVal errNumOverride As Variant, Optional ByVal errDescOverride As Variant, Optional ByVal errLineOverride As Variant)
    Dim lineOut As String
    Dim errNum As Long, errDesc As String, errLine As Long
    On Error Resume Next
    If IsMissing(errNumOverride) Then
        errNum = Err.Number
        errDesc = Err.Description
        errLine = Erl
    Else
        errNum = CLng(errNumOverride)
        If IsMissing(errDescOverride) Then errDesc = "" Else errDesc = CStr(errDescOverride)
        If IsMissing(errLineOverride) Then errLine = 0 Else errLine = CLng(errLineOverride)
    End If
    ' Avoid logging spurious "errors" when there is no active error.
    If errNum = 0 And Len(Trim$(errDesc)) = 0 Then Exit Sub

    lineOut = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | ERROR | " & procName & _
              " | Err=" & CStr(errNum) & " | " & errDesc & " | Erl=" & CStr(errLine) & _
              IIf(Len(mCurrentStep) > 0, " | Step=" & mCurrentStep, "")
    AppendLogLine LOG_PATH, lineOut
    If mInReportRun And Len(mRunLogPath) > 0 Then AppendLogLine mRunLogPath, lineOut
End Sub

Public Sub LogStep(ByVal procName As String, ByVal stepName As String)
    On Error Resume Next
    mCurrentStep = stepName
    LogMsg "STEP | " & procName & " | " & stepName
End Sub

Public Function BuildRunLogPath() As String
    Dim logDir As String
    On Error Resume Next
    logDir = GetLogDir()
    If Len(logDir) = 0 Then logDir = GetProjectRoot() & "\log\"
    BuildRunLogPath = logDir & "reauth_" & Format(Now, "yyyy-mm-dd_HHMMSS") & ".log"
End Function

Public Sub WriteCompletionMarker(ByVal savePathOrMessage As String)
    Dim f As Integer
    On Error Resume Next
    EnsureDirForFile COMPLETION_MARKER_PATH
    f = FreeFile
    Open COMPLETION_MARKER_PATH For Output As #f
    If Err.Number = 0 Then
        Print #f, "completed=" & Format(Now, "yyyy-mm-dd hh:nn:ss")
        Print #f, "saved=" & savePathOrMessage
        Print #f, "run_log=" & mRunLogPath
        Print #f, "step=" & mCurrentStep
        Print #f, "version=" & MACRO_VERSION
        Close #f
    Else
        Dim mkErrNum As Long, mkErrDesc As String
        mkErrNum = Err.Number
        mkErrDesc = Err.Description
        Debug.Print "COMPLETION MARKER WRITE FAILED [" & COMPLETION_MARKER_PATH & "] Err=" & mkErrNum & " " & mkErrDesc
        Err.Clear
        LogMsg "ERROR | completion marker write failed | Err=" & CStr(mkErrNum) & " | " & mkErrDesc & " | path=" & COMPLETION_MARKER_PATH
    End If
    Err.Clear
End Sub

Private Sub EnsureDirForFile(ByVal filePath As String)
    Dim dirPath As String
    On Error Resume Next
    dirPath = Left$(filePath, InStrRev(filePath, "\"))
    If Len(dirPath) = 0 Then Exit Sub
    EnsureFolderExists dirPath
End Sub
