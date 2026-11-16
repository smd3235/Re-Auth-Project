Attribute VB_Name = "modLogging"
Option Explicit

' ============================================================================
'  Module: modLogging - centralized realtime debug/error logging
'  Writes to LOG_PATH (always) and to per-run mRunLogPath during report runs.
' ============================================================================

Public Sub LogMsg(ByVal msg As String)
    Dim f As Integer
    Dim lineOut As String
    On Error Resume Next
    lineOut = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
    EnsureDirForFile LOG_PATH
    f = FreeFile
    Open LOG_PATH For Append As #f
    If Err.Number = 0 Then
        Print #f, lineOut
        Close #f
    End If
    Err.Clear
    If mInReportRun And Len(mRunLogPath) > 0 Then
        EnsureDirForFile mRunLogPath
        f = FreeFile
        Open mRunLogPath For Append As #f
        If Err.Number = 0 Then
            Print #f, lineOut
            Close #f
        End If
        Err.Clear
    End If
End Sub

Public Sub LogErr(ByVal procName As String)
    Dim f As Integer
    Dim lineOut As String
    Dim errNum As Long, errDesc As String, errLine As Long
    On Error Resume Next
    errNum = Err.Number
    errDesc = Err.Description
    errLine = Erl
    lineOut = Format(Now, "yyyy-mm-dd hh:nn:ss") & " | ERROR | " & procName & _
              " | Err=" & CStr(errNum) & " | " & errDesc & " | Erl=" & CStr(errLine)
    EnsureDirForFile LOG_PATH
    f = FreeFile
    Open LOG_PATH For Append As #f
    If Err.Number = 0 Then
        Print #f, lineOut
        Close #f
    End If
    Err.Clear
    If mInReportRun And Len(mRunLogPath) > 0 Then
        EnsureDirForFile mRunLogPath
        f = FreeFile
        Open mRunLogPath For Append As #f
        If Err.Number = 0 Then
            Print #f, lineOut
            Close #f
        End If
        Err.Clear
    End If
End Sub

Public Sub LogStep(ByVal procName As String, ByVal stepName As String)
    On Error Resume Next
    LogMsg "STEP | " & procName & " | " & stepName
End Sub

Public Function BuildRunLogPath() As String
    Dim logDir As String
    On Error Resume Next
    logDir = GetLogDir()
    If Len(logDir) = 0 Then logDir = GetProjectRoot() & "\log\"
    BuildRunLogPath = logDir & "reauth_" & Format(Now, "yyyy-mm-dd_HHMMSS") & ".log"
End Function

Public Sub WriteCompletionMarker(ByVal savePath As String)
    Dim f As Integer
    On Error Resume Next
    EnsureDirForFile COMPLETION_MARKER_PATH
    f = FreeFile
    Open COMPLETION_MARKER_PATH For Output As #f
    If Err.Number = 0 Then
        Print #f, "completed=" & Format(Now, "yyyy-mm-dd hh:nn:ss")
        Print #f, "saved=" & savePath
        Print #f, "run_log=" & mRunLogPath
        Close #f
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
