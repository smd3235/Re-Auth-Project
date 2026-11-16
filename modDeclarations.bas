Attribute VB_Name = "modDeclarations"
Option Explicit

' ============================================================================
'  RE-AUTH REPORT BUILDER
'  Module: modDeclarations
' ============================================================================

' ----------------------------------------------------------------------------
' Sheet / header constants
' ----------------------------------------------------------------------------
Public Const RAW_SHEET As String = "Raw Data"
Public Const MANAGER_HEADER As String = "Assigned Sales Manager"
Public Const AUTH_STATUS_HEADER As String = "Auth Status"
Public Const PATIENT_HEADER As String = "Patient Name"
Public Const COMPANY_HEADER As String = "HC Company"
Public Const REAUTH_DUE_HEADER As String = "Re-auth Due Date"
Public Const AUTH_START_HEADER As String = "Auth Start Date"
Public Const TASK_HEADER As String = "Task Required"
Public Const WCC_HEADER As String = "Wound Care Consultant"
Public Const ASSOC_WCC_HEADER As String = "Associate WCC"
Public Const AUTHORIZATION_STATUS_HEADER As String = "Authorization Status"
Public Const MACRO_VERSION As String = "1.0"
Public Const MANAGER_HEADER_ROW As Long = 15
Public Const DASHBOARD_SHEET As String = "Dashboard"

' ----------------------------------------------------------------------------
' File / log paths (project-local)
' ----------------------------------------------------------------------------
Public Const PROJECT_ROOT As String = "C:\Users\smd32\Downloads\Re-Auth Project"
Public Const LOG_PATH As String = "C:\Users\smd32\Downloads\Re-Auth Project\log\macro-debug.log"
Public Const COMPLETION_MARKER_PATH As String = "C:\Users\smd32\Downloads\Re-Auth Project\log\reauth-complete.txt"
Public Const AUTO_SOURCE_PATH_FILE As String = "C:\Users\smd32\Downloads\Re-Auth Project\log\auto-source.txt"
Public Const AUTO_SKIP_PRIOR_FILE As String = "C:\Users\smd32\Downloads\Re-Auth Project\log\auto-skip-prior.txt"

' ----------------------------------------------------------------------------
' Shared runtime state
' ----------------------------------------------------------------------------
Public mRunLogPath As String
Public mInReportRun As Boolean
Public mReportStartTime As Date
