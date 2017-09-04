VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiEventLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const GMEM_ZEROINIT = &H40  'Initialisiere Speicherbereich mit NULLs
    
Public Enum enumEventType
  enumEventError = 1
  enumEventWarning = 2
  enumEventInformation = 4
End Enum
    
Private Const EVENTLOG_SUCCESS = 0
Private Const EVENTLOG_ERROR_TYPE = 1
Private Const EVENTLOG_WARNING_TYPE = 2
Private Const EVENTLOG_INFORMATION_TYPE = 4
Private Const EVENTLOG_AUDIT_SUCCESS = 8
Private Const EVENTLOG_AUDIT_FAILURE = 10
    
Private Declare Function RegisterEventSource _
                Lib "advapi32.dll" _
                Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, _
                                              ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource _
                Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Private Declare Function ReportEvent _
                Lib "advapi32.dll" _
                Alias "ReportEventA" (ByVal hEventLog As Long, _
                                      ByVal wType As Integer, _
                                      ByVal wCategory As Integer, _
                                      ByVal dwEventID As Long, _
                                      ByVal lpUserSid As Any, _
                                      ByVal wNumStrings As Integer, _
                                      ByVal dwDataSize As Long, _
                                      plpStrings As Long, _
                                      lpRawData As Any) As Boolean
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (hpvDest As Any, _
                                       hpvSource As Any, _
                                       ByVal cbCopy As Long)
Private Declare Function GlobalAlloc _
                Lib "kernel32" (ByVal wFlags As Long, _
                                ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Sub LogNTEvent(strEvent As String, lngLogType As enumEventType, lngEventID As Long)
  Dim hEventLog    As Long
  Dim hEventString As Long
  Dim nSize        As Long
  'lngEventID zeigt auf einen gültigen Eintrag einer Resource-Datei (Exe oder DLL)
  'siehe "The NT EventLog, Part 1: Meet the API" in MSDN
  hEventLog = RegisterEventSource("", app.title)
  nSize = Len(strEvent) + 1
  hEventString = GlobalAlloc(GMEM_ZEROINIT, nSize)
  CopyMemory ByVal hEventString, ByVal strEvent, nSize

  If ReportEvent(hEventLog, lngLogType, 0, lngEventID, 0&, 1&, nSize, hEventString, hEventString) = 0 Then
    MsgBox GetLastError()
  End If

  Call GlobalFree(hEventString)
  DeregisterEventSource (hEventLog)
End Sub

