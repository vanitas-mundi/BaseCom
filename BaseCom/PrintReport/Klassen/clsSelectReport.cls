VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelectReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mOutput As preOutput

Public Sub SelectReport(Optional ByVal intModal)
  If Not IsMissing(intModal) Then
    frmSelectBericht.Show intModal
  Else
    frmSelectBericht.Show
  End If
End Sub

Public Property Get ReportPath() As String
  ReportPath = frmSelectBericht.ReportPath
End Property

Public Property Let ReportPath(ByVal strReportPath As String)
  frmSelectBericht.ReportPath = strReportPath
End Property

Public Property Get SQLPath() As String
  SQLPath = frmSelectBericht.SQLPath
End Property

Public Property Let SQLPath(ByVal strSQLPath As String)
  frmSelectBericht.SQLPath = strSQLPath
End Property

Public Property Get DefaultCopies() As Integer
  DefaultCopies = frmSelectBericht.txtKopien.Text
End Property

Public Property Let DefaultCopies(ByVal intDefaultCopies As Integer)
  frmSelectBericht.txtKopien.Text = intDefaultCopies
End Property

Public Property Get DefaultOutput() As Integer
  If frmSelectBericht.optDruckvorschau.Value Then
    DefaultOutput = 0
  Else
    DefaultOutput = 1
  End If

End Property

Public Property Let DefaultOutput(ByVal intDefaultOutput As Integer)
  If intDefaultOutput = 0 Then
    frmSelectBericht.optDruckvorschau.Value = True
  Else
    frmSelectBericht.optDrucker.Value = True
  End If
End Property

Public Property Let ConnectionString(ByVal strConnectionString As String)
  frmSelectBericht.ConnectionString = strConnectionString
End Property

Public Property Get ConnectionString() As String
  ConnectionString = frmSelectBericht.ConnectionString
End Property

