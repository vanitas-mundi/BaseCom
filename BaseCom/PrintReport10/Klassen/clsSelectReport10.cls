VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelectReport10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mOutput As preOutput

Public Sub SelectReport(Optional ByVal intModal)
  If Not IsMissing(intModal) Then
    frmSelectBericht10.Show intModal
  Else
    frmSelectBericht10.Show
  End If
End Sub

Public Property Get ReportPath() As String
  ReportPath = frmSelectBericht10.ReportPath
End Property

Public Property Let ReportPath(ByVal strReportPath As String)
  frmSelectBericht10.ReportPath = strReportPath
End Property

Public Property Get SQLPath() As String
  SQLPath = frmSelectBericht10.SQLPath
End Property

Public Property Let SQLPath(ByVal strSQLPath As String)
  frmSelectBericht10.SQLPath = strSQLPath
End Property

Public Property Get DefaultCopies() As Integer
  DefaultCopies = frmSelectBericht10.txtKopien.Text
End Property

Public Property Let DefaultCopies(ByVal intDefaultCopies As Integer)
  frmSelectBericht10.txtKopien.Text = intDefaultCopies
End Property

Public Property Get DefaultOutput() As Integer
  If frmSelectBericht10.optDruckvorschau.Value Then
    DefaultOutput = 0
  Else
    DefaultOutput = 1
  End If

End Property

Public Property Let DefaultOutput(ByVal intDefaultOutput As Integer)
  If intDefaultOutput = 0 Then
    frmSelectBericht10.optDruckvorschau.Value = True
  Else
    frmSelectBericht10.optDrucker.Value = True
  End If
End Property

Public Property Let ConnectionString(ByVal strConnectionString As String)
  frmSelectBericht10.ConnectionString = strConnectionString
End Property

Public Property Get ConnectionString() As String
  ConnectionString = frmSelectBericht10.ConnectionString
End Property

