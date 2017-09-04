VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintReport10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDialog As Object
Private mTitle As String
Private mMdiChild As Boolean

Public Property Let EnableAnimationCtrl(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableAnimationCtrl = blnValue
End Property

Public Property Let DisplayBackgroundEdge(ByVal blnValue As Boolean)
  mDialog.crvReport.DisplayBackgroundEdge = blnValue
End Property

Public Property Let DisplayBorder(ByVal blnValue As Boolean)
  mDialog.crvReport.DisplayBorder = blnValue
End Property

Public Property Let DisplayTabs(ByVal blnValue As Boolean)
  mDialog.crvReport.DisplayTabs = blnValue
End Property

Public Property Let DisplayToolbar(ByVal blnValue As Boolean)
  mDialog.crvReport.DisplayToolbar = blnValue
End Property

Public Property Let EnableCloseButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableCloseButton = blnValue
End Property

Public Property Let EnableDrillDown(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableDrillDown = blnValue
End Property

Public Property Let EnableGroupTree(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableGroupTree = blnValue
End Property

Public Property Let EnablePrintButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnablePrintButton = blnValue
End Property

Public Property Let EnableProgressControl(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableProgressControl = blnValue
End Property

Public Property Let EnableRefreshButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableRefreshButton = blnValue
End Property

Public Property Let EnableSearchControl(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableSearchControl = blnValue
End Property

Public Property Let EnableSearchExpertButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableSearchExpertButton = blnValue
End Property

Public Property Let EnableSelectExpertButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableSelectExpertButton = blnValue
End Property

Public Property Let EnableStopButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableStopButton = blnValue
End Property

Public Property Let EnableToolbar(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableToolbar = blnValue
End Property

Public Property Let EnableZoomControl(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableZoomControl = blnValue
End Property

Public Property Let EnableExportButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableExportButton = blnValue
End Property

Public Property Let EnableHelpButton(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableHelpButton = blnValue
End Property

Public Property Let EnableNavigationControls(ByVal blnValue As Boolean)
  mDialog.crvReport.EnableNavigationControls = blnValue
End Property

Public Property Let EnablePopupMenu(ByVal blnValue As Boolean)
  mDialog.crvReport.EnablePopupMenu = blnValue
End Property

Public Sub ShowReport(Optional ByVal intModal)
  mDialog.Output = basGlobal.prScreen
  If Not IsMissing(intModal) Then
    mDialog.SelectAction CInt(intModal)
  Else
    mDialog.SelectAction
  End If
End Sub

Public Sub PrintReport()
  mDialog.Output = prPrinter
  mDialog.SelectAction 1
End Sub

Public Function SetDataSource(ByRef rsData As ADODB.Recordset) As String
  SetDataSource = mDialog.SetDataSource(rsData)
End Function

Public Property Get Copies() As Integer
  Copies = mDialog.Copies
End Property

Public Property Let Copies(ByVal intCopies As Integer)
  mDialog.Copies = intCopies
End Property

Public Property Get ReportFileName() As String
  ReportFileName = mDialog.ReportFileName
End Property

Public Property Let ReportFileName(ByVal strReportFileName As String)
  mDialog.ReportFileName = strReportFileName
End Property

Public Function SetDataSourceSubReport _
(ByVal strSubReportName As String _
, ByRef rsData As ADODB.Recordset) As String

  mDialog.SetDataSourceSubReport strSubReportName, rsData
End Function

Public Sub ExportReportToPDF(ByVal strExportFileName As String)
  mDialog.ExportReportToPDF strExportFileName
End Sub

Public Sub ExportReportToXLS(ByVal strExportFileName As String)
  mDialog.ExportReportToXLS strExportFileName
End Sub

Public Property Get MdiChild() As Boolean
  MdiChild = mMdiChild
End Property

Public Property Let MdiChild(ByVal value As Boolean)
  mMdiChild = value
  Set mDialog = frmBerichte10 'IIf(value, frmBerichte10MdiChild, frmBerichte10) Tja, geht leider nicht als MdiChild - Ich hasse ActiveX ...
End Property

Public Property Get Title() As String
  Title = mTitle
End Property

Public Property Let Title(ByVal value As String)
  mTitle = value
  mDialog.Caption = value
End Property

Private Sub Class_Initialize()
  Me.MdiChild = False
End Sub
