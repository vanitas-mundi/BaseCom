VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let EnableAnimationCtrl(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableAnimationCtrl = blnValue
End Property

Public Property Let DisplayBackgroundEdge(ByVal blnValue As Boolean)
  frmBerichte.crvReport.DisplayBackgroundEdge = blnValue
End Property

Public Property Let DisplayBorder(ByVal blnValue As Boolean)
  frmBerichte.crvReport.DisplayBorder = blnValue
End Property

Public Property Let DisplayTabs(ByVal blnValue As Boolean)
  frmBerichte.crvReport.DisplayTabs = blnValue
End Property

Public Property Let DisplayToolbar(ByVal blnValue As Boolean)
  frmBerichte.crvReport.DisplayToolbar = blnValue
End Property

Public Property Let EnableCloseButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableCloseButton = blnValue
End Property

Public Property Let EnableDrillDown(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableDrillDown = blnValue
End Property

Public Property Let EnableGroupTree(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableGroupTree = blnValue
End Property

Public Property Let EnablePrintButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnablePrintButton = blnValue
End Property

Public Property Let EnableProgressControl(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableProgressControl = blnValue
End Property

Public Property Let EnableRefreshButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableRefreshButton = blnValue
End Property

Public Property Let EnableSearchControl(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableSearchControl = blnValue
End Property

Public Property Let EnableSearchExpertButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableSearchExpertButton = blnValue
End Property

Public Property Let EnableSelectExpertButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableSelectExpertButton = blnValue
End Property

Public Property Let EnableStopButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableStopButton = blnValue
End Property

Public Property Let EnableToolbar(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableToolbar = blnValue
End Property

Public Property Let EnableZoomControl(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableZoomControl = blnValue
End Property

Public Property Let EnableExportButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableExportButton = blnValue
End Property

Public Property Let EnableHelpButton(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableHelpButton = blnValue
End Property

Public Property Let EnableNavigationControls(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnableNavigationControls = blnValue
End Property

Public Property Let EnablePopupMenu(ByVal blnValue As Boolean)
  frmBerichte.crvReport.EnablePopupMenu = blnValue
End Property

Public Sub ShowReport(Optional ByVal intModal)
  frmBerichte.Output = prScreen
  If Not IsMissing(intModal) Then
    frmBerichte.Show intModal
  Else
    frmBerichte.Show
  End If
End Sub

Public Sub PrintReport()
  frmBerichte.Output = prPrinter
  frmBerichte.Show 1
End Sub

Public Function SetDataSource(ByRef rsData As ADODB.Recordset) As String
  frmBerichte.SetDataSource rsData
End Function

Public Property Get Copies() As Integer
  Copies = frmBerichte.Copies
End Property

Public Property Let Copies(ByVal intCopies As Integer)
  frmBerichte.Copies = intCopies
End Property

Public Property Get ReportFileName() As String
  ReportFileName = frmBerichte.ReportFileName
End Property

Public Property Let ReportFileName(ByVal strReportFileName As String)
  frmBerichte.ReportFileName = strReportFileName
End Property

Public Function SetDataSourceSubReport _
(ByVal strSubReportName As String _
, ByRef rsData As ADODB.Recordset) As String

  frmBerichte.SetDataSourceSubReport strSubReportName, rsData
End Function

Public Sub ExportReportToPDF(ByVal strExportFileName As String)
  frmBerichte.ExportReportToPDF strExportFileName
End Sub

Public Sub ExportReportToXLS(ByVal strExportFileName As String)
  frmBerichte.ExportReportToXLS strExportFileName
End Sub

