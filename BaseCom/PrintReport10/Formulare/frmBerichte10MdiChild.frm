VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form frmBerichte10MdiChild 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrStart 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer crvReport 
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      lastProp        =   600
      _cx             =   4260
      _cy             =   2355
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmBerichte10MdiChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'{Moduleigenschaften}
Private mstrReportFileName As String
Private mintCopies As Integer
Private mOutput As preOutput
'{Ende Moduleigenschaften}

'{Modulvariablen}
Private mcrxReport As CRAXDRT.Report
Private mcrxApplication As CRAXDRT.Application
'{Ende Modulvariablen}

Private Sub PrintReport()
On Error GoTo Fehler
  
  mcrxReport.PrintOut False, Me.Copies
  Set mcrxApplication = Nothing
  Set mcrxReport = Nothing
  Unload Me
  Exit Sub

Fehler:
  ShowError "PrintReport"
  Exit Sub
End Sub

Public Sub ExportReportToPDF(ByVal strExportFileName As String)

On Error GoTo Fehler
    
  mcrxReport.ExportOptions.DestinationType = crEDTDiskFile
  mcrxReport.ExportOptions.DiskFileName = strExportFileName
  mcrxReport.ExportOptions.FormatType = crEFTPortableDocFormat
  mcrxReport.ExportOptions.PDFExportAllPages = True
  mcrxReport.Export False
  
  Set mcrxApplication = Nothing
  Set mcrxReport = Nothing
  Unload Me
  Exit Sub

Fehler:
  ShowError "ExportReportToPDF"
  Exit Sub
End Sub

Public Sub ExportReportToXLS(ByVal strExportFileName As String)

On Error GoTo Fehler
    
  mcrxReport.ExportOptions.DestinationType = crEDTDiskFile
  mcrxReport.ExportOptions.DiskFileName = strExportFileName
  mcrxReport.ExportOptions.FormatType = crEFTExcel80
  mcrxReport.Export False
  
  Set mcrxApplication = Nothing
  Set mcrxReport = Nothing
  Unload Me
  Exit Sub

Fehler:
  ShowError "ExportReportToXLS"
  Exit Sub
End Sub

Private Sub ShowReport()
On Error GoTo Fehler

  'hack
  If crvReport Is Nothing Then Exit Sub
  crvReport.ReportSource = mcrxReport
  crvReport.ViewReport
  Set mcrxApplication = Nothing
  Set mcrxReport = Nothing
  Exit Sub
  
Fehler:
  ShowError "ShowReport"
  Exit Sub
End Sub

Private Sub crvReport_ZoomLevelChanged(ByVal ZoomLevel As Integer)
  If ZoomLevel <> 100 Then crvReport.Zoom 100
End Sub

Private Sub Form_Load()
  Me.Height = GetSetting("PrintReport", Me.Name, "Height", Me.Height)
  Me.Width = GetSetting("PrintReport", Me.Name, "Width", Me.Width)
  Me.Top = GetSetting("PrintReport", Me.Name, "top", Me.Top)
  Me.Left = GetSetting("PrintReport", Me.Name, "Left", Me.Left)
End Sub

Private Sub Form_Resize()
  With crvReport
    .Top = 0
    .Left = 0
    .Height = ScaleHeight
    .Width = ScaleWidth
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  SaveSetting "PrintReport", Me.Name, "Height", Me.Height
  SaveSetting "PrintReport", Me.Name, "Width", Me.Width
  SaveSetting "PrintReport", Me.Name, "top", Me.Top
  SaveSetting "PrintReport", Me.Name, "Left", Me.Left
  
  mOutput = prScreen
  mintCopies = 1
  
End Sub

Public Property Get ReportFileName() As String
  ReportFileName = mstrReportFileName
End Property

Public Property Let ReportFileName(ByVal strReportFileName As String)
On Error GoTo Fehler
  
  Screen.MousePointer = 11
  mstrReportFileName = strReportFileName
  
  Set mcrxApplication = New CRAXDRT.Application
  Set mcrxReport = mcrxApplication.OpenReport(strReportFileName)
  Screen.MousePointer = 0
  Exit Property

Fehler:
  ShowError "ReportFileName"
  Exit Property
End Property

Public Property Get Output() As preOutput
  Output = mOutput
End Property

Public Property Let Output(ByVal prOutput As preOutput)
  mOutput = prOutput
End Property

Private Sub tmrStart_Timer()
Dim strParameter As String
Dim strValue As String
Dim i As Integer

  tmrStart.Enabled = False
  
  If Me.Copies = 0 Then Me.Copies = 1
    
  If mOutput = prPrinter Then
    PrintReport
  Else
    ShowReport
  End If
  
End Sub

Public Function SetDataSource(ByRef rsData As ADODB.Recordset) As String
Dim i As Integer

On Error GoTo Fehler

  Screen.MousePointer = 11

  If rsData.EOF Then
    SetDataSource = "Keine Datensätze vorhanden!"
    Screen.MousePointer = 0
    Exit Function
  End If
  
  For i = 0 To rsData.Fields.Count - 1
    If rsData.Fields(i).Value & "" = "" Then
      SetDataSource = "Feld '" & rsData.Fields(i).Name _
      & "' ohne Wert!" & Chr(13) _
      & "Bitte weisen Sie zuvor einen Wert zu!"
      Screen.MousePointer = 0
      Exit Function
    End If
  Next i

  mcrxReport.Database.SetDataSource rsData
  
  SetDataSource = ""
  
  Screen.MousePointer = 0
  Exit Function
  
Fehler:
  ShowError "SetDataSource"
  Exit Function
End Function

Public Property Get Copies() As Integer
  Copies = mintCopies
End Property

Public Property Let Copies(ByVal intCopies As Integer)
  mintCopies = intCopies
End Property

Public Function SetDataSourceSubReport _
(ByVal strSubReportName As String _
, ByRef rsData As ADODB.Recordset) As String

Dim crxUReport As CRAXDRT.Report
Dim i As Integer

On Error GoTo Fehler

  Screen.MousePointer = 11
  
  If rsData.EOF Then
    SetDataSourceSubReport = "Keine Datensätze vorhanden!"
    Screen.MousePointer = 0
    Exit Function
  End If
  
  For i = 0 To rsData.Fields.Count - 1
    If rsData.Fields(i).Value & "" = "" Then
      SetDataSourceSubReport = "Feld '" & rsData.Fields(i).Name _
      & "' ohne Wert!" & Chr(13) _
      & "Bitte weisen Sie zuvor einen Wert zu!"
      Screen.MousePointer = 0
      Exit Function
    End If
  Next i

  Set crxUReport = mcrxReport.OpenSubreport(strSubReportName)
  crxUReport.Database.SetDataSource rsData
  
  SetDataSourceSubReport = True
  Screen.MousePointer = 0
  Exit Function
  
Fehler:
  ShowError "SetDataSourceSubReport"
  Exit Function
End Function


