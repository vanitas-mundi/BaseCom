VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmBerichte 
   Caption         =   "Berichte"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   Icon            =   "frmBerichte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrStart 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5805
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
   End
End
Attribute VB_Name = "frmBerichte"
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
Private mcrxReport As CRAXDDRT.Report
Private mcrxApplication As CRAXDDRT.Application
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
  
'  MsgBox mcrxReport.ExportOptions.ApplicationFileName
'  MsgBox mcrxReport.ExportOptions.CharFieldDelimiter
'  MsgBox mcrxReport.ExportOptions.CharStringDelimiter
'  MsgBox mcrxReport.ExportOptions.DestinationDllName
'  MsgBox mcrxReport.ExportOptions.DestinationType
'  MsgBox mcrxReport.ExportOptions.DiskFileName
'  MsgBox mcrxReport.ExportOptions.ExcelAreaGroupNumber
'  MsgBox mcrxReport.ExportOptions.ExcelAreaType
'  MsgBox mcrxReport.ExportOptions.ExcelConstantColumnWidth
'  MsgBox mcrxReport.ExportOptions.ExcelTabHasColumnHeadings
'  MsgBox mcrxReport.ExportOptions.ExcelUseConstantColumnWidth
'  MsgBox mcrxReport.ExportOptions.ExcelUseTabularFormat
'  MsgBox mcrxReport.ExportOptions.ExcelUseWorksheetFunctions
'  MsgBox mcrxReport.ExportOptions.ExchangeDestinationType
'  MsgBox mcrxReport.ExportOptions.ExchangeFolderPath
'  'MsgBox mcrxReport.ExportOptions.ExchangePassword
'  MsgBox mcrxReport.ExportOptions.ExchangeProfile
'  MsgBox mcrxReport.ExportOptions.FormatDllName
'  MsgBox mcrxReport.ExportOptions.FormatType
'  MsgBox mcrxReport.ExportOptions.HTMLEnableSeparatedPages
'  MsgBox mcrxReport.ExportOptions.HTMLFileName
'  MsgBox mcrxReport.ExportOptions.HTMLHasPageNavigator
'  MsgBox mcrxReport.ExportOptions.LotusDominoComments
'  MsgBox mcrxReport.ExportOptions.LotusDominoDatabaseName
'  MsgBox mcrxReport.ExportOptions.LotusDominoFormName
'  MsgBox mcrxReport.ExportOptions.MailBccList
'  MsgBox mcrxReport.ExportOptions.MailCcList
'  MsgBox mcrxReport.ExportOptions.MailMessage
'  MsgBox mcrxReport.ExportOptions.MailSubject
'  MsgBox mcrxReport.ExportOptions.MailToList
'  MsgBox mcrxReport.ExportOptions.NumberOfLinesPerPage
'  MsgBox mcrxReport.ExportOptions.ODBCDataSourceName
'  'MsgBox mcrxReport.ExportOptions.ODBCDataSourcePassword
'  MsgBox mcrxReport.ExportOptions.ODBCDataSourceUserID
'  MsgBox mcrxReport.ExportOptions.ODBCExportTableName
'  MsgBox mcrxReport.ExportOptions.Parent
'  MsgBox mcrxReport.ExportOptions.PDFExportAllPages
'  MsgBox mcrxReport.ExportOptions.PDFFirstPageNumber
'  MsgBox mcrxReport.ExportOptions.PDFLastPageNumber
'  MsgBox mcrxReport.ExportOptions.RTFExportAllPages
'  MsgBox mcrxReport.ExportOptions.RTFFirstPageNumber
'  MsgBox mcrxReport.ExportOptions.RTFLastPageNumber
'  MsgBox mcrxReport.ExportOptions.UseReportDateFormat
'  MsgBox mcrxReport.ExportOptions.UseReportNumberFormat
'  MsgBox mcrxReport.ExportOptions.XMLAllowMultipleFiles
'  MsgBox mcrxReport.ExportOptions.XMLFileName
'
  
  
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
  'mcrxReport.ExportOptions.PDFExportAllPages = True
    
  mcrxReport.Export False
  
'  MsgBox mcrxReport.ExportOptions.ApplicationFileName
'  MsgBox mcrxReport.ExportOptions.CharFieldDelimiter
'  MsgBox mcrxReport.ExportOptions.CharStringDelimiter
'  MsgBox mcrxReport.ExportOptions.DestinationDllName
'  MsgBox mcrxReport.ExportOptions.DestinationType
'  MsgBox mcrxReport.ExportOptions.DiskFileName
'  MsgBox mcrxReport.ExportOptions.ExcelAreaGroupNumber
'  MsgBox mcrxReport.ExportOptions.ExcelAreaType
'  MsgBox mcrxReport.ExportOptions.ExcelConstantColumnWidth
'  MsgBox mcrxReport.ExportOptions.ExcelTabHasColumnHeadings
'  MsgBox mcrxReport.ExportOptions.ExcelUseConstantColumnWidth
'  MsgBox mcrxReport.ExportOptions.ExcelUseTabularFormat
'  MsgBox mcrxReport.ExportOptions.ExcelUseWorksheetFunctions
'  MsgBox mcrxReport.ExportOptions.ExchangeDestinationType
'  MsgBox mcrxReport.ExportOptions.ExchangeFolderPath
'  'MsgBox mcrxReport.ExportOptions.ExchangePassword
'  MsgBox mcrxReport.ExportOptions.ExchangeProfile
'  MsgBox mcrxReport.ExportOptions.FormatDllName
'  MsgBox mcrxReport.ExportOptions.FormatType
'  MsgBox mcrxReport.ExportOptions.HTMLEnableSeparatedPages
'  MsgBox mcrxReport.ExportOptions.HTMLFileName
'  MsgBox mcrxReport.ExportOptions.HTMLHasPageNavigator
'  MsgBox mcrxReport.ExportOptions.LotusDominoComments
'  MsgBox mcrxReport.ExportOptions.LotusDominoDatabaseName
'  MsgBox mcrxReport.ExportOptions.LotusDominoFormName
'  MsgBox mcrxReport.ExportOptions.MailBccList
'  MsgBox mcrxReport.ExportOptions.MailCcList
'  MsgBox mcrxReport.ExportOptions.MailMessage
'  MsgBox mcrxReport.ExportOptions.MailSubject
'  MsgBox mcrxReport.ExportOptions.MailToList
'  MsgBox mcrxReport.ExportOptions.NumberOfLinesPerPage
'  MsgBox mcrxReport.ExportOptions.ODBCDataSourceName
'  'MsgBox mcrxReport.ExportOptions.ODBCDataSourcePassword
'  MsgBox mcrxReport.ExportOptions.ODBCDataSourceUserID
'  MsgBox mcrxReport.ExportOptions.ODBCExportTableName
'  MsgBox mcrxReport.ExportOptions.Parent
'  MsgBox mcrxReport.ExportOptions.PDFExportAllPages
'  MsgBox mcrxReport.ExportOptions.PDFFirstPageNumber
'  MsgBox mcrxReport.ExportOptions.PDFLastPageNumber
'  MsgBox mcrxReport.ExportOptions.RTFExportAllPages
'  MsgBox mcrxReport.ExportOptions.RTFFirstPageNumber
'  MsgBox mcrxReport.ExportOptions.RTFLastPageNumber
'  MsgBox mcrxReport.ExportOptions.UseReportDateFormat
'  MsgBox mcrxReport.ExportOptions.UseReportNumberFormat
'  MsgBox mcrxReport.ExportOptions.XMLAllowMultipleFiles
'  MsgBox mcrxReport.ExportOptions.XMLFileName
'
  
  
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
  
  Set mcrxApplication = New CRAXDDRT.Application
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

Dim crxUReport As CRAXDDRT.Report
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
