VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportViewerHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'{Moduleigenschaften}
Private mstrReportFileName As String
Private mlngCopies As Long
'{Ende Moduleigenschaften}

'{Modulvariablen}
Private mcrxReport As CRAXDDRT.Report
Private mcrxApplication As CRAXDDRT.Application
'{Ende Modulvariablen}

'{---------------------- Konstruktor und Destruktor -------------------}

Private Sub Class_Initialize()
  mstrReportFileName = ""
  Me.Copies = 1

  Set mcrxReport = New CRAXDDRT.Report
  Set mcrxApplication = New CRAXDDRT.Application

End Sub

Private Sub Class_Terminate()

  Set mcrxReport = Nothing
  Set mcrxApplication = Nothing

End Sub

'{---------------------- End Konstruktor und Destruktor -------------------}



'{--------------- Zugriffmethoden der Klasseneigenschaften --------------}

Public Property Get ReportFileName() As String
Attribute ReportFileName.VB_Description = "Legt den Report fest und weist ihn zur weiteren Bearbeitung zu.\r\nMuss vor SetDataSource und SetDataSourceSubReport geschehen."
  ReportFileName = mstrReportFileName
End Property

Public Property Let ReportFileName(ByVal strReportFileName As String)
On Error GoTo Fehler

  Screen.MousePointer = 11
  mstrReportFileName = strReportFileName

  Set mcrxApplication = CreateObject("CrystalRuntime.Application")
  Set mcrxReport = mcrxApplication.OpenReport(strReportFileName)
  Screen.MousePointer = 0
  Exit Property

Fehler:
  ShowError "ReportFileName"
  Exit Property
End Property

Public Property Get Copies() As Long
Attribute Copies.VB_Description = "Die Anzahl der zu druckenden Kopien."
  Copies = mlngCopies
End Property

Public Property Let Copies(ByVal lngCopies As Long)
  mlngCopies = lngCopies
End Property

'{--------------- Zugriffmethoden der Klasseneigenschaften --------------}




'{----------------- Private Methoden der Klasse ------------------------}
Private Sub ShowError(ByVal strMessage As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strMessage
End Sub

'{----------------- Ende Private Methoden der Klasse ------------------------}



'{------------------- Öffentliche Methoden der Klasse --------------------------}

Public Sub PrintReport()
Attribute PrintReport.VB_Description = "Druckt den aktuellen Report."
On Error GoTo Fehler
  
  mcrxReport.PrintOut False, Me.Copies
  Exit Sub

Fehler:
  ShowError "PrintReport"
  Exit Sub
End Sub

Public Sub ShowReport(ByRef aReportView As CRVIEWERLib.CRViewer)
Attribute ShowReport.VB_Description = "Zeigt den Report an.\r\nFür die Ansicht wird ein Objekt vom Typ rvReport benötigt, welches übergeben werden muss."
On Error GoTo Fehler

  aReportView.ReportSource = mcrxReport
  aReportView.ViewReport
  Exit Sub
  
Fehler:
  ShowError "ShowReport"
  Exit Sub
End Sub

Public Sub ExportReportToPDF(ByVal strExportFileName As String)
Attribute ExportReportToPDF.VB_Description = "Exportiert den Bericht in ein PDF-Dokument."

On Error GoTo Fehler
    
  With mcrxReport.ExportOptions
    .DestinationType = crEDTDiskFile
    .DiskFileName = strExportFileName
    .FormatType = crEFTPortableDocFormat
    .PDFExportAllPages = True
  End With
    
  mcrxReport.Export False
  Exit Sub

Fehler:
  ShowError "ExportReportToPDF"
  Exit Sub
End Sub

Public Sub ExportReportToXLS(ByVal strExportFileName As String)
Attribute ExportReportToXLS.VB_Description = "Exportiert den Bericht in ein Excel-Dokument."

On Error GoTo Fehler
    
  With mcrxReport.ExportOptions
    .DestinationType = crEDTDiskFile
    .DiskFileName = strExportFileName
    .FormatType = crEFTExcel80
  End With
    
  mcrxReport.Export False
  Exit Sub

Fehler:
  ShowError "ExportReportToXLS"
  Exit Sub
End Sub

Private Function CheckFirstRow(ByRef rsData As Object) As Boolean
Dim i As Integer

On Error GoTo Fehler

  CheckFirstRow = True
  
  If rsData.EOF Then
    MsgBox "Keine Datensätze vorhanden!", 48, "Keine Datensätze"
    CheckFirstRow = False
  Else
    For i = 0 To rsData.Fields.Count - 1
      If rsData.Fields(i).value & "" = "" Then
        MsgBox "Feld '" & rsData.Fields(i).name _
        & "' ohne Wert!" & Chr(13) _
        & "Bitte weisen Sie zuvor einen Wert zu!" _
        , 48, "Feld ohne Wert"
        CheckFirstRow = False
        Exit For
      End If
    Next i
  End If
  Exit Function
  
Fehler:
  ShowError "CheckFirstRow"
  Exit Function
End Function

Public Sub SetDataSource(ByRef rsData As Object, Optional ByVal blnSilentMode As Boolean = False)
Attribute SetDataSource.VB_Description = "Legt die Datenquelle des Reports fest."

On Error GoTo Fehler

  Screen.MousePointer = 11


  If blnSilentMode Then
    mcrxReport.Database.SetDataSource rsData
  Else
    If CheckFirstRow(rsData) Then mcrxReport.Database.SetDataSource rsData
  End If
  
  Screen.MousePointer = 0
  Exit Sub
  
Fehler:
  ShowError "SetDataSource"
  Exit Sub
End Sub

Public Sub SetDataSourceSubReport _
(ByVal strSubReportName As String, ByRef rsData As Object _
, Optional ByVal blnSilentMode As Boolean = False)
Attribute SetDataSourceSubReport.VB_Description = "Legt die Datenquelle eines Unterberichtes fest."

Dim crxUReport As CRAXDDRT.Report

On Error GoTo Fehler

  Screen.MousePointer = 11
  
  If blnSilentMode Then
    Set crxUReport = mcrxReport.OpenSubreport(strSubReportName)
    crxUReport.Database.SetDataSource rsData
  Else
    If CheckFirstRow(rsData) Then
      Set crxUReport = mcrxReport.OpenSubreport(strSubReportName)
      crxUReport.Database.SetDataSource rsData
    End If
  End If
  
  Screen.MousePointer = 0
  Exit Sub
  
Fehler:
  ShowError "SetDataSourceSubReport"
  Exit Sub
End Sub

'{------------------- Ende Öffentliche Methoden der Klasse --------------------------}


