VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.UserControl ReportViewerControl 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ReportViewerControl.ctx":0000
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "ReportViewerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Report As CRVIEWERLibCtl.CRViewer
'Public ReportView As clsReportView


'{---------------------- Konstruktor und Destruktor der Klasse -----------------------}
'Private Sub UserControl_Initialize()
'  Set Report = crvReport
'  Set ReportView = New clsReportView
'End Sub
'
'
'Private Sub UserControl_Terminate()
'  Set Report = Nothing
'  Set ReportView = Nothing
'End Sub
'{---------------------- Ende Konstruktor und Destruktor der Klasse -----------------------}



'{------------------ Zugriffsmethoden der Klasseneigenschaften ----------------------}
Public Property Let ReportSource(ByVal crxReport As Object)
  crvReport.ReportSource = crxReport
End Property

'{------------------ Ende Zugriffsmethoden der Klasseneigenschaften ----------------------}


'{------------------------ Private Klassenmethoden --------------------------}
Private Sub UserControl_Resize()
On Error Resume Next
  crvReport.Width = UserControl.ScaleWidth
  crvReport.Height = UserControl.ScaleHeight
End Sub

Private Sub crvReport_ZoomLevelChanged(ByVal ZoomLevel As Integer)
  If ZoomLevel <> 100 Then crvReport.Zoom 100
End Sub
'{------------------------ Ende Private Klassenmethoden --------------------------}



'{----------------------- Öffentliche KLassenmethoden ------------------------}
Public Sub ViewReport()
  crvReport.ViewReport
End Sub
