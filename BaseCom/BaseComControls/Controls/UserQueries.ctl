VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl UserQueriesViewControl 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   9705
   ToolboxBitmap   =   "UserQueries.ctx":0000
   Begin VB.Frame fraSQL 
      Caption         =   "Abfrageergebnis (Datensätze 0)"
      Height          =   4335
      Left            =   3200
      TabIndex        =   3
      Top             =   480
      Width           =   6015
      Begin MSFlexGridLib.MSFlexGrid flexAbfragen 
         Height          =   3975
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7011
         _Version        =   393216
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         FocusRect       =   2
         AllowUserResizing=   1
         BorderStyle     =   0
      End
   End
   Begin MSComctlLib.StatusBar staStatus 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5175
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   4763
            MinWidth        =   4763
            Text            =   "Abfragenübersicht ausblenden"
            TextSave        =   "Abfragenübersicht ausblenden"
            Key             =   "QueryView"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Visible         =   0   'False
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "UserQueries.ctx":0312
            Key             =   "EditComment"
            Object.ToolTipText     =   "Abfragebeschreibung editieren"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12277
            Key             =   "Comment"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblLeiste 
      Align           =   1  'Oben ausrichten
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Steuerdatei öffnen ..."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Steuerdatei speichern .."
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Drucken"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Suchen"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "NewQuery"
            Object.ToolTipText     =   "Neue Abfrage ..."
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "DeleteQuery"
            Object.ToolTipText     =   "Abfrage löschen"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Entwurfsansicht"
            Object.ToolTipText     =   "Entwurfsansicht"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "NewParameter"
            Object.ToolTipText     =   "Neuer Parameter ..."
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "DeleteParameter"
            Object.ToolTipText     =   "Parameter löschen"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "LinkControlFile"
            Object.ToolTipText     =   "Steuerdatei mit MS Word-Serienbrief verknüpfen ..."
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Word"
            Object.ToolTipText     =   "Mit MS Word veröffentlichen"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excel"
            Object.ToolTipText     =   "Analysieren mit MS Excel"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eMail"
            Object.ToolTipText     =   "Sammel-EMail senden"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mail"
                  Text            =   "Sammel-EMail senden"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CallNumber"
                  Text            =   "Anrufen"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SendSMS"
                  Text            =   "SMS senden"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Aufsteigend"
            Object.ToolTipText     =   "Aufsteigend sortieren"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Absteigend"
            Object.ToolTipText     =   "Absteigend sortieren"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Summe"
            Object.ToolTipText     =   "Spaltensumme"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Abfrage"
            Object.ToolTipText     =   "Abfrage ausführen ..."
         EndProperty
      EndProperty
      Begin VB.ComboBox cboLimit 
         Height          =   315
         ItemData        =   "UserQueries.ctx":046C
         Left            =   8520
         List            =   "UserQueries.ctx":04A6
         TabIndex        =   5
         Text            =   "cboLimit"
         ToolTipText     =   "Datensatzlimit"
         Top             =   40
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ilsBilder 
      Left            =   240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   33
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":0508
            Key             =   "Abfrage"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":095A
            Key             =   "OpenQueries"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":0CF4
            Key             =   "DeleteParameter"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":0E4E
            Key             =   "NewParameter"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":0FA8
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":1FFA
            Key             =   "Query"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2154
            Key             =   "Parameter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":22AE
            Key             =   "ParameterInUse"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2408
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2562
            Key             =   "Value"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":26BC
            Key             =   "Operator"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2816
            Key             =   "Parametertyp"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2970
            Key             =   "AdminQueries"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2ACA
            Key             =   "UserQueries"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2C24
            Key             =   "prg"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":2D7E
            Key             =   "QueryFolder"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":3118
            Key             =   "DeleteQuery"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":3272
            Key             =   "NewQuery"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":33CC
            Key             =   "QueryAssistent"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":3BA6
            Key             =   "Entwurfsansicht"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":4DF8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":4F52
            Key             =   "Summe"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":52EC
            Key             =   "eMail"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":5446
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":5798
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":5AEA
            Key             =   "Absteigend"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":5E3C
            Key             =   "Aufsteigend"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":618E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":62E8
            Key             =   "LinkControlFile"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":6682
            Key             =   "Word"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":6A1C
            Key             =   "Excel"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":6DB6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueries.ctx":6F10
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraParameter 
      Height          =   4335
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      Begin MSComctlLib.TreeView tvwBaum 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7011
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   0
      End
   End
   Begin VB.Menu mnuQueries 
      Caption         =   "Queries"
      Begin VB.Menu mnuRenameFolder 
         Caption         =   "Ordner umbenennen"
      End
      Begin VB.Menu mnuDeleteFolder 
         Caption         =   "Ordner löschen"
      End
      Begin VB.Menu mnuLine923 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFolderComment 
         Caption         =   "Ordnerbeschreibung bearbeiten ..."
      End
      Begin VB.Menu mnuFolderComment 
         Caption         =   "Ordnerbeschreibung anzeigen"
      End
      Begin VB.Menu mnuLLine2434 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewFolder 
         Caption         =   "Neuer Ordner"
      End
      Begin VB.Menu mnuNewQuery 
         Caption         =   "Neue Abfrage"
         Begin VB.Menu mnuNewQueryAssistent 
            Caption         =   "Assistent ..."
         End
         Begin VB.Menu mnuNewQuerySQLView 
            Caption         =   "Entwurfsansicht ..."
         End
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "Query"
      Begin VB.Menu mnuRenameQuery 
         Caption         =   "Abfrage umbenennen"
      End
      Begin VB.Menu mnuDeleteQuery 
         Caption         =   "Abfrage löschen"
      End
      Begin VB.Menu mnuLine1223 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQueryManagement 
         Caption         =   "Abfrage-Verwaltung"
         Begin VB.Menu mnuCopyQuery 
            Caption         =   "Abfrage kopieren ..."
         End
         Begin VB.Menu mnuMoveQuery 
            Caption         =   "Abfrage verschieben ..."
         End
         Begin VB.Menu mnuGiveQuery 
            Caption         =   "Abfrage schenken ..."
         End
      End
      Begin VB.Menu mnuLine236 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditQueryComment 
         Caption         =   "Abfragebeschreibung bearbeiten ..."
      End
      Begin VB.Menu mnuQueryComment 
         Caption         =   "Abfragebeschreibung anzeigen"
      End
      Begin VB.Menu mnuShowOwner 
         Caption         =   "Eigentümer anzeigen"
      End
      Begin VB.Menu mnuLine289 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecuteQuery 
         Caption         =   "Abfrage ausführen"
      End
      Begin VB.Menu mnuSQLView 
         Caption         =   "Entwurfsansicht"
      End
      Begin VB.Menu mnuLine232 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewParameter 
         Caption         =   "Neuer Parameter"
      End
   End
   Begin VB.Menu mnuParameter 
      Caption         =   "Parameter"
      Begin VB.Menu mnuRenameParameter 
         Caption         =   "Parameter umbenennen"
      End
      Begin VB.Menu mnuDeleteParameter 
         Caption         =   "Parmeter löschen"
      End
      Begin VB.Menu mnuLie82 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditParameterComment 
         Caption         =   "Parameterbeschreibung bearbeiten ..."
      End
      Begin VB.Menu mnuParameterComment 
         Caption         =   "Parameterbeschreibung anzeigen"
      End
   End
   Begin VB.Menu mnuFlexGrid 
      Caption         =   "FlexGrid"
      Begin VB.Menu mnuOpen 
         Caption         =   "Öffnen ..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Speichern ..."
      End
      Begin VB.Menu mnuLine472384 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Drucken"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Suchen ..."
      End
      Begin VB.Menu mnuLine2742 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLinkControlFile 
         Caption         =   "Mit MS Word-Serienbrief verknüpfen ..."
      End
      Begin VB.Menu mnuWord 
         Caption         =   "Mit MS Word veröffentlichen"
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "Analysieren mit MS Excel"
      End
      Begin VB.Menu mnuLine2782 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKopieren 
         Caption         =   "Kopieren"
      End
      Begin VB.Menu mnuLine23827 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecuteQueryFlexView 
         Caption         =   "Abfrage ausführen"
      End
      Begin VB.Menu mnuSQLViewFlexGrid 
         Caption         =   "Entwurfsansicht"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "UserQueriesViewControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : UserQueriesView
'    Project    : UserQueries
'
'    Description: QuserQueriesControl
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Public Event ButtonClick(ByVal Button As MSComctlLib.Button, ByRef Cancel As Boolean)
Public Event DblClick(ByVal Row As Long, ByVal col As Long, ByVal strValue As String)
Public Event ItemClick(ByVal Row As Long, ByVal col As Long, ByVal strValue As String)
Public Event EnterCell(ByVal Row As Long, ByVal col As Long, ByVal strValue As String)
Public Event LeaveCell(ByVal Row As Long, ByVal col As Long, ByVal strValue As String)
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyPress(ByVal KeyAscii As Integer)
Public Event BeforeExecuteQuery(ByRef Cancel As Boolean, ByRef aQuery As UserQueriesViewQuery)
Public Event AfterExecuteQuery(ByRef aQuery As UserQueriesViewQuery)

Private mButtons As MSComctlLib.Buttons

Private mstrClipDialPath As String
Private mstrPrg As String
Private mstrControlFileFolder As String
Private mstrTemplateFolder As String
Private mrsQueryData As Object
Private mblnEnableSMS As Boolean

'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub UserControl_Initialize()
  PrepareModul
End Sub

Private Sub UserControl_Terminate()
Dim aForm As Form

  CloseQueryView
  
  For Each aForm In Forms
    Unload aForm
  Next aForm
End Sub

'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ClipDialPath() As String
  ClipDialPath = mstrClipDialPath
End Property

Public Property Let ClipDialPath(ByVal value As String)
  mstrClipDialPath = value
End Property

Public Property Get Buttons() As MSComctlLib.Buttons
  Set Buttons = tblLeiste.Buttons
End Property

Public Property Let Buttons(ByVal aButtons As MSComctlLib.Buttons)
  Set tblLeiste.Buttons = aButtons
End Property

Public Property Get EnableSMS() As Boolean
  EnableSMS = mblnEnableSMS
End Property

Public Property Let EnableSMS(ByVal blnEnable As Boolean)
  mblnEnableSMS = blnEnable
End Property

Public Property Get SelectedQuery() As UserQueriesViewQuery
  Select Case True
  Case tvwBaum.SelectedItem Is Nothing
    Set SelectedQuery = New UserQueriesViewQuery
  Case TypeName(tvwBaum.SelectedItem.Tag) <> "UserQueriesViewQuery"
    Set SelectedQuery = New UserQueriesViewQuery
  Case Else
    Set SelectedQuery = tvwBaum.SelectedItem.Tag
  End Select
End Property

Public Property Let SelectedQuery(ByVal aQuery As UserQueriesViewQuery)
  Set tvwBaum.SelectedItem.Tag = aQuery
End Property

Public Property Get Prg() As String
  Prg = mstrPrg
End Property

Public Property Let Prg(ByVal strPrg As String)
  mstrPrg = strPrg
End Property

Private Property Get Comment() As String
  Comment = staStatus.Panels("Comment").Text
End Property

Private Property Let Comment(ByVal strComment As String)
  staStatus.Panels("Comment").Text = strComment
End Property

Public Property Get TemplateFolder() As String
  TemplateFolder = mstrTemplateFolder
End Property

Public Property Let TemplateFolder(ByVal strTemplateFolder As String)
  mstrTemplateFolder = strTemplateFolder
End Property

Public Property Get ControlFileFolder() As String
  ControlFileFolder = mstrControlFileFolder
End Property

Public Property Let ControlFileFolder(ByVal strControlFileFolder As String)
  mstrControlFileFolder = strControlFileFolder
End Property

Public Property Get Text() As String
  Text = flexAbfragen.Text
End Property

Public Property Let Text(ByVal strValue As String)
  flexAbfragen.Text = strValue
End Property

Public Property Get Row() As Long
  Row = flexAbfragen.Row
End Property

Public Property Let Row(ByVal lngRow As Long)
  flexAbfragen.Row = lngRow
End Property

Public Property Get col() As Long
  col = flexAbfragen.col
End Property

Public Property Let col(ByVal lngCol As Long)
  flexAbfragen.col = lngCol
End Property

Public Property Get Rows() As Long
  Rows = flexAbfragen.Rows
End Property

Public Property Let Rows(ByVal lngRows As Long)
  flexAbfragen.Rows = lngRows
End Property

Public Property Get Cols() As Long
  Cols = flexAbfragen.Cols
End Property

Public Property Let Cols(ByVal lngCols As Long)
  flexAbfragen.Cols = lngCols
End Property

Public Property Get RowSel() As Long
  RowSel = flexAbfragen.RowSel
End Property

Public Property Let RowSel(ByVal lngRowSel As Long)
  flexAbfragen.RowSel = lngRowSel
End Property

Public Property Get ColSel() As Long
  ColSel = flexAbfragen.ColSel
End Property

Public Property Let ColSel(ByVal lngColSel As Long)
  flexAbfragen.ColSel = lngColSel
End Property

Public Property Get TopRow() As Long
  TopRow = flexAbfragen.TopRow
End Property

Public Property Let TopRow(ByVal lngTopRow As Long)
  flexAbfragen.TopRow = lngTopRow
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------
Private Sub mnuCopyQuery_Click()
  CopyQuery tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuExcel_Click()
  OpenExcel
End Sub

Private Sub mnuExecuteQuery_Click()
  ExecuteQuery tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuExecuteQueryFlexView_Click()
  ExecuteQuery tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuFind_Click()
  Find
End Sub

Private Sub mnuLinkControlFile_Click()
  LinkControlFile Me.TemplateFolder, Me.ControlFileFolder
End Sub

Private Sub mnuMoveQuery_Click()
  MoveQuery tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuNewQuerySQLView_Click()
  NewQuery tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuOpen_Click()
  OpenControlFile flexAbfragen, Me.ControlFileFolder
End Sub

Private Sub mnuPrint_Click()
  BaseToolKit.Controls.flexGrid.PrintData _
  flexAbfragen, 20, 25, 20, 20, "[Abfragen]", Date, poLandscape
End Sub

Private Sub mnuSave_Click()
  SaveControlFile flexAbfragen, Me.ControlFileFolder
End Sub

Private Sub mnuSQLView_Click()
  ShowEntwurfsAnsicht tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuSQLViewFlexGrid_Click()
  ShowEntwurfsAnsicht tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuWord_Click()
  OpenWord
End Sub

Private Sub tblLeiste_ButtonDropDown(ByVal Button As MSComctlLib.Button)
  tblLeiste.Buttons("eMail").ButtonMenus("SendSMS").Enabled _
  = ((Me.EnableSMS) Or (BaseToolKit.WebService.Authentication.PersonId = 1))
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case Shift
  Case 0
    Select Case KeyCode
    Case 116 '{F5}
      InsertRootNode
    End Select
  End Select
End Sub

Private Sub flexAbfragen_Click()
  RaiseEvent ItemClick(flexAbfragen.Row, flexAbfragen.col, flexAbfragen.Text)
End Sub

Private Sub flexAbfragen_DblClick()
  RaiseEvent DblClick(flexAbfragen.Row, flexAbfragen.col, flexAbfragen.Text)
End Sub

Private Sub flexAbfragen_EnterCell()
  RaiseEvent EnterCell(flexAbfragen.Row, flexAbfragen.col, flexAbfragen.Text)
End Sub

Private Sub flexAbfragen_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
  Select Case KeyCode
  Case 67 '{C}
    Select Case Shift
    Case 2 '{Strg}
      CopyViewToClipBoard
    End Select
  Case 46 'Entf
    RemoveRow
  End Select
End Sub

Private Sub flexAbfragen_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub flexAbfragen_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub flexAbfragen_LeaveCell()
  RaiseEvent LeaveCell(flexAbfragen.Row, flexAbfragen.col, flexAbfragen.Text)
End Sub

Private Sub mnuNewFolder_Click()
  NewFolder tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuFolderComment_Click()
  ShowComment tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuEditFolderComment_Click()
  ShowEditComment tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuDeleteFolder_Click()
  DeleteQueryFolder tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuRenameFolder_Click()
  Rename tvwBaum.SelectedItem.Tag
End Sub

Private Sub tvwBaum_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim aNode As MSComctlLib.Node

  Set aNode = tvwBaum.SelectedItem
  
  If aNode Is Nothing Then Exit Sub
  
  Select Case True
  Case (TypeOf aNode.Tag Is UserQueriesViewQuery)
    aNode.Tag.Name = NewString
    aNode.Tag.SaveQuery
  Case (TypeOf aNode.Tag Is UserQueriesViewQueryFolder) And (aNode.Image = "QueryFolder")
    aNode.Tag.FolderName = NewString
    aNode.Tag.SaveQueryFolder
    aNode.Text = " " & NewString
  Case (TypeOf aNode.Tag Is UserQueriesViewParameter)
    aNode.Tag.Name = NewString
    aNode.Tag.SaveParameter
  End Select

  aNode.Parent.Sorted = True
End Sub

Private Sub tvwBaum_BeforeLabelEdit(Cancel As Integer)
Dim aNode As MSComctlLib.Node

  Set aNode = tvwBaum.SelectedItem
  If aNode Is Nothing Then Exit Sub
  
  aNode.Text = Trim(aNode.Text)
End Sub

Private Sub tvwBaum_KeyDown(KeyCode As Integer, Shift As Integer)
Dim aNode As MSComctlLib.Node

  Set aNode = tvwBaum.SelectedItem
  
  If aNode Is Nothing Then Exit Sub
  
  Select Case True
  '{Rename}
  Case ((KeyCode = vbKeyF2) And (TypeOf aNode.Tag Is UserQueriesViewQuery)) _
  Or ((KeyCode = vbKeyF2) And (TypeOf aNode.Tag Is UserQueriesViewQueryFolder) And (aNode.Image = "QueryFolder")) _
  Or ((KeyCode = vbKeyF2) And (TypeOf aNode.Tag Is UserQueriesViewParameter))
    Rename aNode.Tag
  Case KeyCode = vbKeyDelete
    PrepareDelete aNode
  Case KeyCode = vbKeyInsert
    PrepareInsert aNode
  Case KeyCode = 93
    ShowPopUpMenu tvwBaum.SelectedItem
  Case (Shift = 2) And (KeyCode = 81) And (mnuExecuteQuery.Enabled) '{Strg+Q}
    ExecuteQuery aNode.Tag
  End Select
  
End Sub

Private Sub tblLeiste_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

  Select Case ButtonMenu.Key
  Case "Mail"
    SendEMail
  Case "CallNumber"
    CallNumber
  Case "SendSMS"
    SendSMS
  End Select
End Sub

Private Sub flexAbfragen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button <> 2 Then Exit Sub
  EnableMenu tvwBaum.SelectedItem
  PopupMenu mnuFlexGrid
End Sub

Private Sub mnuGiveQuery_Click()
  GiveQuery tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuKopieren_Click()
  CopyViewToClipBoard
End Sub

Private Sub cboLimit_Validate(Cancel As Boolean)
  If Not IsNumeric(cboLimit.Text) Then
    cboLimit.Text = 0
  Else
    cboLimit.Text = CLng(cboLimit.Text)
  End If
End Sub

Private Sub tvwBaum_DblClick()
Dim SelectedNode As MSComctlLib.Node
Dim aQuery As UserQueriesViewQuery

 On Error GoTo errLabel:

 Set SelectedNode = tvwBaum.SelectedItem
   
 If TypeName(SelectedNode.Tag) <> "String" Then Exit Sub
  
  Select Case SelectedNode.Image
  Case "Field", "Value", "Operator", "Parametertyp"

    Set aQuery = SelectedNode.Parent.Parent.Tag
    
    Select Case aQuery.QueryType
    Case eqqtUserQuery
    Case eqqtOpenQuery
      If (BaseToolKit.WebService.Authentication.PersonId <> 1) _
      And (aQuery.OwnerID <> BaseToolKit.WebService.Authentication.PersonId) Then
        MsgBox "Kopieren Sie bitte die Abfrage!", 16, "Zugriff verweigert"
        Exit Sub
      End If
    Case eqqtAdminQuery
      If BaseToolKit.WebService.Authentication.PersonId <> 1 Then
        MsgBox "Kopieren Sie bitte die Abfrage!", 16, "Zugriff verweigert"
        Exit Sub
      End If
    End Select
  End Select
  
  Select Case SelectedNode.Tag
  Case "Field"
    ChangeField SelectedNode _
    , SelectedNode.Parent.Tag, SelectedNode.Parent.Parent.Tag
  Case "Value"
    ChangeValue SelectedNode, SelectedNode.Parent.Tag
  Case "Operator"
    ChangeOperator SelectedNode, SelectedNode.Parent.Tag
  Case "Parametertyp"
    ChangeParameterType SelectedNode, SelectedNode.Parent.Tag
  End Select
  Exit Sub
  
errLabel:
  ShowError ""
  Exit Sub
End Sub

Private Sub fraParameter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim intDeltaWidth As Integer

On Error GoTo errLabel:
    
  If (x >= fraParameter.Width - 50) Or (Button = 1) Then
    fraParameter.MousePointer = 9
  Else
    fraParameter.MousePointer = 0
  End If

  If (x >= 2000) And (x <= UserControl.Width - 2000) Then
    If (Button = 1) And (fraParameter.MousePointer = 9) Then
      intDeltaWidth = fraParameter.Width - x
      fraParameter.Width = x
      tvwBaum.Width = fraParameter.Width - (tvwBaum.Left * 2)
      
      fraSQL.Left = x + 60
      fraSQL.Width = fraSQL.Width + intDeltaWidth
      
      flexAbfragen.Width = fraSQL.Width - (flexAbfragen.Left * 2)
      
      '{Neuen Prozentwert speichern}
      fraParameter.Tag = Replace(fraParameter.Width / UserControl.ScaleWidth, ",", ".")
    End If
  End If
  Exit Sub
  
errLabel:
  ShowError "ReziseQueryView"
  Exit Sub
End Sub

Private Sub staStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)
  ShowHideParameters Panel
End Sub

Private Sub tvwBaum_NodeClick(ByVal Node As MSComctlLib.Node)
  ActionSelectionQueryTree Node
End Sub

Private Sub UserControl_Resize()
Dim sngBreite As Single

On Error Resume Next

  With UserControl
    If .Width < 4500 Then .Width = 4500
    If .Height < 4500 Then .Height = 4500
  
  
    sngBreite = CSng(Replace(.fraParameter.Tag, ".", ","))
  
    .fraParameter.Width = .ScaleWidth * sngBreite

    .fraParameter.Height = .ScaleHeight - (.tblLeiste.Height + 500)
    .tvwBaum.Height = .fraParameter.Height - ((.tvwBaum.Top * 1.5))
    .tvwBaum.Width = .fraParameter.Width - (.tvwBaum.Left * 2)
    
    
    .fraSQL.Left = (.ScaleWidth * sngBreite) + 60
    .fraSQL.Width = .ScaleWidth - (.fraParameter.Width + 20)
    
    .fraSQL.Width = .ScaleWidth - .fraSQL.Left
    .fraSQL.Height = .ScaleHeight - (.tblLeiste.Height + 500)
    .flexAbfragen.Width = .fraSQL.Width - (.flexAbfragen.Left * 2)
    .flexAbfragen.Height = .fraSQL.Height - (.flexAbfragen.Top * 1.5)
    
    cboLimit.Left = .ScaleWidth - cboLimit.Width
  End With

End Sub

Private Sub tvwBaum_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  fraParameter.MousePointer = 0
End Sub

Private Sub mnuDeleteParameter_Click()
  DeleteParameter tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuDeleteQuery_Click()
  DeleteQuery tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuEditParameterComment_Click()
  ShowEditComment tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuEditQueryComment_Click()
  ShowEditComment tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuNewParameter_Click()
  NewParameter tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuParameterComment_Click()
  ShowComment tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuQueryComment_Click()
  ShowComment tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuRenameParameter_Click()
  Rename tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuRenameQuery_Click()
  Rename tvwBaum.SelectedItem.Tag
End Sub

Private Sub mnuShowOwner_Click()
Dim aQuery As UserQueriesViewQuery

  Set aQuery = tvwBaum.SelectedItem.Tag
  MsgBox aQuery.Owner & " (" & aQuery.OwnerID & ")"
End Sub

Private Sub tblLeiste_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim blnCancel As Boolean

  RaiseEvent ButtonClick(Button, blnCancel)

  If blnCancel Then Exit Sub

  Select Case Button.Key
  Case "Open"
    OpenControlFile flexAbfragen, Me.ControlFileFolder
  Case "Save"
    SaveControlFile flexAbfragen, Me.ControlFileFolder
  Case "Find"
    Find
  Case "NewQuery"
    NewQuery tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
  Case "DeleteQuery"
    DeleteQuery tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
  Case "NewParameter"
    NewParameter tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
  Case "DeleteParameter"
    DeleteParameter tvwBaum.SelectedItem, tvwBaum.SelectedItem.Tag
  Case "Print"
    BaseToolKit.Controls.flexGrid.PrintData _
    flexAbfragen, 20, 25, 20, 20, "[Abfragen]", Date, poLandscape
  Case "eMail"
    SendEMail
  Case "LinkControlFile"
    LinkControlFile Me.TemplateFolder, Me.ControlFileFolder
  Case "Entwurfsansicht"
    ShowEntwurfsAnsicht tvwBaum.SelectedItem.Tag
  Case "Abfrage"
    ExecuteQuery tvwBaum.SelectedItem.Tag
  Case "Aufsteigend"
    flexAbfragen.Sort = 1 '{Aufsteigend sortieren}
  Case "Absteigend"
    flexAbfragen.Sort = 2 '{Absteigend sortieren}
  Case "Summe"
    MsgBox CreateSum, 64, "Summe"
  Case "Word"
    OpenWord
  Case "Excel"
    OpenExcel
  End Select
End Sub

Private Sub tvwBaum_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim SelectedNode As MSComctlLib.Node

  If Button <> 2 Then Exit Sub
  
  Set SelectedNode = tvwBaum.HitTest(x, Y)
  
  If SelectedNode Is Nothing Then Exit Sub
  
  Set tvwBaum.SelectedItem = SelectedNode
  
  ShowPopUpMenu SelectedNode
End Sub
  


'---------------------- Private Methoden der Klasse -----------------------------
Private Function SelectQueryFolder(ByVal strTransfer As String) As UserQueriesViewQueryFolder
Dim aPRGs As UserQueriesViewPrgs
Dim colTemp As Collection
Dim colSelectData As Collection
Dim strPrg As String
Dim aQueryFolders As UserQueriesViewQueryFolders
Dim aQueryFolder As UserQueriesViewQueryFolder
Dim strFolderID As String
Dim strCurrentFolderID As String
Dim strFullPath As String
Dim blnCheckUserID As Boolean

On Error GoTo errLabel:

  '{Progamme anzeigen}
  Set aPRGs = New UserQueriesViewPrgs
  aPRGs.UserFID = BaseToolKit.WebService.Authentication.PersonId
  aPRGs.GetPrgsDB
    
  Set colTemp = aPRGs.GetPrgs
  
  BaseToolKit.Dialog.SelectEntry.Reset
  BaseToolKit.Dialog.SelectEntry.SelectEntry colTemp _
  , "Bitte Programmordner wählen ...", False, True, , Me.Prg
  If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Function
    
  strPrg = BaseToolKit.Dialog.SelectEntry.ValueEntry
  strFullPath = BaseToolKit.Dialog.SelectEntry.ValueEntry & "\"
  '{Ende Progamme anzeigen}
  
  
  '{Hauptordner anzeigen}
  Set colTemp = New Collection
  colTemp.Add "Meine Abfragen"
  colTemp.Add eqqtUserQuery
  colTemp.Add "Administrator-Abfragen"
  colTemp.Add eqqtAdminQuery
  colTemp.Add "Veröffentlichte Abfragen"
  colTemp.Add eqqtOpenQuery
  
  BaseToolKit.Dialog.SelectEntry.Reset
  BaseToolKit.Dialog.SelectEntry.SelectEntry colTemp, strFullPath, True, True
  If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Function

  blnCheckUserID = BaseToolKit.Dialog.SelectEntry.ValueID = eqqtUserQuery
  strFolderID = BaseToolKit.Dialog.SelectEntry.ValueID
  strFullPath = strFullPath & BaseToolKit.Dialog.SelectEntry.ValueEntry & "\"
  '{Ende Hauptordner anzeigen}

  
  '{Unterordner anzeigen bis einer gewählt wurde}
  strCurrentFolderID = strFolderID
  
  Do
    strFolderID = strCurrentFolderID

    Set aQueryFolders = New UserQueriesViewQueryFolders
      
    Select Case blnCheckUserID
    Case True
      aQueryFolders.GetItems strFolderID, strPrg, BaseToolKit.WebService.Authentication.PersonId
    Case False
      aQueryFolders.GetItems strFolderID, strPrg
    End Select
    
    Set colTemp = New Collection
    colTemp.Add "<Hierhin " & strTransfer & ">"
    colTemp.Add "-1"
  
    For Each aQueryFolder In aQueryFolders.Items
      colTemp.Add aQueryFolder.FolderName
      colTemp.Add aQueryFolder.FolderID
    Next aQueryFolder
    
    BaseToolKit.Dialog.SelectEntry.Reset
    BaseToolKit.Dialog.SelectEntry.SelectEntry colTemp, strFullPath, True, True, , , True
    If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Function
    strCurrentFolderID = BaseToolKit.Dialog.SelectEntry.ValueID
    strFullPath = strFullPath & BaseToolKit.Dialog.SelectEntry.ValueEntry & "\"
    
    If Not aQueryFolder Is Nothing Then
      blnCheckUserID = aQueryFolder.QueryType = eqqtUserQuery
    End If
  Loop Until (strCurrentFolderID = "-1")
  '{Ende Unterordner anzeigen bis einer gewählt wurde}
  
  Select Case CLng(strFolderID)
  Case eqqtUserQuery
    Set SelectQueryFolder = tvwBaum.Nodes("Root\" & strPrg & "\" & eqqtUserQuery & "Meine Abfragen").Tag
  Case eqqtAdminQuery
    Set SelectQueryFolder = tvwBaum.Nodes("Root\" & strPrg & "\" & eqqtAdminQuery & "Administrator-Abfragen").Tag
  Case eqqtOpenQuery
    Set SelectQueryFolder = tvwBaum.Nodes("Root\" & strPrg & "\" & eqqtOpenQuery & "Veröffentlichte Abfragen").Tag
  Case Else
    Set SelectQueryFolder = New UserQueriesViewQueryFolder
    SelectQueryFolder.GetQueryFolder strFolderID
  End Select
  Exit Function
  
errLabel:
  ShowError "SelectQueryFolder"
  Exit Function
End Function

Private Function CheckGrants(ByRef obj As Object _
, Optional ByRef blnSilentMode As Boolean = False) As Boolean
  
  If IsAdmin Then
    CheckGrants = True
    Exit Function
  End If
  
  If IsOwner(obj) Then
    CheckGrants = True
    Exit Function
  End If
  
  CheckGrants = False
  If blnSilentMode Then Exit Function
  
  MsgBox "Zugriff verweigert", 48, "Keine Berechtigung!"
  
End Function

Private Function IsAdmin() As Boolean
  IsAdmin = (BaseToolKit.WebService.Authentication.PersonId = 1)
End Function

Private Function IsOwner(ByRef obj As Object) As Boolean
  'IsOwner = (obj.OwnerID = BaseToolKit.WebService.Authentication.PersonId) _
  'Or ((TypeName(obj) = "UserQueriesViewQueryFolder") And (obj.QueryType = eqQueryType.eqqtOpenQuery))
  
  Dim check1Ok As Boolean: check1Ok = False
  Dim check2Ok As Boolean: check2Ok = False
  
  If TypeName(obj) = "UserQueriesViewQueryFolder" Then
    check1Ok = (obj.QueryType = eqQueryType.eqqtOpenQuery)
  End If
  
  check2Ok = (obj.OwnerID = BaseToolKit.WebService.Authentication.PersonId)
  
  IsOwner = check1Ok Or check2Ok
  
End Function

Private Sub MoveQuery(ByRef aQuery As UserQueriesViewQuery)
Dim aNewQuery As UserQueriesViewQuery
Dim colP As Collection
Dim aParameter As UserQueriesViewParameter
Dim aQueryFolder As UserQueriesViewQueryFolder

On Error GoTo errLabel:

  If Not CheckGrants(aQuery) Then Exit Sub

  Set aQueryFolder = SelectQueryFolder("verschieben")
  If aQueryFolder Is Nothing Then Exit Sub
  
  If (aQueryFolder.QueryType = eqqtAdminQuery) And (Not IsAdmin) Then
    MsgBox "Zugriff verweigert", 48, "Keine Berechtigung!"
    Exit Sub
  End If
  
  Set aQuery = tvwBaum.SelectedItem.Tag
  aQuery.GetParametersDB
  Set colP = aQuery.GetParameters
  
  Set aNewQuery = tvwBaum.SelectedItem.Tag
  aNewQuery.QueryFolderFID = aQueryFolder.FolderID
  aNewQuery.Prg = aQueryFolder.Prg
  aNewQuery.OwnerID = BaseToolKit.WebService.Authentication.PersonId
  aNewQuery.QueryType = aQueryFolder.QueryType
  aNewQuery.SaveQuery
  
  For Each aParameter In colP
    aParameter.OwnerID = BaseToolKit.WebService.Authentication.PersonId
    aParameter.QueryFID = aNewQuery.QueryID
    aParameter.SaveParameter
  Next aParameter
  
  InsertRootNode
  Exit Sub
  
errLabel:
  ShowError "MoveQuery"
  Exit Sub
End Sub

Private Sub CopyQuery(ByRef aQuery As UserQueriesViewQuery)
Dim aNewQuery As UserQueriesViewQuery
Dim colP As Collection
Dim aParameter As UserQueriesViewParameter
Dim aQueryFolder As UserQueriesViewQueryFolder

On Error GoTo errLabel:

  Set aQueryFolder = SelectQueryFolder("kopieren")
  If aQueryFolder Is Nothing Then Exit Sub
  
  If (aQueryFolder.QueryType = eqqtAdminQuery) And (Not IsAdmin) Then
    MsgBox "Zugriff verweigert", 48, "Keine Berechtigung!"
    Exit Sub
  End If
  
  Set aQuery = tvwBaum.SelectedItem.Tag
  aQuery.GetParametersDB
  Set colP = aQuery.GetParameters
  
  Set aNewQuery = tvwBaum.SelectedItem.Tag
  aNewQuery.QueryID = "-1"
  aNewQuery.QueryFolderFID = aQueryFolder.FolderID
  aNewQuery.Prg = aQueryFolder.Prg
  aNewQuery.OwnerID = BaseToolKit.WebService.Authentication.PersonId
  aNewQuery.QueryType = aQueryFolder.QueryType
  aNewQuery.SaveQuery
  
  For Each aParameter In colP
    aParameter.ParameterID = "-1"
    aParameter.OwnerID = BaseToolKit.WebService.Authentication.PersonId
    aParameter.QueryFID = aNewQuery.QueryID
    aParameter.SaveParameter
  Next aParameter
  
  InsertRootNode
  Exit Sub
  
errLabel:
  ShowError "CopyQuery"
  Exit Sub
End Sub

Private Sub CopyViewToClipBoard()
Dim strView As String
Dim strRow As String
Dim x As Long
Dim Y As Long
Dim z1 As Long
Dim z2 As Long
Dim s1 As Long
Dim s2 As Long
  
  With flexAbfragen
  
    If .Row > .RowSel Then
      z1 = .RowSel
      z2 = .Row
    Else
      z1 = .Row
      z2 = .RowSel
    End If
    
    If .col > .ColSel Then
      s1 = .ColSel
      s2 = .col
    Else
      s1 = .col
      s2 = .ColSel
    End If
    
    For Y = z1 To z2
      strRow = ""
      For x = s1 To s2
        strRow = strRow & .TextMatrix(Y, x) & vbTab
      Next x
      strView = strView & strRow & vbCrLf
    Next Y
    
    Clipboard.Clear
    Clipboard.SetText strView
  End With

End Sub

Private Sub OpenControlFile _
(ByRef aFlexGrid As MSFlexGrid _
, Optional ByVal strInitialControlFileFolder As String = "")

'Dim Temp As String
'Dim astrCols() As String
'Dim i As Integer

On Error GoTo errLabel:


  With aFlexGrid
    Dim strKey As String: strKey _
    = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

    Dim strHomeDir As String: strHomeDir _
    = BaseToolKit.Win32API.Win32ApiProfessional.SysRegistry.GetRegistryValue _
    (enumHKEY_CURRENT_USER, strKey, "Personal")
    '{Ende Ordner 'Eigene Dateien' ermitteln}

    If strInitialControlFileFolder = "" Then
      strInitialControlFileFolder = strHomeDir
    End If

    '{Steuerdatei abfragen}
    Dim strControlFile As String: strControlFile _
    = BaseToolKit.Win32API.Win32ApiProfessional.SysDialogsEx.GetOpenFileNameEx _
    (strInitialControlFileFolder, "", "txt,csv,xml" _
    , "Bitte Steuerdatei wählen ...", UserControl.hwnd)

    If strControlFile = "" Then Exit Sub
    '{Ende Steuerdatei abfragen}

    Screen.MousePointer = 11

    Dim rs As Object
    If BaseToolKit.FileSystem.io.GetExtensionName(LCase(strControlFile)) = "xml" Then
      Set rs = CreateObject("ADODB.RecordSet"): rs.Open strControlFile
      FillFlexGrid rs
    Else
      .Rows = 0
      
      Dim content As String: content = BaseToolKit.FileSystem.io.ReadAllText(strControlFile)
      Dim colData As Collection: Set colData = BaseToolKit.Convert.SplitCollection(content, vbCrLf)

      If colData.Count > 0 Then
        .Cols = BaseToolKit.Convert.SplitCollection(colData.Item(1), ";").Count
      End If

      Dim R As Variant
      For Each R In colData
        .AddItem Replace$(Replace$(R, ";", vbTab), """", "")
      Next R

      If .Rows = 1 Then
        .AddItem ""
        fraSQL.Caption = "Abfrageergebnis (Datensätze 0)"
      Else
        fraSQL.Caption = "Abfrageergebnis (Datensätze " & .Rows - 1 & ")"
      End If

      .FixedRows = 1

    End If
  End With
  EnableMenu tvwBaum.SelectedItem
  Screen.MousePointer = 0
  Exit Sub

errLabel:
  ShowError "OpenControlFile"
  Exit Sub
End Sub

Private Sub ControlFileToHTML _
(ByVal strOutFileName As String _
, ByVal strControlFile As String _
, ByVal strDelimiter As String)

Const TIEFES_BLAU = "#C6C6C6"
Const WEISS = "#FFFFFF"
Const GRAU = "#E5E5E5"
Const SCHWARZ = "#000000"

'
'

On Error GoTo errLabel:

  Dim strTableName As String: strTableName _
  = Mid(strControlFile, InStrRev(strControlFile, "\") + 1)
  strTableName = Mid(strTableName, 1, Len(strTableName) - 4)

  Dim colData As Collection: Set colData = New Collection

  colData.Add "<HTML>"
  colData.Add "  <HEAD>"
  colData.Add "    <TITLE>"
  colData.Add "      crm: crm"
  colData.Add "    </TITLE>"
  colData.Add "  </HEAD>"

  colData.Add "  <BODY LEFTMARGIN='0' TOPMARGIN='0' BGCOLOR='#FFFFFF' TEXT='#000000' LINK='#0000FF' VLINK='#0000FF' ALINK='#0000FF'>"

  colData.Add "    <table border>"
  colData.Add "    <caption align=top>" & strTableName & "</caption>"

  Dim strBGColor As String: strBGColor = TIEFES_BLAU
  Dim strFColor As String: strFColor = WEISS

  Dim colSource As Collection: Set colSource _
  = BaseToolKit.Convert.SplitCollection _
  (BaseToolKit.FileSystem.io.ReadAllText(strControlFile), vbCrLf)

  Dim intAnzahl As Integer
  
  Dim R As Variant
  For Each R In colSource
  
    intAnzahl = intAnzahl + 1
    colData.Add "      <tr bgcolor='" & strBGColor & "'>"
    
    Dim strZeile As String: strZeile = R
    Dim astrFields() As String: astrFields = Split(strZeile, strDelimiter)
    
    Dim i As Integer
    For i = LBound(astrFields) To UBound(astrFields)
      Dim strField As String
      Dim strAlign As String
      
      strField = Replace(astrFields(i), """", "")
      If strField = "" Then
        strField = "---"
        strAlign = "center"
      Else
        strAlign = "left"
      End If

      colData.Add "          <th align=" & strAlign & ">" & strField & "</th>"
    Next i
    colData.Add "      </tr>"
    
    strFColor = SCHWARZ
    Select Case intAnzahl Mod 2
    Case 1
      strBGColor = WEISS
    Case 0
      strBGColor = GRAU
    End Select
  Next R

  colData.Add "    </table>"
  colData.Add "  </BODY>"
  colData.Add "</HTML>"

  Dim content As String: content = BaseToolKit.Convert.JoinCollection(colData, vbCrLf)
  BaseToolKit.FileSystem.io.WriteAllText strOutFileName, content, False
  Exit Sub

errLabel:
  ShowError "ControlFileToHTML"
  Exit Sub
End Sub

Private Sub RemoveRow()
  With flexAbfragen
    Select Case .Rows
    Case 0 To 1
      Exit Sub
    Case Is = 2
      .FixedRows = 0
      .RemoveItem 1
    Case Else
      .RemoveItem .Row
    End Select
  End With
End Sub

Private Sub Rename(ByRef obj As Object)
  Select Case TypeName(obj)
  Case "UserQueriesViewQueryFolder"
    If CheckGrants(obj, True) Or CheckPublicFolder(obj) Then tvwBaum.StartLabelEdit
  Case Else
    If CheckGrants(obj) Then tvwBaum.StartLabelEdit
  End Select
End Sub

Private Sub PrepareDelete(ByRef SelectedNode As MSComctlLib.Node)
On Error GoTo errLabel:

  Select Case TypeName(SelectedNode.Tag)
  Case "UserQueriesViewQueryFolder"
    DeleteQueryFolder SelectedNode, SelectedNode.Tag
  Case "UserQueriesViewQuery"
    DeleteQuery SelectedNode, SelectedNode.Tag
  Case "UserQueriesViewParameter"
    DeleteParameter SelectedNode, SelectedNode.Tag
  End Select
  Exit Sub
  
errLabel:
  ShowError "PrepareDelete"
  Exit Sub
End Sub
  
Private Sub PrepareInsert(ByRef SelectedNode As MSComctlLib.Node)
On Error GoTo errLabel:

  Select Case TypeName(SelectedNode.Tag)
  Case "UserQueriesViewQueryFolder"
    NewQuery SelectedNode, SelectedNode.Tag
  Case "UserQueriesViewQuery"
    NewParameter SelectedNode, SelectedNode.Tag
  End Select
  Exit Sub
  
errLabel:
  ShowError "PrepareInsert"
  Exit Sub
End Sub

Private Sub Find()
Dim x As Long
Dim Y As Long
Dim blnCaseSensible As Boolean
Dim blnWholeWord As Boolean
Dim strSearchString As String
Dim strTextMatrix As String
Dim blnFound As Boolean
Dim lngNormalBackColor As Long
Dim lngNormalForeColor As Long

  UserQueriesViewFindDialog.Show 1
  
  If UserQueriesViewFindDialog.Cancel = True Then Exit Sub
  
  blnCaseSensible = UserQueriesViewFindDialog.CaseSensible
  blnWholeWord = UserQueriesViewFindDialog.WholeWord
  strSearchString = UserQueriesViewFindDialog.SerachString
  
  With flexAbfragen
      
    For Y = GetStartRow To GetEndRow
      For x = GetStartCol To GetEndCol
      
        strTextMatrix = .TextMatrix(Y, x)
        
        Select Case blnCaseSensible
        Case True
        Case False
          strSearchString = LCase(strSearchString)
          strTextMatrix = LCase(strTextMatrix)
        End Select
        
        Select Case blnWholeWord
        Case True
          blnFound = strTextMatrix = strSearchString
        Case False
          blnFound = InStr(strTextMatrix, strSearchString) > 0
        End Select
        
        If blnFound Then
          .FocusRect = flexFocusHeavy
          .Row = Y
          .RowSel = Y
          .col = x
          .ColSel = x
          .TopRow = Y
          
          lngNormalForeColor = .CellForeColor
          lngNormalBackColor = .CellBackColor
          .CellForeColor = vbHighlightText
          .CellBackColor = vbHighlight
          
          If MsgBox("Weitersuchen?", 36, "Suchen") = vbNo Then
            flexAbfragen.SetFocus
            .CellForeColor = lngNormalForeColor
            .CellBackColor = lngNormalBackColor
            Exit Sub
          End If
          
          .CellForeColor = lngNormalForeColor
          .CellBackColor = lngNormalBackColor
        
        End If
      Next x
    Next Y
  End With
End Sub

Private Function GetFlexData()
Dim c As Long
Dim R As Long
Dim strRow As String
Dim strData As String
Dim strCell As String

On Error GoTo errLabel:
  
  With flexAbfragen
    For R = 0 To .Rows - 1
      strRow = ""
      For c = 0 To .Cols - 1
        strCell = Replace(.TextMatrix(R, c), Chr(10), "")
        strCell = Replace(strCell, Chr(13), " ")
        
        strRow = strRow & strCell & vbTab
      Next c
      strData = strData & strRow & vbCrLf
    Next R
  End With
  
  GetFlexData = strData
  Exit Function
  
errLabel:
  ShowError "GetFlexData"
  Exit Function
End Function

Private Sub OpenExcel()
Dim aExcel As Object
  
On Error GoTo errLabel:
    
  Clipboard.Clear
  Clipboard.SetText GetFlexData

  Set aExcel = CreateObject("excel.application")
  aExcel.Workbooks.Add
  aExcel.Range("A1").Select
  aExcel.ActiveSheet.Paste
  aExcel.Cells.Select
  aExcel.Selection.Columns.AutoFit
  aExcel.Visible = True
  Exit Sub
  
errLabel:
  ShowError "OpenExcel"
  Exit Sub
End Sub

Private Sub OpenWord()
Dim aWord As Object

On Error GoTo errLabel:

  Set aWord = CreateObject("word.application")

  With aWord
    .Documents.Add
    .Visible = True
    Clipboard.Clear
    Clipboard.SetText GetFlexData
    .Selection.Paste
    .Activate
  End With
  Exit Sub
    
errLabel:
  ShowError "OpenWord"
  Exit Sub
End Sub

Private Sub ShowPopUpMenu(ByRef SelectedNode As MSComctlLib.Node)
  
  If SelectedNode Is Nothing Then Exit Sub
  
  Select Case TypeName(SelectedNode.Tag)
  Case "UserQueriesViewPrgs" '{Root-Node}
  Case "String"  '{Prg-Node}
  Case "UserQueriesViewQueryFolder"
    PopupMenu mnuQueries
  Case "UserQueriesViewQueries"
  Case "UserQueriesViewQuery"
    PopupMenu mnuQuery
  Case "UserQueriesViewParameter"
    PopupMenu mnuParameter
  End Select
End Sub

Private Sub ExecuteQuery(ByRef aQuery As UserQueriesViewQuery)

On Error GoTo errLabel:
  
  Dim Cancel As Boolean
  RaiseEvent BeforeExecuteQuery(Cancel, aQuery)
  
  If Cancel Then Exit Sub
   
  Set aQuery = tvwBaum.SelectedItem.Tag
  
  Set mrsQueryData = aQuery.ExecuteQuery(cboLimit.Text)
  
  'Protokoll anlegen
  Dim strSQL As String
  strSQL = "Insert into logs.SQLProtokoll "
  strSQL = strSQL & "(USERID,SQLBefehl,Programm) VALUES "
  strSQL = strSQL & "(" & BaseToolKit.WebService.Authentication.PersonId _
  & ",'" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(aQuery.LastExecutedStatement) _
  & "','" & Me.Prg & "')"
  
  BaseToolKit.Database.ExecuteNonQuery strSQL
  'Ende Protokoll schreiben
  
  If Not mrsQueryData Is Nothing Then
    FillFlexGrid mrsQueryData
    RaiseEvent AfterExecuteQuery(aQuery)
    EnableMenu tvwBaum.SelectedItem
  End If
  Exit Sub
  
errLabel:
  ShowError "ExecuteQuery"
  Exit Sub
End Sub

Private Sub InsertRootNode()
Dim aPRGs As UserQueriesViewPrgs
Dim aQueries As UserQueriesViewQueries
Dim colPrgs As Collection
Dim x As Variant
Dim RootNode As MSComctlLib.Node
Dim aNode As MSComctlLib.Node
Dim aChildNode As MSComctlLib.Node
Dim strNodeKey As String
Dim aQueryFolder As UserQueriesViewQueryFolder
Dim strUser As String

On Error GoTo errLabel:
  
  With UserControl
        
    .tvwBaum.Nodes.Clear
    .tvwBaum.ImageList = .ilsBilder
    
    '{Root-Node einfügen}
    Set aPRGs = New UserQueriesViewPrgs
    aPRGs.UserFID = BaseToolKit.WebService.Authentication.PersonId
    aPRGs.GetPrgsDB
    
    
    Set RootNode = .tvwBaum.Nodes.Add
    RootNode.Text = "BCW-Verwaltungsprogramme"
    Set RootNode.Tag = aPRGs
    RootNode.Key = "Root"
    RootNode.Image = "Root"
    '{Ende Root-Node einfügen}
    
    
    '{Verwaltungprogrammauflistung durchlaufen}
    '{und Ebene in Baum einfügen}
    Set colPrgs = aPRGs.GetPrgs
    
    strUser = BaseToolKit.WebService.Authentication.FullName
    
    For Each x In colPrgs
      '{Programmebene einfügen}
      Set aNode = tvwBaum.Nodes.Add(RootNode.Key, tvwChild)
      aNode.Text = x
      aNode.Key = RootNode.Key & "\" & x
      aNode.Tag = "prg"
      aNode.Image = "prg"
      '{Ende Programmebene einfügen}
      
      '{Knoten Meine Abfragen einfügen}
      Set aQueryFolder = New UserQueriesViewQueryFolder
      aQueryFolder.FolderID = eqqtUserQuery
      aQueryFolder.FolderName = "Meine Abfragen"
      aQueryFolder.ParentFolderID = "-1"
      aQueryFolder.QueryType = eqqtUserQuery
      aQueryFolder.Prg = x
      aQueryFolder.Image = "UserQueries"
      aQueryFolder.OwnerID = BaseToolKit.WebService.Authentication.PersonId
      aQueryFolder.Owner = strUser
      
      Set aChildNode = tvwBaum.Nodes.Add(aNode.Key, tvwChild)
      aChildNode.Text = aQueryFolder.FolderName
      aChildNode.Key = aNode.Key & "\" & aQueryFolder.FolderID & aQueryFolder.FolderName
      aChildNode.Image = aQueryFolder.Image
      Set aChildNode.Tag = aQueryFolder
      '{Ende Knoten Meine Abfragen einfügen}
    
      '{Knoten Administrator-Abfragen einfügen}
      Set aQueryFolder = New UserQueriesViewQueryFolder
      aQueryFolder.FolderID = eqqtAdminQuery
      aQueryFolder.FolderName = "Administrator-Abfragen"
      aQueryFolder.ParentFolderID = "-1"
      aQueryFolder.QueryType = eqqtAdminQuery
      aQueryFolder.Prg = x
      aQueryFolder.Image = "AdminQueries"
      aQueryFolder.OwnerID = "1"
      aQueryFolder.Owner = "root"
      
      Set aChildNode = tvwBaum.Nodes.Add(aNode.Key, tvwChild)
      aChildNode.Text = aQueryFolder.FolderName
      aChildNode.Key = aNode.Key & "\" & aQueryFolder.FolderID & aQueryFolder.FolderName
      aChildNode.Image = aQueryFolder.Image
      Set aChildNode.Tag = aQueryFolder
      '{Ende Knoten Administrator-Abfragen einfügen}
        
      '{Knoten veröffentlichte Abfragen einfügen}
      Set aQueryFolder = New UserQueriesViewQueryFolder
      aQueryFolder.FolderID = eqqtOpenQuery
      aQueryFolder.FolderName = "Veröffentlichte Abfragen"
      aQueryFolder.ParentFolderID = "-1"
      aQueryFolder.QueryType = eqqtOpenQuery
      aQueryFolder.Prg = x
      aQueryFolder.Image = "OpenQueries"
      aQueryFolder.OwnerID = "1"
      aQueryFolder.Owner = "root"
      
      Set aChildNode = tvwBaum.Nodes.Add(aNode.Key, tvwChild)
      aChildNode.Text = aQueryFolder.FolderName
      aChildNode.Key = aNode.Key & "\" & aQueryFolder.FolderID & aQueryFolder.FolderName
      aChildNode.Image = aQueryFolder.Image
      Set aChildNode.Tag = aQueryFolder
      '{Ende Knoten veröffentlichte Abfrtagen einfügen}
                
    Next x
    
    '{Root-Knoten aufklappen}
    RootNode.Expanded = True
    
    strNodeKey = RootNode.Key & "\" & Me.Prg & "\" & "0Meine Abfragen"
    Set tvwBaum.SelectedItem = tvwBaum.Nodes(strNodeKey)
    tvwBaum.SelectedItem.Expanded = True
  End With
  ActionSelectionQueryTree tvwBaum.SelectedItem
  Exit Sub
  
errLabel:
  ShowError "InsertRootNode"
  Exit Sub
End Sub

Private Sub ShowComment(ByRef aObj As Object)

On Error GoTo errLabel:
  
  If aObj.Comment = "" Then Exit Sub
  
  Select Case TypeName(aObj)
  Case "UserQueriesViewQuery"
    MsgBox DivideString(aObj.Comment, 45), 64, "Abfragebeschreibung"
  Case "UserQueriesViewParameter"
    MsgBox DivideString(aObj.Comment, 45), 64, "Parameterbeschreibung"
  Case "UserQueriesViewQueryFolder"
    MsgBox DivideString(aObj.Comment, 45), 64, "Odnerbeschreibung"
  End Select
  Exit Sub
  
errLabel:
  ShowError "ShowComent"
  Exit Sub
End Sub

Private Sub ShowEditComment(ByRef aObj As Object)
On Error GoTo errLabel:

  If Not CheckGrants(aObj) Then Exit Sub

  With UserQueriesViewEditorDialog
    .ShowEditTools = False
    .EditorText = aObj.Comment
    .txtEditor.Text = aObj.Comment
  
    Select Case TypeName(aObj)
    Case "UserQueriesViewQuery"
      .Caption = "Abfragebeschreibung"
    Case "UserQueriesViewParameter"
      .Caption = "Parameterbeschreibung"
    Case "UserQueriesViewQueryFolder"
      .Caption = "Ordnerbeschreibung"
    End Select
    .Show 1
  
    If .Cancel Then Exit Sub
  
    aObj.Comment = .EditorText
    Select Case TypeName(aObj)
    Case "UserQueriesViewQuery"
      aObj.SaveQuery
    Case "UserQueriesViewParameter"
      aObj.SaveParameter
    Case "UserQueriesViewQueryFolder"
      aObj.SaveQueryFolder
    End Select
    
    Comment = aObj.Comment
  End With
  Exit Sub
  
errLabel:
  ShowError "ShowEdit Comment"
  Exit Sub
End Sub

Public Sub ShowHideParameters(ByVal Panel As MSComctlLib.Panel)
Static strParameterWidth As String

On Error GoTo errLabel:

  With Panel
    Select Case .Key
    Case "QueryView"
      Select Case .Tag
      Case "Up"
        fraParameter.Visible = True
        .Picture = ilsBilder.ListImages("Down").Picture
        .Text = "Abfragenübersicht ausblenden"
        .Tag = "Down"
        fraSQL.Left = strParameterWidth
        fraParameter.Tag = strParameterWidth
        strParameterWidth = 0
        UserControl_Resize
      Case "Down"
        fraParameter.Visible = False
        .Picture = ilsBilder.ListImages("Up").Picture
        .Text = "Abfragenübersicht einblenden"
        .Tag = "Up"
        fraSQL.Left = 0
        strParameterWidth = fraParameter.Tag
        fraParameter.Tag = 0
        UserControl_Resize
      End Select
    Case "EditComment"
      
      Select Case TypeName(tvwBaum.SelectedItem.Tag)
      Case "UserQueriesViewQuery", "UserQueriesViewParameter", "UserQueriesViewQueryFolder"
        ShowEditComment tvwBaum.SelectedItem.Tag
      End Select
      
    Case "Comment"
      Select Case TypeName(tvwBaum.SelectedItem.Tag)
      Case "UserQueriesViewQuery", "UserQueriesViewParameter", "UserQueriesViewQueryFolder"
        ShowComment tvwBaum.SelectedItem.Tag
      End Select
    End Select
  End With
  Exit Sub

errLabel:
  ShowError "ShowHideParameters"
  Exit Sub
End Sub

Private Sub ChangeField _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aParameter As UserQueriesViewParameter _
, ByRef aQuery As UserQueriesViewQuery)

On Error GoTo errLabel:
      
  BaseToolKit.Dialog.SelectEntry.Reset
  BaseToolKit.Dialog.SelectEntry.SelectEntry aQuery.GetTablesInSQL _
  , "Bitte Tabelle wählen ...", False, True, False
  If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Sub
  Dim strTable As String: strTable = BaseToolKit.Dialog.SelectEntry.ValueEntry
  
  Dim strAlias As String: strAlias = aQuery.GetTableAlias(strTable)
  
  Dim col As Collection: Set col = New Collection
  
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected("SHOW COLUMNS FROM " & strTable)
  While Not rs.EOF
    col.Add CStr(rs.Fields(0).value)
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  
  BaseToolKit.Dialog.SelectEntry.Reset
  BaseToolKit.Dialog.SelectEntry.SelectEntry col _
  , "Bitte Feld wählen ...", False, True, False
  If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Sub
  Dim strField As String: strField = BaseToolKit.Dialog.SelectEntry.ValueEntry
  
  aParameter.Field = strAlias & "." & strField
  aParameter.SaveParameter

  SelectedNode.Text = "Bezugsfeld: " & aParameter.Field
  Exit Sub
  
errLabel:
  ShowError "ChangeField"
  Exit Sub
End Sub

Private Sub ChangeParameterType _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aParameter As UserQueriesViewParameter)

On Error GoTo errLabel:

  BaseToolKit.Dialog.SelectEntry.Reset
  BaseToolKit.Dialog.SelectEntry.SelectEntry aParameter.GetParameterTypeTextString _
  , "Bitte Parameterdatentyp wählen ...", True, True, False
  If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Sub
      
  aParameter.ParameterType = BaseToolKit.Dialog.SelectEntry.ValueID
  aParameter.SaveParameter
  
  SelectedNode.Text = "Parameterdatentyp: " & aParameter.ParameterTypeText
  tvwBaum.Nodes(SelectedNode.Parent.Key & "\Value").Text = "Parameterwert: " & aParameter.value
  Exit Sub
  
errLabel:
  ShowError "ChangeParameterType"
  Exit Sub
End Sub

Private Sub ChangeValue _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aParameter As UserQueriesViewParameter)

Dim strValue As String

On Error GoTo errLabel:

  strValue = InputBox("Bitte neuen Parameterwert eingeben:" & Chr(13) _
  & "(Mehrere Parameterwerte bitte durch Kommata trennen!)" _
  , aParameter.Name & " " & aParameter.OperatorText _
  & " ...", aParameter.value)

  
  If StrPtr(strValue) = 0 Then Exit Sub
        
  aParameter.value = strValue
  aParameter.SaveParameter
  
  SelectedNode.Text = "Parameterwert: " & aParameter.value
  Exit Sub
  
errLabel:
  ShowError "ChangeValue"
  Exit Sub
End Sub

Private Sub InsertParameterProperties _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aParameter As UserQueriesViewParameter)

Dim aNode As MSComctlLib.Node

On Error GoTo errLabel:

  If SelectedNode.Children > 0 Then Exit Sub

  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Key = SelectedNode.Key & "\Field"
  aNode.Text = "Bezugsfeld: " & aParameter.Field
  aNode.Tag = "Field"
  aNode.Image = "Field"
  
  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Key = SelectedNode.Key & "\Value"
  aNode.Text = "Parameterwert: " & aParameter.value
  aNode.Tag = "Value"
  aNode.Image = "Value"


  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Key = SelectedNode.Key & "\Operator"
  aNode.Text = "Operator: " & aParameter.OperatorText
  aNode.Tag = "Operator"
  aNode.Image = "Operator"

  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Key = SelectedNode.Key & "\Parametertype"
  aNode.Text = "Parameterdatentyp: " & aParameter.ParameterTypeText
  aNode.Tag = "Parametertyp"
  aNode.Image = "Parametertyp"
  Exit Sub
  
errLabel:
  ShowError "InsertParameterProperties"
  Exit Sub
End Sub
 
Private Sub ChangeOperator _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aParameter As UserQueriesViewParameter)


On Error GoTo errLabel:

  BaseToolKit.Dialog.SelectEntry.Reset
  BaseToolKit.Dialog.SelectEntry.SelectEntry aParameter.GetOperatorTextString _
  , "Bitte Vergleichsopertator wählen ...", True, True, False
  If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Sub
  
  aParameter.Operator = BaseToolKit.Dialog.SelectEntry.ValueID
  aParameter.SaveParameter
  
  If aParameter.Operator = pveNoParameter Then
    SelectedNode.Parent.Image = "Parameter"
  Else
    SelectedNode.Parent.Image = "ParameterInUse"
  End If
  
  SelectedNode.Text = "Operator: " & BaseToolKit.Dialog.SelectEntry.ValueEntry
  Exit Sub
  
errLabel:
  ShowError "ChangeOperator"
  Exit Sub
End Sub

Private Sub InsertQueriesFromCollection _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef colQueries As Collection)

Dim aQuery As UserQueriesViewQuery
Dim aNode As MSComctlLib.Node
Dim blnInsertQuery As Boolean
Dim astrName() As String
Dim strUser As String
  
On Error GoTo errLabel:
    
  For Each aQuery In colQueries
    blnInsertQuery = True
    If aQuery.QueryType = eqqtPresentQuery Then
      
      astrName = Split(aQuery.Name, "{#FROM#}")
      aQuery.Name = astrName(0)
      aQuery.SaveQuery
      
      If UBound(astrName) > 0 Then
        strUser = astrName(1)
      Else
        strUser = astrName(0)
      End If
      
      If MsgBox("Der Benutzer '" & strUser & "' hat Ihnen die Abfrage '" _
      & aQuery.Name & "' geschenkt." & Chr(13) _
      & "Wollen Sie das Geschenk annehmen?", 36 _
      , "Geschenkte Abfrage annehmen") = vbYes Then
        aQuery.QueryType = eqqtUserQuery
        aQuery.SaveQuery
      Else
        aQuery.DeleteQuery
        blnInsertQuery = False
      End If
    End If
      
    If blnInsertQuery Then
      Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
      aNode.Key = SelectedNode.Key & "\" & aQuery.Name & aQuery.QueryID
      aNode.Text = aQuery.Name
      aNode.Image = "Query"
      Set aNode.Tag = aQuery
      SelectedNode.Sorted = True
    End If
  Next aQuery
  Exit Sub
  
errLabel:
  ShowError "InsertQueriesFromCollection"
  Exit Sub
End Sub

Private Sub InsertQueryFolders _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aQueryFolder As UserQueriesViewQueryFolder)

Dim aNode As MSComctlLib.Node
Dim aSubQueryFolder As UserQueriesViewQueryFolder

On Error GoTo errLabel:

  If aQueryFolder.QueryType = eqqtUserQuery Then
    aQueryFolder.QueryFolders.GetItems aQueryFolder.FolderID, aQueryFolder.Prg, aQueryFolder.OwnerID
  Else
    aQueryFolder.QueryFolders.GetItems aQueryFolder.FolderID, aQueryFolder.Prg
  End If

  For Each aSubQueryFolder In aQueryFolder.QueryFolders.Items
    Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
    aNode.Text = " " & aSubQueryFolder.FolderName
    aNode.Key = SelectedNode.Key & "\" & aSubQueryFolder.FolderID & aSubQueryFolder.FolderName
    aNode.Image = aSubQueryFolder.Image
    Set aNode.Tag = aSubQueryFolder
  Next aSubQueryFolder
  Exit Sub
  
errLabel:
  ShowError "InsertQueryFolders"
  Exit Sub
End Sub

Private Sub InsertParameters _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aQuery As UserQueriesViewQuery)

Dim colParameters As Collection
Dim aParameter As UserQueriesViewParameter
Dim aNode As MSComctlLib.Node


On Error GoTo errLabel:

  
  If SelectedNode.Children > 0 Then Exit Sub
  
  aQuery.GetParametersDB
  Set colParameters = aQuery.GetParameters
  
  For Each aParameter In colParameters
    Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
    aNode.Key = SelectedNode.Key & "\" & aParameter.Name & aParameter.ParameterID
    aNode.Text = aParameter.Name
    
    If aParameter.Operator = pveNoParameter Then
      aNode.Image = "Parameter"
    Else
      aNode.Image = "ParameterInUse"
    End If
    
    Set aNode.Tag = aParameter
  Next aParameter
  SelectedNode.Sorted = True
  Exit Sub
  
errLabel:
  ShowError "InsertParamters"
  Exit Sub
End Sub

Private Sub DisableMenu()
Dim aButton As MSComctlLib.Button
Dim aButtonMenu As MSComctlLib.ButtonMenu

  Comment = ""
  staStatus.Panels("EditComment").Visible = False

  mnuQueries.Enabled = False
  mnuRenameFolder.Enabled = False
  mnuDeleteFolder.Enabled = False
  mnuEditFolderComment.Enabled = False
  mnuFolderComment.Enabled = False
  mnuNewFolder.Enabled = False
  mnuNewQuery.Enabled = False
  mnuNewQueryAssistent.Enabled = False
  mnuNewQuerySQLView.Enabled = False
  
  mnuQuery.Enabled = False
  mnuRenameQuery.Enabled = False
  mnuDeleteQuery.Enabled = False
  mnuQueryManagement.Enabled = False
  mnuCopyQuery.Enabled = False
  mnuMoveQuery.Enabled = False
  mnuGiveQuery.Enabled = False

  mnuEditQueryComment.Enabled = False
  mnuQueryComment.Enabled = False
  mnuShowOwner.Enabled = False
  mnuSQLView.Enabled = False
  mnuExecuteQuery.Enabled = False
  mnuNewParameter.Enabled = False
  
  mnuParameter.Enabled = False
  mnuRenameParameter.Enabled = False
  mnuDeleteParameter.Enabled = False
  mnuEditParameterComment.Enabled = False
  mnuParameterComment.Enabled = False
  
  mnuFlexGrid.Enabled = False
  mnuOpen.Enabled = False
  mnuSave.Enabled = False
  mnuPrint.Enabled = False
  mnuFind.Enabled = False
  mnuLinkControlFile.Enabled = False
  mnuWord.Enabled = False
  mnuExcel.Enabled = False
  mnuKopieren.Enabled = False
  mnuExecuteQueryFlexView.Enabled = False
  mnuSQLViewFlexGrid.Enabled = False
  
  For Each aButton In tblLeiste.Buttons
    If Mid(aButton.Key, 1, 8) <> "userdef_" Then
      For Each aButtonMenu In aButton.ButtonMenus
        aButtonMenu.Enabled = False
      Next aButtonMenu
      aButton.Enabled = False
    End If
  Next aButton

End Sub

Private Sub EnableMenu(ByRef SelectedNode As MSComctlLib.Node)
Dim aButton As MSComctlLib.Button
Dim aButtonMenu As MSComctlLib.ButtonMenu

  Select Case TypeName(SelectedNode.Tag)
  Case "UserQueriesViewQueryFolder"
    mnuQueries.Enabled = True
    mnuNewQuery.Enabled = True
    mnuFolderComment.Enabled = True
  
    If (IsAdmin Or IsOwner(SelectedNode.Tag)) Then
      mnuEditFolderComment.Enabled = True
      mnuNewFolder.Enabled = True
      
      mnuNewQuery.Enabled = True
      mnuNewQueryAssistent.Enabled = True
      mnuNewQuerySQLView.Enabled = True
      staStatus.Panels("EditComment").Visible = True
      
      If (CLng(SelectedNode.Tag.FolderID) > 3) Then
        mnuRenameFolder.Enabled = True
        mnuDeleteFolder.Enabled = True
      End If
    Else
      mnuNewFolder.Enabled = CheckPublicFolder(SelectedNode.Tag)
      If (CLng(SelectedNode.Tag.FolderID) > 3) _
      And CheckPublicFolder(SelectedNode.Tag) Then
        mnuRenameFolder.Enabled = True
      End If
    End If
      
  Case "UserQueriesViewQuery"
    mnuQuery.Enabled = True
    mnuQueryManagement.Enabled = True
    mnuQueryComment.Enabled = True
    mnuShowOwner.Enabled = True
    mnuExecuteQuery.Enabled = True
    mnuCopyQuery.Enabled = True
    
    If (IsAdmin Or IsOwner(SelectedNode.Tag)) Then
      mnuRenameQuery.Enabled = True
      mnuDeleteQuery.Enabled = True
      mnuMoveQuery.Enabled = True
      mnuGiveQuery.Enabled = True
      
      mnuEditQueryComment.Enabled = True
      mnuSQLView.Enabled = True
      mnuNewParameter.Enabled = True

      staStatus.Panels("EditComment").Visible = True
    End If
  Case "UserQueriesViewParameter"
    mnuParameter.Enabled = True
    mnuParameterComment.Enabled = True
  
    If (IsAdmin Or IsOwner(SelectedNode.Tag)) Then
      mnuRenameParameter.Enabled = True
      mnuDeleteParameter.Enabled = True
      mnuEditParameterComment.Enabled = True
      staStatus.Panels("EditComment").Visible = True
    End If
  End Select
  
  tblLeiste.Buttons("Open").Enabled = True
  tblLeiste.Buttons("Save").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("Print").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("Find").Enabled = flexAbfragen.Rows > 1
  
  tblLeiste.Buttons("NewQuery").Enabled = mnuNewQuery.Enabled
  
  tblLeiste.Buttons("DeleteQuery").Enabled = mnuDeleteQuery.Enabled
  tblLeiste.Buttons("Entwurfsansicht").Enabled = mnuSQLView.Enabled
  tblLeiste.Buttons("NewParameter").Enabled = mnuNewParameter.Enabled
  tblLeiste.Buttons("DeleteParameter").Enabled = mnuDeleteParameter.Enabled
  tblLeiste.Buttons("LinkControlFile").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("Word").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("Excel").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("eMail").Enabled = flexAbfragen.Rows > 1
  
  tblLeiste.Buttons("eMail").ButtonMenus("Mail").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("eMail").ButtonMenus("CallNumber").Enabled = flexAbfragen.Rows > 1
  Me.EnableSMS = (flexAbfragen.Rows > 1) And ((Me.EnableSMS) Or (BaseToolKit.WebService.Authentication.PersonId))

  tblLeiste.Buttons("Aufsteigend").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("Absteigend").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("Summe").Enabled = flexAbfragen.Rows > 1
  tblLeiste.Buttons("Abfrage").Enabled = mnuExecuteQuery.Enabled

  mnuFlexGrid.Enabled = True
  mnuOpen.Enabled = True
  mnuSave.Enabled = flexAbfragen.Rows > 1
  mnuPrint.Enabled = flexAbfragen.Rows > 1
  mnuFind.Enabled = flexAbfragen.Rows > 1
  mnuLinkControlFile.Enabled = flexAbfragen.Rows > 1
  mnuWord.Enabled = flexAbfragen.Rows > 1
  mnuExcel.Enabled = flexAbfragen.Rows > 1
  mnuKopieren.Enabled = flexAbfragen.Rows > 1
  
  mnuExecuteQueryFlexView.Enabled = mnuExecuteQuery.Enabled
  mnuSQLViewFlexGrid.Enabled = mnuSQLView.Enabled

End Sub

Private Sub ActionSelectionQueryTree _
(ByRef SelectedNode As MSComctlLib.Node)

On Error GoTo errLabel:

  DisableMenu
     
  Select Case TypeName(SelectedNode.Tag)
  Case "UserQueriesViewPrgs" '{Root-Node}
  Case "String"  '{Prg-Node}
  
    Select Case SelectedNode.Tag
    Case "prg"
    End Select
  
  Case "UserQueriesViewQueryFolder"
    QueryFolderNodeClick SelectedNode, SelectedNode.Tag
  Case "UserQueriesViewQueries"
  Case "UserQueriesViewQuery"
    QueryNodeClick SelectedNode
  Case "UserQueriesViewParameter"
    ParameterNodeClick SelectedNode
  End Select
  
  EnableMenu SelectedNode
  Exit Sub

errLabel:
  ShowError "ActionSelectionQueryTree"
  Exit Sub
End Sub


Private Sub QueryFolderNodeClick _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aQueryFolder As UserQueriesViewQueryFolder)

On Error GoTo errLabel:
  
  InsertQueries SelectedNode, aQueryFolder
  Comment = aQueryFolder.Comment
  EnableMenu SelectedNode
  Exit Sub
    
errLabel:
  ShowError "QueryFolderNodeClick"
  Exit Sub
End Sub

Private Sub ParameterNodeClick(ByRef SelectedNode As MSComctlLib.Node)
    
On Error GoTo errLabel:
            
  InsertParameterProperties SelectedNode, SelectedNode.Tag
  Comment = SelectedNode.Tag.Comment
  EnableMenu SelectedNode
  Exit Sub
    
errLabel:
  ShowError "ParameterNodeClick"
  Exit Sub
End Sub


Private Sub QueryNodeClick(ByRef SelectedNode As MSComctlLib.Node)
On Error GoTo errLabel:
  
  InsertParameters SelectedNode, SelectedNode.Tag
  Comment = SelectedNode.Tag.Comment
  EnableMenu SelectedNode
  Exit Sub
  
errLabel:
  ShowError "QueryNodeClick"
  Exit Sub
End Sub

Private Sub InsertQueries _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aQueryFolder As UserQueriesViewQueryFolder)

On Error GoTo errLabel:

  If SelectedNode.Children > 0 Then Exit Sub
      
  Select Case aQueryFolder.QueryType
  Case eqQueryType.eqqtUserQuery
    aQueryFolder.Queries.GetQueriesDB aQueryFolder.Prg _
    , BaseToolKit.WebService.Authentication.PersonId, aQueryFolder.QueryType, aQueryFolder.FolderID
    InsertQueriesFromCollection SelectedNode, aQueryFolder.Queries.GetQueries
    
    '{Geschenkte Abfragenn einfügen}
    aQueryFolder.Queries.GetQueriesDB aQueryFolder.Prg _
    , BaseToolKit.WebService.Authentication.PersonId, eqqtPresentQuery, aQueryFolder.FolderID
    InsertQueriesFromCollection SelectedNode, aQueryFolder.Queries.GetQueries
    '{Ende Geschenkte Abfragenn einfügen}
    
  Case eqQueryType.eqqtOpenQuery
    aQueryFolder.Queries.GetQueriesDB aQueryFolder.Prg, "-1", aQueryFolder.QueryType, aQueryFolder.FolderID
    InsertQueriesFromCollection SelectedNode, aQueryFolder.Queries.GetQueries
  Case eqQueryType.eqqtAdminQuery
    aQueryFolder.Queries.GetQueriesDB aQueryFolder.Prg, "-1", aQueryFolder.QueryType, aQueryFolder.FolderID
    InsertQueriesFromCollection SelectedNode, aQueryFolder.Queries.GetQueries
  Case eqQueryType.eqqtPresentQuery
    aQueryFolder.Queries.GetQueriesDB aQueryFolder.Prg _
    , BaseToolKit.WebService.Authentication.PersonId, aQueryFolder.QueryType, aQueryFolder.FolderID
    InsertQueriesFromCollection SelectedNode, aQueryFolder.Queries.GetQueries
  End Select
  
  
  InsertQueryFolders SelectedNode, aQueryFolder
  SelectedNode.Sorted = True
  Exit Sub
  
errLabel:
  ShowError "InsertQueries"
  Exit Sub
End Sub

Private Sub SaveControlFile _
(ByRef aFlexGrid As MSFlexGrid _
, Optional ByVal strInitialControlFileFolder As String = "")

Dim strDelimiter As String
Dim x As Long
Dim Y As Long

'{Variablen für Dateihandling}
Dim Temp As String
Dim strTempFileName As String
Dim intDummyRecords As Integer

'{Variablen für ControlFile-Ermittlung}
Dim strHomeDir As String
Dim strKey As String
Dim strControlFile As String

On Error GoTo errLabel:


  With aFlexGrid

    If .Rows = 0 Then Exit Sub

    '{Ordner 'Eigene Dateien' ermitteln}
    strKey = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

    strHomeDir = BaseToolKit.Win32API.Win32ApiProfessional.SysRegistry.GetRegistryValue _
    (enumHKEY_CURRENT_USER, strKey, "Personal")
    '{Ende Ordner 'Eigene Dateien' ermitteln}

    If strInitialControlFileFolder = "" Then
      strInitialControlFileFolder = strHomeDir
    End If

    '{Steuerdatei abfragen}
    strControlFile = BaseToolKit.Win32API.Win32ApiProfessional.SysDialogsEx.GetSaveFileNameEx( _
    strInitialControlFileFolder, "", "txt,csv,htm,html" _
    , "txt", "Bitte Steuerdatei wählen ...", UserControl.hwnd)

    If strControlFile = "" Then Exit Sub
    '{Ende Steuerdatei abfragen}

    strDelimiter = ";"

    Temp = InputBox _
    ("Wieviele Leerdatensätze sollen hinzugefügt werden?" & Chr(13) _
     & "(Hiermit können Sie steuern an welcher Position" _
     & " ein Etikettendruck beginnen soll)", "Leerdatensätze", 0)

    If Temp = "" Then Exit Sub

    If Not IsNumeric(Temp) Then
      MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
      Exit Sub
    End If
    intDummyRecords = Temp

    Screen.MousePointer = 11

    Dim colData As Collection: Set colData = New Collection
    Dim colColumns As Collection: Set colColumns = New Collection
    
    '{Header einfügen}
    For x = 0 To .Cols - 1
      colColumns.Add """" & .TextMatrix(0, x) & """"
    Next x
    colData.Add BaseToolKit.Convert.JoinCollection(colColumns, strDelimiter)

    Dim colRow As Collection

    '{Leersätze einfügen}
    For Y = 1 To intDummyRecords
      Set colRow = New Collection
      For x = 0 To .Cols - 1
        colRow.Add """"""
      Next x
      colData.Add BaseToolKit.Convert.JoinCollection(colRow, strDelimiter)
    Next Y
    '{Ende Leersätze einfügen}

    '{Alle Datensätze anhängen}
    For Y = 1 To .Rows - 1
      Set colRow = New Collection
      For x = 0 To .Cols - 1
        Dim columnValue As String
        columnValue = """" & Replace(.TextMatrix(Y, x), """", """""") & """"
        colRow.Add columnValue
      Next x
      colData.Add Replace(BaseToolKit.Convert.JoinCollection(colRow, strDelimiter), vbCrLf, vbLf)
    Next Y
  
    Dim content As String: content = BaseToolKit.Convert.JoinCollection(colData, vbCrLf)
    BaseToolKit.FileSystem.io.WriteAllText strControlFile, content, False
  End With

  '{html-Formatierung}
  Select Case LCase(BaseToolKit.FileSystem.io.GetExtensionName(strControlFile))
  Case "htm", "html"

    strTempFileName = Mid(strControlFile, 1, Len(strControlFile) _
    - Len(BaseToolKit.FileSystem.io.GetExtensionName(strControlFile))) & "$$$"

    ControlFileToHTML strTempFileName, strControlFile, ";"
    BaseToolKit.FileSystem.io.DeleteFile strControlFile, True
    BaseToolKit.FileSystem.io.MoveFile strTempFileName, strControlFile
  End Select

  Screen.MousePointer = 0
  Exit Sub

errLabel:
  Screen.MousePointer = 0

  If Err.Number <> 32755 Then '{Abbrechen}
    ShowError "CreateSteuerdatei"
  End If
  Exit Sub
End Sub

Private Sub LinkControlFile _
(Optional ByVal strInitialTemplateFolder As String = "" _
, Optional ByVal strInitialControlFileFolder As String = "")

Dim strHomeDir As String
Dim strKey As String
Dim strControlFile As String
Dim strWordFile As String

On Error GoTo errLabel:


  '{Ordner 'Eigene Dateien' ermitteln}
  strKey = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
  strHomeDir = BaseToolKit.Win32API.Win32ApiProfessional.SysRegistry.GetRegistryValue _
  (enumHKEY_CURRENT_USER, strKey, "Personal")
  '{Ende Ordner 'Eigene Dateien' ermitteln}

  If strInitialTemplateFolder = "" Then
    strInitialTemplateFolder = strHomeDir
  End If
  
  If strInitialControlFileFolder = "" Then
    strInitialControlFileFolder = strHomeDir
  End If

  '{Word-Datei abfragen}
  strWordFile = BaseToolKit.Win32API.Win32ApiProfessional.SysDialogsEx.GetOpenFileNameEx _
  (strInitialTemplateFolder, "", "docx,doc,dot" _
  , "Bitte Serienbrief wählen ...", UserControl.hwnd)
  
  If strWordFile = "" Then Exit Sub
  '{Ende Word-Datei abfragen}


  '{Steuerdatei abfragen}
  strControlFile = BaseToolKit.Win32API.Win32ApiProfessional.SysDialogsEx.GetOpenFileNameEx _
  (strInitialControlFileFolder, "", "txt,csv,*" _
  , "Bitte Steuerdatei wählen ...", UserControl.hwnd)
  
  If strControlFile = "" Then Exit Sub
  '{Ende Steuerdatei abfragen}
  
  OpenSeriendruck strWordFile, strControlFile
  Exit Sub
  
errLabel:
  ShowError "LinkControlFile"
  Exit Sub
End Sub

Private Sub OpenSeriendruck(strFileName, strControlFile)
Dim aWord As Object
  
On Error GoTo errLabel:

  '{Word öffnen}
  Set aWord = CreateObject("word.application")
  With aWord
    .Documents.Open strFileName
    .Visible = True
  End With

  On Error Resume Next
  With aWord.ActiveDocument.MailMerge
    '{Steuerdatei zuweisen}
    .OpenDataSource _
    Name:=strControlFile, _
    ConfirmConversions:=False, _
    ReadOnly:=False, _
    LinkToSource:=True, _
    AddToRecentFiles:=False, _
    PasswordDocument:="", _
    PasswordTemplate:="", _
    WritePasswordDocument:="", _
    WritePasswordTemplate:="", _
    Revert:=False, _
    Format:=0, _
    Connection:="", _
    SQLStatement:="", _
    SQLStatement1:=""
  End With
  Exit Sub
  
errLabel:
  ShowError "OpenSeriendruck"
  Exit Sub
End Sub

Private Function GetStartRow() As Long
  If flexAbfragen.Row <= flexAbfragen.RowSel Then
    GetStartRow = flexAbfragen.Row
  Else
    GetStartRow = flexAbfragen.RowSel
  End If
End Function
    
Private Function GetEndRow() As Long
  If flexAbfragen.Row <= flexAbfragen.RowSel Then
    GetEndRow = flexAbfragen.RowSel
  Else
    GetEndRow = flexAbfragen.Row
  End If
End Function
    
Private Function GetStartCol() As Long
  If flexAbfragen.col <= flexAbfragen.ColSel Then
    GetStartCol = flexAbfragen.col
  Else
    GetStartCol = flexAbfragen.ColSel
  End If
End Function

Private Function GetEndCol() As Long
  If flexAbfragen.col <= flexAbfragen.ColSel Then
    GetEndCol = flexAbfragen.ColSel
  Else
    GetEndCol = flexAbfragen.col
  End If
End Function

Private Sub SendEMail()

On Error GoTo errLabel:

  With flexAbfragen
    If .Rows = 0 Then Exit Sub
    Dim addresses As Collection: Set addresses = New Collection
    
    Dim x As Long
    For x = GetStartRow To GetEndRow
      Dim Y As Long
      For Y = GetStartCol To GetEndCol
        If BaseToolKit.Communication.mail.CheckMailAdress(.TextMatrix(x, Y)) Then
          addresses.Add .TextMatrix(x, Y)
        End If
      Next Y
    Next x
    
  End With

  If addresses.Count = 0 Then
    MsgBox "Keine eMail-Adressen gefunden!", 48, "Keine eMail-Adressen"
    Exit Sub
  End If

  BaseToolKit.Communication.mail.OpenOutlookClient _
  "", "", "", BaseToolKit.Convert.JoinCollection(addresses, ";"), "", "", ""
  Exit Sub

errLabel:
  ShowError "SendEMail"
  Exit Sub
End Sub

Private Sub CallNumber()

On Error GoTo errLabel:

  If Me.ClipDialPath = "" Then
    MsgBox "ClipDialPath nicht gesetzt!", 16, "Anrufen"
    Exit Sub
  End If

  With flexAbfragen
    If .Rows = 0 Then Exit Sub
    BaseToolKit.Communication.Phone.CallNumber .Text, Me.ClipDialPath
  End With
  Exit Sub

errLabel:
  ShowError "CallNumber"
  Exit Sub
End Sub

Private Sub SendSMS()

'Dim lngSum As Long
'Dim lngNotSent As Long

On Error GoTo errLabel:

  With flexAbfragen
    If .Rows = 0 Then Exit Sub

    Dim addresses As Collection: Set addresses = New Collection
    
    Dim x As Long
    For x = GetStartRow To GetEndRow
      Dim Y As Long
      For Y = GetStartCol To GetEndCol
        addresses.Add .TextMatrix(x, Y)
      Next Y
    Next x
    
    If addresses.Count = 0 Then Exit Sub
  End With

  With UserQueriesViewEditorDialog
    .Caption = "SMS"
    .EditorText = "<Bitte SMS-Nachricht eingeben (max. 160 Zeichen)>"
    .txtEditor.Text = "<Bitte SMS-Nachricht eingeben>"
    .txtEditor.MaxLength = 160
    .ShowEditTools = False
    .Show 1
  
    If .Cancel Then Exit Sub

    If Trim(.EditorText) = "" Then
      MsgBox "SMS-Nachrichtentext fehlt!", 16, "SMS senden"
      Exit Sub
    End If

    'lngSum = UBound(Dests) + 1
        
    Dim info As SmsInfoType: info = BaseToolKit.Communication.SMS.CreateSmsInfo
    info.Data = .EditorText

    .txtEditor.MaxLength = 0
    BaseToolKit.Communication.SMS.SendSmsByDestCollection addresses, info
    'lngNotSent = aSMS.MultiSendSMS(Split(strMailAddresses, ","))
  End With
  
  MsgBox "SMS-Nachricht wurde versandt!"
  
'  MsgBox "Von " & lngSum & " SMS-Nachrichten wurden " _
'  & lngSum - lngNotSent & " versandt.", 64, "SMS"
  Exit Sub

errLabel:
  ShowError "SendSMS"
  Exit Sub
End Sub

Private Sub GiveQuery(ByRef aQuery As UserQueriesViewQuery)

On Error GoTo errLabel:

  Dim strSQL As String
  strSQL = strSQL & "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "u.PersonenID as UserID," & vbCrLf & vbTab
  strSQL = strSQL & "CONCAT(u.Nachname,', ',u.Vorname) AS User" & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "datapool.t_personen u" & vbCrLf & vbTab
  
  strSQL = strSQL & "inner join datapool.t_personenstati s" & vbCrLf & vbTab
  strSQL = strSQL & "on u.PersonenID = s.PersonenFID" & vbCrLf & vbTab
  strSQL = strSQL & "AND s.Status = 'Mitarbeiter'" & vbCrLf & vbTab
  
  strSQL = strSQL & "ORDER BY" & vbCrLf & vbTab
  strSQL = strSQL & "u.Nachname, u.Vorname" & vbCrLf
  
  Dim strUser As String
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(strSQL)
  While Not rs.EOF
    strUser = strUser & "#" & rs!User & "#" & rs!UserID
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  
  BaseToolKit.Dialog.SelectEntry.Reset
  BaseToolKit.Dialog.SelectEntry.SelectEntry _
  strUser, "Bitte Benutzer wählen ...", True, True
  If BaseToolKit.Dialog.SelectEntry.ValueEntry = "" Then Exit Sub

  Dim aNewQuery As UserQueriesViewQuery: Set aNewQuery = New UserQueriesViewQuery
  aNewQuery.Comment = aQuery.Comment
  aNewQuery.Name = aQuery.Name & "{#FROM#}" & BaseToolKit.WebService.Authentication.FullName
  aNewQuery.Owner = aQuery.Owner
  aNewQuery.Prg = aQuery.Prg
  aNewQuery.Statement = aQuery.Statement
  aNewQuery.QueryType = eqqtPresentQuery
  aNewQuery.OwnerID = BaseToolKit.Dialog.SelectEntry.ValueID
  aNewQuery.QueryFolderFID = eqqtUserQuery
  aNewQuery.SaveQuery
  
  aQuery.GetParametersDB
  Dim colP As Collection: Set colP = aQuery.GetParameters
  
  Dim aParameter As UserQueriesViewParameter
  For Each aParameter In colP
    Dim aNewParameter As UserQueriesViewParameter: Set aNewParameter = New UserQueriesViewParameter
    With aNewParameter
      .Comment = aParameter.Comment
      .Field = aParameter.Field
      .Name = aParameter.Name
      .Operator = aParameter.Operator
      .OwnerID = BaseToolKit.Dialog.SelectEntry.ValueID
      .ParameterType = aParameter.ParameterType
      .QueryFID = aNewQuery.QueryID
      .value = aParameter.value
      .SaveParameter
    End With
  Next aParameter

  MsgBox "Abfrage '" & aQuery.Name & "' wurde '" _
  & BaseToolKit.Dialog.SelectEntry.ValueEntry & "' geschenkt!" _
  , 64, "Abfrage schenken"
  Exit Sub
  
errLabel:
  ShowError "GiveQuery"
  Exit Sub
End Sub

Private Function CreateSum() As Double

On Error GoTo errLabel:
  
  With flexAbfragen
    Dim i As Long
    For i = 1 To .Rows - 1
      CreateSum = CreateSum + .TextMatrix(i, .col)
    Next i
  End With
  Exit Function
  
errLabel:
  CreateSum = 0
  Exit Function
End Function

Private Sub NewQuery(ByRef SelectedNode As Node, ByRef aQueryFolder As UserQueriesViewQueryFolder)

On Error GoTo errLabel:

  If Not CheckGrants(aQueryFolder) Then Exit Sub

  Dim aQuery As UserQueriesViewQuery: Set aQuery = New UserQueriesViewQuery
  aQuery.Prg = aQueryFolder.Prg
  aQuery.OwnerID = BaseToolKit.WebService.Authentication.PersonId
  aQuery.Owner = BaseToolKit.WebService.Authentication.FullName
  aQuery.QueryFolderFID = aQueryFolder.FolderID
  aQuery.QueryType = SelectedNode.Tag.QueryType
  aQueryFolder.Queries.AddQuery aQuery
  
  '{Baumansicht aktualisieren}
  Dim aNode As Node: Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Key = SelectedNode.Key & "\" & aQuery.Name & aQuery.QueryID
  aNode.Text = aQuery.Name
  aNode.Image = "Query"
  Set aNode.Tag = aQuery

  Set tvwBaum.SelectedItem = aNode
  SelectedNode.Sorted = True
  aNode.EnsureVisible
  EnableMenu aNode
  
  ShowEntwurfsAnsicht aQuery
  Exit Sub
  
errLabel:
  ShowError "NewQuery"
  Exit Sub
End Sub

Private Sub DeleteQuery(ByRef SelectedNode As Node, ByRef aQuery As UserQueriesViewQuery)

On Error GoTo errLabel:

  If Not CheckGrants(aQuery) Then Exit Sub

  If MsgBox("Soll die Abfrgae '" & aQuery.Name _
  & "' wirklich gelöscht werden?", 36 _
  , "Abfrage löschen") <> vbYes Then Exit Sub
  
  aQuery.DeleteQuery
  Set tvwBaum.SelectedItem = SelectedNode.Parent
  EnableMenu tvwBaum.SelectedItem
  tvwBaum.Nodes.Remove SelectedNode.Index
  tvwBaum.SetFocus
  Exit Sub
  
errLabel:
  ShowError "DeleteQuery"
  Exit Sub
End Sub

Private Sub DeleteQueryFolder _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aQueryFolder As UserQueriesViewQueryFolder)

On Error GoTo errLabel:

  Select Case CLng(aQueryFolder.FolderID)
  Case 0 To 3
    Exit Sub
  End Select

  If Not CheckGrants(aQueryFolder) Then Exit Sub

  If MsgBox("Soll der Ordner '" & aQueryFolder.FolderName _
  & "' wirklich gelöscht werden?", 36 _
  , "Ordner löschen") <> vbYes Then Exit Sub
  
  aQueryFolder.DeleteQueryFolder
  Set tvwBaum.SelectedItem = SelectedNode.Parent
  EnableMenu tvwBaum.SelectedItem
  tvwBaum.Nodes.Remove SelectedNode.Index
  tvwBaum.SetFocus
  Exit Sub
  
errLabel:
  ShowError "DeleteQueryFolder"
  Exit Sub
End Sub

Private Sub NewParameter _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aQuery As UserQueriesViewQuery)

Dim aParameter As UserQueriesViewParameter
Dim aNode As MSComctlLib.Node

On Error GoTo errLabel:

  If Not CheckGrants(aQuery) Then Exit Sub

  Set aParameter = New UserQueriesViewParameter
  aParameter.QueryFID = aQuery.QueryID
  aParameter.OwnerID = BaseToolKit.WebService.Authentication.PersonId
  aParameter.SaveParameter
  
  '{Baumansicht aktualisieren}
  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Key = SelectedNode.Key & "\" & aParameter.Name & aParameter.ParameterID
  aNode.Text = aParameter.Name
  aNode.Image = "Parameter"
  Set aNode.Tag = aParameter

  Set tvwBaum.SelectedItem = aNode
  aNode.EnsureVisible
  EnableMenu aNode
  Exit Sub
  
errLabel:
  ShowError "NewParameter"
  Exit Sub
End Sub

Private Function CheckPublicFolder(ByRef aQueryFolder As UserQueriesViewQueryFolder) As Boolean
On Error GoTo errLabel:

  CheckPublicFolder = aQueryFolder.QueryType = eqqtOpenQuery
  Exit Function
  
errLabel:
  ShowError "CheckPublicFolder"
  Exit Function
End Function

Private Sub NewFolder _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aQueryFolder As UserQueriesViewQueryFolder)
  
Dim aNewQueryFolder As UserQueriesViewQueryFolder
Dim aNode As MSComctlLib.Node

On Error GoTo errLabel:

  If (Not CheckGrants(aQueryFolder, True)) _
  And (Not CheckPublicFolder(aQueryFolder)) Then Exit Sub

  Set aNewQueryFolder = New UserQueriesViewQueryFolder
  aNewQueryFolder.ParentFolderID = aQueryFolder.FolderID
  aNewQueryFolder.Prg = aQueryFolder.Prg
  aNewQueryFolder.QueryType = aQueryFolder.QueryType
  aNewQueryFolder.OwnerID = BaseToolKit.WebService.Authentication.PersonId
  aNewQueryFolder.Owner = BaseToolKit.WebService.Authentication.FullName
  aNewQueryFolder.SaveQueryFolder
  
  '{Baumansicht aktualisieren}
  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Key = SelectedNode.Key & "\" & aNewQueryFolder.FolderID & aNewQueryFolder.FolderName
  aNode.Text = " " & aNewQueryFolder.FolderName
  aNode.Image = aNewQueryFolder.Image
  Set aNode.Tag = aNewQueryFolder

  Set tvwBaum.SelectedItem = aNode
  aNode.EnsureVisible
  aNode.Parent.Sorted = True
  EnableMenu aNode
  Exit Sub
  
errLabel:
  ShowError "NewFolder"
  Exit Sub
End Sub

Private Sub DeleteParameter _
(ByRef SelectedNode As MSComctlLib.Node _
, ByRef aParameter As UserQueriesViewParameter)

On Error GoTo errLabel:

  If Not CheckGrants(aParameter) Then Exit Sub

  If MsgBox("Soll der Parameter '" & aParameter.Name _
  & "' wirklich gelöscht werden?", 36 _
  , "Parameter löschen") <> vbYes Then Exit Sub
  
  aParameter.DeleteParameter
  
  Set tvwBaum.SelectedItem = SelectedNode.Parent
  EnableMenu tvwBaum.SelectedItem
  tvwBaum.Nodes.Remove SelectedNode.Index
  tvwBaum.SetFocus
  Exit Sub
  
errLabel:
  ShowError "DeleteParameter"
  Exit Sub
End Sub

Private Sub ShowEntwurfsAnsicht(ByRef aQuery As UserQueriesViewQuery)

On Error GoTo errLabel:

  With UserQueriesViewEditorDialog
    .ShowEditTools = True
    .EditorText = aQuery.Statement
    .txtEditor.Text = aQuery.Statement
    .Show 1
    
    If Not .Cancel Then
      aQuery.Statement = .EditorText
      aQuery.SaveQuery
    End If
  End With
  Exit Sub
  
errLabel:
  ShowError "ShowEntwurfsAnsicht"
  Exit Sub
End Sub

Private Sub FillFlexGrid(ByRef rs As Object)
Dim strFlexItem As String
Dim i As Integer
Dim strFieldValue As String
Dim blnEOF As Boolean

On Error GoTo errLabel:

  With flexAbfragen

    .Rows = 0
    .Cols = rs.Fields.Count
    .ColAlignment(-1) = 1
    
    On Error Resume Next
    blnEOF = rs.EOF
    If Err.Number = 3704 Then Exit Sub
    On Error GoTo errLabel:
    
    Screen.MousePointer = 11
    
    For i = 0 To rs.Fields.Count - 1
      strFlexItem = strFlexItem & rs.Fields(i).Name & vbTab
    Next i
    .AddItem strFlexItem
            
    While Not rs.EOF
      strFlexItem = ""
      For i = 0 To rs.Fields.Count - 1
        strFieldValue = Replace(rs.Fields(i).value & "", vbTab, " ")
        strFlexItem = strFlexItem & strFieldValue & vbTab
      Next i
      .AddItem strFlexItem
      rs.MoveNext
    Wend
    BaseToolKit.Controls.flexGrid.ResizeColumns UserQueriesViewEditorDialog, flexAbfragen
        
    If rs.RecordCount > 0 Then rs.MoveFirst
        
    If .Rows = 1 Then
      .AddItem ""
      fraSQL.Caption = "Abfrageergebnis (Datensätze 0)"
    Else
      fraSQL.Caption = "Abfrageergebnis (Datensätze " & .Rows - 1 & ")"
    End If
    
    .FixedRows = 1
    Screen.MousePointer = 0
    
    Exit Sub
  End With
  
errLabel:
  ShowError "FillFlexGrid"
  Exit Sub
End Sub

Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub

Private Function DivideString _
(ByVal strFlowText As String _
, ByVal intMaxCols As Integer)

Dim strHelp As String
Dim intRest As Integer
Dim i As Integer

On Error GoTo errLabel:

  '{Alle Spalten von strFlowText durchlaufen}
  For i = 1 To Len(strFlowText)

    '{Wenn die Länge von strHelp kleiner als intMaxCols}
    If Len(strHelp) < intMaxCols Then

      '{Wenn strFlowText an der Stelle i kein Steuerzeichen enthält}
      If Mid(strFlowText, i, 1) > Chr(30) Then
        '{Erweiter Hilfsvariable um Zeichen von strFlowText Stelle i}
        strHelp = strHelp & Mid(strFlowText, i, 1)
      Else '{Enthält Steuerzeichen}
        '{Erweiter Hilfsvariable um ein Leerzeichen}
        strHelp = strHelp & " "
      End If
    Else '{Länge von strHelp ist größer gleich intMaxCols geworden}

      '{intRest auf Länge von strHelp setzen}
      intRest = Len(strHelp)

      '{Das letzte Wortende ermitteln}
      While Mid(strHelp, intRest, 1) <> " "
        intRest = intRest - 1
      Wend

      '{Splatenzähler eventuell zurücksetzen wenn nicht Wortende getroffen}
      i = i - (Len(strHelp) - (intRest - 1))
      DivideString = DivideString & Mid(strHelp, 1, intRest) & Chr(13) & Chr(10)
      strHelp = ""
    End If
  Next i
  '{Ende Alle Spalten von strFlowText durchlaufen}
  DivideString = DivideString & strHelp & Chr(13) & Chr(10)
  Exit Function

errLabel:
  DivideString = ""
  Exit Function
End Function

Private Sub PrepareModul()

On Error GoTo errLabel:
  
  With UserControl
      
    '{Gespeicherte Breite vom TreeView aus Registry lesen}
    .fraParameter.Tag = CDbl _
    (GetSetting(App.EXEName, .Name, "TreeWidth", "0,5"))

    '{Bebilderung tblLeiste}
    .tblLeiste.ImageList = .ilsBilder
    
    .tblLeiste.Buttons("Aufsteigend").Image = "Aufsteigend"
    .tblLeiste.Buttons("Absteigend").Image = "Absteigend"
    .tblLeiste.Buttons("Abfrage").Image = "Abfrage"
    .tblLeiste.Buttons("Find").Image = "Find"
        
    .tblLeiste.Buttons("Print").Image = "Print"
    .tblLeiste.Buttons("LinkControlFile").Image = "LinkControlFile"
    .tblLeiste.Buttons("Word").Image = "Word"
    .tblLeiste.Buttons("Excel").Image = "Excel"
    .tblLeiste.Buttons("eMail").Image = "eMail"
    
    .tblLeiste.Buttons("Summe").Image = "Summe"
    
    .tblLeiste.Buttons("NewQuery").Image = "NewQuery"
    .tblLeiste.Buttons("DeleteQuery").Image = "DeleteQuery"
    .tblLeiste.Buttons("Entwurfsansicht").Image = "Entwurfsansicht"
    .tblLeiste.Buttons("NewParameter").Image = "NewParameter"
    .tblLeiste.Buttons("DeleteParameter").Image = "DeleteParameter"
  
    .tblLeiste.Buttons("Open").Image = "Open"
    .tblLeiste.Buttons("Save").Image = "Save"
          
    .staStatus.Panels("QueryView").Picture = .ilsBilder.ListImages("Down").Picture
    
    .cboLimit.ListIndex = 0 '{Kein Limit setzen}
                                       
  End With
  Exit Sub

errLabel:
  ShowError "PrepareModul"
  Exit Sub
End Sub



'---------------------- Öffentliche Methoden der Klasse -------------------------
Public Sub OpenQueryView()

On Error GoTo errLabel:
  
  InsertRootNode
  Me.EnableSMS = (BaseToolKit.WebService.Authentication.PersonId = 1)
  Exit Sub
  
errLabel:
  ShowError "OpenQueryView"
  Exit Sub
End Sub

Public Sub CloseQueryView()
Dim aForm As Form

On Error Resume Next
 
  SaveSetting App.EXEName, UserControl.Name, "TreeWidth" _
  , Replace(UserControl.fraParameter.Tag, ".", ",")
  
  For Each aForm In Forms
    Unload aForm
  Next aForm
End Sub

Public Function GetQueryData _
(Optional ByVal NumRows As Long = -1 _
, Optional ByVal ColumnDelimeter As String = vbTab _
, Optional ByVal RowDelimeter As String = vbCrLf _
, Optional ByVal NullExpr As String = "")

On Error GoTo errLabel:

  Const adClipString = 2

  If mrsQueryData Is Nothing Then
    GetQueryData = ""
  Else
    mrsQueryData.MoveFirst
    GetQueryData = mrsQueryData.GetString _
    (adClipString, NumRows, ColumnDelimeter _
    , RowDelimeter, NullExpr)
  End If
  Exit Function
  
errLabel:
  ShowError "GetQueryData"
  Exit Function
End Function

Public Function GetQueryColData _
(ByVal col As Variant _
, Optional ByVal Delimeter As String = "," _
, Optional ByVal NullExpr As String = "")


  On Error Resume Next
  GetQueryColData = mrsQueryData.Fields(col).Name
  If Err.Number <> 0 Then
    GetQueryColData = ""
    Exit Function
  End If
  On Error GoTo errLabel:
  
  
  Select Case True
  Case mrsQueryData Is Nothing
    GetQueryData = ""
  Case Else
    mrsQueryData.MoveFirst
    
    While Not mrsQueryData.EOF
      GetQueryColData = GetQueryColData _
      & mrsQueryData.Fields(col).value & Delimeter
      
      mrsQueryData.MoveNext
    Wend
    
    If Len(GetQueryColData) > 0 Then
      GetQueryColData = Mid(GetQueryColData, 1, Len(GetQueryColData) - 1)
    End If
  End Select
  Exit Function
  
errLabel:
  ShowError "GetQueryColData"
  Exit Function
End Function

Public Sub AddButton _
(ByVal Key As String _
, Optional ByVal Caption As String = "" _
, Optional ByVal Index As Long = -1 _
, Optional ByVal Style As ButtonStyleConstants = tbrDefault _
, Optional ByVal Image _
, Optional ByVal toolTipText As String = "")
  
'{tbrDefault 0}
'{tbrCheck 1}
'{tbrButtonGroup 2}
'{tbrSeparator 3}
'{tbrPlaceholder 4}
'{tbrDropdown 5}
  
Dim aButton As MSComctlLib.Button

On Error GoTo errLabel:

  If Index = -1 Then Index = tblLeiste.Buttons.Count + 1

  If Not IsMissing(Image) Then ilsBilder.ListImages.Add , "userdef_" & Key, Image

  Set aButton = tblLeiste.Buttons.Add(Index)
  
  aButton.Caption = Caption
  aButton.Key = "userdef_" & Key
  aButton.Style = Style
  If Not IsMissing(Image) Then aButton.Image = "userdef_" & Key
  aButton.toolTipText = toolTipText
  Exit Sub
  
errLabel:
  ShowError "AddButton"
  Exit Sub
End Sub

Public Sub RemoveButton(ByVal Key As String)
    
Dim aButton As MSComctlLib.Button

On Error Resume Next

  Set aButton = tblLeiste.Buttons(Key)
  tblLeiste.Buttons.Remove aButton.Index
  Exit Sub
  
End Sub

Public Sub QueryCopy _
(ByRef DestinationNode As MSComctlLib.Node _
, ByRef SourceQuery As UserQueriesViewQuery)

Dim aQuery As UserQueriesViewQuery
Dim colParameter As Collection
Dim aParameter As UserQueriesViewParameter
Dim SourceParameter As UserQueriesViewParameter
Dim aNode As MSComctlLib.Node

On Error GoTo errLabel:

  Screen.MousePointer = 11

  InsertQueries DestinationNode, DestinationNode.Tag

  Set aQuery = New UserQueriesViewQuery
  
  aQuery.Name = SourceQuery.Name
  aQuery.Prg = SourceQuery.Prg
  aQuery.Statement = SourceQuery.Statement
  aQuery.QueryFolderFID = SourceQuery.QueryFolderFID
  aQuery.Comment = SourceQuery.Comment
  aQuery.QueryType = eqqtUserQuery
  aQuery.OwnerID = BaseToolKit.WebService.Authentication.PersonId
  aQuery.Owner = BaseToolKit.WebService.Authentication.FullName
  
  aQuery.SaveQuery
  aQuery.GetQuery aQuery.QueryID
  
  SourceQuery.GetParametersDB
  Set colParameter = SourceQuery.GetParameters
  
  For Each SourceParameter In colParameter
    Set aParameter = New UserQueriesViewParameter
  
    aParameter.Name = SourceParameter.Name
    aParameter.Field = SourceParameter.Field
    aParameter.value = SourceParameter.value
    aParameter.Operator = SourceParameter.Operator
    aParameter.Comment = SourceParameter.Comment
    aParameter.ParameterType = SourceParameter.ParameterType
    aParameter.QueryFID = aQuery.QueryID
    
    aParameter.SaveParameter
    
  Next SourceParameter
    
  Set aNode = tvwBaum.Nodes.Add(DestinationNode.Key, tvwChild)
  aNode.Key = DestinationNode.Key & "\" & aQuery.Name & aQuery.QueryID
  aNode.Text = aQuery.Name
  aNode.Image = "Query"
  Set aNode.Tag = aQuery
  
  aNode.EnsureVisible
  Screen.MousePointer = 0
  Exit Sub
  
errLabel:
  ShowError "QueryCopy"
  Exit Sub
End Sub

Public Function GetValue(ByVal Row As Long, ByVal col As Long) As String
On Error GoTo errLabel:

  GetValue = flexAbfragen.TextMatrix(Row, col)
  Exit Function
  
errLabel:
  ShowError "GetValue"
  Exit Function
End Function

Public Sub SetValue(ByVal Row As Long, ByVal col As Long, ByVal strValue As String)
On Error GoTo errLabel:

  flexAbfragen.TextMatrix(Row, col) = strValue
  Exit Sub
  
errLabel:
  ShowError "SetValue"
  Exit Sub
End Sub

Public Sub AddItem(ByVal strValue As String)
  flexAbfragen.AddItem strValue
End Sub

Public Sub RemoveItem(ByVal Index As Long)
  flexAbfragen.RemoveItem Index
End Sub

Public Function ItemInColNames(ByVal strItem As String) As Long
Dim c As Long

On Error GoTo errLabel:

  ItemInColNames = -1
  strItem = LCase(strItem)
  For c = 0 To flexAbfragen.Cols - 1
    If LCase(flexAbfragen.TextMatrix(0, c)) = strItem Then
      ItemInColNames = c
      Exit For
    End If
  Next c
  Exit Function
  
errLabel:
  ShowError "ItemInColNames"
  Exit Function
End Function

