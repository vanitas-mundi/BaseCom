VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSelectBericht10 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Bericht wählen ..."
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmSelectBericht10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4710
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdDrucken 
      Height          =   375
      Left            =   4200
      Style           =   1  'Grafisch
      TabIndex        =   5
      ToolTipText     =   "Bericht drucken"
      Top             =   2550
      Width           =   495
   End
   Begin VB.CommandButton cmdAbbrechen 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2550
      Width           =   1095
   End
   Begin VB.Frame Frame 
      Caption         =   "Kopien"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   975
      Begin VB.TextBox txtKopien 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   255
         Left            =   630
         TabIndex        =   3
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Ansicht"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
      Begin VB.OptionButton optDruckvorschau 
         Caption         =   "Druckvorschau"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   200
         Value           =   -1  'True
         Width           =   1650
      End
      Begin VB.OptionButton optDrucker 
         Caption         =   "Drucker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Verfügbare Berichte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.FileListBox filBerichte 
         Height          =   2040
         Left            =   120
         Pattern         =   "*.rpt"
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmSelectBericht10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrReportPath As String
Private mstrSQLPath As String
Private mstrConnectionString As String

Private Sub StartReport()
  If optDruckvorschau.Value Then
    ShowReport
  Else
    PrintReport
  End If
End Sub

Private Sub ShowReport()
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Dim strReportName As String
Dim strSQL As String

On Error GoTo Fehler

  With filBerichte
    If .ListIndex = -1 Then Exit Sub
    
    strReportName = Mid(.List(.ListIndex) _
    , 1, Len(.List(.ListIndex)) - 4)
    
    strSQL = ODBC.GetParameter(ODBC.LoadConst(strReportName))
    
    If strSQL = "" Then Exit Sub
      
    Set con = New ADODB.Connection
    con.ConnectionString = Me.ConnectionString
    con.Open
    Set rs = con.Execute(strSQL)
    
    frmBerichte10.ReportFileName = _
    mstrReportPath & .List(.ListIndex)
    
    frmBerichte10.SetDataSource rs
    
    frmBerichte10.Output = prScreen
    frmBerichte10.Show 1
    con.Close
  End With
  Exit Sub
  
Fehler:
  ShowError "ShowReport"
  Exit Sub
End Sub

Private Sub cmdAbbrechen_Click()
  Unload Me
End Sub

Private Sub cmdDrucken_Click()
  StartReport
End Sub

Private Sub PrintReport()
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Dim strReportName As String
Dim strSQL As String

On Error GoTo Fehler

  With filBerichte
    If .ListIndex = -1 Then Exit Sub
    
    strReportName = Mid(.List(.ListIndex) _
    , 1, Len(.List(.ListIndex)) - 4)
    
    strSQL = ODBC.GetParameter(ODBC.LoadConst(strReportName))
    If strSQL = "" Then Exit Sub

    Set con = New ADODB.Connection
    con.ConnectionString = Me.ConnectionString
    con.Open
    Set rs = con.Execute(strSQL)
    
    frmBerichte10.ReportFileName = _
    mstrReportPath & .List(.ListIndex)
    
    frmBerichte10.SetDataSource rs
    
    frmBerichte10.Copies = txtKopien.Text
    frmBerichte10.Output = prPrinter
    frmBerichte10.Show 1
    con.Close
    
    Unload Me
  End With
  Exit Sub
  
Fehler:
  ShowError "PrintReport"
  Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case 116 'F5
    Aktualisieren
  End Select
End Sub

Public Property Let ConnectionString(ByVal strConnectionString As String)
  mstrConnectionString = strConnectionString
End Property

Public Property Get ConnectionString() As String
  ConnectionString = mstrConnectionString
End Property

Private Sub Form_Load()
  If Right(mstrReportPath, 1) <> "\" Then
    mstrReportPath = mstrReportPath & "\"
  End If
  filBerichte.Path = mstrReportPath
  
  Set ODBC = New ODBCHandling.clsODBCHandling
  ODBC.FillSQLConst Me.SQLPath
End Sub

Private Sub txtKopien_GotFocus()
  txtKopien.SelStart = 1
  txtKopien.SelLength = Len(txtKopien.Text)
  txtKopien.Tag = txtKopien.Text
End Sub

Private Sub txtKopien_LostFocus()
  If Not IsNumeric(txtKopien.Text) Then
    MsgBox "Ungültige Eingabe", 16, "Fornmatfehler"
    txtKopien.Text = txtKopien.Tag
  End If
End Sub

Private Sub UpDown_DownClick()
  If txtKopien.Text > 1 Then
    txtKopien.Text = txtKopien.Text - 1
    SaveSetting "MAV", "Rechnungen", "AnzahlKopien", txtKopien.Text
  Else
    MsgBox "Wert zu niedrig!", 16, "Wertefehler"
  End If
End Sub

Public Property Get ReportPath() As String
  ReportPath = mstrReportPath
End Property

Public Property Let ReportPath(ByVal strReportPath As String)
  mstrReportPath = strReportPath
  If Right(mstrReportPath, 1) <> "\" Then
    mstrReportPath = mstrReportPath & "\"
  End If
End Property

Public Property Get SQLPath() As String
  SQLPath = mstrSQLPath
End Property

Public Property Let SQLPath(ByVal strSQLPath As String)
  mstrSQLPath = strSQLPath
  If Right(mstrSQLPath, 1) <> "\" Then
    mstrSQLPath = mstrSQLPath & "\"
  End If
End Property

Public Sub Aktualisieren()
  filBerichte.Refresh
End Sub

