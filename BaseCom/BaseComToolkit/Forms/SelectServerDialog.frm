VERSION 5.00
Begin VB.Form SelectServerDialog 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Connection..."
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3030
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer formShownTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   0
   End
   Begin VB.Frame FraServerauswahl 
      Caption         =   "Serverauswahl:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdSQLUnitTest 
         Caption         =   "Testdaten (SQL-UnitTest)"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Klicken Sie hier, um eine Verbindung zum Echtdatenbestand zu öffnen. Testversionen sollten IMMER hierrüber gestartet werden!"
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdSQL01 
         Caption         =   "Echtdaten (SQL01)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   1
         ToolTipText     =   "Klicken Sie hier, um eine Verbindung zum Echtdatenbestand zu öffnen. ACHTUNG: Bei Testversion kann Datenverlust entstehen!"
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2640
         Picture         =   "SelectServerDialog.frx":0000
         Top             =   270
         Width           =   240
      End
   End
End
Attribute VB_Name = "SelectServerDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mServerDataType As ServerDataTypes
Private mhInstance As Long
Private mAppPath As String

Public Property Get ServerDataType() As ServerDataTypes
  ServerDataType = mServerDataType
End Property

Public Property Get hInstance() As Long
  hInstance = mhInstance
End Property

Public Property Let hInstance(ByVal Value As Long)
  mhInstance = Value
End Property

Public Property Get appPath() As String
  appPath = mAppPath
End Property

Public Property Let appPath(ByVal Value As String)
  mAppPath = Value
End Property

Private Sub cmdSQL01_Click()
    
  If MsgBox("Sind Sie sicher, dass Sie die Version auf dem Echtdatenbestand ausführen möchten? Dies kann zu DATENVERLUST führen!", vbYesNo, "ACHTUNG!") = vbYes Then
    mServerDataType = RealData
  Else
    mServerDataType = TestData
  End If
  Unload Me
End Sub

Private Sub cmdSQLUnitTest_Click()
  
  mServerDataType = TestData
  Unload Me
End Sub

Private Sub Form_Load()
  formShownTimer.enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode = 0 Then
    mServerDataType = ServerDataTypes.Cancel
  End If
End Sub

Private Sub formShownTimer_Timer()
  formShownTimer.enabled = False
  
  Dim etc As etc: Set etc = New etc
  
  If Not etc.IsTestVersion(Me.appPath) Then
    If Not etc.IsDevelopTime(Me.hInstance) Then
      mServerDataType = RealData
      Unload Me
    End If
  End If
End Sub
