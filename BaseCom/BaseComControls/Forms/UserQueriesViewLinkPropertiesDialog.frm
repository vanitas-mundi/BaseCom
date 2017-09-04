VERSION 5.00
Begin VB.Form UserQueriesViewLinkPropertiesDialog 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Verknüpfungseigenschaften"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtRightTable 
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtLeftTable 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   2175
   End
   Begin VB.OptionButton optRightJoin 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4455
   End
   Begin VB.OptionButton optLeftJoin 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
   End
   Begin VB.OptionButton optInnerJoin 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rechter Tabellenname"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Linker Tabellenname"
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
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "UserQueriesViewLinkPropertiesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrLinkProperty As String
Private mstrLeftTable As String
Private mstrRightTable As String

Public Property Get LinkProperty() As String
  LinkProperty = mstrLinkProperty
End Property

Public Property Let LinkProperty(ByVal strLinkProperty As String)
  mstrLinkProperty = strLinkProperty
End Property

Public Property Get LeftTable() As String
  LeftTable = mstrLeftTable
End Property

Public Property Let LeftTable(ByVal strLeftTable As String)
  mstrLeftTable = strLeftTable
End Property

Public Property Get RightTable() As String
  RightTable = mstrRightTable
End Property

Public Property Let RightTable(ByVal strRightTable As String)
  mstrRightTable = strRightTable
End Property


Private Sub PrepareModul()
Dim Temp As String
On Error GoTo errLabel:

  txtLeftTable.Text = Me.LeftTable
  txtRightTable.Text = Me.RightTable

  Temp = "1: Beinhaltet nur die Datensätze, bei denen die Inhalte der "
  Temp = Temp & "verknüpften Felder beider Tabellen gleich sind."
  optInnerJoin.Caption = Temp
  
  Temp = "2: Beinhaltet ALLE Datensätze aus '" & Me.LeftTable
  Temp = Temp & "' und nur die Datensätze aus '" & Me.RightTable
  Temp = Temp & "', bei denen die Inhalte der verknüpften Felder "
  Temp = Temp & "', beider Tabellen gleich sind."
  optLeftJoin.Caption = Temp
  
  Temp = "3: Beinhaltet ALLE Datensätze aus '" & Me.RightTable
  Temp = Temp & "' und nur die Datensätze aus '" & Me.LeftTable
  Temp = Temp & "', bei denen die Inhalte der verknüpften Felder "
  Temp = Temp & "', beider Tabellen gleich sind."
  optRightJoin.Caption = Temp
  
  optInnerJoin.value = True
  Exit Sub
  
errLabel:
  MsgBox Err.Description, 16, Err.Number
  Exit Sub
End Sub

Private Sub cmdCancel_Click()
  Me.LinkProperty = ""
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  PrepareModul
End Sub

Private Sub optInnerJoin_Click()
  Me.LinkProperty = "="
End Sub

Private Sub optLeftJoin_Click()
  Me.LinkProperty = "<="
End Sub

Private Sub optRightJoin_Click()
  Me.LinkProperty = "=>"
End Sub
