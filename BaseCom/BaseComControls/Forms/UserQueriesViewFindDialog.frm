VERSION 5.00
Begin VB.Form UserQueriesViewFindDialog 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Suchen"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkNotCompareBinary 
      Caption         =   "Groﬂ-/Kleinschreibung beachten"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CheckBox chkGanzesWort 
      Caption         =   "Nur ganzes Wort suchen"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSuchen 
      Caption         =   "&Weitersuchen"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSuchString 
      Height          =   285
      Left            =   1140
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Suchen nach:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "UserQueriesViewFindDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancel As Boolean

Public Property Get Cancel() As Boolean
  Cancel = mblnCancel
End Property

Public Property Let Cancel(ByVal blnCancel As Boolean)
  mblnCancel = blnCancel
End Property

Public Property Get WholeWord() As Boolean
  WholeWord = chkGanzesWort.value = 1
End Property

Public Property Let WholeWord(ByVal blnWholeWord As Boolean)
  If blnWholeWord Then
    chkGanzesWort.value = 1
  Else
    chkGanzesWort.value = 0
  End If
End Property

Public Property Get CaseSensible() As Boolean
  CaseSensible = chkNotCompareBinary.value = 1
End Property

Public Property Let CaseSensible(ByVal blnCaseSensible As Boolean)
  If CaseSensible Then
    chkNotCompareBinary.value = 1
  Else
    chkNotCompareBinary.value = 0
  End If
End Property

Public Property Get SerachString() As String
  SerachString = txtSuchString.Text
End Property

Public Property Let SerachString(ByVal strSearchString As String)
  txtSuchString.Text = strSearchString
End Property

Private Sub cmdCancel_Click()
  Me.Cancel = True
  Me.Hide
End Sub

Private Sub cmdSuchen_Click()
  Me.Cancel = False
  Me.Hide
End Sub

Private Sub Form_Activate()
  txtSuchString.SetFocus
End Sub

Private Sub Form_Load()
  Me.Cancel = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Me.Hide
  Cancel = True
End Sub

Private Sub txtSuchString_GotFocus()
  txtSuchString.SelStart = 0
  txtSuchString.SelLength = Len(txtSuchString.Text)

End Sub

