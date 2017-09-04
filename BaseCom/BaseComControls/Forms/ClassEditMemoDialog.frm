VERSION 5.00
Begin VB.Form ClassEditMemoDialog 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   285
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "ClassEditMemoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrText As String
Private mstrTemp As String
Private mblnCancel As Boolean

Public Property Get Cancel() As String
  Cancel = mblnCancel
End Property

Public Property Get Text() As String
  Text = mstrText
End Property

Private Sub cmdCancel_Click()
  mblnCancel = True
  mstrText = mstrTemp
  Unload Me
End Sub

Private Sub cmdOK_Click()
  mstrText = txtMemo.Text
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Height = GetSetting(App.EXEName, "Memo", "Height", Me.Height)
  Me.Width = GetSetting(App.EXEName, "Memo", "Width", Me.Width)
  Me.Top = GetSetting(App.EXEName, "Memo", "Top", Me.Top)
  Me.Left = GetSetting(App.EXEName, "Memo", "Left", Me.Left)
  mstrText = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  
  If Me.Width < 1500 Then Me.Width = 1500
  If Me.Height < 1500 Then Me.Height = 1500
  
  txtMemo.Height = Me.ScaleHeight - (cmdOK.Height + 40)
  txtMemo.Width = Me.ScaleWidth
  
  cmdOK.Top = Me.ScaleHeight - cmdOK.Height
  cmdOK.Left = Me.ScaleWidth - cmdOK.Width
  
  cmdCancel.Top = cmdOK.Top
  cmdCancel.Left = cmdOK.Left - (cmdCancel.Width + 40)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.EXEName, "Memo", "Height", Me.Height
  SaveSetting App.EXEName, "Memo", "Width", Me.Width
  SaveSetting App.EXEName, "Memo", "Top", Me.Top
  SaveSetting App.EXEName, "Memo", "Left", Me.Left
End Sub

Private Sub txtMemo_GotFocus()
  txtMemo.SelStart = 0
  txtMemo.SelLength = Len(txtMemo.Text)
End Sub

Public Sub ShowMemo _
(ByVal strHeader As String _
, ByVal blnLocked As Boolean _
, Optional strText As String = "")

  Me.Caption = strHeader
  txtMemo.Locked = blnLocked
  txtMemo.Text = strText
  mstrTemp = strText
  Me.Show 1
End Sub
