VERSION 5.00
Begin VB.Form PropertiesViewMemoDialog 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Memo"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtMemo 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "PropertiesViewMemoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMemo As String

Private Sub Form_Resize()
  txtMemo.Height = Me.ScaleHeight
  txtMemo.Width = Me.ScaleWidth
End Sub

Public Property Get Memo() As String
  Memo = mMemo
End Property

Public Property Let Memo(ByVal value As String)
  mMemo = value
  txtMemo.Text = mMemo
End Property

Private Sub txtMemo_Change()
  mMemo = txtMemo.Text
End Sub
