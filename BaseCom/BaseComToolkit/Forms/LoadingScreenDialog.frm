VERSION 5.00
Begin VB.Form LoadingScreenDialog 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
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
   ScaleHeight     =   1575
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label lblBitteWarten 
      AutoSize        =   -1  'True
      Caption         =   "Bitte warten..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   3165
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      Caption         =   "Programm lädt!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3525
   End
End
Attribute VB_Name = "LoadingScreenDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : LoadingScreenDialog
'    Project    : ToolKits
'
'    Description: Formular zum Anzeigen eines Ladebildschirms
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------
Private Sub Form_Load()
  SetFormInForground Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetFormInForground Me, False
End Sub

'---------------------- Private Methoden der Klasse -----------------------------
Private Sub SetFormInForground(ByRef f As Object, ByVal blnForeGround As Boolean)

  If blnForeGround Then
    '{Fenster immer im Vordegrund}
    Call SetWindowPos(f.hwnd, -1, 0, 0, 0, 0, 3)
  Else
    '{Fenster im Normalzustand}
    Call SetWindowPos(f.hwnd, -2, 0, 0, 0, 0, 3)
  End If

End Sub
'---------------------- Öffentliche Methoden der Klasse -------------------------




