VERSION 5.00
Begin VB.UserControl WaitBarControl 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   ScaleHeight     =   1110
   ScaleWidth      =   3765
   ToolboxBitmap   =   "WaitBarControl.ctx":0000
   Begin VB.Timer tmrWaitBarControl 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   0
      Top             =   360
   End
   Begin VB.Label lblWaitBarPointer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   165
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label lblWaitbar 
      AutoSize        =   -1  'True
      Caption         =   "gggggggggg"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   165
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1650
   End
End
Attribute VB_Name = "WaitBarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'{---------------------Private Eigenschaften der Klasse------------------------}
Private mPos As Long
Private mblnMoveUp As Boolean
Private mObjectToStatusBar As WaitbarObjToStatusBar



'{---------------------Konstruktor und Destruktor der Klasse------------------------}

Private Sub UserControl_Initialize()
  Set mObjectToStatusBar = New WaitbarObjToStatusBar
End Sub

Private Sub UserControl_Terminate()
  StopWaitBar
  Set mObjectToStatusBar = Nothing
End Sub



'{---------------------Zugriffsmethoden der Klasseneigenschaften------------------------}
Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Public Property Get WaitPoints() As Long
  WaitPoints = Len(lblWaitbar.Caption)
End Property

Public Property Let WaitPoints(ByVal lngWaitPoints As Long)
  lblWaitbar.Caption = String(lngWaitPoints, "g")
  ResizeForm
End Property

Public Property Get WaitPointHeight() As Long
  WaitPointHeight = lblWaitbar.Font.Size
End Property

Public Property Let WaitPointHeight(ByVal lngWaitPointHeight As Long)
  lblWaitbar.Font.Size = lngWaitPointHeight
  lblWaitBarPointer.Font.Size = lngWaitPointHeight
  ResizeForm
End Property

Public Property Get Interval() As Long
  Interval = tmrWaitBarControl.Interval
End Property

Public Property Let Interval(ByVal lngInterval As Long)
  tmrWaitBarControl.Interval = lngInterval
End Property

Public Property Get WaitPointsColor() As Long
  WaitPointsColor = lblWaitbar.ForeColor
End Property

Public Property Let WaitPointsColor(ByVal lngColor As Long)
  lblWaitbar.ForeColor = lngColor
End Property

Public Property Get WaitPointIndexColor() As Long
  WaitPointIndexColor = lblWaitBarPointer.ForeColor
End Property

Public Property Let WaitPointIndexColor(ByVal lngColor As Long)
  lblWaitBarPointer.ForeColor = lngColor
End Property

Public Property Get BackColor() As Long
  BackColor = lblWaitbar.BackColor
End Property

Public Property Let BackColor(ByVal lngColor As Long)
  lblWaitbar.BackColor = lngColor
  UserControl.BackColor = lngColor
End Property

Public Property Get BorderStyle() As Long
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal lngBorderStyle As Long)
  UserControl.BorderStyle = lngBorderStyle
  ResizeForm
  SetWaitPointerPos
End Property



'{---------------------Private Klassenmethoden------------------------}

Private Sub MoveWaitBar()
  Select Case mblnMoveUp
  Case True
    mPos = mPos + 1
    If mPos >= Len(lblWaitbar.Caption) Then mblnMoveUp = False
  Case False
    mPos = mPos - 1
    If mPos <= 1 Then mblnMoveUp = True
  End Select
  
  lblWaitBarPointer.Left = ((mPos - 1) * lblWaitBarPointer.Width) + lblWaitbar.Left
End Sub

Private Sub SetWaitPointerPos()
  mPos = 1
  lblWaitBarPointer.Left = lblWaitbar.Left
  lblWaitBarPointer.Top = lblWaitbar.Top
  lblWaitBarPointer.ZOrder
End Sub

Private Sub tmrWaitBarControl_Timer()
  MoveWaitBar
End Sub

Private Sub UserControl_Resize()
Dim lngNumber As Long

  lngNumber = (UserControl.Width - (2.5 * lblWaitbar.Left)) / lblWaitBarPointer.Width
  lblWaitbar.Caption = String(lngNumber, "g")
  ResizeForm
End Sub

Private Sub ResizeForm()
  If UserControl.BorderStyle = 1 Then
    lblWaitbar.Left = 60
    lblWaitbar.Top = 60
  Else
    lblWaitbar.Left = 0
    lblWaitbar.Top = 0
  End If
  
  UserControl.Height = lblWaitbar.Height + (2.5 * lblWaitbar.Top)
  UserControl.Width = lblWaitbar.Width + (2.5 * lblWaitbar.Left)
End Sub

'{---------------------Ende Private Klassenmethoden------------------------}



'{---------------------Öffentliche Klassenmethoden------------------------}

Public Sub StartWaitBar()
  SetWaitPointerPos
  mblnMoveUp = True
  lblWaitBarPointer.Visible = True
  tmrWaitBarControl.Enabled = True
End Sub

Public Sub StopWaitBar()
  lblWaitBarPointer.Visible = False
  tmrWaitBarControl.Enabled = False
End Sub

Public Sub ShoInStatusBar _
(ByVal hWnd_SBar As Long _
, Optional ByVal nPanel As Long = 1 _
, Optional ByVal XPos As Long = 0 _
, Optional ByVal YPos As Long = 0 _
, Optional ByVal hWnd_Obj As Long = 0)

  If hWnd_Obj = 0 Then hWnd_Obj = Me.hwnd
  mObjectToStatusBar.SetObjectToStatusBar hWnd_Obj, hWnd_SBar, nPanel, XPos, YPos
End Sub

