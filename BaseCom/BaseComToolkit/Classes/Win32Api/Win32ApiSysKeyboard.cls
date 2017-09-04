VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lngLogicalLeft  As Long  'logische linke Maustaste
Private lngLogicalRight As Long 'logische rechte Maustaste

'Tastaturkonstanten
Private Const VK_CONTROL = &H11     'Steuerungstaste
Private Const VK_MENU = &H12        'Alt-Taste
Private Const VK_SHIFT = &H10       'Shift-Taste
Private Const VK_LCONTROL = &HA2    'linke Steuerungstaste
Private Const VK_LMENU = &HA4       'linke Alt-Taste
Private Const VK_LSHIFT = &HA0      'linke Shift-Taste
Private Const VK_RCONTROL = &HA3    'rechte Steuerungstaste
Private Const VK_RMENU = &HA5       'rechte Alt-Taste
Private Const VK_RSHIFT = &HA1      'rechte Shift-Taste
'Mauskonstanten
Private Const VK_LBUTTON = &H1      'linke Maustaste (physikalisch)
Private Const VK_MBUTTON = &H4      'mittlere Maustaste
Private Const VK_RBUTTON = &H2      'rechte Maustaste (physikalisch)
    
'SystemMetrics-Konstanten
Private Const SM_SWAPBUTTON = 23

'Funktionen
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    
Private Sub Class_Initialize()

  If (GetSystemMetrics(SM_SWAPBUTTON) = 0) Then
    lngLogicalLeft = VK_LBUTTON
    lngLogicalRight = VK_RBUTTON
  Else
    lngLogicalLeft = VK_RBUTTON
    lngLogicalRight = VK_LBUTTON
  End If

End Sub

Public Property Get MouseButtonsSwapped() As Boolean
  MouseButtonsSwapped = (GetSystemMetrics(SM_SWAPBUTTON) <> 0)
End Property

Public Property Get MouseLeftButtonPressed() As Boolean
  MouseLeftButtonPressed = (GetAsyncKeyState(lngLogicalLeft) < 0)
End Property

Public Property Get MouseMiddleButtonPressed() As Boolean
  MouseMiddleButtonPressed = (GetAsyncKeyState(VK_MBUTTON) < 0)
End Property

Public Property Get MouseRightButtonPressed() As Boolean
  MouseRightButtonPressed = (GetAsyncKeyState(lngLogicalRight) < 0)
End Property

Public Property Get ControlKeyPressed() As Boolean
  ControlKeyPressed = (GetAsyncKeyState(VK_CONTROL) < 0)
End Property

Public Property Get ShiftKeyPressed() As Boolean
  ShiftKeyPressed = (GetAsyncKeyState(VK_SHIFT) < 0)
End Property

Public Property Get AltKeyPressed() As Boolean
  AltKeyPressed = (GetAsyncKeyState(VK_MENU) < 0)
End Property

Public Property Get KeyPressed(ByVal lngKeyCode As Long) As Boolean
  KeyPressed = (GetAsyncKeyState(lngKeyCode) < 0)
End Property

