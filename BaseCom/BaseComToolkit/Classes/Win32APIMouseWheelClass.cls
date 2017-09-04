VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32APIMouseWheelClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Win32APIMouseWheel
'    Project    : ToolKits
'
'    Description: Stellt Mausradunterstützung zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum WM_CONST
  WM_ACTIVATEAPP = &H1C
  WM_CAPTURECHANGED = &H215
  WM_CHANGECBCHAIN = &H30D
  WM_CHAR = &H102
  WM_CLOSE = &H10
  WM_COMMAND = &H111
  WM_COMPACTING = &H41
  WM_CONTEXTMENU = &H7B
  WM_COPYDATA = &H4A
  WM_CTLCOLORSCROLLBAR = &H137
  WM_DESTROY = &H2
  WM_DEVMODECHANGE = &H1B
  WM_DEVICECHANGE = &H219
  WM_DISPLAYCHANGE = &H7E
  WM_DRAWCLIPBOARD = &H308
  WM_DROPFILES = &H233
  WM_ENDSESSION = &H16
  WM_ENTERMENULOOP = &H211
  WM_ENTERSIZEMOVE = &H231
  WM_ERASEBKGND = &H14
  WM_EXITMENULOOP = &H212
  WM_EXITSIZEMOVE = &H232
  WM_FONTCHANGE = &H1D
  WM_GETMINMAXINFO = &H24
  WM_HOTKEY = &H312
  WM_HSCROLL = &H114
  WM_KEYDOWN = &H100
  WM_KEYUP = &H101
  WM_KILLFOCUS = &H8
  WM_LBUTTONDBLCLK = &H203
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_MBUTTONDBLCLK = &H209
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
  WM_MENUCHAR = &H120
  WM_MENUSELECT = &H11F
  WM_MOUSEACTIVATE = &H21
  WM_MOUSEMOVE = &H200
  WM_MOUSEWHEEL = &H20A
  WM_MOVE = &H3
  WM_MOVING = &H216
  WM_NCACTIVATE = &H86
  WM_NCHITTEST = &H84
  WM_NCLBUTTONDBLCLK = &HA3
  WM_NCLBUTTONDOWN = &HA1
  WM_NCLBUTTONUP = &HA2
  WM_NCMBUTTONDBLCLK = &HA9
  WM_NCMBUTTONDOWN = &HA7
  WM_NCMBUTTONUP = &HA8
  WM_NCMOUSEMOVE = &HA0
  WM_NCPAINT = &H85
  WM_NCRBUTTONDBLCLK = &HA6
  WM_NCRBUTTONDOWN = &HA4
  WM_NCRBUTTONUP = &HA5
  WM_NOTIFY = &H4E
  WM_OTHERWINDOWCREATED = &H42
  WM_OTHERWINDOWDESTROYED = &H43
  WM_PAINT = &HF
  WM_PALETTECHANGED = &H311
  WM_PALETTEISCHANGING = &H310
  WM_POWER = &H48
  WM_POWERBROADCAST = &H218
  WM_QUERYENDSESSION = &H11
  WM_QUERYNEWPALETTE = &H30F
  WM_QUERYOPEN = &H13
  WM_RBUTTONDBLCLK = &H206
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_SETCURSOR = &H20
  WM_SETFOCUS = &H7
  WM_SETTINGCHANGE = &H1A
  WM_SIZE = &H5
  WM_SIZING = &H214
  WM_SPOOLERSTATUS = &H2A
  WM_SYSCOLORCHANGE = &H15
  WM_SYSCOMMAND = &H112
  WM_SYSKEYDOWN = &H104
  WM_SYSKEYUP = &H105
  WM_TIMECHANGE = &H1E
  WM_USERCHANGED = &H54
  WM_VSCROLL = &H115
  WM_WININICHANGE = &H1A
End Enum

Public Event After(ByRef uMsg As Long, ByRef wParam As Long, _
    ByRef lParam As Long)
Public Event Before(ByRef uMsg As Long, ByRef wParam As Long, _
    ByRef lParam As Long, ByRef retVal As Long)

Private Nr As Long



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
  If Nr Then DoUnhook Nr
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
Public Sub Hook(ByVal hwnd As Long, ParamArray uMsg() As Variant)

  If Nr Then DoUnhook Nr
  Dim v As Variant: v = uMsg
  Nr = DoHook(ObjPtr(Me), hwnd, v)
End Sub

Public Sub RaiseAfter(ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
  RaiseEvent After(uMsg, wParam, lParam)
End Sub

Public Sub RaiseBefore(ByRef uMsg As Long, _
    ByRef wParam As Long, ByRef lParam As Long, ByRef retVal As Long)
  RaiseEvent Before(uMsg, wParam, lParam, retVal)
End Sub

Public Sub Unhook()
  DoUnhook Nr
  Nr = 0
End Sub
'---------'---------'---------'---------'---------'---------'---------



