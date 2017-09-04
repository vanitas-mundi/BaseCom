Attribute VB_Name = "ControlFormHelperModule"
Option Explicit

Private mForm As Object
Private mCtl As Object
Private mPos As Long

Private Const WM_COMMAND = &H111
Private Const WM_SETCURSOR = &H20
Private Const WM_NCPAINT = &H85
Private Const WM_MOVE = &H3
Private Const SWP_FRAMECHANGED = &H20

Private Type CWPSTRUCT
  lParam As Long
  wParam As Long
  Message As Long
  hwnd As Long
End Type

Public Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetParent Lib "user32" (ByVal _
       hwnd As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal _
       hwnd As Long, lpRect As Rect) As Long

Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias _
       "SetWindowsHookExA" (ByVal idHook&, ByVal lpfn&, _
       ByVal hmod&, ByVal dwThreadId&) As Long

Private Function HookProc _
(ByVal nCode As Long, ByVal wParam As Long, Inf As CWPSTRUCT) As Long

  Dim r As Rect
  Static LastParam&

  Select Case True
  
  Case Inf.hwnd = GetParent(mCtl.hwnd)
  
    If Inf.Message = WM_COMMAND Then
      mForm.Caption = "Press"
      Select Case LastParam
      Case mCtl.hwnd
        'Call mForm.Command1_Click - Funktioniert noch nicht richtig bei bedarf bearbeiten
      End Select
    ElseIf Inf.Message = WM_SETCURSOR Then
      LastParam = Inf.wParam
    End If
  
  Case Inf.hwnd = mForm.hwnd
    If (Inf.Message = WM_NCPAINT) Or (Inf.Message = WM_MOVE) Then
      GetWindowRect mForm.hwnd, r
      SetWindowPos mCtl.hwnd, 0, r.Right - (57 + (mPos * 18)), r.Top + 6, 17, 14, SWP_FRAMECHANGED
    End If
  
  End Select
End Function

Public Function SetWindowsHookExModule _
(ByRef aForm As Object, ByRef aControl As Object, ByVal pos As Long) As Long

  Set mForm = aForm
  Set mCtl = aControl
  mPos = pos
  
  SetWindowsHookExModule = SetWindowsHookEx(4, AddressOf HookProc, 0, App.ThreadID)
End Function
