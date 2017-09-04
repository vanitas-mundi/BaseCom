Attribute VB_Name = "Win32APIMouseWheel"
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Win32APIMouseWheelClass
'    Project    : ToolKits
'
'    Description: Setllt Hilfsfunktionen für die Mausradunterstützung zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Private Const GWL_WNDPROC = -4
Private Const MAX_HASH = 257
Private Const MAXHOOKS = 9

Private Type typHook
  hwnd As Long
  MsgHookPtr As Long
  ProcAddr As Long
  WndProc As Long
  uMsgCount As Long
  uMsg(MAX_HASH - 1) As Boolean
  uMsgCol As Collection
End Type

Private Declare Function CallWindowProcA _
                Lib "user32" (ByVal lpPrevWndFunc As Long, _
                              ByVal hwnd As Long, _
                              ByVal msg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (dest As Any, _
                                       source As Any, _
                                       ByVal NumBytes As Long)
Private Declare Function SetWindowLongA _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal nIndex As Long, _
                              ByVal dwNewLong As Long) As Long

Private Initialized As Boolean
Private arrHook()   As typHook
Private NrCol       As Collection

'---------------------- Konstruktor der Klasse ----------------------------------

'---------------------- Zugriffsmethoden der Klasse -----------------------------

'---------------------- Ereignismethoden der Klasse -----------------------------

'---------------------- Private Methoden der Klasse -----------------------------
Private Function Addr2Long(ByVal Addr As Long) As Long
  Addr2Long = Addr
End Function

'---------------------- Öffentliche Methoden der Klasse -------------------------
Public Function DoHook(ByVal ObjPtr As Long, ByVal hwnd As Long, uMsg As Variant) As Long
  
  If Not Initialized Then
    ReDim Preserve arrHook(1 To MAXHOOKS)
    arrHook(1).WndProc = Addr2Long(AddressOf WndProc1)
    arrHook(2).WndProc = Addr2Long(AddressOf WndProc2)
    arrHook(3).WndProc = Addr2Long(AddressOf WndProc3)
    arrHook(4).WndProc = Addr2Long(AddressOf WndProc4)
    arrHook(5).WndProc = Addr2Long(AddressOf WndProc5)
    arrHook(6).WndProc = Addr2Long(AddressOf WndProc6)
    arrHook(7).WndProc = Addr2Long(AddressOf WndProc7)
    arrHook(8).WndProc = Addr2Long(AddressOf WndProc8)
    arrHook(9).WndProc = Addr2Long(AddressOf WndProc9)
    Set NrCol = New Collection
    Initialized = True
  End If

  'Hook-Nr suchen:
  Dim i As Long

  For i = 1 To UBound(arrHook)

    If arrHook(i).MsgHookPtr = 0 Then Exit For
  Next i

  If i > UBound(arrHook) Then
    ReDim Preserve arrHook(1 To 2 * UBound(arrHook))
  End If

  If i > MAXHOOKS Then
    NrCol.Add i, "H" & hwnd
    arrHook(i).WndProc = Addr2Long(AddressOf WndProcX)
  End If

  DoHook = i

  'Hook einrichten:
  arrHook(i).uMsgCount = UBound(uMsg) + 1

  If arrHook(i).uMsgCount Then
    Erase arrHook(i).uMsg
    Set arrHook(i).uMsgCol = New Collection
    Dim j As Long

    For j = LBound(uMsg) To UBound(uMsg)
      arrHook(i).uMsg(uMsg(j) Mod MAX_HASH) = True
      arrHook(i).uMsgCol.Add True, "H" & uMsg(j)
    Next j

  End If

  arrHook(i).hwnd = hwnd
  arrHook(i).MsgHookPtr = ObjPtr
  arrHook(i).ProcAddr = SetWindowLongA(hwnd, GWL_WNDPROC, arrHook(i).WndProc)
End Function

Public Sub DoUnhook(ByVal Nr As Long)

  If Nr Then

    With arrHook(Nr)
      SetWindowLongA .hwnd, GWL_WNDPROC, .ProcAddr
      .MsgHookPtr = 0
      Erase .uMsg
      Set .uMsgCol = Nothing

      If Nr > MAXHOOKS Then NrCol.Remove "H" & .hwnd
    End With

  End If

End Sub

Public Function WndProc(ByVal Nr As Long, _
                        ByVal hwnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
  Dim oMsgHook As Win32APIMouseWheelClass 'MsgHook
  Dim retVal   As Long
  Dim Ok       As Boolean

  With arrHook(Nr)

    If .uMsgCount > 0 Then
      If .uMsg(uMsg Mod MAX_HASH) Then
        On Error Resume Next
        Ok = .uMsgCol("H" & uMsg) And (Err = 0)
        On Error GoTo 0
      End If

    Else
      Ok = True
    End If

    If Ok Then
      CopyMemory oMsgHook, .MsgHookPtr, 4
      oMsgHook.RaiseBefore uMsg, wParam, lParam, retVal

      If uMsg Then
        retVal = CallWindowProcA(.ProcAddr, hwnd, uMsg, wParam, lParam)
        oMsgHook.RaiseAfter uMsg, wParam, lParam
      End If

      CopyMemory oMsgHook, 0&, 4
    Else
      retVal = CallWindowProcA(.ProcAddr, hwnd, uMsg, wParam, lParam)
    End If

  End With

  WndProc = retVal
End Function

Public Function WndProc1(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc1 = WndProc(1, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc2(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc2 = WndProc(2, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc3(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc3 = WndProc(3, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc4(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc4 = WndProc(4, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc5(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc5 = WndProc(5, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc6(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc6 = WndProc(6, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc7(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc7 = WndProc(7, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc8(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc8 = WndProc(8, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProc9(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  WndProc9 = WndProc(9, hwnd, uMsg, wParam, lParam)
End Function

Public Function WndProcX(ByVal hwnd As Long, _
                         ByVal uMsg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
  Dim Nr As Long
  On Error Resume Next
  Nr = NrCol("H" & hwnd)
  On Error GoTo 0

  If Nr Then WndProcX = WndProc(Nr, hwnd, uMsg, wParam, lParam)
End Function

