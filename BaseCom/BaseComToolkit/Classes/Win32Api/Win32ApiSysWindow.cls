VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lngHDCs()  As Long
Private lngHWnds() As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hDC As Long) As Long

Private Sub Class_Initialize()
  ReDim lngHDCs(0)
  ReDim lngHWnds(0)
End Sub

Private Sub Class_Terminate()
  Dim iHandle   As Long
  Dim lngResult As Long

  For iHandle = UBound(lngHDCs) - 1 To 0 Step -1
    lngResult = ReleaseDC(lngHWnds(iHandle), lngHDCs(iHandle))
  Next

End Sub

Public Property Get SystemHWnd() As Long
  SystemHWnd = GetDesktopWindow()
End Property

Public Property Get SystemHDC() As Long
  SystemHDC = GetDC(GetDesktopWindow())
End Property

Public Function GetDC(ByVal hwnd As Long) As Long
  Dim hDC As Long
  hDC = GetWindowDC(hwnd)
  GetDC = hDC
  lngHDCs(UBound(lngHDCs)) = hDC
  lngHWnds(UBound(lngHDCs)) = hwnd
  ReDim Preserve lngHDCs(UBound(lngHDCs) + 1)
  ReDim Preserve lngHWnds(UBound(lngHWnds) + 1)
End Function

