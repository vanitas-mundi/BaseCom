VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaitbarObjToStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' **********************************************************************
' *
' *   Diese Funktion stammt im Ursprung aus dem vbarchiv
' *   von Dieter Otter und wurde von VB-Power.net modifiziert.
' *   Ursprünglicher Tipp: ProgressBar in der StatusBar
' *   von Dietmar G. Bayer, www.vbarchiv.net
' *
' *   SetObjectToStatusBar - beliebiges Objekt in der StatusBar
' *
' *   Dezember 2004, VB-Power.net   www.vb-power.net
' *
' **********************************************************************

' Benötigte API-Deklarationen
Private Declare Function SetParent Lib "user32" ( _
  ByVal hWndChild As Long, _
  ByVal hWndNewParent As Long) As Long

Private Declare Function SendMessageAny Lib "user32" _
  Alias "SendMessageA" ( _
  ByVal hwnd As Long, _
  ByVal msg As Long, _
  ByVal wParam As Long, _
  lParam As Any) As Long

Private Declare Function MoveWindow Lib "user32" ( _
  ByVal hwnd As Long, _
  ByVal x As Long, ByVal Y As Long, _
  ByVal nWidth As Long, ByVal nHeight As Long, _
  ByVal bRepaint As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
   Bottom As Long
End Type

Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

' Objekt in StatusBar setzen
Public Sub SetObjectToStatusBar(ByVal hWnd_Obj As Long, _
  ByVal hWnd_SBar As Long, _
  ByVal nPanel As Long, _
  Optional XPos As Long = 0, _
  Optional YPos As Long = 0)
   
  Dim cXPos As Long
  Dim cYPos As Long
  Dim R As RECT
  
  ' Border in Pixel
  Const cBorder As Long = 1
   
  ' Screenkoordinaten in Pixel umrechnen
  cXPos = XPos / Screen.TwipsPerPixelX
  cYPos = YPos / Screen.TwipsPerPixelY
   
  ' Ausmaße des Panel ermitteln
  SendMessageAny hWnd_SBar, SB_GETRECT, nPanel - 1, R

  ' Objekt ein neues zuhause geben...
  SetParent hWnd_Obj, hWnd_SBar

  ' ... und korrekt positionieren
  MoveWindow hWnd_Obj, R.Left + cXPos + cBorder, _
    R.Top + cYPos + cBorder, _
    R.Right - R.Left - cXPos - cBorder, _
    R.Bottom - R.Top - cYPos - cBorder, True
End Sub





