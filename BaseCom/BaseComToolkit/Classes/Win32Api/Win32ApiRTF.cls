VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiProfessionalRTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private RTF As RichTextBox
Event Scroll(ByVal RelativePosition As Single)
Private lpPrevWndProc As Long       'vorherige Adresse der Window-Prozedur
Private hWnd          As Long
Private ScrInfo       As SCROLLINFO
Private ScrBarInfo    As SCROLLBARINFO
Private lngScrollMax  As Long
Private sngScroll     As Single
Private bAutoSize     As Boolean
Private bSP6          As Boolean             'True, wenn GetScrollBarInfo läuft (ab NT 4 SP 6)
    
Private Type SCROLLINFO
  cbSize As Long
  fMask As Long
  nMin As Long
  nMax As Long
  nPage As Long
  nPos As Long
  nTrackPos As Long
End Type
        
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
    
Private Type rgstateSCROLLBARINFO
  rgstateSCROLLBAR As Long
  rgstateTOPARROW As Long
  rgstatePAGEUP As Long
  rgstateTHUMB As Long
  rgstatePAGEDOWN As Long
  rgstateBOTTOMARROW As Long
End Type
    
Private Type SCROLLBARINFO
  cbSize As Long
  rcScrollBar As RECT
  dxyLineButton  As Long
  xyThumbTop As Long
  xyThumbBottom As Long
  Reserved As Long
  rgstate As rgstateSCROLLBARINFO
End Type

Private Const GWL_WNDPROC = -4
Private Const WM_VSCROLL = &H115
Private Const WM_NCHITTEST = &H84
Private Const SB_ENDSCROLL = 8
Private Const SB_THUMBTRACK = 5
Private Const OBJID_VSCROLL = &HFFFFFFFB
    
Private Const SB_VERT = 1
Private Const SIF_RANGE = 1
Private Const SIF_PAGE = 2
Private Const SIF_POS = 4
Private Const SIF_TRACKPOS = &H10

Private Declare Function GetScrollBarInfo _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal idObject As Long, _
                                  psbi As SCROLLBARINFO) As Long
Private Declare Function GetScrollInfo _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal fnBar As Long, _
                                  lpsi As SCROLLINFO) As Long
Private Declare Function GetScrollPos _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal nBar As Long) As Long
Private Declare Function GetScrollRange _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal nBar As Long, _
                                  lpMinPos As Long, _
                                  lpMaxPos As Long) As Long
Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndProc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long
Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
 
Public Sub Initialize()
  On Error GoTo errHandler
  Set RTF = frmMain.RTF
  hWnd = RTF.hWnd
  Call Hook
  ScrInfo.cbSize = Len(ScrInfo)
  ScrInfo.fMask = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS
  ScrBarInfo.cbSize = Len(ScrBarInfo)
  'Testzugriff auf die versionskritische GetScrollBarInfo
  On Error Resume Next
  GetScrollBarInfo hWnd, SB_VERT, ScrBarInfo

  If Err.number = 0 Then
    bSP6 = True
  End If

  Exit Sub

errHandler:

  If InStr(1, Err.source, ":") = "" Then Err.source = "frmInfo:AutoSizeRTF"
  Err.Raise Err.number
End Sub

Private Sub Class_Terminate()
  Call Unhook
End Sub

Private Sub Hook()
  lpPrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf RTFWindowProc)
End Sub

Private Sub Unhook()
  SetWindowLong hWnd, GWL_WNDPROC, lpPrevWndProc
End Sub

Public Function CallBackProc(ByVal uMsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
  Dim lngResult       As Long
  Dim lngScrollHeight As Long
  Dim lngThumbTop     As Long
  Dim lngThumbHeight  As Long
  On Error GoTo errHandler
  'Originalnachricht weiterleiten
  CallBackProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)

  'Vertical-Scroll-Nachricht suchen
  If uMsg = WM_VSCROLL Then

    'Ende eines Scroll-Vorgangs suchen
    If wParam = SB_ENDSCROLL Then
      'SCROLLINFO-Struktur anfordern
      lngResult = GetScrollInfo(hWnd, SB_VERT, ScrInfo)
      'Scroll-Position ohne Korrektur ermitteln
      sngScroll = ScrInfo.nPos / (ScrInfo.nMax - ScrInfo.nMin)

      If bSP6 Then
        'Scroll-Position mit Korrektur ermitteln
        lngResult = GetScrollBarInfo(hWnd, OBJID_VSCROLL, ScrBarInfo)
        lngScrollHeight = ScrBarInfo.rcScrollBar.Bottom - ScrBarInfo.rcScrollBar.Top - 32
        lngThumbTop = ScrBarInfo.xyThumbTop - 16
        lngThumbHeight = ScrBarInfo.xyThumbBottom - 16 - lngThumbTop
        sngScroll = (lngThumbTop + (lngThumbHeight) * lngThumbTop / (lngScrollHeight - lngThumbHeight)) / lngScrollHeight
      End If

      'Ereignis auslösen
      RaiseEvent Scroll(sngScroll)
    End If
  End If

  Exit Function

errHandler:
  Call Unhook
  Err.Raise Err.number, "frmInfo:AutoSizeRTF"
End Function

Public Sub LoadFile(ByVal strFile As String)
  RTF.LoadFile strFile

  If bAutoSize Then
    Call AutoSizeControl
  End If

End Sub

Private Sub AutoSizeControl()
  Dim hReturn  As Long             'Rückgabe der GetScrollInfo
  Dim lngHOpt  As Long             'Wert, bei dem Scrollbar noch nicht sichtbar war
  Dim lngHGone As Long            'Wert, bei dem ScrollBar verschwand
  Dim lngMin   As Long              'Min-Wert der Scrollbar aus Funktion
  Dim lngMax   As Long              'Max-Wert der Scrollbar aus Funktion
  On Error GoTo errHandler

  Do
    'Range ermitteln
    hReturn = GetScrollRange(RTF.hWnd, SB_VERT, lngMin, lngMax)

    If hReturn = 0 Then
      Err.Raise Err.LastDllError
    End If

    'Prüfen, ob ScrollBar unsichtbar wurde
    If lngMax = 0 Then
      'diese Größe merken ...
      lngHGone = RTF.Height

      '... und schrittweise Höhe reduzieren, bis ...
      Do
        RTF.Height = RTF.Height - 50
        hReturn = GetScrollRange(RTF.hWnd, SB_VERT, lngMin, lngMax)

        If hReturn = 0 Then
          Err.Raise Err.LastDllError
        End If

        If lngMax > 0 Then
          '... wieder erscheint
          'letzte unbedenkliche Höhe merken
          lngHOpt = RTF.Height + 50
          'nacheinander lngHGone und lngHOpt einstellen
          RTF.Height = lngHGone
          RTF.Height = lngHOpt
          Exit Sub
        End If

      Loop

    End If

    'RTF vergrößern
    RTF.Height = RTF.Height + 100
  Loop

  Exit Sub

errHandler:

  If InStr(1, Err.source, ":") = "" Then Err.source = "clsRTF:AutoSizeControl"
  Err.Raise Err.number
End Sub

Public Property Get Scroll() As Long
  Scroll = sngScroll
End Property

Public Property Get AutoSize() As Boolean
  AutoSize = bAutoSize
End Property

Public Property Let AutoSize(ByVal bNewValue As Boolean)
  bAutoSize = bNewValue
End Property

Public Property Let Text(ByVal strNewValue As String)
  RTF.Text = strNewValue

  If bAutoSize Then
    Call AutoSizeControl
  End If

  Call GetScrollMax
End Property

Private Sub GetScrollMax()
  Dim hReturn   As Long             'Rückgabe der GetScrollInfo
  Dim lngMinPos As Long           'Min-Wert der Scrollbar aus Funktion
  On Error GoTo errHandler
  'Range ermitteln
  hReturn = GetScrollRange(hWnd, SB_VERT, lngMinPos, lngScrollMax)

  If hReturn = 0 Then
    Err.Raise Err.LastDllError
  End If

  'lngScrollMax = lngScrollMax - 239
  Exit Sub

errHandler:

  If InStr(1, Err.source, ":") = "" Then Err.source = "clsRTF:GetScrollMax"
  Err.Raise Err.number
End Sub
