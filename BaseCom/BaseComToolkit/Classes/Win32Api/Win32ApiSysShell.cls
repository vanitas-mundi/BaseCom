VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private hWndApp       As Long             'Window-Handle der SATViewers
Private strAppTitle   As String       'Titel in Titelleiste
Private bAppIsrunning As Boolean    'Wenn True, läuft die Anwendung
Private bTopMost      As Boolean         'Wenn True, dann TOPMOST einstellen
Private rectWnd       As RECT             'RECT-Struktur

Private Const HWND_TOPMOST = -1     'oberstes Fenster aktivieren
Private Const HWND_NOTOPMOST = -2   'oberstes Fenster deaktivieren
Private Const SWP_NOSIZE = 1        'Fenster nicht veränderbar
Private Const SWP_NOMOVE = 2        'Fenster nicht bewegbar
Private Const WM_CLOSE = &H10       'Fenster soll sich beenden
Private Const INFINITE = &HFFFFFFFF 'Anweisung an WaitForSingleObject, unendlich lange zu warten
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const WM_NCCALCSIZE = &H83
Private Const WM_PAINT = &HF
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ALLCHILDREN = &H80

'RECT-Struktur für RedrawWindow
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

'verwendete Funktionen
Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long
Private Declare Function FindWindow _
                Lib "user32" _
                Alias "FindWindowA" (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long
Private Declare Function PostMessage _
                Lib "user32" _
                Alias "PostMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Private Declare Function WaitForSingleObject _
                Lib "kernel32" (ByVal hHandle As Long, _
                                ByVal dwMilliseconds As Long) As Long
Private Declare Function RedrawWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              lprcUpdate As RECT, _
                              ByVal hrgnUpdate As Long, _
                              ByVal fuRedraw As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub LaunchApplication(ByVal strApp As String, _
                             ByVal lngWindowState As VbAppWinStyle, _
                             ByVal bSetTopMost As Boolean, _
                             Optional ByVal strParam As String, _
                             Optional strWindowTitle As Variant) '
  '====================================================================================
  '2000-05-02, Klaus Prinz
  '   Start der Applikation mit Übergabe des Arguments strParam
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strApp:         Name der zu startenden Anwendung
  '   lngWindowState: WindowState
  '   bSetTopMost:    wenn True, dann TopMost
  '   strParam:       optional, an strApp zu übergebender String
  '   strWindowTitle: optional, Fenstertitel
  '====================================================================================
  Dim iFind As Long               'Zählt die FindWindow-Anläufe
  On Error GoTo errHandler
  hWndApp = 0
  CheckShellExecuteReturnCode ShellExecute(0, "open", strApp, strParam, vbNullString, lngWindowState), strApp
  bAppIsrunning = True
  bTopMost = bSetTopMost

  If Not IsMissing(strWindowTitle) Then
    strAppTitle = strWindowTitle

    Do While hWndApp = 0
      Sleep 100
      hWndApp = FindWindow(vbNullString, strWindowTitle)
      iFind = iFind + 1
      If iFind > 20 Then Exit Sub
    Loop

    Me.ResizeApplicationWindow
  End If

  Exit Sub

errHandler:
  Err.Raise Err.number, "clsSysShell:LaunchApplication"
End Sub

Public Sub ResizeApplicationWindow()
  '====================================================================================
  '2000-05-19, Klaus Prinz
  '   Resizing des Fensters
  '====================================================================================
  Dim lngResult As Long           'Rückgabe der PostMessage
  Dim lngPos    As Long              'TopMost oder nicht
  Dim lngFlags  As Long            'Flags
  On Error GoTo errHandler

  If hWndApp = 0 Then
    Err.Raise 9999, "Es konnte kein Fenster-Handle ermittelt werden."
  End If

  If Me.AppIsRunning Then
    'Verhindert "Schatten" des Fensters
    lngResult = SendMessage(hWndApp, WM_PAINT, 0&, 0&)

    'TopMost oder nicht einstellen
    If bTopMost Then
      lngPos = HWND_TOPMOST
    Else
      lngPos = HWND_NOTOPMOST
    End If

    'Flags einstellen
    If rectWnd.Right * rectWnd.Bottom = 0 Then
      lngFlags = SWP_NOMOVE Or SWP_NOSIZE
    End If

    'Funktionen aufrufen
    lngResult = SetWindowPos(hWndApp, lngPos, rectWnd.Left / 15, rectWnd.Top / 15, (rectWnd.Right - rectWnd.Left) / 15, (rectWnd.Bottom - rectWnd.Top) / 15, lngFlags)
    lngResult = RedrawWindow(hWndApp, rectWnd, 0&, RDW_UPDATENOW Or RDW_ALLCHILDREN)
  End If

  Exit Sub

errHandler:
  Err.Raise Err.number, "clsSysShell:ResizeApplicationWindow"
End Sub

Public Sub CloseApplication()
  '====================================================================================
  '2000-05-19, Klaus Prinz
  '   Beenden der Applikation
  '====================================================================================
  Dim lngResult As Long           'Rückgabe der PostMessage
  On Error GoTo errHandler

  If hWndApp = 0 Then
    Err.Raise 9999, "Es konnte kein Fenster-Handle ermittelt werden."
  End If

  lngResult = DestroyWindow(hWndApp)
  Exit Sub
  lngResult = PostMessage(hWndApp, WM_CLOSE, 0&, 0&)
  Sleep 100

  If lngResult = 0 Then
    'MsgBox GetLastError()
    Err.Raise 9999, , "Die Nachricht konnte nicht abgeliefert werden."
  End If

  lngResult = WaitForSingleObject(hWndApp, INFINITE)

  If lngResult = 0 Then
    Err.Raise 9999, , "Die Anwendung konnte nicht beendet werden."
  End If

  bAppIsrunning = False
  Exit Sub

errHandler:
  Err.Raise Err.number, "clsSysShell:CloseApplication"
End Sub

Private Sub CheckShellExecuteReturnCode(ByVal hResult As Long, ByVal strFile As String)
  '====================================================================================
  '2000-05-02, Klaus Prinz
  '   Erzeugen eines entsprechenden VB-Fehlers nebst Text
  '====================================================================================
  Dim lngError As Long            'Fehlernummer
  Dim strError As String          'Fehlertext

  If hResult > 32 Then Exit Sub   'Abbruch, wenn Rückgabe OK

  Select Case hResult

    Case 2  'SE_ERR_FNF             Datei nicht gefunden
      lngError = 53
      strError = "The application " & strFile & " could not be found."

    Case 3  'SE_ERR_PNF             Verzeichnis nicht gefunden
      lngError = 76

    Case 5  'SE_ERR_ACCESSDENIED    Zugriff verweigert
      lngError = 70
      strError = "Access to " & strFile & " has been denied."

    Case 8  'SE_ERR_OOM             nicht genügend Speicher
      lngError = 7

    Case 11 'ERROR_BAD_FORMAT       die Anwendung (exe) ist nicht in Ordnung
      lngError = 9999
      strError = "The requestet executable is corrupt."

    Case 26 'SE_ERR_SHARE           Fehler im gemeinsamen Dateizugriff
      lngError = 70

    Case 27 'SE_ERR_ASSOCINCOMPLETE unvollständige oder ungültige Zuordnung der Dateierweiterung ->SE_ERR_NOASSOC
      lngError = 9999
      strError = "The file extension could not be associatet."

    Case 28 'SE_ERR_DDETIMEOUT      es trat ein Timeout bei der zu startenden Anwendung auf
      lngError = 286
      strError = "Timeout while waiting for DDE response"

    Case 29 'SE_ERR_DDEFAIL         die DDE-Operation schlug fehl
      lngError = 9999
      strError = "The DDE operation failed due to unknown reasons."

    Case 30 'SE_ERR_DDEBUSY         andere DDE-Operationen verursachten ein Scheitern der angeforderten Operation
      lngError = 9999
      strError = "The DDE service is busy."

    Case 31 'SE_ERR_NOASSOC         der Dateierweiterung ist keine Anwendung zugeordnet
      lngError = 9999
      strError = "The file extension could not be associatet."

    Case 32 'SE_ERR_DLLNOTFOUND     DLL nicht gefunden
      lngError = 298
      strError = "System DLL could not be loaded."

    Case Else
      lngError = 9999
      strError = "An undefined error has occured."
  End Select

  Err.Raise lngError, , strError
End Sub

Public Property Get AppIsRunning() As Boolean
  '====================================================================================
  '2000-05-02, Klaus Prinz
  '   Ermitteln, ob strAppTitle ein gültiger Fenstertitel ist
  '====================================================================================
  On Error GoTo errHandler

  If strAppTitle = "" Then
    AppIsRunning = False
  Else
    AppIsRunning = FindWindow(vbNullString, strAppTitle) <> 0
  End If

  Exit Sub
    
errHandler:
  Err.Raise Err.number, "clsSysShell:AppIsRunning"
End Property

Public Sub SetWindowDimensions(ByVal lTop As Long, _
                               ByVal lLeft As Long, _
                               ByVal lWidth As Long, _
                               ByVal lHeight As Long)
  '====================================================================================
  '2000-05-19, Klaus Prinz
  '   RECT-Struktur mit übergebenen Werten füllen
  '====================================================================================
  On Error GoTo errHandler
  rectWnd.Top = lTop
  rectWnd.Left = lLeft
  rectWnd.Bottom = lTop + lHeight
  rectWnd.Right = lLeft + lWidth
  Exit Sub
    
errHandler:
  Err.Raise Err.number, "clsSysShell:SetWindowDimensions"
End Sub

