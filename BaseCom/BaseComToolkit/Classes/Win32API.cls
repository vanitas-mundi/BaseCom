VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Win32API
'    Project    : ToolKits
'
'    Description: Stellt API-Funktionalitäten zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mWin32ApiProfessional As Win32ApiProfessional

Private Const WM_PAINT = &HF
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9
'keine Einträge mehr vorhanden
Private Const ERROR_NO_MORE_ITEMS = 259&
'Ungültige Daten
Private Const ERROR_INVALID_DATA = 13&
'erfolgreiche Ausführung, kein Fehler
Private Const ERROR_SUCCESS = 0&
'OLE-Registrierungen, Shell-Erweiterungen, Klasseneinträge
Private Const HKEY_CLASSES_ROOT = &H80000000
'Zugriffsrechte
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RightS_READ = (READ_CONTROL)
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
'Schlüssel lesen (Kombination aus anderen Attributen)
Private Const KEY_READ = ((STANDARD_RightS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

'IsWindowVisible
' ermittelt, ob ein Fenster sichtbar oder verborgen ist.
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

'IsIconic
' ermittelt, ob ein Fenster minimiert ist.
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

'ShowWindow
' setzt den Anzeigestatus eines Fensters.
Private Declare Function ShowWindow _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal nCmdShow As Long) As Long

'BringWindowToTop
' bringt ein Fenster an die oberste Position im Fensterstapel
' und legt es auf anderen überlappenden Fenster auf dem
' Bildschirm ab.
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'SetForegroundWindow
' bringt den Thread, der das vorgegebene Fenster erzeugt hat,
' in den Vordergrund und aktiviert das Fenster.
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess _
                Lib "kernel32" (ByVal hProcess As Long, _
                                lpExitCode As Long) As Long

Private Declare Function OpenProcess _
                Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                ByVal bInheritHandle As Boolean, _
                                ByVal dwProcessID As Long) As Long

Private Declare Function WaitForSingleObject _
                Lib "kernel32" (ByVal hHandle As Long, _
                                ByVal dwMilliseconds As Long) As Long

Private Declare Function LoadLibrary _
                Lib "kernel32" _
                Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetProcAddress _
                Lib "kernel32" (ByVal hModule As Long, _
                                ByVal lpProcName As String) As Long

Private Declare Function CreateThread _
                Lib "kernel32" (lpThreadAttributes As Any, _
                                ByVal dwStackSize As Long, _
                                ByVal lpStartAddress As Long, _
                                ByVal lParameter As Long, _
                                ByVal dwCreationFlags As Long, _
                                lpThreadID As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function GetExitCodeThread _
                Lib "kernel32" (ByVal hThread As Long, _
                                lpExitCode As Long) As Long

'Registrierschlüssel mit bestimmten Zugriffsrechten öffnen
Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       ByRef phkResult As Long) As Long

'Unterschlüssel zum angegebenen Schlüssel ermitteln
Private Declare Function RegEnumKey _
                Lib "advapi32.dll" _
                Alias "RegEnumKeyA" (ByVal hKey As Long, _
                                     ByVal dwIndex As Long, _
                                     ByVal lpName As String, _
                                     ByVal cbName As Long) As Long

'Standardwert eines Schlüssels abfragen
Private Declare Function RegQueryValue _
                Lib "advapi32.dll" _
                Alias "RegQueryValueA" (ByVal hKey As Long, _
                                        ByVal lpSubKey As String, _
                                        ByVal lpValue As String, _
                                        ByRef lpcbValue As Long) As Long

'geöffneten Registrierschlüssel schließen
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliseconds As Long)



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mWin32ApiProfessional = New Win32ApiProfessional
End Sub

Private Sub Class_Terminate()
  Set mWin32ApiProfessional = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get Win32ApiProfessional() As Win32ApiProfessional
  Set Win32ApiProfessional = mWin32ApiProfessional
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RegisterFile
' Description:       Mit der Funktion "RegisterFile" können dll,exe,olb
'                    und tlb-Dateien registriert und deregistriert werden.
'                    Aufgerufen wird das Programm \windows\regtlib.exe
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:40:34
'
' Parameters :       sFile (String)
'                    Register (Boolean)
'--------------------------------------------------------------------------------
Public Function RegisterFile(ByVal sFile As String, Register As Boolean) As Boolean
  'Der Parameter sFile enthält die zu
  'registrierende Datei (inkl. Pfad)
  'Register: True  -> Datei soll registriert werden
  '          False -> Datei soll deregistriert werden
  Dim result As Boolean
  Dim Lib    As Long
  Dim sProc  As String
  Dim r1     As Long
  Dim r2     As Long
  Dim Thread As Long
  On Local Error GoTo RegError
  result = False
  Lib = LoadLibrary(sFile)

  If Lib Then
    sProc = IIf(Register, "DllRegisterServer", "DllUnregisterServer")
    r1 = GetProcAddress(Lib, sProc)

    If r1 Then
      Thread = CreateThread(ByVal 0, 0, ByVal r1, ByVal 0, 0, r2)

      If Thread Then
        r2 = WaitForSingleObject(Thread, 10000)

        If r2 Then
          'Fehler aufgetreten
          FreeLibrary Lib
          r2 = GetExitCodeThread(Thread, r2)
          ExitThread r2

          Exit Function

        End If

        CloseHandle Thread
        'OK
        result = True
      End If
    End If

    FreeLibrary Lib
  End If

RegError:
  RegisterFile = result
  Exit Function
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RegisterFileEx
' Description:       Datei registrieren/deregistrieren
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:41:50
'
' Parameters :       filename (String)
'                    Register (Boolean = True)
'--------------------------------------------------------------------------------
Public Function RegisterFileEx(ByVal filename As String, _
                               Optional ByVal Register As Boolean = True) As Long

  Const PROCESS_QUERY_INFORMATION = &H400
  Dim CmdLine     As String
  Dim dwProcessID As Long
  Dim hProcess    As Long

  If Register Then
    ' Datei registrieren
    CmdLine = "regtlib " & filename
  Else
    ' Datei deregistrieren
    CmdLine = "regtlib -u " & filename
  End If

  ' Prozess-ID
  dwProcessID = VBA.Interaction.Shell(CmdLine)
  ' Prozess öffnen
  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, dwProcessID)

  If hProcess = 0 Then
    ' Fehler!
    RegisterFileEx = Err.LastDllError
    Call ErrMsg("OpenProcess " & CmdLine, RegisterFileEx)

    Exit Function

  End If

  ' Datei registrieren/deregistrieren
  Call WaitForSingleObject(hProcess, 1000)
  Call GetExitCodeProcess(hProcess, RegisterFileEx)

  If RegisterFileEx <> 0 Then Call ErrMsg("GetExitCodeProcess " & CmdLine, RegisterFileEx)
  ' Prozess schliessen
  Call CloseHandle(hProcess)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ErrMsg
' Description:       Fehhlermeldung auswerten
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:42:20
'
' Parameters :       Info (String)
'                    Errno (Long)
'--------------------------------------------------------------------------------
Private Sub ErrMsg(ByVal Info As String, ByVal Errno As Long)
  Dim Errtext As String

  Select Case Errno

    Case 0: Errtext = " ok"

    Case 3: Errtext = " File not found or not a Typelib"

    Case 5: Errtext = " File was not registered"

    Case &H103: Errtext = "Process still active"

    Case Else: Errtext = "??"
  End Select

  MsgBox Info & vbCr & "Error=" & Errno & Errtext, vbExclamation, VB.app.EXEName
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       WindowToTop
' Description:       Fenster in den Vordergrund bringen
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       07.01.2015-13:20:23
'
' Parameters :       hwnd (Long)
'--------------------------------------------------------------------------------
Public Function WindowToTop(ByVal hwnd As Long) As Long
  
  If IsWindowVisible(hwnd) Then
    ShowWindow hwnd, IIf(IsIconic(hwnd), SW_RESTORE, SW_SHOW)
    DoEvents
    BringWindowToTop hwnd
    SetForegroundWindow hwnd
  End If

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetProgramByExtension
' Description:       Sucht nach der Dateiextension und gibt das Kommando zum
'                    Öffnen der Datei zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:42:59
'
' Parameters :       strExtension (String)
'--------------------------------------------------------------------------------
Public Function GetProgramByExtension(ByVal strExtension As String) As String

  Dim hdlKey      As Long
  Dim lngIndex    As Long
  Dim strExt      As String
  Dim lngExtLen   As Long
  Dim strValue    As String
  Dim lngValueLen As Long
  Dim strProg     As String
  Dim lngProgLen  As Long
  Dim lngReturn   As Long
  Dim intPos      As Integer
  Dim intEnv      As Integer
  Dim strEnv      As String
  Dim strWert     As String
  Dim strCommand  As String

  If Left$(strExtension, 1) <> "." Then strExtension = "." & strExtension

  '// Key öffnen
  If ERROR_SUCCESS = RegOpenKeyEx(HKEY_CLASSES_ROOT, vbNullString, 0, KEY_READ, hdlKey) Then

    lngIndex = 0
    'strcommand = B_TH(84051, "Dateierweiterung nicht bekannt")
    '// Variablenpuffer vorbelegen
    strExt = Space$(255)
    lngExtLen = Len(strExt)

    '// Alle .ext durchgehen
    Do While Not ERROR_NO_MORE_ITEMS = RegEnumKey(hdlKey, lngIndex, strExt, lngExtLen)
      '// Abschließende Null (Chr(0)) entfernen
      strExt = Trim$(CStr(strExt))

      If Right$(strExt, 1) = Chr(0) Then
        strExt = Left$(strExt, Len(strExt) - 1)
      End If

      '// Wenn keine Extension gefunden wurde, aufhören
      If Left$(strExt, 1) > "." Then Exit Do

      '// Wenn die richtige Extension gefunden wurde, den Standardwert auslesen
      If UCase(strExtension) = UCase(strExt) Then

        '// Variablenpuffer vorbelegen
        strValue = Space$(255)
        lngValueLen = Len(strValue)

        '// Standardwert des Key's holen
        If ERROR_SUCCESS = RegQueryValue(HKEY_CLASSES_ROOT, strExt, strValue, lngValueLen) Then

          '// Abschließende Null (Chr(0)) entfernen
          strValue = Trim$(CStr(strValue))

          If Right$(strValue, 1) = Chr(0) Then
            strValue = Left$(strValue, lngValueLen - 1)
          End If

          '// Variablenpuffer vorbelegen
          strProg = Space$(255)
          lngProgLen = Len(strProg)

          '// Das open-command der Extension über den vorher gefundenen Wert holen
          lngReturn = RegQueryValue(HKEY_CLASSES_ROOT, strValue & "\shell\open\command", strProg, lngProgLen)

          If lngReturn = ERROR_SUCCESS Then
            '// Abschließende Null (Chr(0)) entfernen
            strProg = Trim$(strProg)

            If Right$(strProg, 1) = Chr(0) Then
              strProg = Left$(strProg, lngProgLen - 1)
            End If

          ElseIf lngReturn = ERROR_INVALID_DATA Then
            '// ERROR_INVALID_DATA kommt evtl. anstatt ERROR_SUCCESS, die Daten
            '// sind aber in Ordnung, nur die Rückgabe der Länge nicht
            '// -> Abschließende Null (Chr(0)) anders entfernen
            strProg = Trim$(strProg)
            strProg = Left$(strProg, InStr(strProg, Chr(0)) - 1)
          Else
            '// Für diese Extension ist kein Open-Befehl definiert
            strCommand = "Befehl zum öffnen nicht gefunden"
            Exit Do
          End If

          '// Dieses Drecks-ProzentEins wegschmeißen!
          '// Das kann entweder in " stehen
          intPos = InStrRev(strProg, Chr(34) & "%1" & Chr(34))

          If intPos > 0 Then
            strProg = RTrim$(Left$(strProg, intPos - 1))
          End If

          '// oder nicht
          intPos = InStrRev(strProg, "%1")

          If intPos > 0 Then
            strProg = RTrim$(Left$(strProg, intPos - 1))
          End If

          '// und die Scheiß Umgebungsvariablen ersetzen
          If InStr(strProg, "%") Then
            strProg = UCase(strProg)
            intEnv = 1

            Do
              strEnv = UCase$(Environ$(intEnv))
              intPos = InStr(strEnv, "=")

              If intPos > 0 Then
                strWert = Right$(strEnv, Len(strEnv) - intPos)
                strEnv = Left$(strEnv, intPos - 1)
              End If

              strProg = Replace(strProg, "%" & strEnv & "%", strWert)
              intEnv = intEnv + 1
            Loop Until strEnv = ""

          End If

          strCommand = strProg
          '// und wenn man dann soweit gekommen ist, kann man die Schleife abbrechen
          Exit Do
        End If
      End If

      '// Variablenpuffer neu vorbelegen
      strExt = Space$(255)
      lngExtLen = Len(strExt)

      If lngIndex > 2000 Then Exit Do
      lngIndex = lngIndex + 1
    Loop

    '// Key schließen
    RegCloseKey hdlKey
  Else
    strCommand = ""
  End If

  GetProgramByExtension = strCommand
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       KeyPressed
' Description:       Prüft, ob die Taste des angegebenen KeyCodes gedrückt wird.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:44:25
'
' Parameters :       KeyCode (Long)
'--------------------------------------------------------------------------------
Public Function KeyPressed(KeyCode As Long) As Boolean
  Dim result As Integer: result = GetAsyncKeyState(KeyCode)
  KeyPressed = (result = -32767) Or (result = 1)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SetFormInForground
' Description:       Setzt ein Formular in den Vordergrund oder hebt dieses Verhalten wieder auf.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:45:25
'
' Parameters :       f (Object)
'                    blnForeGround (Boolean)
'--------------------------------------------------------------------------------
Public Sub SetFormInForground(ByRef f As Object, ByVal blnForeGround As Boolean)

  If blnForeGround Then
    '{Fenster immer im Vordegrund}
    Call SetWindowPos(f.hwnd, -1, 0, 0, 0, 0, 3)
  Else
    '{Fenster im Normalzustand}
    Call SetWindowPos(f.hwnd, -2, 0, 0, 0, 0, 3)
  End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       Delay
' Description:       Friert den Prozess milliseconds lang ein.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       21.01.2015-15:18:07
'
' Parameters :       milliseconds (Long)
'--------------------------------------------------------------------------------
Public Sub Delay(ByVal milliseconds As Long)
  Sleep milliseconds
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RedrawWindow
' Description:       Erzwingt ein Neuzeichner des Formulars.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       21.01.2015-15:19:34
'
' Parameters :       lngHwnd (Long)
'--------------------------------------------------------------------------------
Public Sub RedrawWindow(ByVal lngHwnd As Long)
  SendMessage lngHwnd, WM_PAINT, 0&, 0&
End Sub

