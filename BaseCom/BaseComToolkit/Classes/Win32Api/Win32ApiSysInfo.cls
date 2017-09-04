VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private bInitialized   As Boolean 'True, wenn Komponente initialisiert wurde
'Konstanten und Strukturen für Windows-Version
Private udtVersionInfo As OSVERSIONINFO
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO

  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128

End Type

'Konstanten und Strukturen für IP-Adresse
Private Const SOCKET_ERROR As Long = -1

Private Type HOSTENT

  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLength As Integer
  hAddrList As Long

End Type

Private Type WSADATA

  wVersion As Integer
  wHighVersion As Integer
  szDescription(256) As Byte
  szSystemStatus(128) As Byte
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpszVendorInfo As Long

End Type

'Strukturen für ClassID
Private Type GUID

  bytes(15) As Byte

End Type

'APIs
Private Declare Function GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                          nSize As Long) As Long
Private Declare Function GetVersionEx _
                Lib "kernel32" _
                Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                      nSize As Long) As Long
Private Declare Function GetHostName _
                Lib "WSOCK32.DLL" _
                Alias "gethostname" (ByVal HostName As String, _
                                     ByVal HostLen As Integer) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup _
                Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, _
                                   lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function GetHostByName _
                Lib "WSOCK32.DLL" _
                Alias "gethostbyname" (ByVal strHostname As String) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (hpvDest As Any, _
                                       ByVal hpvSource As Long, _
                                       ByVal cbCopy As Long)
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGUID As GUID) As Long
Private Declare Function StringFromGUID2 _
                Lib "OLE32.DLL" (pGUID As GUID, _
                                 ByVal lpszString As String, _
                                 ByVal lMax As Long) As Long

Public Sub Initialize()

  '====================================================================================
  '2001-01-16, Klaus Prinz
  '   Aufruf der Initialisierungsroutinen
  '====================================================================================
  On Error GoTo errHandler

  Call LoadVersionStructure
  Call InitializeSockets
  bInitialized = True

  Exit Sub

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsSysInfo:Initialize"
  Err.Raise Err.number
End Sub

Private Sub Class_Terminate()

  If bInitialized Then
    Call CleanupSockets
  End If

End Sub

'**************************************************************************************************
'Betriebssysteminformationen
'**************************************************************************************************
Private Sub LoadVersionStructure()
  Dim lngResult As Long           'Rückgabe
  udtVersionInfo.dwOSVersionInfoSize = Len(udtVersionInfo)
  lngResult = GetVersionEx(udtVersionInfo)

  If lngResult = 0 Then
    Err.Raise 440, "clsSysInfo:LoadVersionStructure", "GetVersionEx failed to load OSVERSIONINFO-type."
  End If

End Sub

Public Property Get Win95() As Boolean

  If Not bInitialized Then
    Me.Initialize
  End If

  If udtVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    Win95 = (udtVersionInfo.dwMinorVersion = 0)
  End If

End Property

Public Property Get Win98() As Boolean

  If Not bInitialized Then
    Me.Initialize
  End If

  If udtVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    Win98 = (udtVersionInfo.dwMinorVersion > 0)
  End If

End Property

Public Property Get WinNT() As Boolean

  If Not bInitialized Then
    Me.Initialize
  End If

  WinNT = (udtVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Property

Public Property Get NTVersion() As Long

  If Not bInitialized Then
    Me.Initialize
  End If

  If udtVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then
    NTVersion = udtVersionInfo.dwMajorVersion
  End If

End Property

Public Property Get ServicePack() As String

  If Not bInitialized Then
    Me.Initialize
  End If

  ServicePack = udtVersionInfo.szCSDVersion
End Property

'**************************************************************************************************
'Computer- und Anmeldenamen
'**************************************************************************************************
Public Property Get NetBIOSName() As String
  '====================================================================================
  '1999-11-23, Klaus Prinz
  '   Rückgabe des NetBIOS-Namen (Register Identifikation ->Computername)
  '====================================================================================
  Dim strBuffer As String * 255   'Zwischenvariable für API-Funktion
  Dim nSize     As Long               'Größe von strBuffer
  Dim lngResult As Long           'Rückgabe der Funktion

  On Error GoTo errHandler

  nSize = 255
  lngResult = GetComputerName(strBuffer, nSize)
  NetBIOSName = Left$(strBuffer, nSize)

  Exit Property

errHandler:
  Err.Raise Err.number, "clsSysInfo:NetBIOSName"
End Property

Public Property Get DNSName() As String
  '====================================================================================
  '2001-01-16, Klaus Prinz
  '   Rückgabe des Host-Namen
  '====================================================================================
  Dim strBuffer As String * 255   'Zwischenvariable für API-Funktion
  Dim lngResult As Long           'Rückgabe der Funktion

  On Error GoTo errHandler

  If Not bInitialized Then Call Me.Initialize
  lngResult = GetHostName(strBuffer, 255)

  If lngResult = SOCKET_ERROR Then
    Err.Raise 9999, , "Windows Sockets Error: " & CStr(WSAGetLastError())
  Else
    DNSName = Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
  End If

  Exit Property

errHandler:
  Err.Raise Err.number, "clsSysInfo:DNSName"
End Property

Public Property Get LoginName() As String
  '====================================================================================
  '1999-11-23, Klaus Prinz
  '   Rückgabe des Anmeldenamen
  '====================================================================================
  Dim strBuffer As String * 255   'Zwischenvariable für API-Funktion
  Dim nSize     As Long               'Größe von strBuffer
  Dim lngResult As Long           'Rückgabe der Funktion

  On Error GoTo errHandler

  nSize = 255
  lngResult = GetUserName(strBuffer, nSize)
  LoginName = Left$(strBuffer, nSize - 1)

  Exit Property

errHandler:
  Err.Raise Err.number, "clsSysInfo:LoginName"
End Property

'**************************************************************************************************
'IP-Adresse
'**************************************************************************************************
Private Sub InitializeSockets()
  '====================================================================================
  '2001-02-01, Klaus Prinz
  '   Winsock initialisieren
  '====================================================================================
  Const WSARequVersion As Integer = &H101
  Dim udtWSADATA       As WSADATA       'WSADATA-Struktur
  Dim lngReturn        As Long           'Rückgabe aus WASStartup

  On Error GoTo errHandler

  lngReturn = WSAStartup(WSARequVersion, udtWSADATA)

  If lngReturn <> 0 Then
    Err.Raise 9999, , "Winsock.dll is not responding."
  End If

  If udtWSADATA.wVersion < WSARequVersion Then
    Err.Raise 9999, , "Windows Sockets version is to small."
  End If

  Exit Sub

errHandler:
  Call CleanupSockets
  Err.Raise Err.number, "clsSysInfo:InitializeSockets"
End Sub

Private Sub CleanupSockets()
  '====================================================================================
  '2001-01-16, Klaus Prinz
  '   WinSock-Sitzung beenden
  '====================================================================================
  Dim lngReturn  As Long           'Rückgabe der Funktion
  Dim lngCounter As Long

  On Error Resume Next

  Do
    lngCounter = lngCounter + 1
    lngReturn = WSACleanup()
  Loop Until lngReturn <> 0 Or lngCounter > 10

End Sub
 
Public Property Get IPAddress() As String

  '====================================================================================
  '2001-01-16, Klaus Prinz
  '   Rückgabe der IP-Adresse
  '====================================================================================
  On Error GoTo errHandler

  If Not bInitialized Then Me.Initialize
  IPAddress = HostByName(Me.DNSName)

  Exit Property

errHandler:
  Err.Raise Err.number, "clsSysInfo:strIPAddress"
End Property

Public Function HostByName(ByVal strName As String) As String
  '====================================================================================
  '2001-02-01, Klaus Prinz
  '   Rückgabe der strName zugeordneten IP-Adresse
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strName:    DNS-Name des Rechners
  '====================================================================================
  Dim byteIPSegements(3) As Byte  'enthält die 4 Segmente
  Dim iSegment           As Long            'Zeiger in byteIPSegements
  Dim lngHostAddress     As Long      'Adresse der HOSTENT-Kopie
  Dim lngHostIP          As Long           'Adresse des Byte-Arrays aus
  Dim strIPAddress       As String      'zusammengesetzte IP-Adresse
  Dim udtHOSTENT         As HOSTENT       'HOSTENT-Struktur

  On Error GoTo errHandler

  If Not bInitialized Then Me.Initialize
  'Adresse des Prozess-HOSTENT ermitteln
  lngHostAddress = GetHostByName(strName)

  If lngHostAddress = 0 Then
    Err.Raise 9999, , strName & " konnte keine IP-Adresse zugeordnet werden."
  End If

  'Prozess-HOSTENT duplizieren
  Call CopyMemory(udtHOSTENT, lngHostAddress, LenB(udtHOSTENT))
  'Network Byte Order über Long in Byte-Array kopieren
  Call CopyMemory(lngHostIP, udtHOSTENT.hAddrList, 4)
  Call CopyMemory(byteIPSegements(0), lngHostIP, 4)

  'String aus Byte-Array zusammenstellen
  For iSegment = 0 To 3
    strIPAddress = strIPAddress & byteIPSegements(iSegment) & "."
  Next

  HostByName = Left$(strIPAddress, Len(strIPAddress) - 1)

  Exit Function

errHandler:
  Err.Raise Err.number, "clsSysInfo:HostByName"
End Function

'**************************************************************************************************
'Class-ID
'**************************************************************************************************
Public Function GetNewGUID() As String
  '====================================================================================
  '2000-08-30, Klaus Prinz
  '   neuen GUID erzeugen
  '====================================================================================
  Dim udtGUID As GUID             'GUID-Struktur
  Dim strGUID As String           'GUID-String
  Dim nSize   As Long               'Bruttolänge des Unicode-GUID

  On Error GoTo errHandler

  strGUID = Space$(78)
  'GUID erzeugen
  CoCreateGuid udtGUID
  'nach ANSI konvertieren
  nSize = StringFromGUID2(udtGUID, strGUID, Len(strGUID))
  strGUID = Left$(StrConv(strGUID, vbFromUnicode), nSize - 1)
  'geschweifte Klammern entfernen
  GetNewGUID = Mid$(strGUID, 2, Len(strGUID) - 2)

  Exit Function

errHandler:
  Err.Raise Err.number, "clsSysInfo:GetNewGUID"
End Function

