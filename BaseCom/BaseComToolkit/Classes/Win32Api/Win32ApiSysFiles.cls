VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MAX_PATH = 260
    
Private Const SHGFI_ICON = &H100
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_TYPENAME = &H400

'Error-Konstanten
Private Const INVALID_HANDLE_VALUE = -1
Private Const ERROR_SHARING_VIOLATION = &H20
Private Const ERROR_FILE_NOT_FOUND = 2
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

'Access-Modi
Private Const DELETE = &H10000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

'Share-Modi
Private Const FILE_SHARE_READ = 1
Private Const FILE_SHARE_WRITE = 2
Private Const FILE_SHARE_DELETE = 4

'Creation-Modi
Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

'Datei-Attribute
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_ENCRYPTED = &H40
Private Const FILE_ATTRIBUTE_HIDDEN = 2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000
Private Const FILE_ATTRIBUTE_READONLY = 1
Private Const FILE_ATTRIBUTE_SYSTEM = 4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

'Struktur für SHGEtFileInfo
Private Type SHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

'APIs
Private Declare Function CreateFile _
                Lib "kernel32" _
                Alias "CreateFileA" (ByVal lpFileName As String, _
                                     ByVal dwDesiredAccess As Long, _
                                     ByVal dwShareMode As Long, _
                                     ByVal lpSecurityAttributes As Long, _
                                     ByVal dwCreationDisposition As Long, _
                                     ByVal dwFlagsAndAttributes As Long, _
                                     ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SHGetFileInfo _
                Lib "shell32.dll" _
                Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                        ByVal dwFileAttributes As Long, _
                                        psfi As SHFILEINFO, _
                                        ByVal cbFileInfo As Long, _
                                        ByVal uFlags As Long) As Long
    
'Fehlerbehandlung
Private Declare Function FormatMessage _
                Lib "kernel32" _
                Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                        ByVal lpSource As Long, _
                                        ByVal dwMessageId As Long, _
                                        ByVal dwLanguageId As Long, _
                                        ByVal lpBuffer As String, _
                                        nSize As Long, _
                                        Arguments As Long) As Long

Public Property Get FileInUse(ByVal strFile As String) As Boolean
  '====================================================================================
  '2000-09-18, Klaus Prinz
  '   Prüfen, ob übergebene Datei bereits geöffnet ist
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strFile:    zu prüfende Datei
  '====================================================================================
  Dim hFile    As Long               'Datei-Handle
  Dim lngError As Long            'Fehler-Code
  On Error GoTo errHandler
  hFile = CreateFile(strFile, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

  If hFile = INVALID_HANDLE_VALUE Then
    lngError = Err.LastDllError

    If lngError = ERROR_SHARING_VIOLATION Then
      FileInUse = True
    Else
      Err.Raise 10000 + lngError, , GetErrorMessage(lngError)
    End If

  Else
    CloseHandle hFile
  End If

  Exit Property

errHandler:

  If hFile = INVALID_HANDLE_VALUE Then CloseHandle hFile
  Err.Raise Err.number, "clsFiles:FileInUse"
End Property

Public Property Get IconHandle(ByVal strFile As String) As Long
  '====================================================================================
  '2000-09-18, Klaus Prinz
  '   Icon der übergebenen Datei zurückgeben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strFile:    zu prüfende Datei
  '====================================================================================
  Dim lngReturn     As Long           'Rückgabe der Funktion
  Dim udtSHFILEINFO As SHFILEINFO 'SHFILEINFO-Struktur
  On Error GoTo errHandler
  lngReturn = SHGetFileInfo(strFile, 0, udtSHFILEINFO, LenB(udtSHFILEINFO), SHGFI_ICON)
  IconHandle = udtSHFILEINFO.hIcon
  Exit Property

errHandler:
  Err.Raise Err.number, "clsFiles:IconHandle"
End Property

Public Property Get DisplayName(ByVal strFile As String) As String
  '====================================================================================
  '2000-09-18, Klaus Prinz
  '   Anzeigename (Dateiname) der übergebenen Datei ermitteln
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strFile:    zu prüfende Datei
  '====================================================================================
  Dim lngReturn     As Long           'Rückgabe der Funktion
  Dim udtSHFILEINFO As SHFILEINFO 'SHFILEINFO-Struktur
  On Error GoTo errHandler
  lngReturn = SHGetFileInfo(strFile, 0, udtSHFILEINFO, LenB(udtSHFILEINFO), SHGFI_DISPLAYNAME)
  DisplayName = Left(udtSHFILEINFO.szDisplayName, InStr(1, udtSHFILEINFO.szDisplayName, vbNullChar) - 1)
  Exit Property

errHandler:
  Err.Raise Err.number, "clsFiles:DisplayName"
End Property

Public Property Get TypeName(ByVal strFile As String) As String
  '====================================================================================
  '2000-09-18, Klaus Prinz
  '   Typenamen der übergebenen Datei ermitteln
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strFile:    zu prüfende Datei
  '====================================================================================
  Dim lngReturn     As Long           'Rückgabe der Funktion
  Dim udtSHFILEINFO As SHFILEINFO 'SHFILEINFO-Struktur
  On Error GoTo errHandler
  lngReturn = SHGetFileInfo(strFile, 0, udtSHFILEINFO, LenB(udtSHFILEINFO), SHGFI_TYPENAME)
  TypeName = Left(udtSHFILEINFO.szTypeName, InStr(1, udtSHFILEINFO.szTypeName, vbNullChar) - 1)
  Exit Property

errHandler:
  Err.Raise Err.number, "clsFiles:TypeName"
End Property

Private Function GetErrorMessage(ByVal lngMessageID As Long) As String
  '====================================================================================
  '2000-08-23, Klaus Prinz
  '   Fehlertext ermitteln
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngMessageID:  FehlerCode in System Message Table
  '====================================================================================
  Dim strError  As String * 256    'Fehlertext
  Dim lngResult As Long           'Rückgabe
  Dim nSize     As Long               'Länge von strError [IN]
  On Error GoTo errHandler
  nSize = 256
  lngResult = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngMessageID, 0, strError, nSize, 0)

  If lngResult = 0 Then
    Err.Raise 9999
  Else
    GetErrorMessage = Left(strError, lngResult - 1)
  End If

  Exit Function

errHandler:
  Err.Raise 9999, , "Der Fehlertext zu ReturnCode " & lngResult & " konnte nicht ermittelt werden."
End Function

