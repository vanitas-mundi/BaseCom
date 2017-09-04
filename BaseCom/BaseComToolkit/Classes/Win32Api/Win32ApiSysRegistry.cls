VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private strAppName  As String    'Programmname für Settings
Private strSection  As String    'aktuelle Section der Settings
Private lngDataType As Long     'aktueller Datentyp
'KEY-Konstanten
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Public Enum enumHKEY
  enumHKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
  enumHKEY_CURRENT_USER = HKEY_CURRENT_USER
  enumHKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
  enumHKEY_USERS = HKEY_USERS
  enumHKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
  enumHKEY_DYN_DATA = HKEY_DYN_DATA
End Enum

'Error-Konstanten
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_FILE_NOT_FOUND = 2
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
'Datentyp-Konstanten
Private Const REG_SZ = 1        'Unicode String
Private Const REG_EXPAND_SZ = 2 'Unicode Null-terminated String
Private Const REG_BINARY = 3    'binäre Daten
Private Const REG_DWORD = 4     'C-Long
Private Const REG_MULTI_SZ = 7  'Multiple Unicode Strings

Public Enum enumDataType

  enumREG_SZ = 1
  enumREG_EXPAND_SZ = 2
  enumREG_BINARY = 3
  enumREG_DWORD = 4
  enumREG_MULTI_SZ = 7

End Enum

'FileTime-Struktur für RegEnumKeyEx
Private Type FILETIME

  dwLowDateTime As Long
  dwHighDateTime As Long

End Type

'KEY_READ-Komponenten
Private Const STANDARD_RightS_ALL = &H1F0000
Private Const STANDARD_RightS_READ = &H20000
Private Const KEY_QUERY_VALUE = 1
Private Const KEY_ENUMERATE_SUB_KEYS = 8
Private Const KEY_NOTIFY = &H10&
Private Const SYNCHRONIZE = &H100000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = 4
Private Const KEY_SET_VALUE = 2
Private Const KEY_READ = (STANDARD_RightS_READ Or KEY_QUERY_VALUE)
Private Const KEY_WRITE = ((STANDARD_RightS_ALL Or KEY_SET_VALUE) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RightS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'APIs
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          ByRef lpType As Long, _
                                          lpData As Any, _
                                          ByRef lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx _
                Lib "advapi32.dll" _
                Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                         ByVal lpSubKey As String, _
                                         ByVal Reserved As Long, _
                                         ByVal lpClass As String, _
                                         ByVal dwOptions As Long, _
                                         ByVal samDesired As Long, _
                                         ByVal lpSecurityAttributes As Long, _
                                         phkResult As Long, _
                                         ByVal lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal Reserved As Long, _
                                        ByVal dwType As Long, _
                                        lpValue As Any, _
                                        ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey _
                Lib "advapi32.dll" _
                Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue _
                Lib "advapi32.dll" _
                Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                         ByVal lpValueName As String) As Long
Private Declare Function RegQueryInfoKey _
                Lib "advapi32.dll" _
                Alias "RegQueryInfoKeyA" (ByVal hKey As Long, _
                                          ByVal lpClass As Long, _
                                          lpcbClass As Long, _
                                          ByVal lpReserved As Long, _
                                          lpcSubKeys As Long, _
                                          lpcbMaxSubKeyLen As Long, _
                                          lpcbMaxClassLen As Long, _
                                          lpcValues As Long, _
                                          lpcbMaxValueNameLen As Long, _
                                          lpcbMaxValueLen As Long, _
                                          lpcbSecurityDescriptor As Long, _
                                          lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumKeyEx _
                Lib "advapi32.dll" _
                Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpName As String, _
                                       lpcbName As Long, _
                                       ByVal lpReserved As Long, _
                                       ByVal lpClass As Long, _
                                       lpcbClass As Long, _
                                       lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue _
                Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       ByVal lpData As String, _
                                       lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings _
                Lib "kernel32" _
                Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, _
                                                   ByVal lpDst As String, _
                                                   ByVal nSize As Long) As Long
Private Declare Function ExtractIcon _
                Lib "shell32.dll" _
                Alias "ExtractIconA" (ByVal hInst As Long, _
                                      ByVal lpszExeFileName As String, _
                                      ByVal nIconIndex As Long) As Long
Private Declare Function FormatMessage _
                Lib "kernel32" _
                Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                        ByVal lpSource As Long, _
                                        ByVal dwMessageId As Long, _
                                        ByVal dwLanguageId As Long, _
                                        ByVal lpBuffer As String, _
                                        nSize As Long, _
                                        Arguments As Long) As Long
Private Declare Function StrLen _
                Lib "kernel32" _
                Alias "lstrlenA" (ByVal Ptr As String) As Long

Public Function HasKey(ByVal lngHKEY As enumHKEY, _
                       ByVal strKey As String, _
                       Optional ByVal strValue As Variant) As Boolean
  '====================================================================================
  '2000-08-22, Klaus Prinz
  '   Prüfen, ob der angeforderte Schlüssel und evt. auch der Wert existiert
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngHKey:    KEY
  '   strKey:     Key
  '   strValue:   Name des Wertes
  '====================================================================================
  Dim strData As String           'Rückgabe bei String- und Binary-Werten
  Dim nSize   As Long               'Länge der Rückgabe
  Dim hResult As Long             'Rückgabe der API
  Dim hKey    As Long                'Key-Handle
  Dim lngType As Long             'Datentyp

  On Error GoTo errHandler

  'Schlüssel öffnen
  hResult = RegOpenKeyEx(lngHKEY, strKey, 0&, KEY_READ, hKey)

  If hResult = 0 Then
    If Not IsMissing(strValue) Then
      hResult = RegQueryValueEx(hKey, strValue, 0&, lngType, ByVal strData, nSize)

      If hResult = 0 Then
        HasKey = True
      Else
        HasKey = False
      End If

    Else
      HasKey = True
    End If

  Else
    HasKey = False
  End If

  RegCloseKey hKey

  Exit Function

errHandler:
  RegCloseKey hKey
  Err.Raise Err.number, "cSysRegistry:HasKey"
End Function

Public Function GetRegistryValue(ByVal lngHKEY As enumHKEY, _
                                 ByVal strKey As String, _
                                 ByVal strValue As String) As Variant
  '====================================================================================
  '2000-02-24, Klaus Prinz
  '   Rückgabe eines Registry-Wertes
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngHKey:    KEY
  '   strKey:     Key
  '   strValue:   Name des Wertes
  '====================================================================================
  Dim strData As String           'Rückgabe bei REG_SZ und REG_EXPAND_SZ
  Dim lngData As Long             'Rückgabe bei REG_DWORD
  Dim nSize   As Long               'Länge der Rückgabe
  Dim hResult As Long             'Rückgabe der API
  Dim hKey    As Long                'Key-Handle
  Dim lngType As Long             'Datentyp

  On Error GoTo errHandler

  'Schlüssel öffnen
  hResult = RegOpenKeyEx(lngHKEY, strKey, 0&, KEY_READ, hKey)

  If hResult = ERROR_FILE_NOT_FOUND Then
    GetRegistryValue = Empty

    Exit Function

  ElseIf hResult <> ERROR_SUCCESS Then
    Err.Raise hResult, , GetErrorMessage(hResult)
  End If

  'Datentyp ermitteln (mit leerem strData und nSize = 0)
  hResult = RegQueryValueEx(hKey, strValue, 0&, lngType, ByVal strData, nSize)

  If hResult = 2 Then
    GetRegistryValue = Empty

    Exit Function

  ElseIf hResult <> ERROR_SUCCESS Then
    Err.Raise hResult, , GetErrorMessage(hResult)
  End If

  'Rückgabe ermitteln
  Select Case lngType

    Case REG_SZ         'Text
      strData = Space$(nSize)
      hResult = RegQueryValueEx(hKey, strValue, 0&, lngType, ByVal strData, nSize)
      GetRegistryValue = Left$(strData, nSize - 1)

    Case REG_EXPAND_SZ  'potientiell zu expandierender Text
      strData = Space$(nSize)
      hResult = RegQueryValueEx(hKey, strValue, 0&, lngType, ByVal strData, nSize)
      GetRegistryValue = ExpandString(Left$(strData, nSize - 1))

    Case REG_MULTI_SZ   'NULL-getrennte Strings
      strData = Space$(nSize)
      hResult = RegQueryValueEx(hKey, strValue, 0&, lngType, ByVal strData, nSize)
      GetRegistryValue = Split(Left$(strData, nSize - 2), Chr$(0))

    Case REG_BINARY     'binäre Daten
      strData = Space$(nSize)
      hResult = RegQueryValueEx(hKey, strValue, 0&, lngType, ByVal strData, nSize)
      GetRegistryValue = strData

    Case REG_DWORD      'Long
      hResult = RegQueryValueEx(hKey, strValue, 0&, lngType, lngData, 4)
      GetRegistryValue = lngData
  End Select

  If hResult <> ERROR_SUCCESS Then
    Err.Raise hResult, , GetErrorMessage(hResult)
  End If

  'Schlüssel schließen
  RegCloseKey hKey

  Exit Function

errHandler:
  RegCloseKey hKey

  If InStr(1, Err.source, ":") = "" Then Err.source = "cSysRegistry:GetRegistryValue"
  Err.Raise Err.number
End Function

Private Function ExpandString(ByVal strPath As String) As String
  '====================================================================================
  '2000-04-22, Klaus Prinz
  '   Übergebenen String expandieren (%SystemRoot% wird beispielsweise zu "c:\Temp\WINNT")
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strPath:    zu expandierender String
  '====================================================================================
  Dim nSize       As Long               'Länge der Unicode-Ausführung von strExpanded
  Dim strExpanded As String       'expandierter Pfad

  On Error GoTo errHandler

  If InStr(1, strPath, "%") = 0 Then
    ExpandString = strPath
  Else
    strExpanded = " "
    strPath = strPath & Chr$(0)
    nSize = ExpandEnvironmentStrings(strPath, strExpanded, 0)
    strExpanded = Space$(nSize)
    nSize = ExpandEnvironmentStrings(strPath, strExpanded, nSize)
    ExpandString = Left$(strExpanded, StrLen(strExpanded))
  End If

  Exit Function

errHandler:
  Err.Raise Err.number, "cSysRegistry:ExpandString"
End Function

'Schreiben
Public Sub SetRegistryValue(ByVal lngHKEY As enumHKEY, _
                            ByVal strKey As String, _
                            ByVal strValueName As String, _
                            lngType As enumDataType, _
                            ByVal varValue As Variant)
  '====================================================================================
  '2000-08-23, Klaus Prinz
  '   übergebenen Wert in Registry schreiben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngHKey:    KEY
  '   strKey:     Key
  '   strValue:   Name des Wertes
  '   lngType:    Datentyp
  '   varValue:   zu schreibender Wert
  '====================================================================================
  Dim lngResult As Long   'Rückgabe aus Funktion
  Dim hKey      As Long        'Handle des zu öffnenden Keys

  On Error GoTo errHandler

  'Schlüssel öffnen
  lngResult = RegOpenKeyEx(lngHKEY, strKey, 0, KEY_WRITE, hKey)

  If lngResult = ERROR_FILE_NOT_FOUND Then
    'Schlüssel erzeugen
    lngResult = RegCreateKeyEx(lngHKEY, strKey, 0&, vbNullString, 0&, KEY_WRITE, 0&, hKey, 0&)

    If lngResult <> ERROR_SUCCESS Then
      Err.Raise lngResult, , GetErrorMessage(lngResult)
    End If

  ElseIf lngResult <> ERROR_SUCCESS Then
    Err.Raise lngResult, , GetErrorMessage(lngResult)
  End If

  'Wert je nach Datentyp schreiben
  Select Case lngType

    Case REG_SZ, REG_EXPAND_SZ
      varValue = varValue & Chr$(0)
      lngResult = RegSetValueEx(hKey, strValueName, 0&, lngType, ByVal CStr(varValue), Len(varValue))

    Case REG_MULTI_SZ
      varValue = Join(varValue, Chr$(0)) & Chr$(0) & Chr$(0)
      lngResult = RegSetValueEx(hKey, strValueName, 0&, lngType, ByVal CStr(varValue), Len(varValue))

    Case REG_BINARY
      lngResult = RegSetValueEx(hKey, strValueName, 0&, lngType, ByVal CStr(varValue), Len(varValue))

    Case REG_DWORD
      lngResult = RegSetValueEx(hKey, strValueName, 0&, lngType, CLng(varValue), 4)
  End Select

  If lngResult <> ERROR_SUCCESS Then
    Err.Raise lngResult, , GetErrorMessage(lngResult)
  End If

  'Key schließen
  RegCloseKey hKey

  Exit Sub

errHandler:
  RegCloseKey hKey

  If InStr(1, Err.source, ":") = "" Then Err.source = "cSysRegistry:SetRegistryValue"
  Err.Raise Err.number
End Sub

Public Sub DeleteRegistryValue(ByVal lngHKEY As enumHKEY, _
                               ByVal strKey As String, _
                               ByVal strValueName As String)
  '====================================================================================
  '2000-08-23, Klaus Prinz
  '   übergebenen Wert in Registry schreiben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngHKey:        KEY
  '   strKey:         Key
  '   strValueName:   Name des Wertes
  '====================================================================================
  Dim lngResult As Long   'Rückgabe aus Funktion
  Dim hKey      As Long        'Handle des zu öffnenden Keys

  On Error GoTo errHandler

  'Schlüssel öffnen
  lngResult = RegOpenKeyEx(lngHKEY, strKey, 0, KEY_WRITE, hKey)

  If lngResult = ERROR_SUCCESS Then
    lngResult = RegDeleteValue(hKey, strValueName)

    If lngResult <> ERROR_SUCCESS And lngResult <> 2 Then
      Err.Raise lngResult, , GetErrorMessage(lngResult)
    End If

    RegCloseKey hKey
  End If

  Exit Sub

errHandler:
  RegCloseKey hKey

  If InStr(1, Err.source, ":") = "" Then Err.source = "cSysRegistry:DeleteRegistryValue"
  Err.Raise Err.number
End Sub

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
  lngResult = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lngMessageID, 0&, strError, nSize, 0&)

  If lngResult = 0 Then
    If Err.number = 0 Then
      Err.Raise 9999
    Else
      GetErrorMessage = Err.description & vbCrLf & Err.LastDllError
    End If

  Else
    GetErrorMessage = Left$(strError, lngResult - 1)
  End If

  Exit Function

errHandler:
  Err.Raise 9999, , "Der Fehlertext zu ReturnCode " & lngResult & " konnte nicht ermittelt werden."
End Function

'**************************************************************************************************
'Enum-Routinen
'**************************************************************************************************
Public Function EnumKeys(ByVal lngHKEY As enumHKEY, ByVal strKey As String) As Variant
  '====================================================================================
  '2000-08-29, Klaus Prinz
  '   SubKeys von strKey zurückgeben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngHKey:        KEY
  '   strKey:         Key
  '====================================================================================
  Dim lngResult  As Long   'Rückgabe aus Funktion
  Dim hKey       As Long        'Handle des zu öffnenden Keys
  Dim nSubKeys   As Long    'Anzahl der SubKeys
  Dim iSubKey    As Long     'Zeiger in nSubKeys
  Dim strSubKey  As String 'jeweiliger SubKey
  Dim nSize      As Long       'Länge von strKey
  Dim lngMaxSize As Long  'größte Länge der SubKeys
  Dim udtTime    As FILETIME 'FILETIME-Struktur
  Dim lngDummy   As Long    'Dummy für alle nicht benötigten Rückgaben
  Dim strKeys()  As String 'temporäres Array

  On Error GoTo errHandler

  'Schlüssel öffnen
  lngResult = RegOpenKeyEx(lngHKEY, strKey, 0, KEY_ALL_ACCESS, hKey)

  If lngResult = ERROR_SUCCESS Then
    'Größen ermitteln
    lngResult = RegQueryInfoKey(hKey, 0&, 0&, 0&, nSubKeys, lngMaxSize, lngDummy, lngDummy, lngDummy, lngDummy, lngDummy, udtTime)

    If lngResult <> ERROR_SUCCESS Then
      Err.Raise lngResult, , GetErrorMessage(lngResult)
    End If

    'Abbruch, wenn keine SubKeys vorhanden
    If nSubKeys = 0 Then
      EnumKeys = Empty
      RegCloseKey hKey

      Exit Function

    End If

    ReDim strKeys(nSubKeys - 1)

    'SubKeys ermitteln
    For iSubKey = 0 To nSubKeys - 1
      'Variablen vorbelegen
      strSubKey = Space$(lngMaxSize + 1)
      nSize = lngMaxSize + 1
      'Funktionsaufruf
      lngResult = RegEnumKeyEx(hKey, iSubKey, strSubKey, nSize, 0&, 0&, 0&, udtTime)

      If lngResult <> ERROR_SUCCESS Then
        Err.Raise lngResult, , GetErrorMessage(lngResult)
      End If

      'SubKey in Array schreiben
      strKeys(iSubKey) = Left$(strSubKey, nSize)
    Next

    EnumKeys = strKeys
  Else
    EnumKeys = Empty
  End If

  'Schlüssel schließen
  RegCloseKey hKey

  Exit Function

errHandler:
  RegCloseKey hKey

  If InStr(1, Err.source, ":") = "" Then Err.source = "cSysRegistry:EnumKeys"
  Err.Raise Err.number
End Function

Public Function EnumValues(ByVal lngHKEY As enumHKEY, ByVal strKey As String) As Variant
  '====================================================================================
  '2000-08-29, Klaus Prinz
  '   Values von strKey zurückgeben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngHKey:        KEY
  '   strKey:         Key
  '====================================================================================
  Dim lngResult       As Long           'Rückgabe aus Funktion
  Dim hKey            As Long                'Handle des zu öffnenden Keys
  Dim nValues         As Long             'Anzahl der Values
  Dim iValue          As Long              'Zeiger in nValues
  Dim lngMaxNameSize  As Long      'größte Länge der ValueNames
  Dim strValueName    As String      'jeweiliger ValueName
  Dim nValueNameSize  As Long      'Länge von strValueName
  Dim lngMaxValueSize As Long     'größte Länge der Values
  Dim strValue        As String          'jeweiliger Value
  Dim nValueSize      As Long          'Länge von strValue
  Dim lngType         As Long             'Registry-Datentyp
  Dim udtTime         As FILETIME         'FILETIME-Struktur
  Dim lngDummy        As Long            'Dummy für alle nicht benötigten Rückgaben
  Dim strValueNames() As String   'temporäres Array

  On Error GoTo errHandler

  'Schlüssel öffnen
  lngResult = RegOpenKeyEx(lngHKEY, strKey, 0, KEY_ALL_ACCESS, hKey)

  If lngResult = ERROR_SUCCESS Then
    'Größen ermitteln
    lngResult = RegQueryInfoKey(hKey, 0&, 0&, 0&, lngDummy, lngDummy, lngDummy, nValues, lngMaxNameSize, lngMaxValueSize, lngDummy, udtTime)

    If lngResult <> ERROR_SUCCESS Then
      Err.Raise lngResult, , GetErrorMessage(lngResult)
    End If

    'Abbruch, wenn keine Werte vorhanden
    If nValues = 0 Then
      EnumValues = Empty
      RegCloseKey hKey

      Exit Function

    End If

    ReDim strValueNames(1, nValues - 1)

    'Werte ermitteln
    For iValue = 0 To nValues - 1
      'Variablen vorbelegen
      strValueName = Space$(lngMaxNameSize + 1)
      nValueNameSize = lngMaxNameSize + 1
      strValue = Space$(lngMaxValueSize + 1)
      nValueSize = lngMaxValueSize + 1
      'Funktionsaufruf
      lngResult = RegEnumValue(hKey, iValue, strValueName, nValueNameSize, 0&, lngType, strValue, nValueSize)

      If lngResult <> ERROR_SUCCESS Then
        Err.Raise lngResult, , GetErrorMessage(lngResult)
      End If

      'Werte in Array schreiben
      strValueNames(0, iValue) = Left$(strValueName, nValueNameSize)
      strValueNames(1, iValue) = lngType
    Next

    EnumValues = strValueNames
  Else
    EnumValues = Empty
  End If

  'Schlüssel schließen
  RegCloseKey hKey

  Exit Function

errHandler:
  RegCloseKey hKey

  If InStr(1, Err.source, ":") = "" Then Err.source = "cSysRegistry:EnumValues"
  Err.Raise Err.number
End Function

'**************************************************************************************************
'ExtensionTo-Routinen
'**************************************************************************************************
Public Function ExtensionToContentType(ByVal strExtension As String) As String

  '====================================================================================
  '2000-08-21, Klaus Prinz
  '   MIME-Content-Type ermitteln
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strExtension:   Dateierweiterung
  '====================================================================================
  On Error GoTo errHandler

  If Me.HasKey(enumHKEY_CLASSES_ROOT, "." & strExtension, "Content Type") Then
    ExtensionToContentType = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, "." & strExtension, "Content Type")
  Else
    ExtensionToContentType = "application/octet-stream"
  End If

  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "cSysRegistry:ExtensionToContentType"
  Err.Raise Err.number
End Function

Public Function ExtensionToFileType(ByVal strExtension As String) As Variant
  '====================================================================================
  '2000-04-20, Klaus Prinz
  '   Dateityp der übergebenen Dateierweiterung zurückgeben und Empty, wenn er nicht
  '   existiert.
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strExtension:   Dateierweiterung
  '====================================================================================
  Dim strClass As Variant         'Klassenname

  On Error GoTo errHandler

  strClass = GetRegistryValue(HKEY_CLASSES_ROOT, "." & strExtension, "")

  If IsEmpty(strClass) Then
    ExtensionToFileType = Empty

    Exit Function

  End If

  ExtensionToFileType = GetRegistryValue(HKEY_CLASSES_ROOT, strClass, "")

  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "cSysRegistry:ExtensionToFileType"
  Err.Raise Err.number
End Function

Public Function FileToIconHandle(ByVal strFile As String, _
                                 Optional ByVal iIcon As Long = 0) As Long
  '====================================================================================
  '2000-08-21, Klaus Prinz
  '   Icon-Handle des übergebenen Dateityps zurückgeben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strFile:    Dateierweiterung
  '   iIcon:      Icon-Index
  '====================================================================================
  Dim strExtension As String      'Dateierweiterung
  Dim strClass     As String          'Klassenname
  Dim strCLSID     As Variant         'CLSID
  Dim strServer    As String         'Server-Applikation

  On Error GoTo errHandler

  strExtension = Right$(strFile, Len(strFile) - InStrRev(strFile, "."))
  strClass = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, "." & strExtension, "")

  If Me.HasKey(enumHKEY_CLASSES_ROOT, strClass & "\DefaultIcon") Then
    'Default-Icon bereits in Dateierweiterungseintrag
    strServer = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, strClass & "\DefaultIcon", "")

    If InStr(1, strServer, ",") > 0 Then
      'Server und Default-Icon angegeben
      iIcon = CLng(Right$(strServer, Len(strServer) - InStr(1, strServer, ",")))
      strServer = Left$(strServer, InStr(1, strServer, ",") - 1)
    Else
      'nur Default-Icon angegeben
      strCLSID = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, strClass & "\CLSID", "")

      If IsEmpty(strCLSID) Then

        'keine CLSID -> vermutlich eine Exe
        If IsNumeric(strServer) Then
          iIcon = CLng(strServer)
        Else
          iIcon = 0
        End If

        strServer = strFile
      Else

        'CLSID vorhanden
        If Me.HasKey(enumHKEY_CLASSES_ROOT, "CLSID\" & strCLSID & "\DefaultIcon") Then
          strServer = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, "CLSID\" & strCLSID & "\DefaultIcon", "")
        Else
          strServer = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, "CLSID\" & strCLSID & "\InProcServer32", "")
        End If

        If InStr(1, strServer, ",") > 0 Then
          iIcon = CLng(Right$(strServer, Len(strServer) - InStr(1, strServer, ",")))
          strServer = Left$(strServer, InStr(1, strServer, ",") - 1)
        End If
      End If
    End If

    FileToIconHandle = ExtractIcon(app.hInstance, strServer, iIcon)
  Else

    'Default-Icon in CLSID-Sektion suchen
    If Me.HasKey(enumHKEY_CLASSES_ROOT, strClass & "\CLSID") Then
      'CLSID-Sektion vorhanden
      strCLSID = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, strClass & "\CLSID", "")

      If Me.HasKey(enumHKEY_CLASSES_ROOT, "CLSID\" & strCLSID & "\LocalServer32") Then
        strServer = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, "CLSID\" & strCLSID & "\LocalServer32", "")
      ElseIf Me.HasKey(enumHKEY_CLASSES_ROOT, "CLSID\" & strCLSID & "\InProcServer32") Then
        strServer = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, "CLSID\" & strCLSID & "\InProcServer32", "")
      End If

      'strServer = ExpandString$(strServer)
      FileToIconHandle = ExtractIcon(app.hInstance, strServer, iIcon)
    Else
      'keine CLSID-Sektion vorhanden
      FileToIconHandle = 0
    End If
  End If

  Exit Function

errHandler:

  If Err.number = 2 Then
    Err.description = strExtension & " ist ein unbekannter Dateityp."
  End If

  If InStr(1, Err.source, ":") = 0 Then Err.source = "cSysRegistry:ExtensionToIconHandle"
  Err.Raise Err.number
End Function

Public Property Get WebBrowserExe() As String
  '====================================================================================
  '2000-10-11, Klaus Prinz
  '   Rückgabe des vollständigen WebBrowser-Pfades
  '====================================================================================
  Dim strClass   As String          'Klassennamen hinter .htm
  Dim strBrowser As String

  On Error GoTo errHandler

  strClass = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, ".htm", "")
  strBrowser = Me.GetRegistryValue(enumHKEY_CLASSES_ROOT, strClass & "\shell\open\command", "")

  If Left$(strBrowser, 1) = Chr$(34) Then strBrowser = Right$(strBrowser, Len(strBrowser) - 1)
  strBrowser = Left$(strBrowser, InStr(1, UCase$(strBrowser), ".EXE") + 3)

  If LenB(Dir$(strBrowser)) = 0 Then
    Err.Raise 9999, , "No Browser found"
  Else
    WebBrowserExe = strBrowser
  End If

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = "" Then Err.source = "cSysRegistry:WebBrowserExe"
  Err.Raise Err.number
End Property

'**************************************************************************************************
'Settings-Routinen
'**************************************************************************************************
Public Property Get AppName() As String
  AppName = strAppName
End Property

Public Property Let AppName(ByVal NewAppName As String)
  strAppName = NewAppName
End Property

Public Property Let Section(ByVal NewSection As String)
  strSection = NewSection
End Property

Public Property Get Section() As String
  Section = strSection
End Property

Public Property Let DataType(ByVal NewDataType As enumDataType)
  lngDataType = NewDataType
End Property

Public Sub SaveSetting(ByVal ValueName As String, ByVal value As Variant)
  '====================================================================================
  '2000-08-24, Klaus Prinz
  '   Key zusammenstellen und übergebenen Wert speichern
  '------------------------------------------------------------------------------------
  'Argumente:
  '   ValueName:  Name des Wertes
  '   Value:      Neuer Wert
  '====================================================================================
  Me.SetRegistryValue enumHKEY_CURRENT_USER, "Software\" & strAppName & "\" & strSection, ValueName, lngDataType, value
End Sub

Public Function GetSetting(ByVal ValueName As String, _
                           Optional ByVal DefaultValue As Variant) As Variant
  '====================================================================================
  '2000-08-24, Klaus Prinz
  '   Key zusammenstellen und übergebenen Wert speichern
  '------------------------------------------------------------------------------------
  'Argumente:
  '   ValueName:      Name des Wertes
  '   DefaultValue:   Default-Wert: wird zurückgegeben, wenn Eintrag nicht existiert
  '====================================================================================
  Dim varSetting As Variant       'Zwischenwert

  On Error GoTo errHandler

  varSetting = Me.GetRegistryValue(enumHKEY_CURRENT_USER, "Software\" & strAppName & "\" & strSection, ValueName)

  If IsEmpty(varSetting) And Not IsMissing(DefaultValue) Then
    varSetting = DefaultValue
  End If

  GetSetting = varSetting

  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = "" Then Err.source = "cSysRegistry:GetSetting"
  Err.Raise Err.number
End Function

Public Function DeleteSetting(ByVal ValueName As String)
  '====================================================================================
  '2000-08-24, Klaus Prinz
  '   Key zusammenstellen und DeleteRegistryValue-Methode aufrufen
  '------------------------------------------------------------------------------------
  'Argumente:
  '   ValueName:      Name des Wertes
  '====================================================================================
  Me.DeleteRegistryValue enumHKEY_CURRENT_USER, "Software\" & strAppName & "\" & strSection, ValueName
End Function

