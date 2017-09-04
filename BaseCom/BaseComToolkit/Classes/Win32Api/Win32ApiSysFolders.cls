VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'© 2000 Klaus Prinz EDV-Dienstleistungen
'       Diese Klasse dient als Schnittstelle zu allgemeinen Registry-Einträgen
'==================================================================================================
'Methoden:
'--------------------------------------------------------------------------------------------------
'   GetRegistryValue:   Private
'       Rückgabe eines Registry-Wertes (Private)
'==================================================================================================
'Eigenschaften:
'--------------------------------------------------------------------------------------------------
'   DSNDirectory:       Read
'       Verzeichnis der ODBC-DSN-Dateien
'   LoginName:          Read
'       Anmeldename des Anwenders
'==================================================================================================

Private Const MAX_PATH = 260

'Type zur Parametrierung der SHBrowseForFolder
Private Type ShortItemId
  cb As Long
  abID As Byte
End Type

Private Type ITEMIDLIST
  mkid As ShortItemId
End Type

Private Const CSIDL_DESKTOP = &H0           'virtuelles DeskTop-Verzeichnis
Private Const CSIDL_PROGRAMS = &H2          'Programmverzeichnis in Profilen
Private Const CSIDL_CONTROLS = &H3          'virtuelles Verzeichnis der Systemsteuerung
Private Const CSIDL_PRINTERS = &H4          'virtuelles Druckerverzeichnis
Private Const CSIDL_PERSONAL = &H5          'Eigene Dateien in Profilen
Private Const CSIDL_FAVORITES = &H6         'Favoritenverzeichnis in Profilen
Private Const CSIDL_STARTUP = &H7           'Startverzeichnis in Profilen
Private Const CSIDL_RECENT = &H8            'RecentFiles-Verzeichnis
Private Const CSIDL_SENDTO = &H9            'SendTo-Verzeichnis
Private Const CSIDL_BITBUCKET = &HA         'virtuelles Verzeichnis des Papierkorbs
Private Const CSIDL_STARTMENU = &HB         'Startmenüverzeichnis in Profilen
Private Const CSIDL_DESKTOPDIRECTORY = &H10 'virtuelles Desktop-Verzeichnis
Private Const CSIDL_DRIVES = &H11           'virtuelles Geräteverzeichnis
Private Const CSIDL_NETWORK = &H12          'virtuelles Netzwerkverzeichnis
'    ssfNETHOOD  = 0x13'
'    ssfFONTS    = 0x14,
Private Const CSIDL_TEMPLATES = &H15        'Vorlagenverzeichnis
'    ssfCOMMONSTARTMENU  = 0x16,
'    ssfCOMMONPROGRAMS   = 0x17,
'    ssfCOMMONSTARTUP    = 0x18,
'    ssfCOMMONDESKTOPDIR = 0x19,
'    ssfAPPDATA  = 0x1a,
'    ssfPRINTHOOD    = 0x1b,
'    ssfLOCALAPPDATA = 0x1c,
'    ssfALTSTARTUP   = 0x1d,
'    ssfCOMMONALTSTARTUP = 0x1e,
'    ssfCOMMONFAVORITES  = 0x1f,
'    ssfINTERNETCACHE    = 0x20,
'    ssfCOOKIES  = 0x21,
'    ssfHISTORY  = 0x22,
'    ssfCOMMONAPPDATA    = 0x23,
'    ssfWINDOWS  = 0x24,
'    ssfSYSTEM   = 0x25,
'    ssfPROGRAMFILES = 0x26,
'    ssfMYPICTURES   = 0x27,
'    ssfPROFILE  = 0x28

'APIs
Private Declare Function SHGetSpecialFolderLocation _
                Lib "shell32.dll" (ByVal hwndOwner As Long, _
                                   ByVal nFolder As Long, _
                                   pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList _
                Lib "shell32" (ByVal pidList As Long, _
                               ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectory _
                Lib "kernel32" _
                Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                             ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory _
                Lib "kernel32" _
                Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
                                              ByVal nSize As Long) As Long

'Enumeration für DriveState
Public Enum enumDriveState
  enumDriveIsOK = 1
  enumDriveDoesNotExist = 2
  enumDriveAintReady = 3
  enumDriveHasUnexpectedError = 4
End Enum

'Enumeration für DriveType
Public Enum enumDriveTypeFolders
  enumDriveNotIdentified = 0
  enumDriveNotExisting = 1
  enumDriveRemoveable = 2
  enumDriveFixed = 3
  enumDriveRemote = 4
  enumDriveCDRom = 5
  enumDriveRamDisk = 6
End Enum
'API
Private Declare Function GetDriveType _
                Lib "kernel32" _
                Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Function DriveState(ByVal strDriveLetter As String) As enumDriveState
  '====================================================================================
  '1999-12-30, Klaus Prinz
  '   Rückgabe Laufwerkszustands
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strDriveLetter:     Laufwerksbuchstabe
  '====================================================================================
  Dim lngDriveType As enumDriveTypeFolders   'DriveType aus API
  On Error GoTo errHandler
  'API fragen, welchen Eindruck es von dem Laufwerk hat
  lngDriveType = GetDriveType(strDriveLetter)

  'Ergebnis auswerten
  Select Case lngDriveType

    Case enumDriveNotIdentified, enumDriveNotExisting
      'Laufwerk nicht vorhanden
      DriveState = enumDriveDoesNotExist

    Case Else
      'Laufwerk zwar vorhanden, aber möglicherweise nicht
      'bereit, -> Testzugriff und Fehler auswerten
      On Error Resume Next
      Dir strDriveLetter

      If Err.number = 0 Then
        DriveState = enumDriveIsOK
      Else

        'bei diesen beiden ist das Laufwerk nicht bereit
        If Err.number = 52 Or Err.number = 71 Then
          DriveState = enumDriveAintReady
        Else
          'unwahrscheinlich, aber man weiß ja nie ...
          DriveState = enumDriveHasUnexpectedError
        End If

        Err.Clear
      End If

  End Select

  Exit Function

errHandler:
  Err.Raise Err.number, "modUtil:DriveState"
End Function

Public Function DriveType(ByVal strDriveLetter As String) As enumDriveTypeFolders
  '====================================================================================
  '1999-12-30, Klaus Prinz
  '   Rückgabe Laufwerkstyps
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strDriveLetter:     Laufwerksbuchstabe
  '====================================================================================
  On Error GoTo errHandler
  DriveType = GetDriveType(strDriveLetter)
  Exit Function

errHandler:
  Err.Raise Err.number, "modUtil:DriveType"
End Function

Private Function GetSpecialFolder(ByVal lngCSIDL As Long) As String
  '====================================================================================
  '2000-02-24, Klaus Prinz
  '   Systemverzeichnis ermitteln
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngCSIDL:   Konstante für Systempfad
  '====================================================================================
  Dim lngResult As Long   'Rückgabe der Funktionen
  Dim strPath   As String   'gesuchter Pfad
  Dim IDL       As ITEMIDLIST   'Item-ID-List
  On Error GoTo errHandler
  lngResult = SHGetSpecialFolderLocation(0, lngCSIDL, IDL)

  If lngResult = 0 Then
    strPath = Space(260)
    lngResult = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath)

    If lngResult <> 0 Then
      GetSpecialFolder = Left$(strPath, InStr(strPath, Chr(0)) - 1) & "\"
    End If
  End If

  Exit Function

errHandler:
  Err.Raise Err.number + vbObjectError, "clsRegistry:GetSpecialFolder"
End Function

Public Property Get PersonalFolder() As String
  '====================================================================================
  '2000-02-24, Klaus Prinz
  '   Rückgabe des PersonalFolders
  '====================================================================================
  PersonalFolder = GetSpecialFolder(CSIDL_PERSONAL)
End Property

Public Property Get SystemFolder() As String
  '====================================================================================
  '2000-08-14, Klaus Prinz
  '   Rückgabe des Systemverzeichnisses
  '====================================================================================
  Dim strFolder As String         'Rückgabe der Funktion
  Dim lngReturn As Long           'Nettolänge der Rückgabe
  On Error GoTo errHandler
  strFolder = Space(MAX_PATH)
  lngReturn = GetSystemDirectory(strFolder, MAX_PATH)

  If lngReturn <> 0 Then
    SystemFolder = Left(strFolder, lngReturn) & "\"
  End If

  Exit Function

errHandler:
  Err.Raise Err.LastDllError, "clsRegistry:Property Get SystemFolder"
End Property

Public Property Get WindowsFolder() As String
  '====================================================================================
  '2000-08-14, Klaus Prinz
  '   Rückgabe des Windows-Verzeichnisses
  '====================================================================================
  Dim strFolder As String         'Rückgabe der Funktion
  Dim lngReturn As Long           'Nettolänge der Rückgabe
  On Error GoTo errHandler
  strFolder = Space(MAX_PATH)
  lngReturn = GetWindowsDirectory(strFolder, MAX_PATH)

  If lngReturn <> 0 Then
    WindowsFolder = Left(strFolder, lngReturn) & "\"
  End If

  Exit Property

errHandler:
  Err.Raise Err.LastDllError, "clsRegistry:Property Get SystemFolder"
End Property
