VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private bClassReady           As Boolean          'True, wenn Class_Initialize durchlaufen wurde
Private strDrives()           As String           'Laufwerksbezeichner, z.B. "C:\"
Private strDriveTypes(2 To 6) As String 'Laufwerkstypbezeichner

'Enumeration für DriveType
Public Enum enumDriveTypeDrives
  enumDriveTypeNotExisting = 1
  enumDriveTypeRemoveable = 2
  enumDriveTypeLocal = 3
  enumDriveTypeRemote = 4
  enumDriveTypeCDRom = 5
  enumDriveTypeRamDisk = 6
End Enum

'Konstanten für Laufwerkstyp
Private Const DRIVE_UNKNOWN = 0
Private Const DRIVE_NO_ROOT_DIR = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
    
'APIs
Private Declare Function GetModuleHandle _
                Lib "kernel32" _
                Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress _
                Lib "kernel32" (ByVal hModule As Long, _
                                ByVal lpProcName As String) As Long
Private Declare Function GetDiskFreeSpace _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
                                           lpSectorsPerCluster As Long, _
                                           lpBytesPerSector As Long, _
                                           lpNumberOfFreeClusters As Long, _
                                           lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, _
                                             lpFreeBytesAvailable As Any, _
                                             lpTotalNumberOfBytes As Any, _
                                             lpTotalNumberOfFreeBytes As Any) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetLogicalDriveStrings _
                Lib "kernel32" _
                Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                 ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType _
                Lib "kernel32" _
                Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Sub Class_Initialize()
  Call InitializeArrays
End Sub

Private Sub InitializeArrays()
  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Arrays initialisieren
  '====================================================================================
  Dim strBuffer As String         'String mit allen verfügbaren Laufwerksbezeichnern
  Dim nSize     As Long               'Bruttogröße der Laufwerksbezeichnern
  On Error GoTo errHandler
  'Drives-Array füllen
  strBuffer = Space(200)
  nSize = GetLogicalDriveStrings(200, strBuffer)
  strDrives = Split(Left(strBuffer, nSize - 1), vbNullChar)
  'DriveTypes-Array füllen
  strDriveTypes(2) = "FloppyDrive"
  strDriveTypes(3) = "LocalDrive"
  strDriveTypes(4) = "NetDrive"
  strDriveTypes(5) = "CDDrive"
  strDriveTypes(6) = "LocalDrive"
  bClassReady = True
  Exit Sub

errHandler:
  Err.Raise Err.number, "clsSysDrives:InitializeArrays"
End Sub

Public Property Get DrivesCount() As Long

  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Anzahl der Laufwerke
  '====================================================================================
  If Not bClassReady Then Call InitializeArrays
  DrivesCount = UBound(strDrives) + 1
End Property

Public Property Get DriveName(ByVal Index As Long) As String

  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Name des Laufwerks
  '====================================================================================
  If Not bClassReady Then Call InitializeArrays
  DriveName = strDrives(Index - 1)
End Property

Public Property Get DriveType(ByVal Index As Long) As enumDriveTypeDrives

  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Typ des Laufwerks als Enumeration
  '====================================================================================
  If Not bClassReady Then Call InitializeArrays
  DriveType = GetDriveType(strDrives(Index - 1))
End Property

Public Property Get DriveTypeName(ByVal Index As Long) As String

  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Typ des Laufwerks im Klartext, um z.B. ein Image einer ImageList anzusprechen
  '====================================================================================
  If Not bClassReady Then Call InitializeArrays
  DriveTypeName = strDriveTypes(GetDriveType(strDrives(Index - 1)))
End Property

Public Property Get DriveIndex(ByVal strDrive As String) As Long
  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Index des Laufwerks
  '====================================================================================
  Dim iDrive As Long              'Zeiger in strDrives()
  On Error GoTo errHandler

  If Not bClassReady Then Call InitializeArrays

  For iDrive = 0 To UBound(strDrives)

    If strDrive = strDrives(iDrive) Then
      DriveIndex = iDrive + 1
      Exit Property
    End If

  Next

  Exit Property
    
errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsSysDrives:DriveIndex"
  Err.Raise Err.number
End Property

Public Property Get DriveReady(ByVal Index As Long) As Boolean

  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Bereitschaft des Laufwerks
  '====================================================================================
  If Not bClassReady Then Call InitializeArrays
  On Error Resume Next
  Dir strDrives(Index - 1)
  DriveReady = (Err.number = 0)
End Property

Public Property Get CDDrivesCount() As Long
  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Anzahl der CD-Laufwerke
  '====================================================================================
  Dim iDrive  As Long              'Zeiger in strDrives()
  Dim nDrives As Long             'Zwischenvariable
  On Error GoTo errHandler

  If Not bClassReady Then Call InitializeArrays

  For iDrive = 0 To UBound(strDrives)

    If GetDriveType(strDrives(iDrive)) = DRIVE_CDROM Then
      nDrives = nDrives + 1
    End If

  Next

  CDDrivesCount = nDrives
  Exit Property
    
errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsSysDrives:CDDrivesCount"
  Err.Raise Err.number
End Property

Public Property Get CDDriveName(ByVal Index As Long) As String
  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Name des CD-Laufwerks mit dem übergebenen Index
  '====================================================================================
  Dim iDrive  As Long              'Zeiger in strDrives()
  Dim nDrives As Long             'Zwischenvariable
  On Error GoTo errHandler

  If Not bClassReady Then Call InitializeArrays
  If Index > Me.CDDrivesCount Then Err.Raise 9

  For iDrive = 0 To UBound(strDrives)

    If GetDriveType(strDrives(iDrive)) = DRIVE_CDROM Then
      nDrives = nDrives + 1

      If nDrives = Index Then
        CDDriveName = strDrives(iDrive)
        Exit Property
      End If
    End If

  Next

  Exit Property
    
errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsSysDrives:CDDriveName"
  Err.Raise Err.number
End Property

Public Property Get DriveFreeSpace(ByVal Index As Long) As Variant
  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Freier Platz
  '====================================================================================
  Dim lngResult          As Long           'Rückgabe aus Funktionen
  'Variablen für Win 95 vor SR 2 (GetDiskFreeSpace)
  Dim nSectorsPerCluster As Long  'Anzahl der Sektoren je Cluster
  Dim nBytesPerSector    As Long     'Anzahl der Bytes je Sektor
  Dim nFreeClusters      As Long       'Anzahl der freien Clusters
  Dim nTotalClusters     As Long      'Gesamtzahl der Cluster
  'Variablen für Rest (GetDiskFreeSpaceEx)
  Dim byteAvail(7)       As Byte        'Array zur Aufnahme der verfügbaren Bytes
  Dim byteFree(7)        As Byte         'Array zur Aufnahme der freien Kapazität
  Dim byteTotal(7)       As Byte        'Array zur Aufnahme der Größe des Datenträgers
  Dim iByte              As Long               'Zeiger in byteFree()
  On Error GoTo errHandler

  If Not bClassReady Then Call InitializeArrays
  'Prüfen, ob GetDiskFreeSpaceEx existiert
  lngResult = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA")

  If lngResult = 0 Then
    'Win 95 vor SR 2
    lngResult = GetDiskFreeSpace(strDrives(Index - 1), nSectorsPerCluster, nBytesPerSector, nFreeClusters, nTotalClusters)
    DriveFreeSpace = CDec(nFreeClusters * nSectorsPerCluster * nBytesPerSector)
  Else
    'der Rest
    lngResult = GetDiskFreeSpaceEx(strDrives(Index - 1), byteAvail(0), byteTotal(0), byteFree(0))

    For iByte = 0 To 7
      DriveFreeSpace = CDec(DriveFreeSpace + byteFree(iByte) * 2 ^ (8 * iByte))
    Next

  End If

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsSysDrives:DriveFreeSpace"
  Err.Raise Err.number
End Property

Public Property Get DriveTotalSpace(ByVal Index As Long) As Variant
  '====================================================================================
  '2000-09-14, Klaus Prinz
  '   Freier Platz
  '====================================================================================
  Dim lngResult          As Long           'Rückgabe aus Funktionen
  'Variablen für Win 95 vor SR 2 (GetDiskFreeSpace)
  Dim nSectorsPerCluster As Long  'Anzahl der Sektoren je Cluster
  Dim nBytesPerSector    As Long     'Anzahl der Bytes je Sektor
  Dim nFreeClusters      As Long       'Anzahl der freien Clusters
  Dim nTotalClusters     As Long      'Gesamtzahl der Cluster
  'Variablen für Rest (GetDiskFreeSpaceEx)
  Dim byteAvail(7)       As Byte        'Array zur Aufnahme der verfügbaren Bytes
  Dim byteFree(7)        As Byte         'Array zur Aufnahme der freien Kapazität
  Dim byteTotal(7)       As Byte        'Array zur Aufnahme der Größe des Datenträgers
  Dim iByte              As Long               'Zeiger in byteTotal()
  On Error GoTo errHandler

  If Not bClassReady Then Call InitializeArrays
  'Prüfen, ob GetDiskFreeSpaceEx existiert
  lngResult = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetDiskFreeSpaceExA")

  If lngResult = 0 Then
    'Win 95 vor SR 2
    lngResult = GetDiskFreeSpace(strDrives(Index - 1), nSectorsPerCluster, nBytesPerSector, nFreeClusters, nTotalClusters)
    DriveTotalSpace = CDec(nTotalClusters * nSectorsPerCluster * nBytesPerSector)
  Else
    'der Rest
    lngResult = GetDiskFreeSpaceEx(strDrives(Index - 1), byteAvail(0), byteTotal(0), byteFree(0))

    For iByte = 0 To 7
      DriveTotalSpace = CDec(DriveTotalSpace + byteTotal(iByte) * 2 ^ (8 * iByte))
    Next

  End If

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsSysDrives:DriveTotalSpace"
  Err.Raise Err.number
End Property

