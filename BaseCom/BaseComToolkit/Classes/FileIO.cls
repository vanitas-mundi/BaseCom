VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : FileIO
'    Project    : ToolKits
'
'    Description: Stellt I/O Routinen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------
' Noch nicht implementiert
' fso.Drives
' fso.GetDrive
' fso.GetFile
' fso.GetFolder


'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum ShellRunWindowStyles
  '''<summary>Hides the window and activates another window.</summary>
  HideAndNotActivate = 0
  '''<summary>
  ''' Activates and displays a window.
  ''' If the window is minimized or maximized, the system restores
  ''' it to its original size and position. An application should specify
  ''' this flag when displaying the window for the first time.
  '''</summary>
  DisplayWindowAndActivateLastWindowState = 1 '
  '''<summary>Activates the window and displays it as a minimized window.</summary>
  DisplayWindowAndActivateMinimized = 2
  '''<summary>Activates the window and displays it as a maximized window.</summary>
  DisplayWindowAndActivateMaximized = 3
  '''<summary>
  ''' Displays a window in its most recent size and position.
  ''' The active window remains active.
  '''</summary>
  DisplayWindowAndActivateRecentSize = 4
  '''<summary>Activates the window and displays it in its current size and position.</summary>
  DisplayWindowAndActivateCurrentSize = 5
  '''<summary>Minimizes the specified window and activates the next top-level window in the Z order.</summary>
  MinimizeWindowAndActivateNextTopLevelWindow = 6
  '''<summary>Displays the window as a minimized window. The active window remains active.</summary>
  DisplayWindowNotActivate = 7
  '''<summary>Displays the window in its current state. The active window remains active.</summary>
  DisplayWindowCurrentStateNotActivate = 8
  '''<summary>
  ''' Activates and displays the window.
  ''' If the window is minimized or maximized, the system restores it to its original size and position.
  ''' An application should specify this flag when restoring a minimized window.
  '''</summary>
  DisplayWindowAndActivateOriginalSize = 9
  '''<summary>Sets the show state based on the state of the program that started the application.</summary>
  BasedOnStateOfStartedProgram = 10
End Enum

Public Enum SpecialDirectories
  WindowsFolder = 0
  SystemFolder = 1
  TemporaryFolder = 2
End Enum

Public Enum SizeTypes
  bytes = 1
  KBytes = 2
  MBytes = 3
  GBytes = 4
End Enum

Public Enum ShellExecuteViewModes
  Hide = 0
  Normal = 1
End Enum

Private mFso As Object



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mFso = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
  Set mFso = Nothing
End Sub

'---------------------- Zugriffsmethoden der Klasse -----------------------------
'---------------------- Ereignismethoden der Klasse -----------------------------
'---------------------- Private Methoden der Klasse -----------------------------
'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       WriteAllBytes
' Description:       Wandelt ein ByteArray in eine Datei um
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       1.28.2016-14:24:30
'
' Parameters :       strByteArray() (String)
'                    strFileName (String)
'--------------------------------------------------------------------------------
Public Function WriteAllBytes(ByRef strByteArray As Variant, _
                              ByVal strFileName As String) As Boolean

    '<EhHeader>
    On Error GoTo WriteAllBytes_Err

    '</EhHeader>

    Dim fnum As Integer
    fnum = FreeFile
    Open strFileName For Binary As #fnum
    Put #fnum, 1, strByteArray
    Close fnum
    
    WriteAllBytes = True

    '<EhFooter>
    Exit Function

WriteAllBytes_Err:
    WriteAllBytes = False
    'Err.Raise vbObjectError, "ToolKits.FileIO.WriteAllBytes", "FileIO component failure"

    '</EhFooter>

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReadAllText
' Description:       Lädt eine Textdatei und gibt deren Inhalt als String zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:27:00
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Function ReadAllText(ByVal filename As String) As String
  Const ForReading = 1
  
  Dim ts     As Object: Set ts = mFso.OpenTextFile(filename, ForReading)
  
  Dim result As Collection: Set result = New Collection
  While Not ts.AtEndOfStream
    result.Add ts.ReadLine
  Wend
  ts.Close
  
  Dim conv As Convert: Set conv = New Convert
  ReadAllText = conv.JoinCollection(result, vbCrLf)
  
  Set ts = Nothing
  Set conv = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       WriteAllText
' Description:       Schreibt den übergebenen Text in eine Textdatei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:27:41
'
' Parameters :       filename (String)
'                    text (String)
'                    append (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub WriteAllText(ByVal filename As String, _
                        ByVal Text As String, _
                        Optional ByVal append As Boolean = False)

  Const ForWriting = 2
  Const ForAppending = 8
  
  Dim Create As Boolean: Create = Not append
  
  Dim ts     As Object: Set ts = mFso.OpenTextFile(filename, IIf(append, ForAppending, ForWriting), Create)
  
  Dim conv   As Convert: Set conv = New Convert
  Dim lines  As Collection: Set lines = conv.SplitCollection(Text, vbCrLf)
  
  Dim X      As Variant

  For Each X In lines
    ts.WriteLine CStr(X)
  Next X

  ts.Close
  
  Set ts = Nothing
  Set conv = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FileExists
' Description:       Prüft, ob die angegebene Datei existiert.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:28:15
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Function FileExists(ByVal filename As String) As Boolean
  
  FileExists = mFso.FileExists(filename)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DirectoryExists
' Description:       Prüft, ob das angegebene Verzeichnis existiert.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:28:35
'
' Parameters :       path (String)
'--------------------------------------------------------------------------------
Public Function DirectoryExists(ByVal path As String) As Boolean
  
  DirectoryExists = mFso.FolderExists(path)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CombinePath
' Description:       Setzt Pfad und Dateiname zu einem Gesamt-Dateinamen zusammen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:29:04
'
' Parameters :       path (String)
'                    filename (String)
'--------------------------------------------------------------------------------
Public Function CombinePath(ByVal path As String, ByVal filename As String) As String
  
  CombinePath = mFso.BuildPath(path, filename)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CreateDirectory
' Description:       Legt das angegebene Verzeichnis an.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:29:41
'
' Parameters :       path (String)
'--------------------------------------------------------------------------------
Public Sub CreateDirectory(ByVal path As String)
    If Not Me.DirectoryExists(path) Then
        mFso.CreateFolder path
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetTempFileName
' Description:       Liefert einen temporären Filename im Temp-Verzeichnis.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:29:58
'
' Parameters :       extension (String)
'--------------------------------------------------------------------------------
Public Function GetTempFileName(ByVal extension As String) As String

        GetTempFileName = Me.CombinePath(Me.GetSpecialDirectory(TemporaryFolder), mFso.GetTempName & Replace$("." & extension, "..", "."))
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DeleteFile
' Description:       Löscht die angegebene Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:31:32
'
' Parameters :       filename (String)
'                    force (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub DeleteFile(ByVal filename As String, Optional ByVal force As Boolean = False)
    If Me.FileExists(filename) Then
        mFso.DeleteFile filename, force
    End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DeleteDirectory
' Description:       Löscht das angegebenen Verzeichnis.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:31:48
'
' Parameters :       path (String)
'                    force (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub DeleteDirectory(ByVal path As String, Optional ByVal force As Boolean = False)

    On Error Resume Next
        If Me.DirectoryExists(path) Then
                mFso.DeleteFolder path, force
        End If
    On Error GoTo 0
    
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetExtensionName
' Description:       Liefert den Namen der Extension.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:32:22
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Function GetExtensionName(ByVal filename As String) As String
  GetExtensionName = mFso.GetExtensionName(filename)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetFileVersion
' Description:       Liefert die Version der angegebenen Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:32:39
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Function GetFileVersion(ByVal filename As String) As String
  GetFileVersion = mFso.GetFileVersion(filename)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSpecialDirectory
' Description:       Liefert den Pfad diverser Systemordner.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:32:56
'
' Parameters :       directory (SpecialDirectories)
'--------------------------------------------------------------------------------
Public Function GetSpecialDirectory(ByVal directory As SpecialDirectories) As String
  GetSpecialDirectory = mFso.GetSpecialFolder(directory)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MoveFile
' Description:       Verschiebt die angegebene Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:33:29
'
' Parameters :       sourceFileName (String)
'                    destinationFileName (String)
'--------------------------------------------------------------------------------
Public Sub MoveFile(ByVal sourceFileName As String, ByVal destinationFileName As String)
  mFso.MoveFile sourceFileName, destinationFileName
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MoveDirectory
' Description:       Verschiebt das angegebene Verzeichnis.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:33:54
'
' Parameters :       sourcePath (String)
'                    destinationPath (String)
'--------------------------------------------------------------------------------
Public Sub MoveDirectory(ByVal sourcePath As String, ByVal destinationPath As String)
  mFso.MoveFolder sourcePath, destinationPath
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CopyFile
' Description:       Kopiert die angegebene Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:34:11
'
' Parameters :       sourceFileName (String)
'                    destinationFileName (String)
'                    override (Boolean = True)
'--------------------------------------------------------------------------------
Public Sub CopyFile(ByVal sourceFileName As String, _
                    ByVal destinationFileName As String, _
                    Optional override As Boolean = True)

  mFso.CopyFile sourceFileName, destinationFileName, override
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CopyDirectory
' Description:       Kopiert das angegebene Verzeichnis.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:34:33
'
' Parameters :       sourcePath (String)
'                    destinationPath (String)
'                    override (Boolean = True)
'--------------------------------------------------------------------------------
Public Sub CopyDirectory(ByVal sourcePath As String, _
                         ByVal destinationPath As String, _
                         Optional override As Boolean = True)

  mFso.CopyFolder sourcePath, destinationPath, override
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetFileNameWithoutPath
' Description:       Liefert aus filename den Namen der Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:35:22
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Function GetFileNameWithoutPath(ByVal filename As String) As String
  GetFileNameWithoutPath = mFso.GetFileName(filename)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetDriveName
' Description:       Liefert den Volume-Name eines Laufwerks.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:35:54
'
' Parameters :       drive (String)
'--------------------------------------------------------------------------------
Public Function GetDriveName(ByVal drive As String) As String
  GetDriveName = mFso.GetDriveName(drive)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetParentDirectoryName
' Description:       Liefert das Vorgängerverzeichnis eins Pfads.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:36:24
'
' Parameters :       path (String)
'--------------------------------------------------------------------------------
Public Function GetParentDirectoryName(ByVal path As String) As String
  GetParentDirectoryName = mFso.GetParentFolderName(path)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetAbsolutePathName
' Description:       Liefert den absoluten Pfadnamen.
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:37:03
'
' Parameters :       path (String)
'--------------------------------------------------------------------------------
Public Function GetAbsolutePathName(ByVal path As String) As String
  GetAbsolutePathName = mFso.GetAbsolutePathName(path)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetBaseName
' Description:       Liefert den Basisnamen eines Pfads.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:37:23
'
' Parameters :       path (String)
'--------------------------------------------------------------------------------
Public Function GetBaseName(ByVal path As String) As String
  GetBaseName = mFso.GetBaseName(path)
End Function
  
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetFileSize
' Description:       Liefert die Größe einer Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:37:51
'
' Parameters :       filename (String)
'                    SizeType (SizeTypes)
'--------------------------------------------------------------------------------
Public Function GetFileSize(ByVal filename As String, _
                            Optional SizeType As SizeTypes) As Long
  Dim b As Long: b = mFso.GetFile(filename).Size

  Select Case SizeType

    Case SizeTypes.bytes
      GetFileSize = b

    Case SizeTypes.KBytes
      GetFileSize = b \ 1024

    Case SizeTypes.MBytes
      GetFileSize = (b / 1024) \ 1024

    Case SizeTypes.GBytes
      GetFileSize = ((b / 1024) / 1024) \ 1024

    Case Else
  End Select

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ShellExecute
' Description:       Run a script or application in the Windows Shell.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       06.02.2015-13:56:56
'
' Parameters :       fileName (String) - The file to execute (required)
'                    args (String = "" )- Arguments for the executable
'                    workingDirectory (String = "" ) - Working directory
'                    viewMode (ShellExecuteViewModes = Normal)
'--------------------------------------------------------------------------------
Public Sub ShellExecute _
(ByVal filename As String _
, Optional args As String = "" _
, Optional ByVal workingDirectory As String = "" _
, Optional ByVal verb As String = "open" _
, Optional ByVal viewMode As ShellExecuteViewModes = Normal)

  Dim sh As Object: Set sh = CreateObject("Shell.Application")
  Select Case True
  Case (args <> "") And (workingDirectory <> "")
    sh.ShellExecute filename, args, workingDirectory, verb, viewMode
  Case (args <> "")
    sh.ShellExecute filename, args, , verb, viewMode
  Case (workingDirectory <> "")
    sh.ShellExecute filename, , workingDirectory, verb, viewMode
  Case Else
    sh.ShellExecute filename, , , verb, viewMode
  End Select
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ShellRun
' Description:       Führt eine Applikation in eiem neuen Prozess aus und wartet
'                    ggf. auf die Beendigung.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       12.03.2015-12:33:58
'
' Parameters :       filename (String)
'                    windowStyle (ShellRunWindowStyles = HideAndNotActivate )
'                    waitOnReturn (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub ShellRun _
(ByVal filename As String _
, Optional ByVal windowStyle As ShellRunWindowStyles = HideAndNotActivate _
, Optional ByVal waitOnReturn As Boolean = False)

  Dim sh As Object
  Set sh = CreateObject("WScript.Shell")
  sh.Run filename, windowStyle
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ShellRun
' Description:       Liest eine Datei binär ein und liefert den Hinhalt als Byte-Array.
'
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       06.10.2015-12:33:58
'
' Parameters :       fileName (String)
'--------------------------------------------------------------------------------
Public Function GetByteArrayFromFile(ByVal filename As String) As Byte()

  Dim nFile As Integer
  
  nFile = FreeFile
  Open filename For Binary Access Read As #nFile
    If LOF(nFile) > 0 Then
      ReDim GetByteArrayFromFile(0 To LOF(nFile) - 1)
      Get nFile, , GetByteArrayFromFile
    End If
  Close #nFile
End Function

