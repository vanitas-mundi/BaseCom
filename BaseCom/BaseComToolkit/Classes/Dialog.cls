VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Dialog
'    Project    : ToolKits
'
'    Description: Setllt Standarddialoge zur Verfügung
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum ServerDataTypes
  TestData = 1
  RealData = 2
  Cancel = 3
End Enum

Private Enum GetFileNameTypes
  SaveFileName = 1
  OpenFilename = 2
End Enum

Private mSelectEntry          As DialogSelectEntry
Private mLoadingScreen        As LoadingScreenDialog
Private mLoadingScreenVisible As Boolean

'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mSelectEntry = New DialogSelectEntry
End Sub

Private Sub Class_Terminate()
  Set mSelectEntry = Nothing
End Sub

'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get SelectEntry() As DialogSelectEntry
  Set SelectEntry = mSelectEntry
End Property

Public Property Get LoadingScreenVisible() As Boolean
  LoadingScreenVisible = mLoadingScreenVisible
End Property

Public Property Let LoadingScreenVisible(ByVal value As Boolean)
  
  If (value) And (Not mLoadingScreenVisible) Then
    Set mLoadingScreen = New LoadingScreenDialog
    mLoadingScreen.Show
  Else

    If Not mLoadingScreen Is Nothing Then
      Unload mLoadingScreen
      Set mLoadingScreen = Nothing
    End If
  End If
  
  mLoadingScreenVisible = value
End Property

'---------------------- Ereignismethoden der Klasse -----------------------------

'---------------------- Private Methoden der Klasse -----------------------------
Private Function GetFileNameBase(ByVal getFileNameType As GetFileNameTypes, _
                                 ByVal title As String, _
                                 ByVal initialPath As String, _
                                 ByVal initialFileName As String, _
                                 ByVal restoreSelectedPath As Boolean, _
                                 ParamArray filterExtensions() As Variant) As String

  Static lastPath As String: If lastPath = "" Then lastPath = "c:\"
  
  If initialPath = "" Then initialPath = lastPath
    
  Dim etc    As etc: Set etc = New etc
  Dim fe     As Variant: fe = filterExtensions(0)
  Dim filter As String
  
  filter = IIf(etc.ArrayCount(fe) = 0, "*", Join(fe, ";"))
  
  Dim defaultExtension As String
  defaultExtension = IIf(etc.ArrayCount(fe) = 0, "", fe(0))
  
  Set etc = Nothing
  
  Dim dialogsEx As Win32ApiDialogsEx
  Set dialogsEx = New Win32ApiDialogsEx
    
  Select Case getFileNameType

    Case GetFileNameTypes.OpenFilename

      If title = "" Then title = "GetOpenFileName"
      GetFileNameBase = dialogsEx.GetOpenFileNameEx(initialPath, initialFileName, filter, title)

    Case GetFileNameTypes.SaveFileName

      If title = "" Then title = "GetSaveFileName"
      GetFileNameBase = dialogsEx.GetSaveFileNameEx(initialPath, initialFileName, filter, defaultExtension, title)

    Case Else
      GetFileNameBase = ""
  End Select
    
  If restoreSelectedPath Then
    Dim io As fileIO: Set io = New fileIO
    lastPath = io.GetParentDirectoryName(GetFileNameBase)
    Set io = Nothing
  End If
  
  Set dialogsEx = Nothing
End Function

Private Function GetMultiFileNameBase _
                                (ByVal getFileNameType As GetFileNameTypes, _
                                 ByVal title As String, _
                                 ByVal initialPath As String, _
                                 ByVal initialFileName As String, _
                                 ByVal restoreSelectedPath As Boolean, _
                                 ParamArray filterExtensions() As Variant) As Collection


  Static lastPath As String: If lastPath = "" Then lastPath = "c:\"
  
  Dim files As Collection: Set files = New Collection
  
  If initialPath = "" Then initialPath = lastPath
    
  Dim etc    As etc: Set etc = New etc
  Dim fe     As Variant: fe = filterExtensions(0)
  Dim filter As String
  Dim result As String
  
  filter = IIf(etc.ArrayCount(fe) = 0, "*", Join(fe, ";"))
  
  Dim defaultExtension As String
  defaultExtension = IIf(etc.ArrayCount(fe) = 0, "", fe(0))
  
  Set etc = Nothing
  
  Dim dialogsEx As Win32ApiDialogsEx
  Set dialogsEx = New Win32ApiDialogsEx
    
  Select Case getFileNameType

    Case GetFileNameTypes.OpenFilename

      If title = "" Then title = "GetOpenFileName"
      result = dialogsEx.GetOpenFileNamesEx(initialPath, initialFileName, filter, title)

    Case GetFileNameTypes.SaveFileName

      If title = "" Then title = "GetSaveFileName"
      result = dialogsEx.GetSaveFileNamesEx(initialPath, initialFileName, filter, defaultExtension, title)

    Case Else
      result = ""
  End Select
  
  Set dialogsEx = Nothing

  If result <> "" Then
    Dim conv As Convert: Set conv = New Convert
    Set files = conv.ArrayToCollection(Split(result, "|"))
    Set conv = Nothing
  
    If restoreSelectedPath And files.count > 0 Then
    Dim io As fileIO: Set io = New fileIO
      lastPath = io.GetParentDirectoryName(files.item(1))
      Set io = Nothing
    End If
  
  End If

  Set GetMultiFileNameBase = files
  
  

'  Const cdlOFNAllowMultiselect = &H200
'  Const cdlOFNExplorer = &H80000
'
'  Dim dialog As Object: Set dialog = CreateObject("MSComDlg.CommonDialog")
'
'  Dim files As Collection: Set files = New Collection
'  Set GetMultiFileNameBase = files
'
'  Dim flag As Boolean
'  Dim aa As String
'  Dim bb As String
'  Dim path As String
'  Dim mem As Boolean
'  Dim pos As Long
'  Dim X As Long
'
'  On Error GoTo Break
'
'  With dialog
'    .DialogTitle = title
'    .filter = filter
'    .DefaultExt = defaultExtention
'    .filename = ""
'    .CancelError = True
'    .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
'
'    Select Case getFileNameType
'    Case OpenFilename
'      .ShowOpen
'    Case SaveFileName
'      .ShowSave
'    End Select
'
'    If flag Then 'Es wurde 'Abbrechen' gewählt!
'      Exit Function
'    Else
'      'Schauen ob evt. nur eine Datei ausgewählt wurde
'      If InStr(.filename, Chr$(0)) <> 0 Then
'        aa = .filename & Chr$(0)
'        pos = 1
'          Do
'            'Chr$(0) ist Trennzeichen
'            X = InStr(pos, aa, Chr$(0))
'
'            If X = 0 Then Exit Do
'              bb = Mid$(aa, pos, X - pos)
'
'              If Not mem Then
'              'Im ersten Teil befindet sich immer der Pfad
'              path = bb
'              mem = True
'            Else
'              'Pfad ist bereits bekannt, es folgen Dateinamen
'              files.Add path & "\" & bb
'            End If
'            pos = X + 1
'          Loop
'        Else
'        'Es wurde nur eine Datei ausgewählt
'        files.Add .filename
'      End If
'    End If
'  End With
'  Exit Function
'
'Break:
'  If Err.number = 32755 Then
'    flag = True
'    Resume Next
'  ElseIf Err.number = 20476 Then flag = True
'    MsgBox "Es wurden zu viele Dateien ausgewählt!"
'  End If
End Function



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetServerDataType
' Description:       Wählt Echt- oder Testdaten aus.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-16:48:35
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function GetServerDataType(ByVal appPath As String, ByVal hInstance As Long) As ServerDataTypes
  Dim f As SelectServerDialog: Set f = New SelectServerDialog
  f.appPath = appPath
  f.hInstance = hInstance
  f.Show 1
  GetServerDataType = f.ServerDataType
  Set f = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSaveFileName
' Description:       Öffnet einen SaveFileDialog mit Multiselect.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-16:48:52
'
' Parameters :       title (String)
'                    initialPath (String)
'                    initialFileName (String)
'                    restoreSelectedPath (Boolean)
'                    filterExtensions() (Variant)
'--------------------------------------------------------------------------------
Public Function GetSaveFileNames(ByVal title As String, _
                                ByVal initialPath As String, _
                                ByVal initialFileName As String, _
                                ByVal restoreSelectedPath As Boolean, _
                                ParamArray filterExtensions() As Variant) As Collection

  Set GetSaveFileNames = GetMultiFileNameBase(SaveFileName, title, initialPath, initialFileName, restoreSelectedPath, filterExtensions)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSaveFileName
' Description:       Öffnet einen SaveFileDialog
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-16:48:52
'
' Parameters :       title (String)
'                    initialPath (String)
'                    initialFileName (String)
'                    restoreSelectedPath (Boolean)
'                    filterExtensions() (Variant)
'--------------------------------------------------------------------------------
Public Function GetSaveFileName(ByVal title As String, _
                                ByVal initialPath As String, _
                                ByVal initialFileName As String, _
                                ByVal restoreSelectedPath As Boolean, _
                                ParamArray filterExtensions() As Variant) As String

  GetSaveFileName = GetFileNameBase(SaveFileName, title, initialPath, initialFileName, restoreSelectedPath, filterExtensions)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetOpenFileNames
' Description:       Öffnet einen OpenFileDialog mit Multiselect.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-16:48:56
'
' Parameters :       title (String)
'                    initialPath (String)
'                    initialFileName (String)
'                    restoreSelectedPath (Boolean)
'                    filterExtensions() (Variant)
'--------------------------------------------------------------------------------
Public Function GetOpenFileNames(ByVal title As String, _
                                ByVal initialPath As String, _
                                ByVal initialFileName As String, _
                                ByVal restoreSelectedPath As Boolean, _
                                ParamArray filterExtensions() As Variant) As Collection

  Set GetOpenFileNames = GetMultiFileNameBase(OpenFilename, title, initialPath, initialFileName, restoreSelectedPath, filterExtensions)
End Function


'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetOpenFileName
' Description:       Öffnet einen OpenFileDialog.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-16:48:56
'
' Parameters :       title (String)
'                    initialPath (String)
'                    initialFileName (String)
'                    restoreSelectedPath (Boolean)
'                    filterExtensions() (Variant)
'--------------------------------------------------------------------------------
Public Function GetOpenFileName(ByVal title As String, _
                                ByVal initialPath As String, _
                                ByVal initialFileName As String, _
                                ByVal restoreSelectedPath As Boolean, _
                                ParamArray filterExtensions() As Variant) As String

  GetOpenFileName = GetFileNameBase(OpenFilename, title, initialPath, initialFileName, restoreSelectedPath, filterExtensions)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetMultiSaveFileName (veraltet)
' Description:       Öffnet einen SaveFileDialog mit MultiSelect.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       23.01.2015-10:12:00
'
' Parameters :       title (String)
'                    filter (String = "Alle Dateien (*.*)
'                    *.*" )
'                    defaultExtention (String = "")
'--------------------------------------------------------------------------------
Public Function GetMultiSaveFileName _
(ByVal title As String _
, Optional ByVal filter As String = "Alle Dateien (*.*)|*.*" _
, Optional ByVal defaultExtention As String = "") As Collection
  
  MsgBox "Funktion veraltet benutzen Sie GetSaveFileNames"
  Set GetMultiSaveFileName = New Collection
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetMultiOpenFileName (veraltet)
' Description:       Öffnet einen OpenFileDialog mit MultiSelect.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       23.01.2015-10:11:56
'
' Parameters :       title (String)
'                    filter (String = "Alle Dateien (*.*)
'                    *.*" )
'                    defaultExtention (String = "")
'--------------------------------------------------------------------------------
Public Function GetMultiOpenFileName _
(ByVal title As String _
, Optional ByVal filter As String = "Alle Dateien (*.*)|*.*" _
, Optional ByVal defaultExtention As String = "") As Collection
  
  MsgBox "Funktion veraltet benutzen Sie GetOpenFileNames"
  Set GetMultiOpenFileName = New Collection
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetPath
' Description:       Öffnet einen OpenPathDialog
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       21.01.2015-15:29:05
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function GetPath() As String
  Dim dialogsEx As Win32ApiDialogsEx
  Set dialogsEx = New Win32ApiDialogsEx
  GetPath = dialogsEx.GetPath
  Set dialogsEx = Nothing
End Function

