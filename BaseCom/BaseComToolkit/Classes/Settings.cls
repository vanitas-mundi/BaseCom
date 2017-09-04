VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Settings
'    Project    : ToolKits
'
'    Description: Stellt einer Applikation Programmeinstellungen zur Verfügung.
'
'                 Beispiel:
'                 Dem Projekt ein Modul basGlobal hinzufügen.
'
'                 Option Explicit
'
'                 Public Enum SettingNames
'                   SettingName = 1
'                   SettingUser = 2
'                   SettingComputer = 3
'                 End Enum
'
'                 Public BaseToolkit As New ToolKits.BaseToolKitVb6
'
'                 Public Sub InitializeSettings()
'                   BaseToolkit.Settings.InitializeSettings SettingName _
'                   , "Max Mustermann", "max.mustermann", "pc-mustermann"
'                 End Sub
'
'                 Public Function Settings(ByVal key As SettingNames)
'                   Settings = BaseToolkit.Settings.Item(key)
'                 End Function
'
'                 ---
'
'                 Nun kann von überall auf die Settings zugegriffen werden:
'
'                 MsgBox Settings(SettingName)
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mCrypt As Cryptography
Private mEtc As etc
Private mIni As FileIni
Private mSettings As Collection



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mCrypt = New Cryptography
  Set mEtc = New etc
  Set mIni = New FileIni
  Set mSettings = New Collection
End Sub

Private Sub Class_Terminate()
  Set mCrypt = Nothing
  Set mEtc = Nothing
  Set mIni = Nothing
  Set mSettings = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       Register
' Description:       Registriert ein neues Setting.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:43:50
'
' Parameters :       value (String)
'                    key (Integer)
'--------------------------------------------------------------------------------
Public Sub Register(ByVal value As String, ByVal key As Integer, Optional ByVal crypt As Boolean)
  mSettings.Add IIf(crypt, mCrypt.EncryptString(value), value), "key" & CStr(key)
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       UnRegister
' Description:       Entfernt die Registrierung des angegebenen Settings.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:44:48
'
' Parameters :       key (Integer)
'--------------------------------------------------------------------------------
Public Sub UnRegister(ByVal key As Integer)
  mSettings.Remove "key" & CStr(key)
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReRegister
' Description:       Registriert einen neuen Wert unter einem bereits vorhandenen Setting-Key.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:45:26
'
' Parameters :       value (String)
'                    key (Integer)
'--------------------------------------------------------------------------------
Public Sub ReRegister(ByVal value As String, ByVal key As Integer, Optional ByVal crypt As Boolean)
  Dim k As String: k = "key" & CStr(key)
  If mEtc.CollectionContainsKey(mSettings, k) Then mSettings.Remove k
  Register value, key, crypt
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       UnRegisterAll
' Description:       Entfernt alle Settings aus der Registrierung
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:46:16
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub UnRegisterAll()
  Set mSettings = New Collection
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       InitializeSettings
' Description:       Initialisiert die Settings mit den angegebenen Einstellungen.
'                    StartIndex ist der erste Wert der SettingsEnum.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:46:48
'
' Parameters :       startKey (Integer)
'                    args() (Variant)
'--------------------------------------------------------------------------------
Public Sub InitializeSettings(ByVal startKey As Integer, ParamArray args() As Variant)
  Dim i As Integer
  Dim key As Integer: key = startKey
  
  For i = LBound(args) To UBound(args)
    Me.Register CStr(args(i)), key
    key = key + 1
  Next i
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       Item
' Description:       Liefert den Wert des angegebenen Settings
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:47:57
'
' Parameters :       key (Integer)
'                    crypt (Boolean = False)
'--------------------------------------------------------------------------------
Public Function item(ByVal key As Integer, Optional ByVal crypt As Boolean = False) As String
  Dim k As String: k = "key" & CStr(key)
  
  If mEtc.CollectionContainsKey(mSettings, k) Then
    If crypt Then
      item = mCrypt.DecryptString(mSettings.item(k))
    Else
      item = mSettings.item(k)
    End If
  Else
    item = ""
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       LoadSettings
' Description:       Lädt Settings aus der angegebenen Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:50:13
'
' Parameters :       filename (String)
'                    isCrypted (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub LoadSettings(ByVal filename As String, Optional ByVal isCrypted As Boolean = False)

  Dim valueString As String: valueString = mIni.ReadCounterEntryValues(filename, "Settings", "keyCounter", "key", vbCrLf)
  Dim values() As String
  values = Split(valueString, vbCrLf)
  
  Dim i As Integer
  For i = LBound(values) To UBound(values)
    Dim value As String: value = Replace(IIf(isCrypted, mCrypt.DecryptString(values(i)), values(i)), "~~", vbCrLf)
    Me.Register value, i + 1
  Next i
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SaveSettings
' Description:       Speichert Settings in der angegebenen Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:51:15
'
' Parameters :       filename (String)
'                    isCrypted (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub SaveSettings(ByVal filename As String, Optional ByVal isCrypted As Boolean = False)
  
  Dim io As FileIO: Set io = New FileIO
  
  If io.FileExists(filename) Then
    mIni.WriteEntryValue filename, "Settings", "keyCounter", "0"
  End If
  
  Dim i As Integer
  For i = 1 To mSettings.count
    Dim value As String: value = Replace$(mSettings.item("key" & i), vbCrLf, "~~")
    If isCrypted Then value = mCrypt.EncryptString(value)
    
    mIni.WriteCounterEntryValues filename, "Settings", "keyCounter", "key", value
  Next i
  
  Set io = Nothing
End Sub

