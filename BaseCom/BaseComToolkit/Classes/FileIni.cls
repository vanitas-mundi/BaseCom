VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : FileIni
'    Project    : ToolKits
'
'    Description: Stellt Routinen zur Ini-Datei-Bearbeitung zur Verfügung.
'                 ;Kommentar
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Declare Function GetPrivateProfileString _
                Lib "kernel32" _
                Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                  ByVal lpKeyName As Any, _
                                                  ByVal lpDefault As String, _
                                                  ByVal lpReturnedString As String, _
                                                  ByVal nSize As Long, _
                                                  ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString _
                Lib "kernel32" _
                Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Function ReadEntryValueWrapper(ByVal filename As String, _
                                       ByVal sectionName As String, _
                                       ByVal entryName As String) As String
                              
  Dim Temp$: Temp$ = String$(255, 0)
  
  Dim X As Long:  X = GetPrivateProfileString(sectionName, entryName, "", Temp$, 255, filename)
  
  ReadEntryValueWrapper = Left$(Temp$, X)
End Function

Private Sub WriteEntryValueWrapper(filename As String, _
                                   sectionName As String, _
                                   entryName As String, _
                                   value As String)
                          
  WritePrivateProfileString sectionName, entryName, value, filename
End Sub



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReadEntryValue
' Description:       Liefert einen Wert aus einer Ini-Datei.
' Created by :       Sascha Glika
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:06:46
'
' Parameters :       filename (String)
'                    sectionName (String)
'                    entryName (String)
'--------------------------------------------------------------------------------
Public Function ReadEntryValue(ByVal filename As String, _
                               ByVal sectionName As String, _
                               ByVal entryName As String) As String
                            
  ReadEntryValue = ReadEntryValueWrapper(filename, sectionName, entryName)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       WriteEntryValue
' Description:       Schreibt einen Wert in einer Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:12:54
'
' Parameters :       filename (String)
'                    sectionName (String)
'                    entryName (String)
'                    value (String)
'--------------------------------------------------------------------------------
Public Sub WriteEntryValue(ByVal filename As String, _
                           ByVal sectionName As String, _
                           ByVal entryName As String, _
                           ByVal value As String)
                             
  WriteEntryValueWrapper filename, sectionName, entryName, value
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DeleteEntry
' Description:       Löscht einen Wert aus einer Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:13:29
'
' Parameters :       filename (String)
'                    sectionName (String)
'                    entryName (String)
'--------------------------------------------------------------------------------
Public Sub DeleteEntry(ByVal filename As String, _
                       sectionName As String, _
                       ByVal entryName As String)
                          
  WritePrivateProfileString sectionName, entryName, ByVal 0&, filename
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DeleteSection
' Description:       Löscht eine Sektion aus einer Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:13:47
'
' Parameters :       filename (String)
'                    sectionName (String)
'--------------------------------------------------------------------------------
Public Sub DeleteSection(ByVal filename As String, ByVal sectionName As String)

  WritePrivateProfileString sectionName, ByVal 0&, ByVal 0&, filename
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReadCounterEntryValues
' Description:       Liest Werte aus einer Sektion einer Ini-Datei über einen
'                    Counter-Wert.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:14:11
'
' Parameters :       filename (String)
'                    sectionName (String)
'                    entryNameCounter (String)
'                    entryName (String)
'                    Delimiter (String = "#")
'--------------------------------------------------------------------------------
Public Function ReadCounterEntryValues(ByVal filename As String, _
                                       ByVal sectionName As String, _
                                       ByVal entryNameCounter As String, _
                                       ByVal entryName As String, _
                                       Optional ByVal Delimiter As String = "#") As String
  
  Dim counter As Long: counter = ReadEntryValue(filename, sectionName, entryNameCounter)

  Dim conv    As Convert: Set conv = New Convert
  Dim result  As Collection: Set result = New Collection
  
  Dim i       As Integer

  For i = 1 To counter
    result.Add ReadEntryValue(filename, sectionName, entryName & i)
  Next i

  ReadCounterEntryValues = conv.JoinCollection(result, Delimiter)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       WriteCounterEntryValues
' Description:       Schreibt Werte in eine Sektion einer Ini-Datei über einen
'                    Counter-Wert.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:15:11
'
' Parameters :       filename (String)
'                    sectionName (String)
'                    entryNameCounter (String)
'                    entryName (String)
'                    value (String)
'--------------------------------------------------------------------------------
Public Sub WriteCounterEntryValues(ByVal filename As String, _
                                   ByVal sectionName As String, _
                                   ByVal entryNameCounter As String, _
                                   ByVal entryName As String, _
                                   ByVal value As String)
                                    
  Dim countervalue As String: countervalue = ReadEntryValue(filename, sectionName, entryNameCounter)
  Dim counter      As Long

  If countervalue = "" Then
    counter = 1
  Else
    counter = CInt(countervalue) + 1
  End If
  
  WriteEntryValue filename, sectionName, entryNameCounter, CStr(counter)
  WriteEntryValue filename, sectionName, entryName & counter, value
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DeleteCounterEntry
' Description:       Löscht Werte aus einer Sektion einer Ini-Datei über einen
'                    Counter-Wert.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:15:38
'
' Parameters :       filename (String)
'                    sectionName (String)
'                    entryNameCounter (String)
'                    entryName (String)
'                    value (String)
'--------------------------------------------------------------------------------
Public Sub DeleteCounterEntry(ByVal filename As String, _
                              ByVal sectionName As String, _
                              ByVal entryNameCounter As String, _
                              ByVal entryName As String, _
                              ByVal value As String)

  Dim counter    As Long: counter = ReadEntryValue(filename, sectionName, entryNameCounter)
  Dim newCounter As Long: newCounter = 0
  
  Dim i          As Integer

  For i = 1 To counter
    Dim iniValue As String: iniValue = ReadEntryValue(filename, sectionName, entryName & i)
    DeleteEntry filename, sectionName, entryName & i

    If iniValue <> value Then '{Eintrag neu anlegen}
      newCounter = newCounter + 1
      WriteEntryValue filename, sectionName, entryName & newCounter, iniValue
    End If

  Next i

  WriteEntryValue filename, sectionName, entryNameCounter, CStr(newCounter)
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSections
' Description:       Liefert sämtliche Sektions-Namen einer Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:16:00
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Function GetSections(ByVal filename As String) As Collection
  
  Set GetSections = New Collection
  
  Dim io      As FileIO: Set io = New FileIO
  Dim iniFile As String: iniFile = io.ReadAllText(filename)
  
  Dim Temp()  As String: Temp = Split(iniFile, "[")
  
  Dim i       As Integer

  For i = 1 To UBound(Temp)
    Dim temp2() As String: temp2 = Split(Temp(i), "]")
    GetSections.Add temp2(0)
  Next i

  Set io = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSectionEntryNames
' Description:       Liefert sämtliche Wertenamen einer Sektion einer Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:16:28
'
' Parameters :       filename (String)
'                    sectionName (String)
'--------------------------------------------------------------------------------
Public Function GetSectionEntryNames(ByVal filename As String, _
                                     ByVal sectionName As String) As Collection
                                  
  Set GetSectionEntryNames = New Collection
  sectionName = LCase$(sectionName)
  
  Dim io      As FileIO: Set io = New FileIO
  Dim iniFile As String: iniFile = io.ReadAllText(filename)
  Dim Temp()  As String: Temp = Split(iniFile, "[")
  
  Dim i       As Integer

  For i = 1 To UBound(Temp)
    Dim temp2()            As String: temp2 = Split(Temp(i), "]")
    Dim currentSectionName As String: currentSectionName = LCase(temp2(0))

    If currentSectionName = sectionName Then
      Dim entrieNamesString As String
      entrieNamesString = Replace(Replace(temp2(1), vbCrLf, vbLf), Chr(13), vbLf)
      
      Dim temp3() As String: temp3 = Split(entrieNamesString, vbLf)
      Dim X       As Integer

      For X = LBound(temp3) To UBound(temp3)
        Dim entry As String: entry = temp3(X)
        Dim pos   As Integer: pos = InStr(entry, "=") - 1
        
        If pos > 0 Then
          GetSectionEntryNames.Add Trim(Mid$(entry, 1, pos))
        End If

      Next X

      Exit Function
    End If

  Next i
  
  Set io = Nothing
End Function

