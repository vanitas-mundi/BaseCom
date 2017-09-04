VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSqlGroupFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : FileSqlGroupFile
'    Project    : ToolKits
'
'    Description: Stellt Funktionen für den SqlGroupFile-Zugriff zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mDatabaseIntern As Database

Private mStatements As Collection
Private mConv As Convert
Private mEtc As etc
Private mError As ToolkitError
Private mResolver As ParameterResolve
Private mDatabae As Database


'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mDatabae = New Database
  Set mStatements = New Collection
  Set mConv = New Convert
  Set mEtc = New etc
  Set mError = New ToolkitError
  Set mResolver = New ParameterResolve
End Sub

Private Sub Class_Terminate()
  Set mDatabae = Nothing
  Set mStatements = Nothing
  Set mConv = Nothing
  Set mEtc = Nothing
  Set mError = Nothing
  Set mResolver = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get DatabaseIntern() As Database
  Set DatabaseIntern = mDatabaseIntern
End Property

Public Property Let DatabaseIntern(ByVal value As Database)
  Set mDatabaseIntern = value
End Property

Public Property Get LastExecutedStatement() As String
  LastExecutedStatement = mDatabaseIntern.LastExecutedStatement
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Function GetStatementName(ByVal statement As String) As String
  Dim pos As Long: pos = InStr(statement, "*/") - 1
  If pos > 0 Then
    GetStatementName = mEtc.TrimWhiteSpace(Replace(Mid$(statement, 1, pos), "/*", ""))
  Else
    GetStatementName = ""
  End If
End Function

Private Function GetStatementBody _
(ByVal statement As String _
, ByVal statementName As String) As String
    
  GetStatementBody = mEtc.TrimWhiteSpace(Replace$(statement, "/*" & statementName & "*/", ""))
End Function

Private Function StatementContainsStatementNameComment _
(ByVal statementName As String, ByVal statement As String) As Boolean

  Dim nameComment As String: nameComment = "/*" & statementName & "*/"
  StatementContainsStatementNameComment = mEtc.Contains(statement, nameComment, True)
End Function

Private Function AddStatementNameComment _
(ByVal statementName As String, ByVal statement As String) As String
  
  Dim nameComment As String: nameComment = "/*" & statementName & "*/"
  AddStatementNameComment = nameComment & vbCrLf & statement
End Function

Private Sub LoadSqlGroupFileBase(ByVal filename As String, ByVal append As Boolean, Optional ByVal isCrypted As Boolean = False)
  
  If Not append Then Set mStatements = New Collection

  Dim io As FileIO: Set io = New FileIO
  Dim content As String: content = io.ReadAllText(filename)
  
  If isCrypted Then
    Dim crypt As Cryptography: Set crypt = New Cryptography
    content = crypt.DecryptString(content)
    Set crypt = Nothing
  End If

  Dim statements As Collection: Set statements = mConv.SplitCollection(content, "<SQL Name=")

  Dim X As Variant
  For Each X In statements
    Dim currentStatement As String
    currentStatement = mEtc.TrimWhiteSpace(Replace$(CStr(X), "</SQL>", ""))

    Dim pos As Long: pos = InStr(currentStatement, ">") - 1
    If pos > 0 Then
      Dim statementName As String
      statementName = mEtc.TrimWhiteSpace(Mid$(currentStatement, 1, pos))
      pos = pos + 2
      Dim statement As String: statement = mEtc.TrimWhiteSpace(Mid$(currentStatement, pos))

      AddStatement statementName, statement
    End If
  Next X

  Set io = Nothing
End Sub

Private Function BaseResolveStatement(ByVal statementName As String _
, ParamArray parameters() As Variant) As String

  If Me.ExistsStatement(statementName) Then
    Dim statement As String: statement = mStatements.item(LCase(statementName))
    If mEtc.ArrayCount(parameters) > 0 Then
      BaseResolveStatement = mResolver.ReplaceParameters(statement, parameters(0))
    Else
      BaseResolveStatement = mResolver.ReplaceParameters(statement, parameters)
    End If
  Else
    BaseResolveStatement = ""
  End If
End Function



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReadConnectionStringFromIniFile
' Description:       Liest einen ConnectionString aus einer Ini-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-10:42:06
'
' Parameters :       filename (String)
'                    sectionName (String)
'                    isCrypted (Boolean = True)
'--------------------------------------------------------------------------------
Public Function ReadConnectionStringFromIniFile(ByVal filename As String, _
                                                ByVal sectionName As String, _
                                                Optional ByVal isCrypted As Boolean = True) As String

  Dim FileSystem As FileSystem: Set FileSystem = New FileSystem
  
  Dim ConnectionString As String
  ConnectionString = FileSystem.ini.ReadCounterEntryValues _
  (filename, sectionName, "ConnectInfos", "ConnectInfo", "")
  
  If Not isCrypted Then
    ReadConnectionStringFromIniFile = ConnectionString
  Else
    Dim crypt As Cryptography: Set crypt = New Cryptography
    ReadConnectionStringFromIniFile = crypt.EncryptString(ConnectionString)
    Set crypt = Nothing
  End If
  
  Set FileSystem = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReplaceEscape
' Description:       Maskiert in einem String MySQL geschütze Zeichen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:18:07
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function ReplaceEscape(ByVal value As String) As String

  ReplaceEscape = mConv.MySqlReplaceEscape(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       PointToComma
' Description:       Wandelt in einem String Punkte in Kommata um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:19:30
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function PointToComma(ByVal value As String) As String
  PointToComma = mConv.PointToComma(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CommaToPoint
' Description:       Wandelt in einem String in KommataPunkte um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:19:50
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function CommaToPoint(ByVal value As String) As String
  CommaToPoint = mConv.CommaToPoint(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FormatYMD
' Description:       Formatiert einen Datum-String in das Format JJJJ-MM-TT um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:20:06
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function FormatYMD(ByVal value As String) As String
  FormatYMD = mConv.FormatYMD(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FormatYMDHMS
' Description:       Formatiert einen Datum-String in das Format JJJJ-MM-TT HH:MM:SS um.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.04.2015-12:20:06
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function FormatYMDHMS(ByVal value As String) As String
  FormatYMDHMS = mConv.FormatYMDHMS(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FormatDMY
' Description:       Formatiert einen Datum-String in das Format TT-MM-JJJJ um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:20:28
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function FormatDMY(ByVal value As String) As String
  FormatDMY = mConv.FormatDMY(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       TimeSpanToTimeRay
' Description:       Liefert einen String welcher alle Minuten zwischen zwei
'                    Uhrzeiten enthält.
'                    Der Intervall beträgt eine Minute, kann aber über
'                    den optionalen Parameter geändert werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:21:03
'
' Parameters :       timeFrom (String)
'                    timeTo (String )
'                    intervalInMinutes (Integer = 1)
'--------------------------------------------------------------------------------
Public Function TimeSpanToTimeRay _
(ByVal timeFrom As String _
, ByVal timeTo As String _
, Optional intervalInMinutes As Integer = 1) As String

  TimeSpanToTimeRay = mConv.TimeSpanToTimeRay(timeFrom, timeTo, intervalInMinutes)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DateSpanToDateRay
' Description:       Liefert einen String welcher alle Tage zwischen zwei
'                    Daten enthält.
'                    Der Intervall beträgt einen Tag, kann aber über
'                    den optionalen Parameter geändert werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:21:29
'
' Parameters :       dateFrom (String)
'                    dateTo (String )
'                    intervalInDays (Integer = 1)
'--------------------------------------------------------------------------------
Public Function DateSpanToDateRay _
(ByVal dateFrom As String _
, ByVal dateTo As String _
, Optional ByVal intervalInDays As Integer = 1) As String

  DateSpanToDateRay = mConv.DateSpanToDateRay(dateFrom, dateTo, intervalInDays)
End Function


'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ClearStatements
' Description:       Entfernt alle Statements aus der Statement-Auflistung.
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:21:44
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub ClearStatements()
  Set mStatements = New Collection
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       StatementNames
' Description:       Liefert alle Statementnamen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:22:11
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function StatementNames() As Collection

  Set StatementNames = New Collection

  Dim X As Variant
  For Each X In mStatements
    StatementNames.Add GetStatementName(CStr(X))
  Next X
End Function

'Public Function RemoveStatementName(ByVal statement As String) As String
'
'  Dim statementName As String: statementName = GetStatementName(statement)
'  RemoveStatementName = Replace$(statement, "/*" & statementName & "*/", "")
'End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExistsStatement
' Description:       Prüft, ob der Statementname in der Statement-Auflistung vorhanden ist.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:24:43
'
' Parameters :       statementName (String)
'--------------------------------------------------------------------------------
Public Function ExistsStatement(ByVal statementName As String) As Boolean
  
  ExistsStatement = mEtc.CollectionContainsKey(mStatements, LCase(statementName))
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       StatementCount
' Description:       Liefert die Anzahl der Statements in der Statement-Auflistung.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:25:20
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function StatementCount() As Long
  StatementCount = mStatements.count
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       AddStatement
' Description:       Fügt das Statement der Staement-Auflistung hinzu.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:25:44
'
' Parameters :       statementName (String)
'                    statement (String)
'--------------------------------------------------------------------------------
Public Sub AddStatement(ByVal statementName As String, ByVal statement As String)
  
  If Not StatementContainsStatementNameComment(statementName, statement) Then
    statement = AddStatementNameComment(statementName, statement)
  End If
  
  If Not Me.ExistsStatement(statementName) Then
    mStatements.Add statement, statementName
  Else
    Err.Raise ErrorNumbers.SqlGroupFileStatementNameAlreadyExists, "FileSqlGroupFile.AddStatement" _
    , mError.GetErrorMessage(ErrorNumbers.SqlGroupFileStatementNameAlreadyExists)
  End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RenameStatement
' Description:       Benennt das Statement in der Statement-Auflistung um.
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:26:19
'
' Parameters :       oldStatementName (String)
'                    newStatementName (String)
'--------------------------------------------------------------------------------
Public Sub RenameStatement _
(ByVal oldStatementName As String _
, ByVal newStatementName As String)

  If Me.ExistsStatement(oldStatementName) And (Not Me.ExistsStatement(newStatementName)) Then
    Dim Temp As String: Temp = Me.GetStatement(oldStatementName)
    Dim statementName As String: statementName = GetStatementName(Temp)
    Dim statement As String: statement = GetStatementBody(Temp, statementName)
    
    RemoveStatement oldStatementName
    AddStatement newStatementName, statement
  End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RemoveStatement
' Description:       Entfernt das Statement aus der Statement-Auflistung.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:26:48
'
' Parameters :       statementName (String)
'--------------------------------------------------------------------------------
Public Sub RemoveStatement(ByVal statementName As String)

  If Me.ExistsStatement(statementName) Then
    mStatements.Remove LCase(statementName)
  Else
    Err.Raise SqlGroupFileStatementNotExists, "FileSqlGroupFile.RemoveStatement" _
    , mError.GetErrorMessage(SqlGroupFileStatementNotExists)
  End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       LoadSqlGroupFile
' Description:       Lädt eine sgf-Datei in die Statement-Auflistung
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:27:11
'
' Parameters :       filename (String)
'                    isCrypted (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub LoadSqlGroupFile(ByVal filename As String, Optional ByVal isCrypted As Boolean = False)

  LoadSqlGroupFileBase filename, False, isCrypted
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       AppendSqlGroupFile
' Description:       Lädt eine sgf-Datei und fügt sie an die bestehende Statement-Auflistung an.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:43:56
'
' Parameters :       filename (String)
'                    isCrypted (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub AppendSqlGroupFile(ByVal filename As String, Optional ByVal isCrypted As Boolean = False)

  LoadSqlGroupFileBase filename, True, isCrypted
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SaveSqlGroupFile
' Description:       Speicher die aktuelle Statement-Auflistung als sgf-Datei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:44:43
'
' Parameters :       filename (String)
'                    isCrypted (Boolean = False)
'--------------------------------------------------------------------------------
Public Sub SaveSqlGroupFile _
(ByVal filename As String, Optional ByVal isCrypted As Boolean = False)

  If mStatements.count = 0 Then Exit Sub

  Dim conv As Convert: Set conv = New Convert
  Dim content As Collection: Set content = New Collection

  Dim X As Variant
  For Each X In mStatements
    Dim statementName As String: statementName = GetStatementName(CStr(X))
    Dim statement As String: statement = GetStatementBody(CStr(X), statementName)
    
    Dim tag As Collection: Set tag = New Collection
    tag.Add "<SQL Name=" & statementName & ">"
    tag.Add ""
    tag.Add statement
    tag.Add ""
    tag.Add "</SQL>"
    tag.Add "": tag.Add ""
    content.Add conv.JoinCollection(tag, vbCrLf)
  Next X
    
  Dim fileContent As String: fileContent = conv.JoinCollection(content, vbCrLf)
  If isCrypted Then
    Dim crypt As Cryptography: Set crypt = New Cryptography
    fileContent = crypt.EncryptString(fileContent)
    Set crypt = Nothing
  End If
  
  Dim io As FileIO: Set io = New FileIO
  io.WriteAllText filename, fileContent, False
  
  Set io = Nothing
  Set conv = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SaveStatementsAsStatementFiles
' Description:       Speichert die Statements der Auflistung als einzelne Textdateien.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:45:20
'
' Parameters :       path (String)
'--------------------------------------------------------------------------------
Public Sub SaveStatementsAsStatementFiles(ByVal path As String)

  If mStatements.count = 0 Then Exit Sub

  Dim io As FileIO: Set io = New FileIO

  Dim X As Variant
  For Each X In mStatements
    Dim statementName As String: statementName = GetStatementName(CStr(X))
    Dim statement As String: statement = GetStatementBody(CStr(X), statementName)
    Dim filename As String: filename = io.CombinePath(path, statementName & ".sql")
    io.WriteAllText filename, statement, False
  Next X
  
  Set io = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetStatement
' Description:       Liefert das Statement zum angegebenen Statementnamen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:46:30
'
' Parameters :       statementName (String)
'--------------------------------------------------------------------------------
Public Function GetStatement _
(ByVal statementName As String) As String

  If Me.ExistsStatement(statementName) Then
    Dim statement As String: statement = mStatements.item(LCase(statementName))
    GetStatement = statement
  Else
    GetStatement = ""
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SetStatement
' Description:       Ersetzt das Statement mit dem angegebenen Statementnamen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:47:07
'
' Parameters :       statementName (String)
'                    statement (String)
'--------------------------------------------------------------------------------
Public Function SetStatement _
(ByVal statementName As String, ByVal statement As String) As String

  If Me.ExistsStatement(statementName) Then
    If Not StatementContainsStatementNameComment(statementName, statement) Then
      statement = AddStatementNameComment(statementName, statement)
    End If
    mEtc.SetCollectionItem mStatements, LCase(statementName), statement
  Else
    Err.Raise ErrorNumbers.SqlGroupFileStatementNotExists, "FileSqlGroupFile.SetStatement" _
    , mError.GetErrorMessage(SqlGroupFileStatementNotExists)
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ResolveStatement
' Description:       Liefert das Statement zum angegebenen Statementnamen
'                    und löst die übergebenen Parameter auf.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-12:48:03
'
' Parameters :       statementName (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ResolveStatement _
(ByVal statementName As String, ParamArray parameters() As Variant) As String

  ResolveStatement = BaseResolveStatement(statementName, parameters)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExecuteNonQuery
' Description:       Sendet ein DML-Statement zur Datenbank.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       07.01.2015-10:57:20
'
' Parameters :       statement (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ExecuteNonQuery(ByVal statementName As String, _
                                ParamArray parameters() As Variant) As Long
  
  If Me.ExistsStatement(statementName) Then
    ExecuteNonQuery = Me.DatabaseIntern.ExecuteNonQuery _
    (BaseResolveStatement(statementName, parameters))
  Else
    ExecuteNonQuery = -1
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExecuteScalar
' Description:       Liefert genau einen Wert aus einer Datentabelle.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:14:01
'
' Parameters :       statement (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ExecuteScalar(ByVal statementName As String, _
                              ParamArray parameters() As Variant) As Variant

  If Me.ExistsStatement(statementName) Then
    ExecuteScalar = Me.DatabaseIntern.ExecuteScalar _
    (BaseResolveStatement(statementName, parameters))
  Else
    ExecuteScalar = ""
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExecuteIsTrue
' Description:       Liefert genau einen Wert aus einer Datentabelle und prüft,
'                    ob die Rückgabe true oder false ist. Statement muss einen
'                    boolschen Wert liefern.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       26.01.2015-12:48:27
'
' Parameters :       statement (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ExecuteIsTrue(ByVal statementName As String, _
                              ParamArray parameters() As Variant) As Boolean
  
  If Me.ExistsStatement(statementName) Then
    ExecuteIsTrue = Me.DatabaseIntern.ExecuteIsTrue(BaseResolveStatement _
    (statementName, parameters))
  Else
    ExecuteIsTrue = ""
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExecuteReader
' Description:       Führt ein Select-Statement aus und liefert das Ergebnis
'                    als verbindungslosen RecordSet. Die Feldtypen werden als
'                    Strings zurückgegeben.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:14:45
'
' Parameters :       statement (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ExecuteReader(ByVal statementName As String, _
                              ParamArray parameters() As Variant) As Object

  If Me.ExistsStatement(statementName) Then
    Set ExecuteReader = Me.DatabaseIntern.ExecuteReader(BaseResolveStatement _
    (statementName, parameters))
  Else
    Set ExecuteReader = CreateObject("ADODB.RecordSet")
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetRecordSet
' Description:       Führt ein Select-Statement aus und liefert einen RecordSet.
'                    Die Verbindung des RecordSets ist bei Rückgabe noch offen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:16:11
'
' Parameters :       statement (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ExecuteReaderConnected(ByVal statementName As String, _
                                       ParamArray parameters() As Variant) As Object

  If Me.ExistsStatement(statementName) Then
    Set ExecuteReaderConnected = Me.DatabaseIntern.ExecuteReaderConnected _
    (BaseResolveStatement(statementName, parameters))
  Else
    Set ExecuteReaderConnected = CreateObject("ADODB.RecordSet")
  End If
End Function
