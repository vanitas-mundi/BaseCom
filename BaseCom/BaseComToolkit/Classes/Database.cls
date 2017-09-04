VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Database
'    Project    : ToolKits
'
'    Description: Setllt Datenbankroutinen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mCon                   As Object
Private mConnectionString      As String
Private mConv                  As Convert
Private mResolver              As ParameterResolve
Private mLastExecutedStatement As String



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mConv = New Convert
  Set mCon = CreateObject("ADODB.Connection")
  Set mResolver = New ParameterResolve
End Sub

Private Sub Class_Terminate()
  Set mConv = Nothing
  CloseConnection mCon
  Set mResolver = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ConnectionString() As String
  ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal value As String)
  mConnectionString = value
  CloseConnection mCon
  Set mCon = CreateObject("ADODB.Connection")
  mCon.ConnectionString = value
  mCon.Open
End Property

Public Property Get LastExecutedStatement() As String
  LastExecutedStatement = mLastExecutedStatement
End Property

Public Property Get ConnectionState() As Integer
  If mCon Is Nothing Then
    ConnectionState = -1
  Else
    ConnectionState = mCon.State
  End If
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CloseConnection
' Description:       Schliesst die aktive Verbindung.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       22.01.2015-15:09:18
'
' Parameters :       rs (Object)
'--------------------------------------------------------------------------------
Public Sub CloseConnection(ByVal con As Object)
  On Error Resume Next

  con.Close
  Set con = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CloseRecordSetAndConnection
' Description:       Schliesst die aktive Verbindung.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:12:05
'
' Parameters :       rs (Object)
'--------------------------------------------------------------------------------
Public Sub CloseRecordSetAndConnection(ByVal rs As Object)
  On Error Resume Next

  rs.Close
  rs.ActiveConnection.Close

  Set rs = Nothing
  Set rs.ActiveConnection = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CloseRecordSet
' Description:       Schliesst Recordset.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:12:05
'
' Parameters :       rs (Object)
'--------------------------------------------------------------------------------
Public Sub CloseRecordSet(ByVal rs As Object)
  On Error Resume Next
  rs.Close
  Set rs = Nothing
End Sub

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
Public Function ExecuteNonQuery(ByVal statement As String, _
                                ParamArray parameters() As Variant) As Long
                                
  mLastExecutedStatement = mResolver.ReplaceParameters(statement, parameters)
  mCon.Execute mLastExecutedStatement
  
  If InStr(LCase(statement), "insert into") > 0 Then
    Dim rs As Object
    Set rs = mCon.Execute("SELECT last_insert_id()")
    ExecuteNonQuery = CLng(rs.Fields.item(0).value)
  Else
    ExecuteNonQuery = 0
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
Public Function ExecuteScalar(ByVal statement As String, _
                              ParamArray parameters() As Variant) As Variant

  mLastExecutedStatement = mResolver.ReplaceParameters(statement, parameters)
  Dim rs As Object: Set rs = mCon.Execute(mLastExecutedStatement)
  If rs.EOF Then
    ExecuteScalar = ""
  Else
    ExecuteScalar = rs.Fields.item(0).value & ""
  End If
  Me.CloseRecordSet rs
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
Public Function ExecuteIsTrue(ByVal statement As String, _
                              ParamArray parameters() As Variant) As Boolean
  
  mLastExecutedStatement = mResolver.ReplaceParameters(statement, parameters)
  Dim rs As Object: Set rs = mCon.Execute(mLastExecutedStatement)
  If rs.EOF Then
    ExecuteIsTrue = False
  Else
    ExecuteIsTrue = CBool(rs.Fields.item(0).value & "")
  End If
  Me.CloseRecordSet rs
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
Public Function ExecuteReader(ByVal statement As String, _
                              ParamArray parameters() As Variant) As Object

  Const adBstr As Integer = 8

  mLastExecutedStatement = mResolver.ReplaceParameters(statement, parameters)
  Dim rs As Object: Set rs = mCon.Execute(mLastExecutedStatement)
  
  Dim resultRs As Object: Set resultRs = CreateObject("ADODB.RecordSet")
  
  Dim i        As Integer
  For i = 0 To rs.Fields.count - 1
    resultRs.Fields.append rs.Fields.item(i).name, adBstr
  Next i
  
  resultRs.Open
  
  While Not rs.EOF
    resultRs.AddNew
    
    For i = 0 To rs.Fields.count - 1
      resultRs.Fields.item(i).value = rs.Fields.item(i).value & ""
    Next i
    
    resultRs.Update
    rs.MoveNext
  Wend
  
  If Not resultRs.EOF Then resultRs.MoveFirst
  Set ExecuteReader = resultRs
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
Public Function ExecuteReaderConnected(ByVal statement As String, _
                                       ParamArray parameters() As Variant) As Object

  mLastExecutedStatement = mResolver.ReplaceParameters(statement, parameters)
  Set ExecuteReaderConnected = mCon.Execute(mLastExecutedStatement)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetConnection
' Description:       Liefert ein geöffnetes ConnectionObject
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       22.01.2015-15:08:30
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function GetConnection() As Object

  Set GetConnection = CreateObject("ADODB.Connection")
  GetConnection.Open Me.ConnectionString
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetValueString
' Description:       Liefert aus einem RecordSet alle Werte des Feldes fieldName
'                    als String, getrennt durch delimiter
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:17:33
'
' Parameters :       rs (Object)
'                    fieldName (String)
'                    Delimiter (String = ",")
'                    quoteChar (String = "")
'--------------------------------------------------------------------------------
Public Function GetValueString(ByRef rs As Object, _
                               ByVal fieldName As String, _
                               Optional ByVal Delimiter As String = ",", _
                               Optional ByVal quoteChar As String = "") As String

  Dim result As Collection: Set result = New Collection

  While Not rs.EOF
    result.Add quoteChar & CStr(rs.Fields(fieldName).value) & quoteChar
    rs.MoveNext
  Wend
  
  Dim conv As Convert: Set conv = New Convert
  GetValueString = conv.JoinCollection(result, Delimiter)
  Set conv = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetValuesString
' Description:       Liefert aus einem RecordSet alle Werte aller Felder
'                    als String, getrennt durch delimiter
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:17:33
'
' Parameters :       rs (Object)
'                    Delimiter (String = ",")
'                    quoteChar (String = "")
'--------------------------------------------------------------------------------
Public Function GetValuesString(ByRef rs As Object, _
                               Optional ByVal Delimiter As String = ",", _
                               Optional ByVal quoteChar As String = "") As String

  Dim result As Collection: Set result = New Collection

  While Not rs.EOF
    Dim i   As Integer
    For i = 0 To rs.Fields.count - 1
      result.Add quoteChar & CStr(rs.Fields(i).value) & quoteChar
    Next i
    rs.MoveNext
  Wend
  
  Dim conv As Convert: Set conv = New Convert
  GetValuesString = conv.JoinCollection(result, Delimiter)
  Set conv = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetValueArray
' Description:       Liefert aus einem RecordSet alle Werte des Feldes fieldName
'                    als String, getrennt durch delimiter in einem Array
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:18:46
'
' Parameters :       rs (Object)
'                    fieldName (String)
'                    quoteChar (String = "")
'--------------------------------------------------------------------------------
Public Function GetValueArray(ByRef rs As Object, _
                              ByVal fieldName As String, _
                              Optional ByVal quoteChar As String = "") As Variant

  GetValueArray = Split(GetValueString(rs, fieldName, "~,;", quoteChar), "~,;")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetValuesArray
' Description:       Liefert aus einem RecordSet alle Werte aller Felder
'                    als String, getrennt durch delimiter in einem Array
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:18:46
'
' Parameters :       rs (Object)
'                    quoteChar (String = "")
'--------------------------------------------------------------------------------
Public Function GetValuesArray(ByRef rs As Object, _
                               Optional ByVal quoteChar As String = "") As Variant

  GetValuesArray = Split(GetValuesString(rs, "~,;", quoteChar), "~,;")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetValueCollection
' Description:       Liefert aus einem RecordSet alle Werte des Feldes fieldName
'                    als String, getrennt durch delimiter in einer Collection.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:19:29
'
' Parameters :       rs (Object)
'                    fieldName (String)
'                    quoteChar (String = "")
'--------------------------------------------------------------------------------
Public Function GetValueCollection(ByRef rs As Object, _
                                   ByVal fieldName As String, _
                                   Optional ByVal quoteChar As String = "") As Collection

  Dim conv   As Convert: Set conv = New Convert
  Dim values As String: values = GetValueString(rs, fieldName, "~,;", quoteChar)
  Set GetValueCollection = conv.SplitCollection(values, "~,;")
  Set conv = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetValuesCollection
' Description:       Liefert aus einem RecordSet alle Werte aller Felder
'                    als String, getrennt durch delimiter in einer Collection.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:19:29
'
' Parameters :       rs (Object)
'                    quoteChar (String = "")
'--------------------------------------------------------------------------------
Public Function GetValuesCollection(ByRef rs As Object, _
                                    Optional ByVal quoteChar As String = "") As Collection

  Dim conv   As Convert: Set conv = New Convert
  Dim values As String: values = GetValuesString(rs, "~,;", quoteChar)
  Set GetValuesCollection = conv.SplitCollection(values, "~,;")
  Set conv = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RecordSetFieldsToString
' Description:       Liefert sämtliche Feldnamen als String, getrennt durch delimiter.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:20:26
'
' Parameters :       rs (Object)
'                    Delimiter (String = vbTab)
'--------------------------------------------------------------------------------
Public Function RecordSetFieldsToString(ByVal rs As Object, _
                                        Optional ByVal Delimiter As String = vbTab) As String

  Dim col As Collection: Set col = New Collection
  
  Dim i   As Integer

  For i = 0 To rs.Fields.count - 1
    col.Add CStr(rs.Fields.item(i).name)
  Next i
  
  RecordSetFieldsToString = mConv.JoinCollection(col, Delimiter)
  Set col = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RecordSetRowToString
' Description:       Liefert eine RecordSet-Zeile als String, getrennt durch delimiter.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:21:11
'
' Parameters :       rs (Object)
'                    Delimiter (String = vbTab)
'--------------------------------------------------------------------------------
Public Function RecordSetRowToString(ByVal rs As Object, _
                                     Optional ByVal Delimiter As String = vbTab) As String

  If rs.EOF Then
    RecordSetRowToString = ""
  Else
    Dim col As Collection: Set col = New Collection
    
    Dim i   As Integer

    For i = 0 To rs.Fields.count - 1
      col.Add CStr(rs.Fields.item(i).value)
    Next i

    rs.MoveNext
    
    RecordSetRowToString = mConv.JoinCollection(col, Delimiter)
    Set col = Nothing
  End If

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RecordSetToString
' Description:       Liefert einen RecordSet als String, getrennt durch crlf und delimiter.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:24:41
'
' Parameters :       rs (Object)
'                    includeFieldNames (Boolean)
'                    Delimiter (String = vbTab)
'--------------------------------------------------------------------------------
Public Function RecordSetToString(ByVal rs As Object, _
                                  ByVal includeFieldNames As Boolean, _
                                  Optional ByVal Delimiter As String = vbTab) As String

  Dim col As Collection: Set col = New Collection
  
  If includeFieldNames Then
    col.Add Me.RecordSetFieldsToString(rs, Delimiter)
  End If
  
  While Not rs.EOF
    col.Add Me.RecordSetRowToString(rs, Delimiter)
  Wend
  
  RecordSetToString = mConv.JoinCollection(col, vbCrLf)
  Set col = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RecordSetToStringArray
' Description:       Liefert einen RecordSet als String-Array, getrennt durch delimiter.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:25:40
'
' Parameters :       rs (Object)
'                    includeFieldNames (Boolean)
'                    Delimiter (String = vbTab) As String()
'--------------------------------------------------------------------------------
Public Function RecordSetToStringArray(ByVal rs As Object, _
                                       ByVal includeFieldNames As Boolean, _
                                       Optional ByVal Delimiter As String = vbTab) As String()

  Dim result As String: result = Me.RecordSetToString(rs, includeFieldNames, Delimiter)
  RecordSetToStringArray = Split(result, vbCrLf)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       InsertImageIntoDatabase
' Description:       Speichert eine Bilddatei in eine Datentabelle.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:26:40
'
' Parameters :       databaseName (String)
'                    tableName (String)
'                    blobFieldName (String)
'                    imageFileName (String)
'                    useFunctionOption (Boolean = True)
'--------------------------------------------------------------------------------
Public Function InsertImageIntoDatabase(ByVal databaseName As String, _
                                        ByVal tableName As String, _
                                        ByVal blobFieldName As String, _
                                        ByVal imageFileName As String, _
                                        Optional ByVal useFunctionOption As Boolean = True) As Long

  Const adTypeBinary = 1
  Const adOpenStatic = 3
  Const adLockOptimistic = 3

  Dim con As Object
  Set con = CreateObject("ADODB.Connection")
  '{Sollen erweiterte Optionen benutzt werden}
  con.Open Me.ConnectionString & IIf(useFunctionOption, ";OPTION=16427", "")
  
  '{Stream-Objekt öffnen}
  Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
  stream.Type = adTypeBinary
  stream.Open
  
  '{Recordset-Objekt öffnen}
  mLastExecutedStatement = "SELECT " & blobFieldName & " FROM " & databaseName & "." & tableName & " WHERE (1 = 0)"
  
  Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")
  rs.Open mLastExecutedStatement, con, adOpenStatic, adLockOptimistic

  '{Alle Bilddatei in DB schreiben}
  stream.LoadFromFile imageFileName
  rs.AddNew
  rs.Fields(blobFieldName).value = stream.Read
  rs.Update '{DB aktualisieren}
  
  '{Letzten eingeführten Autoinkremtenwert zurückgeben}
  Set rs = con.Execute("SELECT LAST_INSERT_ID() AS Id")
  InsertImageIntoDatabase = CLng(rs!Id)

  stream.Close '{Stream-Objekt schließen}
  rs.Close '{Recordset-Objekt schließen}
  con.Close '{Connection-Objekt schließen}

  Set stream = Nothing
  Set rs = Nothing
  Set con = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetImageFromDatabaseAsFileName
' Description:       Liefert ein Bild aus einer Datentabelle, speichert sie und
'                    gibt deren Dateinamen zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:31:55
'
' Parameters :       databaseName (String)
'                    tableName (String)
'                    idFieldName (String)
'                    blobFieldName (String)
'                    idValue (Long)
'                    useFunctionOption (Boolean = True)
'--------------------------------------------------------------------------------
Public Function GetImageFromDatabaseAsFileName(ByVal databaseName As String, _
                                               ByVal tableName As String, _
                                               ByVal idFieldName As String, _
                                               ByVal blobFieldName As String, _
                                               ByVal idValue As Long, _
                                               Optional ByVal useFunctionOption As Boolean = True) As String
                             
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
                             
  Dim con As Object: Set con = CreateObject("ADODB.Connection")
  '{Sollen erweiterte Optionen benutzt werden}
  con.Open Me.ConnectionString & IIf(useFunctionOption, ";OPTION=16427", "")
    
  '{Stream-Objekt öffnen}
  Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
  stream.Type = adTypeBinary
  stream.Open
  
  '{Recordset-Objekt öffnen}
  mLastExecutedStatement = "SELECT " & blobFieldName & " FROM " & databaseName & "." & tableName & " WHERE " & idFieldName & " = " & idValue
  Dim rs As Object: Set rs = con.Execute(mLastExecutedStatement)
  
  '{Ende Tempfilenamen erzeugen}
  '{Alle Datensätze durchlaufen}
  If Not rs.EOF Then
    '{Fehlerüberwachung ausschalten weil    }
    '{Bildinformationen in der DB fehlerhaft}
    '{oder nicht vorhanden sein könnten     }
    On Error Resume Next
    stream.Write rs.Fields(blobFieldName).value '& &HD9
    Dim isError As Boolean: isError = Err.number <> 0
    On Error GoTo 0
    
    Dim io       As FileIO: Set io = New FileIO
    Dim filename As String: filename = io.GetTempFileName("jpg")
    
    '{Wenn Bildinformationen OK}
    If Not isError Then
      '{Stream auf Platte zwischenspeichern}
      stream.SaveToFile filename, adSaveCreateOverWrite
      GetImageFromDatabaseAsFileName = filename
    End If
    
    Set io = Nothing
  End If

  stream.Close '{Stream-Objekt schließen}
  rs.Close '{Recordset-Objekt schließen}
  con.Close '{Connection-Objekt schließen}

  Set stream = Nothing
  Set rs = Nothing
  Set con = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetImageFromDatabase
' Description:       Liefert ein Bild aus einer Datentabelle und gibt es als
'                    Picture-Objekt zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:33:38
'
' Parameters :       databaseName (String)
'                    tableName (String)
'                    idFieldName (String)
'                    blobFieldName (String)
'                    idValue (Long)
'                    useFunctionOption (Boolean = True)
'--------------------------------------------------------------------------------
Public Function GetImageFromDatabase(ByVal databaseName As String, _
                                     ByVal tableName As String, _
                                     ByVal idFieldName As String, _
                                     ByVal blobFieldName As String, _
                                     ByVal idValue As Long, _
                                     Optional ByVal useFunctionOption As Boolean = True) As Picture
                             
  Dim filename As String: filename = GetImageFromDatabaseAsFileName(databaseName, tableName, idFieldName, blobFieldName, idValue, useFunctionOption)
                             
  If filename <> "" Then
    Set GetImageFromDatabase = LoadPicture(filename)
    '{Tempfile löschen}
    Dim io As FileIO: Set io = New FileIO
    io.DeleteFile filename, True
    Set io = Nothing
  End If

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       UpdateImageInDatabase
' Description:       Aktualisiert ein Bild in einer Datentabelle.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:34:32
'
' Parameters :       databaseName (String)
'                    tableName (String)
'                    idFieldName (String)
'                    blobFieldName (String)
'                    idValue (Long)
'                    imageFileName (String)
'                    useFunctionOption (Boolean = True)
'--------------------------------------------------------------------------------
Public Sub UpdateImageInDatabase(ByVal databaseName As String, _
                                 ByVal tableName As String, _
                                 ByVal idFieldName As String, _
                                 ByVal blobFieldName As String, _
                                 ByVal idValue As Long, _
                                 ByVal imageFileName As String, _
                                 Optional ByVal useFunctionOption As Boolean = True)
                              
  Const adUseClient = 3
  Const adTypeBinary = 1
  Const adOpenStatic = 3
  Const adLockOptimistic = 3
                            
  Dim con As Object: Set con = CreateObject("ADODB.Connection")
  '{Sollen erweiterte Optionen benutzt werden}
  Dim cs  As String: cs = Me.ConnectionString & IIf(useFunctionOption, ";OPTION=16427", "")
  con.ConnectionString = cs & IIf(InStr(LCase$(cs), "database=") = 0, ";DATABASE=" & databaseName, "")
  con.CursorLocation = adUseClient
  con.Open
    
  '{Stream-Objekt öffnen}
  Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
  stream.Type = adTypeBinary
  stream.Open
  stream.LoadFromFile imageFileName
  
  '{Recordset-Objekt öffnen}
  mLastExecutedStatement = "SELECT * FROM " & databaseName & "." & tableName & " WHERE " & idFieldName & " = " & idValue
  
  Dim rs As Object: Set rs = CreateObject("ADODB.RecordSet")
  rs.Open mLastExecutedStatement, con, adOpenStatic, adLockOptimistic
  
  rs.Fields(blobFieldName).value = stream.Read
  rs.Update
  
  rs.Close
  stream.Close
  con.Close

  Set stream = Nothing
  Set rs = Nothing
  Set con = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ResolveStatementParameters
' Description:       Führt eine Parameterauflösung durch und liefert das
'                    Statement zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       28.01.2015-11:38:22
'
' Parameters :       statement (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ResolveStatementParameters(ByVal statement As String, _
                                           ParamArray parameters() As Variant) As Variant

  ResolveStatementParameters = mResolver.ReplaceParameters(statement, parameters)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ResolveStatementParameters
' Description:       Führt eine Parameterauflösung durch und liefert das
'                    Statement zurück (nicht übergebene Parameter werden nicht
'                    abgefragt).
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       21.04.2015-11:38:22
'
' Parameters :       statement (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function ResolveStatementParametersWithoutMissingParameters(ByVal statement As String, _
                                           ParamArray parameters() As Variant) As Variant

  ResolveStatementParametersWithoutMissingParameters = mResolver.ReplaceParametersWithoutMissingParameters(statement, parameters)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetServerNameFromConnectionString
' Description:       Ermittelt den Servernamen aus dem Connectionstring
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       9.30.2015-16:15:27
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function GetServerNameFromConnectionString() As String
'Beispiel Connectionstring: DRIVER={MYSQL ODBC 5.1 Driver};SERVER=sql-unittest;UID=apps;PWD=***;PORT=3306;
    
    Dim regexp As Object: Set regexp = CreateObject("VBScript.RegExp")
        regexp.Pattern = "(?:;\w+=)(.+)(?:;U)"
        GetServerNameFromConnectionString = regexp.Execute(ConnectionString).item(0).SubMatches(0)
    Set regexp = Nothing
End Function

Public Function GetServerNameFromConnectionString2(ByVal strConnectionString) As String
'Beispiel Connectionstring: DRIVER={MYSQL ODBC 5.1 Driver};SERVER=sql-unittest;UID=apps;PWD=***;PORT=3306;
    
    Dim regexp As Object: Set regexp = CreateObject("VBScript.RegExp")
        regexp.Pattern = "(?:;\w+=)(.+)(?:;U)"
        GetServerNameFromConnectionString2 = regexp.Execute(strConnectionString).item(0).SubMatches(0)
    Set regexp = Nothing
End Function
