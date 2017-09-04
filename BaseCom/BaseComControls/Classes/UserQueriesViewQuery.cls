VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserQueriesViewQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'{Eigenschaften der Klasse}
Private mstrQueryID As String
Private mstrOwnerID As String
Private mstrPrg As String
Private mstrQueryFolderFID As String
Private mstrName As String
Private mstrStatement As String
Private mstrComment As String
Private mQueryType As eqQueryType
Private mstrOwner As String
Private mstrLastExecutedStatement As String
Private mblnResolveParamterBeforeExecute As Boolean
'{Ende Eigenschaften der Klasse}


'{Auflistung für Klausuren}
Private mcolParameters As Collection
Private mcolLastResolvedParameters As Collection


'{---------- Konstruktor und Destrunktor der Klasse ----------}
Private Sub Class_Initialize()
  '{Konstruktor der Klasse}
  Set mcolParameters = New Collection
  Set mcolLastResolvedParameters = New Collection
  
  Me.QueryID = "-1"
  Me.OwnerID = "-1"
  Me.Prg = ""
  Me.QueryFolderFID = "-1"
  Me.Name = "NeueAbfrage"
  
  Me.Statement = _
  "SELECT" & vbCrLf & vbTab _
  & "*" & vbCrLf & vbCrLf _
  & "FROM" & vbCrLf & vbTab & vbCrLf & vbCrLf _
  & "WHERE" & vbCrLf & vbTab _
  & "(1 = 1)" & vbCrLf & vbTab _
  & "[WHERE]"
  
  Me.Comment = ""
  Me.QueryType = eqqtUserQuery
  Me.Owner = "o.A."
  mstrLastExecutedStatement = ""
End Sub

Private Sub Class_Terminate()
  '{Destruktor der Klasse}
  Set mcolParameters = Nothing
  Set mcolLastResolvedParameters = Nothing

End Sub
'{---------- Ende Konstruktor und Destrunktor der Klasse ----------}


'{---------- Zugriffsmethoden der Eigenschaften der Klasse ----------}
Public Property Get QueryID() As String
  QueryID = mstrQueryID
End Property

Public Property Let QueryID(ByVal strQueryID As String)
  mstrQueryID = strQueryID
End Property

Public Property Get OwnerID() As String
  OwnerID = mstrOwnerID
End Property

Public Property Let OwnerID(ByVal strOwnerID As String)
  mstrOwnerID = strOwnerID
End Property

Public Property Get Prg() As String
  Prg = mstrPrg
End Property

Public Property Let Prg(ByVal strPrg As String)
  mstrPrg = strPrg
End Property

Public Property Get QueryFolderFID() As String
  QueryFolderFID = mstrQueryFolderFID
End Property

Public Property Let QueryFolderFID(ByVal strQueryFolderFID As String)
  mstrQueryFolderFID = strQueryFolderFID
End Property

Public Property Get Name() As String
  Name = mstrName
End Property

Public Property Let Name(ByVal strName As String)
  mstrName = strName
End Property

Public Property Get Statement() As String
  Statement = mstrStatement
End Property

Public Property Let Statement(ByVal strStatement As String)
  mstrStatement = strStatement
End Property

Public Property Get Comment() As String
  Comment = mstrComment
End Property

Public Property Let Comment(ByVal strComment As String)
  mstrComment = strComment
End Property

Public Property Get QueryType() As eqQueryType
  QueryType = mQueryType
End Property

Public Property Let QueryType(ByVal QueryType As eqQueryType)
  mQueryType = QueryType
End Property

Public Property Get Owner() As String
  Owner = mstrOwner
End Property

Public Property Let Owner(ByVal strOwner As String)
  mstrOwner = strOwner
End Property

Public Property Get LastExecutedStatement() As String
  LastExecutedStatement = mstrLastExecutedStatement
End Property


Public Property Get LastResolvedParameters() As Collection
  Set LastResolvedParameters = mcolLastResolvedParameters
End Property

'{---------- Ende Zugriffsmethoden der Eigenschaften der Klasse ----------}


'{---------- Private Methoden der Klasse ----------}
Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub
'{---------- Ende Private Methoden der Klasse ----------}


'{---------- Öffentliche Methoden der Klasse ----------}
Public Function AddParameter(ByRef aPar As UserQueriesViewParameter _
, Optional ByVal blnTemporary As Boolean = False) As Long

On Error GoTo errLabel:

  AddParameter = 0

  If Not blnTemporary Then
    AddParameter = aPar.SaveParameter
  End If

  If AddParameter = 0 Then
    mcolParameters.Add aPar, "PAR" & aPar.ParameterID
  End If
  Exit Function

errLabel:
  AddParameter = -1
  ShowError "AddParameter"
  Exit Function
End Function

Public Function RemoveParameter _
(ByRef aPar As UserQueriesViewParameter _
, Optional ByVal blnTemporary As Boolean = False) As Long

On Error GoTo errLabel:

  RemoveParameter = 0

  mcolParameters.Remove "PAR" & aPar.ParameterID

  If Not blnTemporary Then
    RemoveParameter = aPar.DeleteParameter
  End If
  Exit Function

errLabel:
  RemoveParameter = -1
  ShowError "RemoveParameter"
  Exit Function
End Function

Public Function GetParameter _
(ByVal strParameterID As String) As UserQueriesViewParameter

On Error GoTo errLabel:

  strParameterID = "PAR" & strParameterID
  Set GetParameter = mcolParameters.Item(strParameterID)
  Exit Function

errLabel:
  ShowError "GetParameter"
  Exit Function
End Function

Public Function GetParameters() As Collection
On Error GoTo errLabel:

  Set GetParameters = mcolParameters
  Exit Function

errLabel:
  ShowError "GetParameters"
  Exit Function
End Function

Public Function GetParametersDB() As Long

On Error GoTo errLabel:

  GetParametersDB = 0
  
  Set mcolParameters = New Collection

  Dim strSQL As String
  strSQL = strSQL & "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "p.ParameterID" & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_parameter p" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "p.QueryFID = '" & Me.QueryID & "'"

  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(strSQL)
  
  While Not rs.EOF
    Dim aPar As UserQueriesViewParameter: Set aPar = New UserQueriesViewParameter
    aPar.GetParameter rs!ParameterID
    Me.AddParameter aPar, True
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  Exit Function

errLabel:
  GetParametersDB = -1
  ShowError "GetParametersDB"
  Exit Function
End Function

Public Function GetQuery(ByVal strQueryID As String) As Long

On Error GoTo errLabel:

  GetQuery = 0
  
  Dim strSQL As String
  strSQL = strSQL & "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "q.QueryID," & vbCrLf & vbTab
  strSQL = strSQL & "q.UserFID," & vbCrLf & vbTab
  strSQL = strSQL & "IF(u.Nachname IS NULL,'gelöscht',CONCAT(u.Nachname,', ',u.Vorname)) AS Owner," & vbCrLf & vbTab
  strSQL = strSQL & "q.Prg," & vbCrLf & vbTab
  strSQL = strSQL & "q.QueryFolderFID," & vbCrLf & vbTab
  strSQL = strSQL & "q.Name," & vbCrLf & vbTab
  strSQL = strSQL & "q.Statement," & vbCrLf & vbTab
  strSQL = strSQL & "q.Comment," & vbCrLf & vbTab
  strSQL = strSQL & "q.QueryType" & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_user_queries q" & vbCrLf & vbTab
  strSQL = strSQL & "LEFT JOIN datapool.t_personen u" & vbCrLf & vbTab
  strSQL = strSQL & "ON q.UserFID = u.PersonenID" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "q.QueryID = '" & strQueryID & "'"

  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(strSQL)

  Set mcolParameters = New Collection
  
  If Not rs.EOF Then
    Me.QueryID = rs!QueryID & ""
    Me.OwnerID = rs!UserFID & ""
    Me.Prg = rs!Prg & ""
    Me.Name = rs!Name & ""
    Me.Statement = rs!Statement & ""
    Me.Comment = rs!Comment & ""
    Me.QueryType = rs!QueryType
    Me.Owner = rs!Owner
    Me.QueryFolderFID = rs!QueryFolderFID & ""
  End If
  BaseToolKit.Database.CloseRecordSet rs
  Exit Function

errLabel:
  GetQuery = -1
  ShowError "GetQuery"
  Exit Function
End Function

Public Function SaveQuery() As Long

On Error GoTo errLabel:

  SaveQuery = 0

  If Me.QueryID = "-1" Then '{Neuanlegen}
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbCrLf & vbTab
    strSQL = strSQL & "queries.t_user_queries" & vbCrLf & vbTab
    strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "UserFID," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Prg," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "QueryFolderFID," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Name," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Statement," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Comment," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "QueryType" & vbCrLf & vbTab
    strSQL = strSQL & ")" & vbCrLf
    strSQL = strSQL & "VALUES" & vbCrLf & vbTab
    strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.OwnerID & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.Prg & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.QueryFolderFID & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Name) & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'[statement]'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Comment) & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.QueryType & "'" & vbCrLf & vbTab
    strSQL = strSQL & ")"
    
    Me.QueryID = BaseToolKit.Database.ExecuteNonQuery _
    (strSQL, "statement:" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Statement))
  Else '{Aktualisieren}
  
    strSQL = strSQL & "UPDATE" & vbCrLf & vbTab
    strSQL = strSQL & "queries.t_user_queries" & vbCrLf
    strSQL = strSQL & "SET" & vbCrLf & vbTab
    strSQL = strSQL & "UserFID = '" & Me.OwnerID & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Prg = '" & Me.Prg & "'," & vbCrLf & vbTab
    strSQL = strSQL & "QueryFolderFID = '" & Me.QueryFolderFID & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Name = '" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Name) & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Statement = '[statement]'," & vbCrLf & vbTab
    strSQL = strSQL & "Comment = '" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Comment) & "'," & vbCrLf & vbTab
    strSQL = strSQL & "QueryType = '" & Me.QueryType & "'" & vbCrLf
    strSQL = strSQL & "WHERE" & vbCrLf & vbTab
    strSQL = strSQL & "QueryID = '" & Me.QueryID & "'"
    
    BaseToolKit.Database.ExecuteNonQuery strSQL _
    , "statement:" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Statement)
  End If
  Exit Function

errLabel:
  SaveQuery = -1
  ShowError "SaveQuery"
  Exit Function
End Function

Public Function DeleteQuery() As Long

On Error GoTo errLabel:

  DeleteQuery = 0

  Me.GetParametersDB

  Dim aPar As UserQueriesViewParameter
  For Each aPar In mcolParameters
    DeleteQuery = aPar.DeleteParameter
  Next aPar

  Dim strSQL As String
  strSQL = strSQL & "DELETE FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_user_queries" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "QueryID = '" & Me.QueryID & "'"

  BaseToolKit.Database.ExecuteNonQuery strSQL
  Exit Function

errLabel:
  DeleteQuery = -1
  ShowError "DeleteQuery"
  Exit Function
End Function

Private Function GetOperatorSyntax _
(ByRef aParameter As UserQueriesViewParameter) As String

Dim strAusdruck As String
Dim astrAusdruck() As String
Dim astrFields() As String
Dim i As Integer
Dim strTimeRay As String

On Error GoTo errLabel:

  GetOperatorSyntax = "<error name=undeferror>"

  If aParameter.Field = "" Then
    MsgBox "Bitte geben Sie ein Feld für den Parameter '" _
    & aParameter.Name & "' an!", 16, "GetOperatorSyntax"
    Exit Function
  End If

  astrAusdruck = Split(aParameter.value, ",")
  
  For i = LBound(astrAusdruck) To UBound(astrAusdruck)
    astrAusdruck(i) = Replace(astrAusdruck(i), "<escape name=comma>", ",")
  Next i

  Select Case aParameter.Operator
  Case eParameterOperator.pveEqual  'entspricht
    strAusdruck = "AND (" & aParameter.Field & " = '" _
    & astrAusdruck(0) & "')" & vbCrLf & vbTab
    
  Case eParameterOperator.pveNotEqual 'entspricht nicht
    strAusdruck = "AND (" & aParameter.Field & " <> '" _
    & astrAusdruck(0) & "')" & vbCrLf & vbTab
    
  Case eParameterOperator.pveMemberOf 'ist einer von
    For i = LBound(astrAusdruck) To UBound(astrAusdruck)
      astrAusdruck(i) = "'" & Trim(astrAusdruck(i)) & "'"
    Next i
    strAusdruck = "AND (" & aParameter.Field & " IN (" _
    & Join(astrAusdruck, ",") & "))" & vbCrLf & vbTab
  
  Case eParameterOperator.pveNotMemberOf  'ist nicht einer von
    For i = LBound(astrAusdruck) To UBound(astrAusdruck)
      astrAusdruck(i) = "'" & Trim(astrAusdruck(i)) & "'"
    Next i
    strAusdruck = "AND (" & aParameter.Field & " NOT IN (" _
    & Join(astrAusdruck, ",") & "))" & vbCrLf & vbTab
    
  Case eParameterOperator.pveLesser 'ist kleiner als
    strAusdruck = "AND (" & aParameter.Field & " < '" _
    & astrAusdruck(0) & "')" & vbCrLf & vbTab
    
  Case eParameterOperator.pveLesserEqual 'ist kleiner oder gleich
    strAusdruck = "AND (" & aParameter.Field & " <= '" _
    & astrAusdruck(0) & "')" & vbCrLf & vbTab
    
  Case eParameterOperator.pveGreater 'ist größer als
    strAusdruck = "AND (" & aParameter.Field & " > '" _
    & astrAusdruck(0) & "')" & vbCrLf & vbTab
    
  Case eParameterOperator.pveGreaterEqual '"ist größer oder gleich"
    strAusdruck = "AND (" & aParameter.Field & " >= '" _
    & astrAusdruck(0) & "')" & vbCrLf & vbTab

    
  Case eParameterOperator.pveBetween 'liegt zwischen
    If LBound(astrAusdruck) = UBound(astrAusdruck) Then
      MsgBox "Fehlende Argumente!", 16, aParameter.Name
      Exit Function
    End If
        
    strAusdruck = "AND (" & aParameter.Field _
    & " BETWEEN '" & Trim(astrAusdruck(0)) & "' AND '" _
    & Trim(astrAusdruck(1)) & "')" & vbCrLf & vbTab
    
  Case eParameterOperator.pveNotBetween 'liegt nicht zwischen
    If LBound(astrAusdruck) = UBound(astrAusdruck) Then
      MsgBox "Fehlende Argumente!", 16, aParameter.Name
      Exit Function
    End If
    
    strAusdruck = "AND (" & aParameter.Field _
    & " NOT BETWEEN '" & Trim(astrAusdruck(0)) & "' AND '" _
    & Trim(astrAusdruck(1)) & "')" & vbCrLf & vbTab
    
    
  Case eParameterOperator.pveLike 'wie
    strAusdruck = "AND (" & aParameter.Field & " LIKE '" _
    & astrAusdruck(0) & "')" & vbCrLf & vbTab
  
  Case eParameterOperator.pveIntervalMinute, eParameterOperator.pveIntervalHour
    If LBound(astrAusdruck) = UBound(astrAusdruck) Then
      MsgBox "Fehlende Argumente!", 16, aParameter.Name
      Exit Function
    End If

    If InStr(aParameter.Field, ",") = 0 Then
      MsgBox "Fehlendes Bezugsfeld!", 16, aParameter.Name
      Exit Function
    End If

    astrFields = Split(aParameter.Field, ",")
    
    Select Case aParameter.Operator
    Case eParameterOperator.pveIntervalMinute
      strTimeRay = BaseToolKit.FileSystem.SqlGroupFile.TimeSpanToTimeRay _
      (astrAusdruck(0), astrAusdruck(1), "00:01")
    Case eParameterOperator.pveIntervalHour
      strTimeRay = BaseToolKit.FileSystem.SqlGroupFile.TimeSpanToTimeRay _
      (astrAusdruck(0), astrAusdruck(1), "01:00")
    End Select

    strAusdruck = "AND (((" & astrFields(0) & " IN (" & strTimeRay & "))" _
    & " OR (" & astrFields(1) & " IN(" & strTimeRay & ")))" _
    & " OR ((" & astrFields(0) & " <= '" & astrAusdruck(0) & "')" _
    & " AND (" & astrFields(1) & " >= '" & astrAusdruck(1) & "')))"

  Case eParameterOperator.pveIsNULL '"ist NULL"
    strAusdruck = "AND (" & aParameter.Field & " IS NULL)"

  Case eParameterOperator.pveIsNotNULL '"ist nicht NULL"
    strAusdruck = "AND (" & aParameter.Field & " IS NOT NULL)"

  End Select
  
  strAusdruck = Replace(strAusdruck, "= '#NULL#'", "IS NULL")
  strAusdruck = Replace(strAusdruck, "<> '#NULL#'", "IS NOT NULL")
  GetOperatorSyntax = strAusdruck
  Exit Function
  
errLabel:
  ShowError "GetOperatorSyntax"
  Exit Function
End Function

Public Function ResolveParameter() As String

On Error GoTo errLabel:

  Set mcolLastResolvedParameters = New Collection
  Me.GetParametersDB
    
  Dim aParameter As UserQueriesViewParameter
  For Each aParameter In mcolParameters
    
    If (aParameter.value = "") Then
      Select Case aParameter.Operator
      Case pveNoParameter, pveIsNULL, pveIsNotNULL
      Case Else
        aParameter.value = InputBox _
        ("Bitte neuen Parameterwert eingeben:" & Chr(13) _
        & "(Mehrere Parameterwerte bitte durch Kommata trennen!)" _
        , aParameter.Name & " " & aParameter.OperatorText _
        & " ...", aParameter.value)
        
        If aParameter.value = "" Then
          ResolveParameter = "-1"
          Exit Function
        End If
        
      End Select
    End If
    
    Dim strOperatorSyntax As String: strOperatorSyntax = GetOperatorSyntax(aParameter)
    
    On Error Resume Next
    mcolLastResolvedParameters.Add ResolveParameter & strOperatorSyntax, aParameter.Name
    Dim blnOK As Boolean: blnOK = (Err.Number = 0)
    On Error GoTo 0
    
    If blnOK Then
      ResolveParameter = ResolveParameter & strOperatorSyntax & vbCrLf & vbTab
    Else
      MsgBox "Parametername '" & aParameter.Name & "' nicht eindeutig!" & Chr(13) & Chr(13) _
      & "Bitte vergeben Sie eindeutige Parameternamen!" _
      , 16, "ResolveParameter"
      ResolveParameter = ResolveParameter & "<error name=undeferror>"
    End If
  
  Next aParameter
  Exit Function
  
errLabel:
  ShowError "ResolveParameter"
  Exit Function
End Function

Public Property Get ResolvedQuery() As String

  Dim strQuery As String
  Dim intPos As Integer
  Dim strParameter As String
  
  ResolvedQuery = ""
  
  intPos = InStr(LCase(Me.Statement), "[where]")
  
  If intPos > 0 Then
  
    strParameter = ResolveParameter
      
    If InStr(strParameter, "<error name=undeferror>") > 0 Then Exit Property
    
    If strParameter = "-1" Then Exit Property
    
    strQuery = Mid(Me.Statement, 1, intPos - 1)
    
    strQuery = strQuery & strParameter
    
    strQuery = strQuery & Mid(Me.Statement, intPos + Len("[where]"))
    
  Else
    If Me.CountParameters > 0 Then
      MsgBox "Ihre angelegten Parameter werden nur ausgewertet," & Chr(13) _
      & "wenn im SQL-Statement der Abfrage folgende Zeichenfolge existiert:" _
      & Chr(13) & Chr(13) & "WHERE" & vbCrLf & vbTab & "(1 = 1)" _
      & vbCrLf & vbTab & "[WHERE]", 48, "ExecuteQuery"
    End If
    strQuery = Me.Statement
  End If

  ResolvedQuery = strQuery
End Property

Public Function ExecuteQuery(Optional ByVal lngLimit As Long = 0) As Object

On Error GoTo errLabel:

  Set ExecuteQuery = Nothing
    
  Dim strQuery As String: strQuery = ResolvedQuery
  
  If lngLimit > 0 Then
    strQuery = strQuery & vbCrLf & "LIMIT " & lngLimit
  End If
  
  strQuery = BaseToolKit.Database.ResolveStatementParameters(strQuery)
    
  Dim astrQuery() As String: astrQuery = Split(strQuery, "-- {nextSQL}")
    
  Dim varQuery As Variant
  For Each varQuery In astrQuery
    Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(varQuery)
  Next varQuery
  
  Set ExecuteQuery = rs
  mstrLastExecutedStatement = strQuery
  
  Me.GetQuery Me.QueryID
  mblnResolveParamterBeforeExecute = False
  Exit Function
  
errLabel:
  If Err.Number <> ErrorNumbers.ParameterInputCanceledError Then
    ShowError "ExecuteQuery"
  End If
  Exit Function
End Function

Public Function CountParameters() As Long

On Error GoTo errLabel:

  Dim strSQL As String
  strSQL = "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "COUNT(1) AS CountParameters" & vbCrLf & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_parameter p" & vbCrLf & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "p.QueryFID = " & Me.QueryID
  CountParameters = BaseToolKit.Database.ExecuteScalar(strSQL)
  Exit Function
  
errLabel:
  ShowError "CountParameters"
  Exit Function
End Function

Public Function GetTableAlias(ByVal strTable As String) As String
Dim lngPos As Long
Dim strSQL As String
Dim astrSQL() As String
Dim astrParts() As String
Dim x As Long
  
On Error GoTo errLabel:

  astrSQL = Split(SeparateTables, "join", , vbTextCompare)
    
  For x = LBound(astrSQL) To UBound(astrSQL)
    lngPos = InStr(1, astrSQL(x), " on ", vbTextCompare)
    If lngPos > 0 Then astrSQL(x) = Mid(astrSQL(x), 1, lngPos - 1)
    
    astrParts = Split(Trim(astrSQL(x)), " ")
           
    If astrParts(0) = strTable Then
      Select Case True
      Case UBound(astrParts) = 0
        GetTableAlias = astrParts(0)
      Case (UBound(astrParts) = 1) And (Trim(astrParts(1)) = "")
        GetTableAlias = astrParts(0)
      Case (UBound(astrParts) = 1) And (Trim(astrParts(1)) <> "")
        GetTableAlias = astrParts(1)
      End Select
      Exit Function
    End If
  
  Next x
  Exit Function
  
errLabel:
  ShowError "GetTableAlias"
  Exit Function
End Function

Private Function SeparateTables() As String
Dim lngPos As Long
Dim lngPos2 As Long
Dim strSQL As String
    
On Error GoTo errLabel:

  strSQL = Me.Statement
  
  lngPos = InStr(strSQL, "/*")
  
  While lngPos > 0
    
    lngPos2 = InStr(strSQL, "*/")
    
    '{Hier könnte Kommentarende fehlen}
    
    strSQL = Mid(strSQL, 1, lngPos - 1) & Mid(strSQL, lngPos2 + 2)
  
    lngPos = InStr(strSQL, "/*")
    
  Wend
  
  strSQL = Replace(strSQL, vbTab, " ")
  lngPos = InStr(strSQL, "-- ")
  
  
  While lngPos > 0
    
    lngPos2 = InStr(Mid(strSQL, lngPos), vbCrLf)
    
    '{Hier könnte Kommentarende fehlen}
    strSQL = Mid(strSQL, 1, lngPos - 1) & Mid(strSQL, lngPos + lngPos2 + 2)
  
    lngPos = InStr(strSQL, "-- ")
  Wend
   
  strSQL = Replace(strSQL, vbCrLf, " ")
  
  lngPos = InStr(LCase(strSQL), " from ")
  
  If lngPos = 0 Then
    MsgBox "FROM-Zweig fehlt!", 48, ""
    Exit Function
  End If
  
  strSQL = Mid(strSQL, lngPos + 6)
  
  lngPos = InStr(LCase(strSQL), " where ")
  If lngPos > 0 Then strSQL = Mid(strSQL, 1, lngPos)
  
  lngPos = InStr(LCase(strSQL), " group by ")
  If lngPos > 0 Then strSQL = Mid(strSQL, 1, lngPos)
  
  lngPos = InStr(LCase(strSQL), " having ")
  If lngPos > 0 Then strSQL = Mid(strSQL, 1, lngPos)
  
  lngPos = InStr(LCase(strSQL), " order by ")
  If lngPos > 0 Then strSQL = Mid(strSQL, 1, lngPos)
  
  strSQL = Replace(strSQL, "inner", "", , , vbTextCompare)
  strSQL = Replace(strSQL, "left", "", , , vbTextCompare)
  strSQL = Replace(strSQL, "right", "", , , vbTextCompare)
  strSQL = Replace(strSQL, "outer", "", , , vbTextCompare)
  strSQL = Replace(strSQL, "union", "", , , vbTextCompare)
  SeparateTables = strSQL
  Exit Function
  
errLabel:
  ShowError "SeparateTables"
  Exit Function
End Function

Public Function GetTablesInSQL() As Collection
Dim lngPos As Long
Dim strSQL As String
Dim astrSQL() As String
Dim astrParts() As String
Dim x As Long
  
On Error GoTo errLabel:

  Set GetTablesInSQL = New VBA.Collection

  astrSQL = Split(SeparateTables, "join", , vbTextCompare)
    
  For x = LBound(astrSQL) To UBound(astrSQL)
    lngPos = InStr(1, astrSQL(x), " on ", vbTextCompare)
    If lngPos > 0 Then astrSQL(x) = Mid(astrSQL(x), 1, lngPos - 1)
    
    astrParts = Split(Trim(astrSQL(x)), " ")
       
    GetTablesInSQL.Add astrParts(0)
  Next x
  Exit Function
  
errLabel:
  ShowError "GetTablesInSQL"
  Exit Function
End Function

Public Function ResolveParamterBeforeExecute _
(ByVal strParameter As String _
, ByVal strValue As String)

On Error GoTo errLabel:

  Me.Statement = Replace _
  (Me.Statement, "[" & strParameter & "]", strValue, , , vbTextCompare)
  mblnResolveParamterBeforeExecute = True
  Exit Function
  
errLabel:
  ShowError "ResolveParameterBeforeExecute"
  Exit Function
End Function
'{---------- Ende Öffentliche Methoden der Klasse ----------}


