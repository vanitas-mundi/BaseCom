VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserQueriesViewQueries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eqQueryType
  eqqtUserQuery = 0
  eqqtOpenQuery = 1
  eqqtAdminQuery = 2
  eqqtPresentQuery = 3
End Enum

'{Eigenschaften der Klasse}
Private mstrUserID As String
Private mstrPrg As String
Private mTypeOfQueries As eqQueryType
'{Ende Eigenschaften der Klasse}


'{Auflistung für Klausuren}
Private mcolQueries As Collection


'{---------- Konstruktor und Destrunktor der Klasse ----------}
Private Sub Class_Initialize()
  '{Konstruktor der Klasse}
  Set mcolQueries = New Collection
  
  mstrUserID = -1
  mstrPrg = ""
  mTypeOfQueries = eqqtUserQuery
  
End Sub

Private Sub Class_Terminate()
  '{Destruktor der Klasse}
  Set mcolQueries = Nothing
End Sub
'{---------- Ende Konstruktor und Destrunktor der Klasse ----------}


'{---------- Zugriffsmethoden der Eigenschaften der Klasse ----------}
Public Property Get UserID() As String
  UserID = mstrUserID
End Property

Public Property Get Prg() As String
  Prg = mstrPrg
End Property

Public Property Get TypeOfQueries() As eqQueryType
  TypeOfQueries = mTypeOfQueries
End Property
'{---------- Ende Zugriffsmethoden der Eigenschaften der Klasse ----------}


'{---------- Private Methoden der Klasse ----------}
Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub
'{---------- Ende Private Methoden der Klasse ----------}


'{---------- Öffentliche Methoden der Klasse ----------}
Public Function AddQuery(ByRef aQuery As UserQueriesViewQuery _
, Optional ByVal blnTemporary As Boolean = False) As Long

On Error GoTo errLabel:

  AddQuery = 0

  If Not blnTemporary Then
    AddQuery = aQuery.SaveQuery
  End If

  If AddQuery = 0 Then
    mcolQueries.Add aQuery, "Q" & aQuery.QueryID
  End If
  Exit Function

errLabel:
  AddQuery = -1
  ShowError "AddQuery"
  Exit Function
End Function

Public Function RemoveQuery(ByRef aQuery As UserQueriesViewQuery _
, Optional ByVal blnTemporary As Boolean = False) As Long

On Error GoTo errLabel:

  RemoveQuery = 0

  mcolQueries.Remove "Q" & aQuery.QueryID

  If Not blnTemporary Then
    RemoveQuery = aQuery.DeleteQuery
  End If
  Exit Function

errLabel:
  RemoveQuery = -1
  ShowError "RemoveQuery"
  Exit Function
End Function

Public Function GetQuery _
(ByVal strQueryID As String) As UserQueriesViewQuery

On Error GoTo errLabel:

  strQueryID = "Q" & strQueryID
  Set GetQuery = mcolQueries.Item(strQueryID)
  Exit Function

errLabel:
  ShowError "GetQuery"
  Exit Function
End Function

Public Function GetQueries() As Collection
On Error GoTo errLabel:

  Set GetQueries = mcolQueries
  Exit Function

errLabel:
  ShowError "GetQueries"
  Exit Function
End Function

Public Function GetQueriesDB _
(Optional ByVal Prg As String = "" _
, Optional ByVal UserID As String = -1 _
, Optional ByVal QueryType As eqQueryType = eqqtUserQuery _
, Optional ByVal strQueryFolderID As String = "0") As Long

On Error GoTo errLabel:

  GetQueriesDB = 0

  Set mcolQueries = New Collection

  Dim strSQL As String
  strSQL = strSQL & "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "q.QueryID" & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_user_queries q" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  
  strSQL = strSQL & "(q.QueryType = '" & QueryType & "')" & vbCrLf & vbTab
  
  strSQL = strSQL & "AND (q.QueryFolderFID = '" & strQueryFolderID & "')" & vbCrLf & vbTab
  
  If UserID <> -1 Then
    strSQL = strSQL & "AND (q.UserFID = '" & UserID & "')" & vbCrLf & vbTab
  End If
  
  If Prg <> "" Then
    strSQL = strSQL & "AND (q.Prg = '" & Prg & "')" & vbCrLf & vbTab
  End If
    
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(strSQL)

  While Not rs.EOF
    Dim aQuery As UserQueriesViewQuery: Set aQuery = New UserQueriesViewQuery
    aQuery.GetQuery rs!QueryID
    Me.AddQuery aQuery, True
    rs.MoveNext
  Wend
  
  BaseToolKit.Database.CloseRecordSet rs
  
  mstrPrg = Prg
  mstrUserID = UserID
  mTypeOfQueries = QueryType
  Exit Function

errLabel:
  GetQueriesDB = -1
  ShowError "GetQueriesDB"
  Exit Function
End Function
'{---------- Ende Öffentliche Methoden der Klasse ----------}





