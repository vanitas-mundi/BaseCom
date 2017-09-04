VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserQueriesViewQueryFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'{Eigenschaften der Klasse}
Public FolderID As String
Public ParentFolderID As String
Public FolderName As String
Public Prg As String
Public OwnerID As String
Public Owner As String
Public QueryType As eqQueryType
Public Image As String
Public Comment As String
'{Ende Eigenschaften der Klasse}

'{Auflistung für Klausuren}
Public Queries As UserQueriesViewQueries
Public QueryFolders As UserQueriesViewQueryFolders

Private Sub Class_Initialize()
  Set Queries = New UserQueriesViewQueries
  Set QueryFolders = New UserQueriesViewQueryFolders

  Me.FolderID = "-1"
  Me.ParentFolderID = "-1"
  Me.FolderName = "Neuer Ordner"
  Me.Prg = ""
  Me.OwnerID = "-1"
  Me.Owner = ""
  Me.QueryType = eqqtUserQuery
  Me.Image = "QueryFolder"
  Me.Comment = ""

End Sub

Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub

Public Function GetQueryFolder(ByVal strQueryFolderID As String) As Long

On Error GoTo errLabel:

  GetQueryFolder = 0

  Dim strSQL As String
  strSQL = strSQL & "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "qf.QueryFolderID," & vbCrLf & vbTab
  strSQL = strSQL & "qf.ParentFolderFID," & vbCrLf & vbTab
  strSQL = strSQL & "qf.Prg," & vbCrLf & vbTab
  strSQL = strSQL & "u.PersonenID AS OwnerFID," & vbCrLf & vbTab
  strSQL = strSQL & "qf.QueryType," & vbCrLf & vbTab
  strSQL = strSQL & "qf.FolderName," & vbCrLf & vbTab
  strSQL = strSQL & "qf.Image," & vbCrLf & vbTab
  strSQL = strSQL & "qf.Comment," & vbCrLf & vbTab
  strSQL = strSQL & "IF(u.Nachname IS NULL,'gelöscht',CONCAT(u.Nachname,', ',u.Vorname)) AS Owner" & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_query_folders qf" & vbCrLf & vbTab
  strSQL = strSQL & "LEFT JOIN datapool.t_personen u" & vbCrLf & vbTab
  strSQL = strSQL & "ON qf.OwnerFID = u.PersonenID" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "qf.QueryFolderID = '" & strQueryFolderID & "'"

  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(strSQL)

  If Not rs.EOF Then
    Set Me.Queries = New UserQueriesViewQueries
    Set Me.QueryFolders = New UserQueriesViewQueryFolders
  
  
    Me.FolderID = rs!QueryFolderID & ""
    Me.ParentFolderID = rs!ParentFolderFID & ""
    Me.Prg = rs!Prg & ""
    Me.OwnerID = rs!OwnerFID & ""
    Me.Owner = rs!Owner & ""
    Me.QueryType = rs!QueryType & ""
    Me.FolderName = rs!FolderName & ""
    Me.Image = rs!Image & ""
    Me.Comment = rs!Comment & ""
  End If
  
  BaseToolKit.Database.CloseRecordSet rs
  Exit Function

errLabel:
  GetQueryFolder = -1
  ShowError "GetQueryFolder"
  Exit Function
End Function

Public Function SaveQueryFolder() As Long

On Error GoTo errLabel:

  SaveQueryFolder = 0

  If Me.FolderID = "-1" Then '{Neuanlegen}
    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbCrLf & vbTab
    strSQL = strSQL & "queries.t_query_folders" & vbCrLf & vbTab
    strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
    
    strSQL = strSQL & "ParentFolderFID," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Prg," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "OwnerFID," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "QueryType," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "FolderName," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Image," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Comment" & vbCrLf & vbTab & vbTab
        
    strSQL = strSQL & ")" & vbCrLf
    strSQL = strSQL & "VALUES" & vbCrLf & vbTab
    strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
    
    strSQL = strSQL & "'" & Me.ParentFolderID & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.Prg & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.OwnerID & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.QueryType & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.FolderName & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.Image & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Comment) & "'" & vbCrLf & vbTab
    
    strSQL = strSQL & ")"
    Me.FolderID = BaseToolKit.Database.ExecuteNonQuery(strSQL)
  Else '{Aktualisieren}
    strSQL = strSQL & "UPDATE" & vbCrLf & vbTab
    strSQL = strSQL & "queries.t_query_folders" & vbCrLf
    strSQL = strSQL & "SET" & vbCrLf & vbTab
    strSQL = strSQL & "ParentFolderFID = '" & Me.ParentFolderID & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Prg = '" & Me.Prg & "'," & vbCrLf & vbTab
    strSQL = strSQL & "OwnerFID = '" & Me.OwnerID & "'," & vbCrLf & vbTab
    strSQL = strSQL & "QueryType = '" & Me.QueryType & "'," & vbCrLf & vbTab
    strSQL = strSQL & "FolderName = '" & Me.FolderName & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Image = '" & Me.Image & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Comment = '" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Comment) & "'" & vbCrLf
    strSQL = strSQL & "WHERE" & vbCrLf & vbTab
    strSQL = strSQL & "QueryFolderID = '" & Me.FolderID & "'"
    BaseToolKit.Database.ExecuteNonQuery strSQL
  End If
  Exit Function

errLabel:
  SaveQueryFolder = -1
  ShowError "SaveQueryFolder"
  Exit Function
End Function

Public Function DeleteQueryFolder() As Long

On Error GoTo errLabel:

  DeleteQueryFolder = 0
  
  '{Unterordner löschen}
  Me.QueryFolders.GetItems Me.FolderID, Me.Prg
  
  Dim aQueryFolder As UserQueriesViewQueryFolder
  For Each aQueryFolder In Me.QueryFolders.Items
    aQueryFolder.DeleteQueryFolder
  Next aQueryFolder
  '{Ende Unterordner löschen}

  '{Queries löschen}
  Dim strSQL As String
  strSQL = "DELETE FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_user_queries" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "QueryFolderFID = '" & Me.FolderID & "'"
  BaseToolKit.Database.ExecuteNonQuery strSQL
  '{Ende Queries löschen}
  
  strSQL = "DELETE FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_query_folders" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "QueryFolderID = '" & Me.FolderID & "'"
  BaseToolKit.Database.ExecuteNonQuery strSQL
  Exit Function

errLabel:
  DeleteQueryFolder = -1
  ShowError "DeleteQueryFolder"
  Exit Function
End Function

