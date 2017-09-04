VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserQueriesViewQueryFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Items As Collection

Private Sub Class_Initialize()
  Set Items = New Collection
End Sub

Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub

Public Function GetItems _
(ByVal strParentFolderID As String _
, ByVal strPrg As String _
, Optional ByVal strOwnerID As String = "-1")

On Error GoTo errLabel:

  Set Items = New Collection

  Dim strSQL As String
  strSQL = strSQL & "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "qf.QueryFolderID" & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_query_folders qf" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "(qf.ParentFolderFID = '" & strParentFolderID & "')" & vbCrLf & vbTab
  strSQL = strSQL & "AND (qf.Prg = '" & strPrg & "')" & vbCrLf & vbTab
  
  If strOwnerID <> "-1" Then
    strSQL = strSQL & "AND (qf.OwnerFID = '" & strOwnerID & "')"
  End If
  
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(strSQL)
  
  While Not rs.EOF
    Dim aQueryFolder As UserQueriesViewQueryFolder: Set aQueryFolder = New UserQueriesViewQueryFolder
    aQueryFolder.GetQueryFolder rs!QueryFolderID
    
    Items.Add aQueryFolder, "QF" & aQueryFolder.FolderID
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  Exit Function
  
errLabel:
  ShowError "GetItems"
  Exit Function
End Function
