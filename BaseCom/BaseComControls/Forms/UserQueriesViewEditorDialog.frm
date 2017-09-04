VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form UserQueriesViewEditorDialog 
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   Icon            =   "UserQueriesViewEditorDialog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   1680
   End
   Begin VB.Frame fraEditTools 
      Height          =   1575
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      Begin MSComctlLib.TreeView tvwBaum 
         Height          =   1005
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1773
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   265
         LabelEdit       =   1
         PathSeparator   =   "."
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdAbbrechen 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtEditor 
      Height          =   1575
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ilsBilder 
      Left            =   1200
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":014A
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":02A4
            Key             =   "User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":03FE
            Key             =   "Databases"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":0558
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":0872
            Key             =   "Unique"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":09CC
            Key             =   "Primary"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":0B26
            Key             =   "Multi"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":0C80
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":0DDA
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":0F34
            Key             =   "Point"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":12CE
            Key             =   "Keywords"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":1668
            Key             =   "Entwurf"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":28BA
            Key             =   "Editor"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserQueriesViewEditorDialog.frx":2A14
            Key             =   "Syntax"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "UserQueriesViewEditorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrEditorText As String
Private mblnShowEditTools As Boolean
Private mblnCancel As Boolean

Public Property Get EditorText() As String
  EditorText = mstrEditorText
End Property

Public Property Let EditorText(ByVal strEditorText As String)
  mstrEditorText = strEditorText
End Property

Public Property Get ShowEditTools() As Boolean
  ShowEditTools = mblnShowEditTools
End Property

Public Property Let ShowEditTools(ByVal blnShowEditTools As Boolean)
  mblnShowEditTools = blnShowEditTools
End Property

Public Property Get Cancel() As Boolean
  Cancel = mblnCancel
End Property

Private Sub cmdAbbrechen_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  mblnCancel = False
  Me.EditorText = CheckFormalSyntax(txtEditor.Text)
  Unload Me
End Sub

Private Function CheckFormalSyntax(ByVal strText As String) As String
Dim aNode As MSComctlLib.Node
Dim colChars As Collection
Dim varPräfix As Variant
Dim varSuffix As Variant
  
On Error Resume Next
  
  With txtEditor
  
    Set colChars = New Collection
    colChars.Add " "
    colChars.Add ","
    colChars.Add "("
    colChars.Add ")"
    colChars.Add vbCrLf
  
    If Me.ShowEditTools Then
  
      strText = strText & " "
      For Each aNode In tvwBaum.Nodes
        If aNode.Image = "Point" Then
          
          For Each varPräfix In colChars
            For Each varSuffix In colChars
              strText = Replace(strText, varPräfix & aNode.Text _
              & varSuffix, varPräfix & UCase(aNode.Text) & varSuffix, , , vbTextCompare)
            Next varSuffix
          Next varPräfix
          
        End If
      Next aNode
      strText = Trim(strText)
      CheckFormalSyntax = strText
    
    Else
      CheckFormalSyntax = .Text
    End If
  End With
  
End Function

Private Sub Form_Load()
  mblnCancel = True
  Me.EditorText = ""
  Me.Top = GetSetting(App.EXEName, Me.Name, "Top", Me.Top)
  Me.Left = GetSetting(App.EXEName, Me.Name, "Left", Me.Left)
  Me.Height = GetSetting(App.EXEName, Me.Name, "Height", Me.Height)
  Me.Width = GetSetting(App.EXEName, Me.Name, "Width", Me.Width)
  Me.WindowState = GetSetting(App.EXEName, Me.Name, "WindowsState", Me.WindowState)
  txtEditor.Text = Me.EditorText
  tmrStart.Enabled = True
End Sub

Private Sub Form_Resize()
On Error Resume Next

  If Me.Width < 2000 Then Me.Width = 2000
  If Me.Height < 1500 Then Me.Height = 1500
  
  txtEditor.Height = Me.ScaleHeight - (cmdOK.Height + 50)
  
  If Me.ShowEditTools Then
    fraEditTools.Height = txtEditor.Height
    txtEditor.Width = (Me.ScaleWidth * 0.7) - 20
    fraEditTools.Width = (Me.ScaleWidth * 0.3) - 20
    fraEditTools.Left = txtEditor.Width + 40
    tvwBaum.Width = fraEditTools.Width - (2 * tvwBaum.Left)
    tvwBaum.Height = fraEditTools.Height - (1.5 * tvwBaum.Top)
  Else
    txtEditor.Width = Me.ScaleWidth
  End If
  
  cmdOK.Left = Me.ScaleWidth - cmdOK.Width
  cmdOK.Top = Me.ScaleHeight - cmdOK.Height
  
  cmdAbbrechen.Left = cmdOK.Left - (cmdAbbrechen.Width + 50)
  cmdAbbrechen.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  SaveSetting App.EXEName, Me.Name, "Top", Me.Top
  SaveSetting App.EXEName, Me.Name, "Left", Me.Left
  SaveSetting App.EXEName, Me.Name, "Height", Me.Height
  SaveSetting App.EXEName, Me.Name, "Width", Me.Width
  SaveSetting App.EXEName, Me.Name, "WindowsState", Me.WindowState
  Me.ShowEditTools = False
End Sub

Private Sub tmrStart_Timer()
  tmrStart.Enabled = False
  fraEditTools.Visible = Me.ShowEditTools
  If Me.ShowEditTools Then
    BaueTeilbaumRoot tvwBaum, ilsBilder
    Me.Icon = ilsBilder.ListImages("Entwurf").Picture
    Me.Caption = "Entwurfsansicht"
  Else
    Me.Icon = ilsBilder.ListImages("Editor").Picture
  End If
  
End Sub

Public Sub BaueTeilbaumRoot _
(ByRef aBaum As TreeView _
, ByRef aImageList As ImageList)
Dim aNode As Node

On Error GoTo errLabel:
  
  tvwBaum.Nodes.Clear
  aBaum.ImageList = aImageList
    
  Set aNode = aBaum.Nodes.Add
  aNode.Text = "Databases"
  aNode.Tag = ""
  aNode.Key = "Databases"
  aNode.Image = "Databases"
  
  Set aNode = aBaum.Nodes.Add
  aNode.Text = "Schlüsselwörter"
  aNode.Tag = ""
  aNode.Key = "Keywords"
  aNode.Image = "Keywords"
  
  BaueTeilbaumKeywords tvwBaum.Nodes("Keywords")
  
  Set aNode = aBaum.Nodes.Add
  aNode.Text = "Syntax"
  aNode.Tag = ""
  aNode.Key = "Syntax"
  aNode.Image = "Syntax"
  
  BaueTeilbaumSyntax tvwBaum.Nodes("Syntax")
  Exit Sub
  
errLabel:
  MsgBox "(" & Err.Number & ") " & Err.Description _
  , 16, "BaueTeilbaumRoot"
  Exit Sub
End Sub

Public Sub BaueTeilbaumDatabases _
(ByRef SelectedNode As MSComctlLib.Node)

On Error GoTo errLabel:
    
  If SelectedNode.Children > 0 Then Exit Sub
    
  Screen.MousePointer = 11
  
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected("SHOW DATABASES")
  While Not rs.EOF
    Dim aNode As Node: Set aNode = tvwBaum.Nodes.Add("Databases", tvwChild)
    aNode.Text = rs!Database
    aNode.Tag = rs!Database
    aNode.Key = rs!Database
    aNode.Image = "Database"
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  
  Screen.MousePointer = 0
  SelectedNode.Sorted = True
  Exit Sub
  
errLabel:
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description _
  , 16, "BaueTeilbaumDatabases"
  Exit Sub
End Sub

Public Sub BaueTeilbaumTables _
(ByRef SelectedNode As MSComctlLib.Node, ByVal strDatabase As String)

On Error GoTo errLabel:
  
  If SelectedNode.Children > 0 Then Exit Sub
  
  Screen.MousePointer = 11
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected _
  ("SHOW TABLES FROM " & strDatabase)
  
  While Not rs.EOF
    Dim aNode As Node: Set aNode = tvwBaum.Nodes.Add(strDatabase, tvwChild)
    aNode.Text = rs.Fields(0).value
    aNode.Tag = rs.Fields(0).value
    aNode.Key = strDatabase & "/" & rs.Fields(0).value
    aNode.Image = "Table"
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  
  Screen.MousePointer = 0
  SelectedNode.Sorted = True
  Exit Sub
  
errLabel:
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description _
  , 16, "BaueTeilbaumTables"
  Exit Sub
End Sub

Public Sub BaueTeilbaumFields(ByRef SelectedNode As MSComctlLib.Node _
, ByVal strDatabase As String, ByVal strTable As String)

On Error GoTo errLabel:
  
  If SelectedNode.Children > 0 Then Exit Sub
  
  Screen.MousePointer = 11
    
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected _
  ("SHOW COLUMNS FROM " & strDatabase & "." & strTable)
  While Not rs.EOF
    
    Dim aNode As Node: Set aNode = tvwBaum.Nodes.Add _
    (strDatabase & "/" & strTable, tvwChild)
    With aNode
      .Text = rs!Field & " [" & rs!Type & "]"
      .Tag = rs!Field
      .Key = strDatabase & "/" & strTable & "/" & rs!Field
      If rs!Key & "" = "" Then
        .Image = "Field"
      Else
        Select Case rs!Key
        Case "PRI"
          .Image = "Primary"
        Case "UNI"
          .Image = "Unique"
        Case "MUL"
          .Image = "Multi"
        End Select
      End If
    End With
    
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  
  Screen.MousePointer = 0
  Exit Sub
  
errLabel:
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description _
  , 16, "BaueTeilbaumFields"
  Exit Sub
End Sub

Public Sub BaueTeilbaumKeywords _
(ByRef SelectedNode As MSComctlLib.Node)

On Error GoTo errLabel:
  
  If SelectedNode.Children > 0 Then Exit Sub
  
  Screen.MousePointer = 11
  
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected _
  ("SELECT k.Keyword FROM queries.t_keywords k ORDER BY k.Keyword")
  
  While Not rs.EOF
    Dim aNode As Node: Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
    aNode.Text = rs.Fields(0).value
    aNode.Tag = rs.Fields(0).value
    aNode.Key = SelectedNode.Key & "/" & rs.Fields(0).value
    aNode.Image = "Point"
    rs.MoveNext
  Wend
  BaseToolKit.Database.CloseRecordSet rs
  
  Screen.MousePointer = 0
  SelectedNode.Sorted = True
  Exit Sub
  
errLabel:
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description _
  , 16, "BaueTeilbaumKeywords"
  Exit Sub
End Sub

Public Sub BaueTeilbaumSyntax _
(ByRef SelectedNode As MSComctlLib.Node)

Dim aNode As Node
Dim strSQL As String

On Error GoTo errLabel:
  
  If SelectedNode.Children > 0 Then Exit Sub
  
  Screen.MousePointer = 11
    
  '{SELECT}
  strSQL = "SELECT" & vbCrLf & vbTab & "[columns]" & vbCrLf & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab & "[tables]" & vbCrLf & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab & "[search_conditions]" & vbCrLf & vbCrLf
  strSQL = strSQL & "GROUP BY" & vbCrLf & vbTab & "[columns]" & vbCrLf & vbCrLf
  strSQL = strSQL & "HAVING" & vbCrLf & vbTab & "[search_conditions]" & vbCrLf & vbCrLf
  strSQL = strSQL & "ORDER BY" & vbCrLf & vbTab & "[sort_orders]"

  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Text = "SELECT"
  aNode.Tag = strSQL
  aNode.Key = SelectedNode.Key & "/SyntaxSELECT"
  aNode.Image = "Point"
  '{Ende SELECT}
  
  '{INSERT INTO}
  strSQL = "INSERT INTO" & vbCrLf & vbTab & "[tablename]" & vbCrLf & vbCrLf & vbTab
  strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
  strSQL = strSQL & "[column1], [column2] ..." & vbCrLf & vbTab
  strSQL = strSQL & ")" & vbCrLf & vbCrLf
  strSQL = strSQL & "VALUES" & vbCrLf & vbCrLf & vbTab
  strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
  strSQL = strSQL & "[value1], [value2] ..." & vbCrLf & vbTab
  strSQL = strSQL & ")"

  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Text = "INSERT INTO"
  aNode.Tag = strSQL
  aNode.Key = SelectedNode.Key & "/SyntaxINSERTINTO"
  aNode.Image = "Point"
  '{Ende INSERT INTO}
  
  '{UPDATE}
  strSQL = "UPDATE" & vbCrLf & vbTab & "[tablename]" & vbCrLf & vbCrLf
  strSQL = strSQL & "SET" & vbCrLf & vbTab
  strSQL = strSQL & "[column1] = [value1]," & vbCrLf & vbTab
  strSQL = strSQL & "[column2] = [value2] ..." & vbCrLf & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab & "[search_conditions]"

  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Text = "UPDATE"
  aNode.Tag = strSQL
  aNode.Key = SelectedNode.Key & "/SyntaxUPDATE"
  aNode.Image = "Point"
  '{Ende UPDATE}
    
  '{DELETE}
  strSQL = "DELETE FROM" & vbCrLf & vbTab & "[tablename]" & vbCrLf & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab & "[search_conditions]"

  Set aNode = tvwBaum.Nodes.Add(SelectedNode.Key, tvwChild)
  aNode.Text = "DELETE"
  aNode.Tag = strSQL
  aNode.Key = SelectedNode.Key & "/SyntaxDELETE"
  aNode.Image = "Point"
  '{Ende DELETE}
    
  Screen.MousePointer = 0
  SelectedNode.Sorted = True
  Exit Sub
  
errLabel:
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description _
  , 16, "BaueTeilbaumSyntax"
  Exit Sub
End Sub

Private Sub tvwBaum_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then InsertNodeTagToEditor
End Sub

Private Sub InsertNodeTagToEditor()
Dim aNode As MSComctlLib.Node
Dim Temp As String
Dim lngSelStart As Long

  Set aNode = tvwBaum.SelectedItem
  If aNode Is Nothing Then Exit Sub

  If InStr(aNode.Key, "Syntax") > 0 Then txtEditor.Text = ""

  Temp = Mid(txtEditor.Text, 1, txtEditor.SelStart)
  Temp = Temp & aNode.Tag
  lngSelStart = Len(Temp)
  Temp = Temp & Mid(txtEditor.Text, txtEditor.SelStart + txtEditor.SelLength + 1)
  
  txtEditor.Text = Temp
  txtEditor.SelStart = lngSelStart
End Sub

Private Sub tvwBaum_NodeClick(ByVal Node As MSComctlLib.Node)

  Select Case Node.Image
  Case "Databases"
    BaueTeilbaumDatabases Node
  Case "Database"
    BaueTeilbaumTables Node, Node.Text
  Case "Table"
    BaueTeilbaumFields Node, Node.Parent.Text, Node.Text
  Case "Keywords"
    BaueTeilbaumKeywords Node
  End Select
  txtEditor.SetFocus
End Sub

Private Sub txtEditor_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case Shift
  Case 2 '{Strg}
    Select Case KeyCode
    Case 65 '{A}
      txtEditor.SelStart = 0
      txtEditor.SelLength = Len(txtEditor.Text)
    End Select
  End Select
End Sub
