VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SelectEntryDiaolg 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Bitte wählen ..."
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdCheck 
      Height          =   375
      Left            =   480
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdAbbrechen 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   0
      Picture         =   "SelectEntryDialog.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   3240
      Width           =   375
   End
   Begin MSComctlLib.ImageList ilsBilder 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelectEntryDialog.frx":014A
            Key             =   "CloseFolder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelectEntryDialog.frx":04E4
            Key             =   "Point"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelectEntryDialog.frx":087E
            Key             =   "OpenFolder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelectEntryDialog.frx":0C18
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelectEntryDialog.frx":0D72
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraVorgabe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4095
      Begin MSComctlLib.TreeView tvwVorgabe 
         Height          =   2775
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4895
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         Style           =   7
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "SelectEntryDiaolg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnCancelCheck As Boolean

Private mstrValueEntry As String
Private mstrValueID As String
Private mcolValueEntries As Collection
Private mcolValueIDs As Collection
Private mcolPreChecked As Collection

Private mblnSortInFolder As String
Private mstrDataString As String
Private mcolDataCollection As Collection
Private mblnReturnID As Boolean
Private mblnSorted As Boolean
Private mstrDelimiter As String
Private mstrPreSelectValue As String
Private mblnReturnSingleEntryAutomatically As Boolean
Private mblnMultiSelect As Boolean
Private mblnPreSelect As Boolean
Private mblnUseFirstLetter As Boolean
Private mblnIsMarkDemarkAllVisible As Boolean

Public Property Get PreSelect() As Boolean
Attribute PreSelect.VB_Description = "Besitzt PreSelect den Wert True, wird versucht den Eintrag von PreSelectValue zu selektieren."
  PreSelect = mblnPreSelect
End Property

Public Property Let PreSelect(ByVal blnPreSelect As Boolean)
  mblnPreSelect = blnPreSelect
End Property

Public Property Get ValueEntry() As String
Attribute ValueEntry.VB_Description = "Rückgabewert. Beinhaltet die getroffene Benutzerauswahl. Bei Rückgabe eines Leerstrings wurde keine Auswahl getroffen."
  ValueEntry = mstrValueEntry
End Property

Public Property Let ValueEntry(ByVal strValueEntry As String)
  mstrValueEntry = strValueEntry
End Property

Public Property Get ValueID() As String
Attribute ValueID.VB_Description = "Optionaler Rückgabewert. Wird nur gesetzt wenn ReturnID den Wert ""True"" besitzt. Liefert ansonsten einen Leerstring, als wenn keine Auswahl getroffen wurde."
  ValueID = mstrValueID
End Property

Public Property Let ValueID(ByVal strValueID As String)
  mstrValueID = strValueID
End Property

Public Property Get ValueEntries() As Collection
Attribute ValueEntries.VB_Description = "Eine Collection, welche die ausgewählten Elemente in einer Mehrfachauswahl enthält."
  Set ValueEntries = mcolValueEntries
End Property

Public Property Get ValueIDs() As Collection
Attribute ValueIDs.VB_Description = "Eine Collection, welche die IDs der ausgewählten Elemente in einer Mehrfachauswahl enthält."
  Set ValueIDs = mcolValueIDs
End Property

Public Property Get SortInFolder() As Boolean
Attribute SortInFolder.VB_Description = "Eigenschaft welche bestimmt, ob alle Einträge in Unterordner einsortiert werden sollen. Die Unterordner gruppieren die Einträge nach ihrem ersten Buchstaben. Vorbelegung ist ""False""."
  SortInFolder = mblnSortInFolder
End Property

Public Property Let SortInFolder(ByVal blnSortInFolder As Boolean)
  mblnSortInFolder = blnSortInFolder
End Property

Public Property Get UseFirstLetter() As Boolean
Attribute UseFirstLetter.VB_Description = "Beimn Wert True wird der erste Buchstabe für die Gruppierung benutzt. Bei False wird das jeweils dritte Element aus DataCollection benutzt."
  UseFirstLetter = mblnUseFirstLetter
End Property

Public Property Let UseFirstLetter(ByVal blnUseFirstLetter As Boolean)
  mblnUseFirstLetter = blnUseFirstLetter
End Property

Public Property Get DataString() As String
Attribute DataString.VB_Description = "Beinhaltet alle Einträge, welche in der Auswahlliste erscheinen sollen. Die einzelnen Einträge sind durch ein Trennzeichen (Delimiter) getrennt."
  DataString = mstrDataString
End Property

Public Property Let DataString(ByVal strDataString As String)
  mstrDataString = strDataString
  Set mcolDataCollection = TransIntoCollection(strDataString, Me.ReturnID, Me.SortInFolder)
End Property

Public Property Get DataCollection() As Collection
Attribute DataCollection.VB_Description = "Beinhaltet alle Auswahlmöglichkeiten"
  Set DataCollection = mcolDataCollection
End Property

Public Property Let DataCollection(ByVal colDataCollection As Collection)
  Set mcolDataCollection = colDataCollection
End Property

Public Property Get PreChecked() As Collection
Attribute PreChecked.VB_Description = "Bestimmt in einer Mehrfachauswahl welche Wahlmöglichkeiten vorgehakt sind."
  Set PreChecked = mcolPreChecked
End Property

Public Property Let PreChecked(ByRef colPreChecked As Collection)
  Set mcolPreChecked = colPreChecked
End Property

Public Property Get ReturnID() As Boolean
Attribute ReturnID.VB_Description = "Bestmmt ob nur die Textansicht in der Auswahlliste gefüllt wird oder auch ein ID-Wert, welcher bei ReturnID=True zusätzlich zu ValueEntry in ValueID zurück gegeben wird. Vorbelegt ist ReturnID mit 'False'. Beispiel für DataString: ""#Value1#ID1...""."
  ReturnID = mblnReturnID
End Property

Public Property Let ReturnID(ByVal blnReturnID As Boolean)
  mblnReturnID = blnReturnID
End Property

Public Property Get IsMarkDemarkAllVisible() As Boolean
  IsMarkDemarkAllVisible = mblnIsMarkDemarkAllVisible
End Property

Public Property Let IsMarkDemarkAllVisible(ByVal blnIsMarkDemarkAllVisible As Boolean)
  mblnIsMarkDemarkAllVisible = blnIsMarkDemarkAllVisible
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Eigenschaft welche angibt ob die Einträge in der Auswahlliste alphanumerisch sortiert werden sollen. Vorbelegt mit ""True""."
  Sorted = mblnSorted
End Property

Public Property Let Sorted(ByVal blnSorted As Boolean)
  mblnSorted = blnSorted
End Property

Public Property Get Header() As String
Attribute Header.VB_Description = "Überschrift des Auswahlfensters. Vorbelegt mit ""'Bitte wählen ...""."
  Header = Me.Caption
End Property

Public Property Let Header(ByVal strHeader As String)
  Me.Caption = strHeader
End Property

Public Property Get Delimiter() As String
Attribute Delimiter.VB_Description = "Trennzeichen zwischen den einzelnen Einträgen in DataString. Vorbelegt mit ""#"", kann jedoch bei Bedarf umgesetzt werden."
  Delimiter = mstrDelimiter
End Property

Public Property Let Delimiter(ByVal strDelimiter As String)
  mstrDelimiter = strDelimiter
End Property

Public Property Get PreSelectValue() As String
Attribute PreSelectValue.VB_Description = "Vorbelegung in der Auswahlliste. Ist PreSelectValue angegeben, wird SelectEntry versuchen, nach dem Füllen der Auswahlliste, einen Eintrag vorzubelegen, welcher mit PreSelectValue übereinstimmt."
  PreSelectValue = mstrPreSelectValue
End Property

Public Property Let PreSelectValue(ByVal strPreSelectValue As String)
  mstrPreSelectValue = strPreSelectValue
End Property

Public Property Get ReturnSingleEntryAutomatically() As Boolean
Attribute ReturnSingleEntryAutomatically.VB_Description = "Ein einzelner Eintrag wird automatisch ausgewählt und in ValueEntry und ggf. in ValueID zurück gegeben, wenn ReturnSingleEntryAutomatically den Wert ""True"" enthält. Vorbelegung ist ""False""."
  ReturnSingleEntryAutomatically = mblnReturnSingleEntryAutomatically
End Property

Public Property Let ReturnSingleEntryAutomatically(ByVal blnReturnSingleEntryAutomatically As Boolean)
  mblnReturnSingleEntryAutomatically = blnReturnSingleEntryAutomatically
End Property

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Schaltet bei True in den Mehrfachauswahl-Modus, bei False ist nur eine Einzelwahl möglich."
  MultiSelect = mblnMultiSelect
End Property

Public Property Let MultiSelect(ByVal blnMultiSelect As Boolean)
  mblnMultiSelect = blnMultiSelect
End Property

Public Sub SetDataSource(ByRef rs As Object)
Attribute SetDataSource.VB_Description = "Liest Daten aus einem ADO-RecordSet und belegt den DataString. Ein Aufruf von SetDataSource hat die gleichen Auswirkungen, wie eine direkte DataString-Zuweisung. Die Einträge werden durch Delimiter getrennt. Rs benötigt das Feld Value opt. ID und Folder"
Dim strFieldName As String

On Error GoTo Fehler

  Set mcolDataCollection = New Collection

  rs.MoveFirst

  While Not rs.EOF
      
    Me.DataCollection.Add rs.Fields("Value").value & ""
    
    strFieldName = ""
    On Error Resume Next
    strFieldName = rs.Fields("ID").name
    On Error GoTo Fehler
    
    If strFieldName <> "" Then Me.DataCollection.Add rs.Fields("ID").value
    
    strFieldName = ""
    On Error Resume Next
    strFieldName = rs.Fields("Folder").name
    On Error GoTo Fehler
    
    If strFieldName <> "" Then Me.DataCollection.Add rs.Fields("Folder").value
    
    rs.MoveNext
  Wend
  Exit Sub
  
Fehler:
  ShowError "GetDataSource"
  Exit Sub
End Sub

Public Sub SetIniFileSource _
(ByVal strFileName As String _
, ByVal strSection As String)
Attribute SetIniFileSource.VB_Description = "Liest Daten aus einer Ini-Datei und belegt den DataString. SetIniFileSource hat die gleichen Auswirkungen, wie eine direkte DataString-Zuweisung. Die Einträge werden durch Delimiter getrennt. Sektion benötigt Entries, Valuex und opt. IDx und Folderx"

Dim ini As FileIni: Set ini = New FileIni
Dim astrValues() As String
Dim astrIDs() As String
Dim astrFolders()  As String
Dim i As Integer

On Error GoTo Fehler


  astrValues = Split(ini.ReadCounterEntryValues(strFileName, strSection, "Entries", "Value"), "#")
    
  On Error Resume Next
  astrIDs = Split(ini.ReadCounterEntryValues(strFileName, strSection, "Entries", "ID"), "#")
  On Error GoTo Fehler
  
  On Error Resume Next
  astrFolders = Split(ini.ReadCounterEntryValues(strFileName, strSection, "Entries", "Folder"), "#")
  On Error GoTo Fehler
  
  Me.DataString = ""
  For i = 1 To UBound(astrValues)
    Me.DataString = Me.DataString & Me.Delimiter & astrValues(i)
    
    If astrIDs(i) <> "" Then
      Me.DataString = Me.DataString & Me.Delimiter & astrIDs(i)
    End If
  
    If astrFolders(i) <> "" Then
      Me.DataString = Me.DataString & Me.Delimiter & astrFolders(i)
    End If
  
  Next i
  
  Set ini = Nothing
  Exit Sub
  
Fehler:
  ShowError "SetIniFileSource"
  Exit Sub
End Sub

Public Sub SelectEntry _
(ByVal DataCollection As Collection _
, Optional ByVal strHeader As String = "" _
, Optional ByVal blnReturnID As Boolean = False _
, Optional ByVal blnSorted As Boolean = True _
, Optional ByVal blnSortInFolder As Boolean = False _
, Optional ByVal strPreSelectValue As String = "" _
, Optional ByVal blnReturnSingleEntryAutomatically As Boolean = False _
, Optional ByVal blnUnload As Boolean = False)
Attribute SelectEntry.VB_Description = "Startet das Auswahlfenster."

Dim i As Integer
Dim intStep As Integer
Dim aNode As MSComctlLib.Node
Dim strFirstLetter As String
Dim lngEntries As Long
Dim X As Variant

Const vbKeyÄ = 196
Const vbKeyÖ = 214
Const vbKeyÜ = 220
Const vbKeyMinus = 45

On Error GoTo Fehler

  tvwVorgabe.Checkboxes = Me.MultiSelect
  Set mcolDataCollection = New Collection
  Set mcolValueEntries = New Collection
  Set mcolValueIDs = New Collection

  Screen.MousePointer = 11
  Me.Caption = strHeader
  
  tvwVorgabe.Nodes.Clear
  tvwVorgabe.Sorted = blnSorted
  tvwVorgabe.ImageList = ilsBilder
    
  For i = 1 To DataCollection.count
    If Len(DataCollection.item(i)) > 0 Then

      If Not blnSortInFolder Then
        Set aNode = tvwVorgabe.Nodes.Add
      Else
      
        If (Me.UseFirstLetter) Then
          strFirstLetter = UCase(Mid(DataCollection.item(i), 1, 1))
        Else
          If blnReturnID Then
            strFirstLetter = DataCollection.item(i + 2)
          Else
            strFirstLetter = DataCollection.item(i + 1)
          End If
        End If
     
        Select Case Asc(strFirstLetter)
        Case vbKeyMinus
          Set aNode = tvwVorgabe.Nodes.Add
        Case vbKey0 To vbKey9, vbKeyA To vbKeyZ, vbKeyÄ, vbKeyÖ, vbKeyÜ, 97 To 122 '{Zahlen, alle Buchstaben}
          
          On Error Resume Next:
          Set aNode = tvwVorgabe.Nodes.Add(, , "K#" & strFirstLetter, strFirstLetter, "CloseFolder", "OpenFolder")
          aNode.Sorted = blnSorted
          On Error GoTo Fehler
          Set aNode = tvwVorgabe.Nodes.Add("K#" & strFirstLetter, tvwChild)
        Case Else
          On Error Resume Next
          Set aNode = tvwVorgabe.Nodes.Add _
          (, , "K#+", "- Sonderzeichen -", "CloseFolder", "OpenFolder")
          aNode.Sorted = blnSorted
          On Error GoTo Fehler
          Set aNode = tvwVorgabe.Nodes.Add("K#+", tvwChild)
        
          aNode.Sorted = blnSorted
        End Select
      End If
      
      aNode.Text = DataCollection.item(i)
      aNode.Image = "Point"
      aNode.Checked = Me.PreSelect
      If blnReturnID Then
        i = i + 1
        aNode.tag = DataCollection.item(i)
      End If
      
      aNode.Sorted = blnSorted
      lngEntries = lngEntries + 1
    Else
      If blnReturnID Then i = i + 1
    End If
    
    If (Not Me.UseFirstLetter) And (blnSortInFolder) Then i = i + 1
  Next i
  Screen.MousePointer = 0
  
  '{Damit auch Focus erscheint beim ersten Aufruf}
  '{auf einem modal geöffneten Form}
  If blnUnload Then
    On Error Resume Next
      Me.Show
      Unload Me
    On Error GoTo Fehler
  Else
  
    If (lngEntries = 1) And (blnReturnSingleEntryAutomatically) Then
      If tvwVorgabe.Nodes.count > lngEntries Then
        Set tvwVorgabe.SelectedItem = tvwVorgabe.Nodes(1).Child
      Else
        Set tvwVorgabe.SelectedItem = tvwVorgabe.Nodes(1)
      End If
      Ok
      Exit Sub
    End If
    
    If strPreSelectValue <> "" Then
      FindValue strPreSelectValue
    End If

    Me.fraVorgabe.Caption = "(" & lngEntries & ")"
  
    If Me.MultiSelect Then
        If IsMarkDemarkAllVisible Then
            cmdCheck.Visible = True
        Else
            cmdCheck.Visible = False
        End If
      If Me.PreSelect Then
        Set cmdCheck.Picture = ilsBilder.ListImages("UnCheck").Picture
        cmdCheck.ToolTipText = "Alle demarkieren"
      Else
        Set cmdCheck.Picture = ilsBilder.ListImages("Check").Picture
        cmdCheck.ToolTipText = "Alle markieren"
      End If
    Else
      cmdCheck.Visible = False
    End If
  
    For Each X In Me.PreChecked
      For Each aNode In tvwVorgabe.Nodes
        If aNode.Text = X Then aNode.Checked = True
      Next aNode
    Next X
    Me.Show 1
  
  End If
  Exit Sub

Fehler:
  ShowError "SelectEntry"
  Exit Sub
End Sub

Private Sub FindValue(Optional ByVal strPeSelectValue As String = "")
Dim Temp As String
Dim i As Long
Dim intLenTemp As Integer


On Error GoTo Fehler

  If strPeSelectValue = "" Then
    Temp = InputBox("Bitte Suchkriterium eingeben:", "Suchen")
    If Temp = "" Then Exit Sub
  Else
    Temp = strPeSelectValue
  End If
  
  Temp = LCase(Temp)
  intLenTemp = Len(Temp)
  
  With tvwVorgabe
    For i = 1 To .Nodes.count
      Select Case strPeSelectValue
      Case ""
        If intLenTemp <= Len(.Nodes(i)) Then
            
          If (InStr(LCase(.Nodes(i).Text), Temp) > 0) _
          And (.Nodes(i).Children = 0) Then
          
            Set .SelectedItem = .Nodes(i)
            If MsgBox("Weitersuchen?", 36, "Suchen") <> 6 Then Exit Sub
            
          End If
        End If
        
      Case Else
        If (LCase(.Nodes(i).Text) = Temp) And (.Nodes(i).Children = 0) Then
          Set .SelectedItem = .Nodes(i)
          Exit Sub
        End If
        
      End Select
    Next i
  End With
  
  If strPeSelectValue <> "" Then Exit Sub
  
  MsgBox "Es wurden keine weiteren Einträge gefunden!", 64, "Suchen"
  Exit Sub
  
Fehler:
  ShowError "FindValue"
  Exit Sub
End Sub

Private Sub GetWindowPos(ByRef aForm As Form)
  With aForm
    .Top = GetSetting(app.EXEName, .name, "Top", .Top)
    .Left = GetSetting(app.EXEName, .name, "Left", .Left)
    .Height = GetSetting(app.EXEName, .name, "Height", .Height)
    .Width = GetSetting(app.EXEName, .name, "Width", .Width)
  End With
End Sub


Private Sub PutWindowPos(ByRef aForm As Form)
  With aForm
    SaveSetting app.EXEName, .name, "Top", .Top
    SaveSetting app.EXEName, .name, "Left", .Left
    SaveSetting app.EXEName, .name, "Height", .Height
    SaveSetting app.EXEName, .name, "Width", .Width
  End With
End Sub

Private Sub ShowError(ByVal strMessage As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.number & ") " & Err.description, 16, strMessage
End Sub

Private Sub Cancel()
  Me.ValueEntry = ""
  Me.ValueID = ""
  Unload Me
End Sub

Private Sub Ok()
Dim aNode As MSComctlLib.Node

  Select Case Me.MultiSelect
  Case False
    Set aNode = tvwVorgabe.SelectedItem
    If aNode Is Nothing Then Exit Sub
    If aNode.Children > 0 Then Exit Sub
    Me.ValueEntry = aNode.Text
    Me.ValueID = aNode.tag
  Case True
    For Each aNode In tvwVorgabe.Nodes
      If (aNode.Checked) And (aNode.Children = 0) Then
        mcolValueEntries.Add aNode.Text
        mcolValueIDs.Add aNode.tag
      End If
    Next aNode
  End Select
  Unload Me
End Sub

Private Sub cmdCheck_Click()

  If cmdCheck.ToolTipText = "Alle demarkieren" Then
    Set cmdCheck.Picture = ilsBilder.ListImages("Check").Picture
    cmdCheck.ToolTipText = "Alle markieren"
    DeMark False
  Else
    Set cmdCheck.Picture = ilsBilder.ListImages("UnCheck").Picture
    cmdCheck.ToolTipText = "Alle demarkieren"
    DeMark True
  End If

End Sub

Private Sub DeMark(ByVal blnMark As Boolean)
Dim aNode As MSComctlLib.Node

  For Each aNode In tvwVorgabe.Nodes
    aNode.Checked = blnMark
  Next aNode
  
End Sub

Private Sub cmdFind_Click()
  FindValue
End Sub

Private Sub cmdOK_Click()
  Ok
End Sub

Private Sub Form_Initialize()
  Set mcolDataCollection = New Collection
  Set mcolPreChecked = New Collection
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case Shift
  Case 2 '{Strg}
    Select Case KeyCode
    Case 70 '{Taste-f}
      FindValue
    End Select
  End Select
End Sub

Private Sub tvwVorgabe_DblClick()
  If Me.MultiSelect Then Exit Sub
  Ok
End Sub

Private Sub tvwVorgabe_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case 13 '{Enter}
    Ok
  End Select
End Sub

Private Sub cmdAbbrechen_Click()
  Cancel
End Sub

Private Sub Form_Load()
  Me.ValueEntry = ""
  Me.ValueID = ""
  GetWindowPos Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
  
  fraVorgabe.Width = Me.ScaleWidth
  fraVorgabe.Height = Me.ScaleHeight - (Me.cmdOK.Height + 100)
  
  tvwVorgabe.Width = fraVorgabe.Width - 250
  tvwVorgabe.Height = fraVorgabe.Height - 400
    
  cmdOK.Top = fraVorgabe.Height + 100
  cmdOK.Left = Me.ScaleWidth - cmdOK.Width
  
  cmdAbbrechen.Top = fraVorgabe.Height + 100
  cmdAbbrechen.Left = cmdOK.Left - (cmdAbbrechen.Width + 100)
  
  cmdFind.Top = fraVorgabe.Height + 100
  cmdCheck.Top = fraVorgabe.Height + 100

End Sub

Private Sub Form_Unload(Cancel As Integer)
  PutWindowPos Me
End Sub

Public Function TransIntoCollection _
(ByVal strDataString As String _
, ByVal blnReturnID As Boolean _
, Optional ByVal blnSortInFolder As Boolean = False) As Collection
Attribute TransIntoCollection.VB_Description = "Wandelt den DataString in die DataCollection um."

Dim colDataCollection As Collection
Dim astrDataString() As String
Dim intStep As Integer
Dim i As Long
Dim lngUbound As Long

On Error GoTo Fehler

  Screen.MousePointer = 11
  
  Set TransIntoCollection = New Collection
  
  If blnReturnID Then
    intStep = 2
  Else
    intStep = 1
  End If

  If (Not Me.UseFirstLetter) And (blnSortInFolder) Then intStep = intStep + 1

  astrDataString = Split(strDataString, Me.Delimiter)
  
  lngUbound = UBound(astrDataString)
  
  For i = 1 To lngUbound Step intStep
    TransIntoCollection.Add astrDataString(i)
    If blnReturnID Then TransIntoCollection.Add astrDataString(i + 1)
    If (Not Me.UseFirstLetter) And (blnSortInFolder) Then
      TransIntoCollection.Add astrDataString(i + 2)
    End If
  Next i
  Screen.MousePointer = 0
  Exit Function
  
Fehler:
  ShowError "TransIntoCollection"
  Exit Function
End Function

Private Sub tvwVorgabe_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not mblnCancelCheck Then Exit Sub
  tvwVorgabe.SelectedItem.Checked = False
End Sub

Private Sub tvwVorgabe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Set tvwVorgabe.SelectedItem = tvwVorgabe.HitTest(X, Y)
End Sub

Private Sub tvwVorgabe_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Not mblnCancelCheck Then Exit Sub
  tvwVorgabe.SelectedItem.Checked = False
End Sub

Private Sub tvwVorgabe_NodeCheck(ByVal Node As MSComctlLib.Node)
  If Node.Children > 0 Then
    mblnCancelCheck = True
  Else
    mblnCancelCheck = False
  End If
End Sub
