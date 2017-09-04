VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ClassEditControl 
   BackColor       =   &H80000005&
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   PropertyPages   =   "ClassEditControl.ctx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   3375
   ToolboxBitmap   =   "ClassEditControl.ctx":0013
   Begin VB.CommandButton cmdDown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1320
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1080
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   1800
   End
   Begin MSComctlLib.ImageList ilsBilder 
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":0545
            Key             =   "Memo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":08DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":0C79
            Key             =   "Point"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":1013
            Key             =   "URL"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":13AD
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":1747
            Key             =   "PhoneNumber"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":1AE1
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":1E7B
            Key             =   "Password"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":2215
            Key             =   "Blank"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":25AF
            Key             =   "FullAccess"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":2949
            Key             =   "Writeable"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ClassEditControl.ctx":2CE3
            Key             =   "Readable"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwClassEdit 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Kein
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "ClassEditControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'{API-Definitionen}
Private Const INVOKE_CONST = 32
Private Const INVOKE_EVENTFUNC = 16
Private Const INVOKE_FUNC = 1
Private Const INVOKE_PROPERTYGET = 2
Private Const INVOKE_PROPERTYPUT = 4
Private Const INVOKE_PROPERTYPUTREF = 8
Private Const INVOKE_UNKNOWN = 0

Private Declare Function SetParent Lib "user32" ( _
  ByVal hWndChild As Long, _
  ByVal hWndNewParent As Long) As Long
'{Ende API-Definitionen}

'{Enumerationen}
Public Enum eceSort
  ceItemName = 0
  ceSortIndex = 1
End Enum
'{Ende Enumerationen}

'{Definition der Klassenereignisse}
Public Event ItemClick(ByRef Item As ClassEditInfos, ByVal Key As String)
Public Event KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
Public Event KeyPress(ByVal KeyAscii As Integer)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Public Event DblClick(ByRef Item As ClassEditInfos)
Public Event BeforeUpdate(ByRef Item As ClassEditInfos, ByVal NewValue As Variant, ByRef Cancel As Boolean)
Public Event AfterUpdate(ByRef Item As ClassEditInfos)
'{Ende Definition der Klassenereignisse}

'{Klasseneigenschaften}
Public Items As Collection
Attribute Items.VB_VarDescription = "Auflistung, welche alle Eigenschaften des Objektes verwaltet."
Private mblnAutoDetectSpecialType As Boolean
Private mblnAutoCallSaveMethod As Boolean
Private mstrSaveMethodName As String
Private mvarSaveMethodParams As Variant
Private mObj As Object
Private mlngScrollPos As Long
Private mClipDialPath As String
'{Ende Klasseneigenschaften}

'{Konstruktor und Destruktor}
Private Sub UserControl_Initialize()
  Set Items = New Collection
  Me.AutoDetectSpecialType = True
  Me.AutoCallSaveMethod = True
  mstrSaveMethodName = ""
  BuildListViewHeader
End Sub

Private Sub UserControl_Terminate()
  Set Items = Nothing
End Sub
'{Ende Konstruktor Destruktor}


'{Zugriffsmethoden der Klasseneigenschaften}
Public Property Get SelectedItem() As MSComctlLib.ListItem
Attribute SelectedItem.VB_Description = "Gibt eine Referenz auf die aktuell selektierte Eigenschaft zurück."
  Set SelectedItem = lvwClassEdit.SelectedItem
End Property

Public Property Get AutoDetectSpecialType() As Boolean
Attribute AutoDetectSpecialType.VB_Description = "Versucht zu erkennen, ob der Wert einer Eigenschaft dem Format einer URL oder\r\nEMail-Adresse entspricht und formatiert sie ggf. als SpecialType\r\n"
Attribute AutoDetectSpecialType.VB_ProcData.VB_Invoke_Property = "Allgemein"
  AutoDetectSpecialType = mblnAutoDetectSpecialType
End Property

Public Property Let AutoDetectSpecialType(ByVal blnAutoDetectSpecialType As Boolean)
  mblnAutoDetectSpecialType = blnAutoDetectSpecialType
End Property

Public Property Get SaveMethodName() As String
Attribute SaveMethodName.VB_Description = "Gibt den Namen der Speicherroutine zurück, sofern sie mit SetSaveMethod angegeben wurde."
  SaveMethodName = mstrSaveMethodName
End Property

Public Property Get AutoCallSaveMethod() As Boolean
Attribute AutoCallSaveMethod.VB_Description = "Ruft automatsich, bei jeder Änderung, die mit SetSaveMethod gesetzte Speicherroutine des Objektes auf."
Attribute AutoCallSaveMethod.VB_ProcData.VB_Invoke_Property = "Allgemein"
  AutoCallSaveMethod = mblnAutoCallSaveMethod
End Property

Public Property Let AutoCallSaveMethod(ByVal blnAutoCallSaveMethod As Boolean)
  mblnAutoCallSaveMethod = blnAutoCallSaveMethod
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Legt Schriftattribute fest."
  Set Font = lvwClassEdit.Font
End Property

Public Property Set Font(ByVal Value As Font)
  Set lvwClassEdit.Font = Value
  'PropertyChanged "Font"
End Property
'{Zugriffsmethoden der Klasseneigenschaften}


'{Private Klassen-Methoden}
Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub

Private Sub ScrollUp()
  
  With lvwClassEdit
    Set .SelectedItem = .GetFirstVisible
    .SelectedItem.EnsureVisible
    cmdDown.Enabled = True
    
    
    If .SelectedItem.Index <= 1 Then Exit Sub
    Set .SelectedItem = .ListItems.Item(.SelectedItem.Index - 1)
    .SelectedItem.EnsureVisible
    If .SelectedItem.Index = 1 Then cmdUp.Enabled = False
  End With
End Sub

Private Sub cmdDown_Click()
  ScrollDown
End Sub

Private Sub cmdUp_Click()
  ScrollUp
End Sub

Private Sub ScrollDown()
On Error GoTo Fehler

  With lvwClassEdit
    Set .SelectedItem = ToolKitsModule.BaseToolKit.Controls.ListView.GetLastVisibleListItem(lvwClassEdit)
    .SelectedItem.EnsureVisible
    cmdUp.Enabled = True
    
    If .SelectedItem.Index >= .ListItems.Count Then Exit Sub
    
    Set .SelectedItem = .ListItems.Item(.SelectedItem.Index + 1)
    .SelectedItem.EnsureVisible
  
    If .SelectedItem.Index = .ListItems.Count Then cmdDown.Enabled = False
  
  End With
  Exit Sub
  
Fehler:
  ShowError "ScrollDown"
  Exit Sub
End Sub

Private Function CheckInputOK(ByRef aItem As ClassEditInfos, ByVal strValue As Variant) As Boolean
On Error GoTo Fehler

  CheckInputOK = True

  If (Trim(strValue) = "") And (Not aItem.AllowNullString) Then
    CheckInputOK = False
    Exit Function
  End If

  Select Case aItem.ValueType
  Case evtString
  Case evtReal
    CheckInputOK = IsReal(aItem, strValue)
  Case evtInteger
    CheckInputOK = IsInteger(aItem, strValue)
  Case evtDate
    CheckInputOK = IsDateEx(aItem, strValue)
  Case evtTime
    CheckInputOK = IsTime(aItem, strValue)
  End Select
  Exit Function
  
Fehler:
  ShowError "CheckInput"
  Exit Function
End Function

Private Function CheckMinMaxOK(ByRef aItem As ClassEditInfos, ByVal strValue As Variant) As Boolean
On Error GoTo Fehler

  CheckMinMaxOK = True

  Select Case aItem.ValueType
  Case evtString
    If strValue <> "" Then strValue = CStr(strValue)
    If aItem.Min <> "" Then aItem.Min = CStr(aItem.Min)
    If aItem.Max <> "" Then aItem.Max = CStr(aItem.Max)
  Case evtReal
    If strValue <> "" Then strValue = CDbl(strValue)
    If aItem.Min <> "" Then aItem.Min = CDbl(aItem.Min)
    If aItem.Max <> "" Then aItem.Max = CDbl(aItem.Max)
  Case evtInteger
    If strValue <> "" Then strValue = CLng(strValue)
    If aItem.Min <> "" Then aItem.Min = CLng(aItem.Min)
    If aItem.Max <> "" Then aItem.Max = CLng(aItem.Max)
  Case evtDate
    If strValue <> "" Then strValue = CDate(strValue)
    If aItem.Min <> "" Then aItem.Min = CDate(aItem.Min)
    If aItem.Max <> "" Then aItem.Max = CDate(aItem.Max)
  Case evtTime
    If strValue <> "" Then strValue = Format(strValue, "hh:mm")
    If aItem.Min <> "" Then aItem.Min = Format(aItem.Min, "hh:mm")
    If aItem.Max <> "" Then aItem.Max = Format(aItem.Max, "hh:mm")
  End Select
  
  '{Unteren Bereich prüfen}
  If Trim(aItem.Min) <> "" Then
    If strValue < aItem.Min Then
      MsgBox "Untere Grenze von " & aItem.Min _
      & " unterschritten!", 16, "Wertunterschreitung"
      txtInput.Text = Format(aItem.Value, aItem.Format)
      CheckMinMaxOK = False
      Exit Function
    End If
  End If

  '{Oberen Bereich prüfen}
  If Trim(aItem.Max) <> "" Then
    If strValue > aItem.Max Then
      MsgBox "Obere Grenze von " & aItem.Max _
      & " überschritten!", 16, "Wertüberschreitung"
      txtInput.Text = Format(aItem.Value, aItem.Format)
      CheckMinMaxOK = False
      Exit Function
    End If
  End If
  Exit Function
  
Fehler:
  ShowError "CheckMinMaxOK"
  Exit Function
End Function

Private Sub lvwClassEdit_DblClick()
  If lvwClassEdit.SelectedItem Is Nothing Then Exit Sub
  RaiseEvent DblClick(lvwClassEdit.SelectedItem.Tag)
End Sub

Private Sub lvwClassEdit_ItemClick(ByVal Item As MSComctlLib.ListItem)
  RaiseEvent ItemClick(Item.Tag, Item.Tag.Name)
  PrepareItemEnter
End Sub

Private Sub lvwClassEdit_KeyDown(KeyCode As Integer, Shift As Integer)
Dim aItem As ClassEditInfos

  If Not lvwClassEdit.SelectedItem Is Nothing Then
    Set aItem = lvwClassEdit.SelectedItem.Tag
  End If
  
  RaiseEvent KeyDown(KeyCode, Shift)
  
  If aItem Is Nothing Then Exit Sub
  Select Case KeyCode
  Case 13 '{Enter}
    Select Case aItem.SpecialType
    Case eceiSSelect
      DoSelect aItem
    Case eceiSMemo
      ShowMemo aItem
    Case Else
      ShowInputBox lvwClassEdit.SelectedItem
      'PrepareItemEnter
      txtInput.SetFocus
    End Select
  Case 40, 38 '{Down,Up}
    CheckUpDown
  End Select
End Sub

Private Sub lvwClassEdit_KeyPress(KeyAscii As Integer)
Dim aItem As ClassEditInfos

  If Not lvwClassEdit.SelectedItem Is Nothing Then
    Set aItem = lvwClassEdit.SelectedItem.Tag
  End If
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lvwClassEdit_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim aItem As ClassEditInfos
'
'  If Not lvwClassEdit.SelectedItem Is Nothing Then
'    Set aItem = lvwClassEdit.SelectedItem.Tag
'  End If
'
  RaiseEvent KeyUp(KeyCode, Shift)
  
'  If aItem Is Nothing Then Exit Sub
'  Select Case KeyCode
'  Case 13 '{Enter}
'
'    Select Case aItem.SpecialType
'    Case eceiSSelect
'      DoSelect aItem
'    Case eceiSMemo
'      ShowMemo aItem
'    Case Else
'      ShowInputBox lvwClassEdit.SelectedItem
'      txtInput.SetFocus
'    End Select
'  Case 40, 38 '{Down,Up}
'    CheckUpDown
'  End Select
End Sub

Private Sub lvwClassEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  ActionSelection x, Y
  RaiseEvent MouseDown(Button, Shift, x, Y)
  
End Sub

Private Sub lvwClassEdit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub lvwClassEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub ActionSelection(ByVal x As Single, ByVal Y As Single)
Dim aListItem As ListItem
Dim aListSubItem As ListSubItem
Dim nPart As ListItemHitTestPartConstants
Dim nResult As ListItemHitTestSuccessConstants
    
On Error GoTo Fehler
    
  nResult = ToolKitsModule.BaseToolKit.Controls.ListView.ListItemHitTest _
  (lvwClassEdit, x, Y, aListItem, aListSubItem, nPart, True)
    
  Set lvwClassEdit.SelectedItem = aListItem
  ShowInputBox aListItem
    
  Select Case True
  Case (nResult = lhsNone) Or (nResult = lhsListItem)
  '{Wenn Spalte Check und Icon geklickt}
  Case (aListSubItem.Key = "Value") And (nPart = lhpIcon)
    Select Case aListSubItem.ReportIcon
    Case "Blank"
      ShowInputBox aListItem
    Case "Memo" '{*}
      ShowMemo aListItem.Tag
    Case "Select" '{*}
      DoSelect aListItem.Tag
    Case "URL"
      OpenURL aListSubItem.Text
    Case "Mail"
      SendMail aListSubItem.Text
    Case "PhoneNumber"
      CallNumber aListSubItem.Text, Me.ClipDialPath
    End Select
  Case (aListSubItem.Key = "Value")
    RaiseEvent ItemClick(aListItem.Tag, aListItem.Tag.Name)
    ShowInputBox aListItem
    txtInput.SetFocus
  End Select
  Exit Sub
  
Fehler:
  ShowError "ActionSelection"
  Exit Sub
End Sub

Private Sub ShowMemo(ByVal aItem As ClassEditInfos)
  Dim f As ClassEditMemoDialog: Set f = New ClassEditMemoDialog
  f.ShowMemo aItem.Name, txtInput.Locked, aItem.ShownValue
  PrepareUpdate aItem, f.Text, , True
  
End Sub

Public Sub SendMail(ByVal strTo As String)
      
  ToolKitsModule.BaseToolKit.Communication.mail.OpenOutlookClient "", strTo, "", "", "", "", ""
End Sub

Public Sub CallNumber(ByVal strPhoneNumber As String, ByVal ClipDialPath As String)

On Error GoTo Fehler

  ToolKitsModule.BaseToolKit.Communication.Phone.CallNumber strPhoneNumber, ClipDialPath
  Exit Sub
  
Fehler:
  Exit Sub
End Sub

Public Sub OpenURL(ByVal strURL As String)
  Shell ToolKitsModule.BaseToolKit.Win32API.Win32ApiProfessional.SysRegistry.WebBrowserExe & " " & strURL, vbNormalFocus
End Sub

Private Sub DoSelect(ByRef aItem As ClassEditInfos)

  With ToolKitsModule.BaseToolKit.Dialog.SelectEntry
    .Reset
    .SelectEntry aItem.SelectData, aItem.Alias, aItem.ReturnID, True
    If .ValueEntry = "" Then Exit Sub
    PrepareUpdate aItem, .ValueEntry, .ValueID, True
  End With
End Sub

Private Sub ShowInputBox(ByRef SelectedListItem As MSComctlLib.ListItem)

Dim lngColumnTextHeight As Long
Dim lngCellTextHeight As Long
Dim lngVisibleCell As Long
Dim i As Long

Dim lngTop As Long
Dim lngLeft As Long
Dim lngHeight As Long
Dim lngWidth As Long
Dim lngBorderSize As Long

On Error GoTo Fehler

  If SelectedItem Is Nothing Then Exit Sub

  Select Case lvwClassEdit.Font.Size
  Case 1 To 9
    lngBorderSize = 25
  Case Is > 9
    lngBorderSize = 55
  End Select

  lngColumnTextHeight = SelectedListItem.Height + lngBorderSize
  lngCellTextHeight = SelectedListItem.Height
  
  For i = lvwClassEdit.GetFirstVisible.Index To lvwClassEdit.ListItems.Count
    If SelectedListItem.Index = lvwClassEdit.ListItems(i).Index Then Exit For
    lngVisibleCell = lngVisibleCell + 1
  Next i
 
  lngTop = (lngVisibleCell * lngCellTextHeight) + lngColumnTextHeight
  lngLeft = lvwClassEdit.ColumnHeaders("Value").Left + 330
  lngHeight = lngCellTextHeight
  lngWidth = lvwClassEdit.ColumnHeaders("Value").Width + 30

  SetParent txtInput.hwnd, lvwClassEdit.hwnd

  txtInput.Top = lngTop
  txtInput.Left = lngLeft
  txtInput.Height = lngHeight - 30
  txtInput.Width = lngWidth
  txtInput.Font.Size = lvwClassEdit.Font.Size
  
  txtInput.Text = SelectedListItem.ListSubItems("Value").Text
  txtInput.Visible = True
  
  Set txtInput.Font = lvwClassEdit.Font
  
  If SelectedListItem.Tag.SpecialType = eceiSPassword Then
    txtInput.PasswordChar = "*"
  Else
    txtInput.PasswordChar = ""
  End If
  Exit Sub
  
Fehler:
  ShowError "ShowInputBOx"
  Exit Sub
End Sub

Private Sub tmrResize_Timer()
Dim lngLeft As Long
Dim lngWidth As Long
  
  lngLeft = lvwClassEdit.ColumnHeaders("Value").Left + 330
  lngWidth = lvwClassEdit.ColumnHeaders("Value").Width + 30
  txtInput.Left = lngLeft
  txtInput.Width = lngWidth

End Sub

Private Sub txtInput_GotFocus()
'Prepare???
End Sub

Private Sub PrepareItemEnter()
Dim aItem As ClassEditInfos

  Set aItem = lvwClassEdit.SelectedItem.Tag

  txtInput.MaxLength = aItem.MaxLenght
  txtInput.Text = aItem.Value
  txtInput.SelStart = 0
  txtInput.SelLength = Len(txtInput.Text)
  
  Select Case True
  Case (aItem.SelectOnlyFromSelectData) And (aItem.SpecialType = eceiSSelect)
    txtInput.Locked = True
  Case Not aItem.Writeable
    txtInput.Locked = True
  Case Else
    txtInput.Locked = False
  End Select
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
Dim aItem As ClassEditInfos

  Select Case KeyCode
  Case 27 '{esc}
    Set aItem = lvwClassEdit.SelectedItem.Tag
    txtInput.Text = aItem.ShownValue
    txtInput.SelStart = Len(txtInput.Text)
  Case 13 '{Enter}
    Set aItem = lvwClassEdit.SelectedItem.Tag
    PrepareUpdate aItem, txtInput.Text
    lvwClassEdit.SetFocus
  End Select
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
  PrepareUpdate lvwClassEdit.SelectedItem.Tag, txtInput.Text
End Sub

Private Sub PrepareUpdate _
(ByRef aItem As ClassEditInfos _
, ByVal strValue As String _
, Optional ByVal strValueID As String = "" _
, Optional blnHideInput As Boolean = False)

Dim Cancel As Boolean

On Error GoTo Fehler

  RaiseEvent BeforeUpdate(aItem, txtInput.Text, Cancel)
  
  Select Case Cancel
  Case False
    
    Select Case aItem.Writeable
    Case True '{Wenn nicht Schreibgeschützt}
      Select Case CheckInputOK(aItem, strValue) '{Datentyp prüfen}
      Case True '{Datentyp OK}
        Select Case CheckMinMaxOK(aItem, strValue) '(Grenzen prüfen)
        Case True '{Eingabegrenzen OK}
                            
          '{Mit Berücksichtigung der Formatierung Wert übernehmen}
          
          If aItem.Format = "" Then
            aItem.Value = strValue
          Else
            aItem.Value = Trim(Format(strValue, aItem.Format))
          End If
          
          aItem.ValueID = Trim(strValueID)
          
          If aItem.ReturnID Then
            CallByName mObj, aItem.Name, VbLet, strValueID '{Wenn ID zurückgegeben werden soll}
          Else
            CallByName mObj, aItem.Name, VbLet, strValue '{Wenn Wert zurückgegeben werden soll}
          End If
          
          '{Aufruf der Savemethode, wenn SaveMethod angegeben}
          '{und AutoCallSaveMethod auf True}
          If Me.AutoCallSaveMethod Then CallSaveMethod
            
        End Select
      End Select
    End Select
    
    txtInput.Text = aItem.ShownValue
    txtInput.Visible = False ' Not blnHideInput
    lvwClassEdit.ListItems(aItem.Name).ListSubItems("Value").Text = aItem.ShownValue
      
    RaiseEvent AfterUpdate(aItem)
  End Select

  Exit Sub
  
Fehler:
  ShowError "PrepareUpdate"
  Exit Sub
End Sub

Private Sub UserControl_Resize()

  CheckUpDown
  
  cmdDown.Left = UserControl.ScaleWidth - (cmdDown.Width + 40)
  cmdUp.Left = cmdDown.Left - (cmdUp.Width + 10)
  
  lvwClassEdit.Width = UserControl.ScaleWidth
  
  If lvwClassEdit.ListItems.Count = 0 Then Exit Sub
  lvwClassEdit.Height = UserControl.ScaleHeight
End Sub

Private Sub CheckUpDown()
Dim aListItem As MSComctlLib.ListItem
  
  Set aListItem = lvwClassEdit.SelectedItem
  
  If aListItem Is Nothing Then
    cmdUp.Enabled = False
    cmdDown.Enabled = False
  Else
    cmdUp.Enabled = aListItem.Index > 1
    cmdDown.Enabled = ToolKitsModule.BaseToolKit.Controls.ListView.GetLastVisibleListItem _
    (lvwClassEdit).Index < lvwClassEdit.ListItems.Count
  End If
End Sub

Private Sub QuickSort _
(ByRef avntArray As Variant _
, ByVal lngVon As Long _
, ByVal lngBis As Long _
, ByVal Sort As eceSort)

Dim i As Long
Dim j As Long
Dim vntTestWert As Variant
Dim intMitte As Long
Dim vntTemp As Variant
Dim astrSortIndex() As String
Dim astrSortIndexF() As String

  If lngVon < lngBis Then
    intMitte = (lngVon + lngBis) \ 2
    
    vntTestWert = avntArray(intMitte)
    
    i = lngVon
    j = lngBis
    
    Do
    
      If Sort = ceSortIndex Then
      
        astrSortIndex = Split(vntTestWert, "{#~#}")
        astrSortIndexF = Split(avntArray(i), "{#~#}")
        
        While CLng(astrSortIndexF(0)) < CLng(astrSortIndex(0))
          i = i + 1
          astrSortIndexF = Split(avntArray(i), "{#~#}")
        Wend
        
        astrSortIndexF = Split(avntArray(j), "{#~#}")
        
        While CLng(astrSortIndexF(0)) > CLng(astrSortIndex(0))
          j = j - 1
          astrSortIndexF = Split(avntArray(j), "{#~#}")
        Wend
      Else
        While CStr(avntArray(i)) < CStr(vntTestWert)
          i = i + 1
        Wend
        While CStr(avntArray(j)) > CStr(vntTestWert)
          j = j - 1
        Wend
      End If
      
      If i <= j Then
        vntTemp = avntArray(j)
        avntArray(j) = avntArray(i)
        avntArray(i) = vntTemp
        i = i + 1
        j = j - 1
      End If
    Loop Until i > j
    
    If j <= intMitte Then
      QuickSort avntArray, lngVon, j, Sort
      QuickSort avntArray, i, lngBis, Sort
    Else
      QuickSort avntArray, i, lngBis, Sort
      QuickSort avntArray, lngVon, j, Sort
    End If
  End If
  
End Sub

Private Sub BuildListViewHeader()
  lvwClassEdit.ColumnHeaders.Clear
  lvwClassEdit.ColumnHeaders.Add 1, "Member", "Eigenschaft"
  lvwClassEdit.ColumnHeaders.Add 2, "Value", "Wert"
  
  lvwClassEdit.ListItems.Clear
  lvwClassEdit.SmallIcons = ilsBilder
  lvwClassEdit.Height = 0
End Sub

Private Function IsReal(ByRef aItem As ClassEditInfos, ByVal strValue As Variant) As Boolean
  Select Case IsNumeric(strValue)
  Case True
    IsReal = True
  Case False
    IsReal = False
    MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
    txtInput.Text = Format(aItem.Value, aItem.Format)
  End Select

End Function

Private Function IsInteger(ByRef aItem As ClassEditInfos, ByVal strValue As Variant) As Boolean
  Select Case IsNumeric(strValue)
  Case True
    IsInteger = True
    txtInput.Text = CInt(txtInput.Text)
  Case False
    IsInteger = False
    MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
    txtInput.Text = Format(aItem.Value, aItem.Format)
  End Select
End Function


Private Function IsDateEx(ByRef aItem As ClassEditInfos, ByVal strValue As Variant) As Boolean
    
  Select Case True
  Case IsDate(strValue)
    IsDateEx = True
    
  Case (aItem.AllowNullString) And (strValue = "")
    IsDateEx = True
  
  Case Else
    IsDateEx = False
    MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
    txtInput.Text = Format(aItem.Value, aItem.Format)
  End Select
End Function

Private Function IsTime(ByRef aItem As ClassEditInfos, ByVal strValue As Variant) As Boolean
  Select Case IsDate(strValue)
  Case True
    IsTime = True
    txtInput.Text = Format(txtInput.Text, "hh:mm")
  Case False
    IsTime = False
    MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
    txtInput.Text = Format(Format(aItem.Value, "hh:mm"), aItem.Format)
  End Select
End Function
'{Ende Private Klassen-Methoden}


'{Öffentliche Klassen-Methoden}
Public Sub SetSaveMethod(ByVal strSaveMethodName As String, ParamArray Parameters())
Attribute SetSaveMethod.VB_Description = "Legt die Speicherroutine des Objektes fest, welche bei jeder Änderung der Eigenschaften aufgerufen wird. Es können auch alle notwendigen Parameter für die Speicherroutine angegeben werden."
  mstrSaveMethodName = strSaveMethodName
  mvarSaveMethodParams = Parameters
End Sub

Public Sub GetItems(ByRef obj As Object)
Attribute GetItems.VB_Description = "Lädt alle Eigenschaften eines Objektes in die Eigenschaften-Auflistung."
    Dim AllMemberInfos As ToolKits.ReflectionAllMemberInfos
    Dim MemberInfos    As ToolKits.ReflectionMemberInfos
    Dim aItem          As ClassEditInfos
    Dim lngIndex       As Long
    Dim blnNewMember   As Boolean

    On Error GoTo Fehler

    Set Items = New Collection
    Set AllMemberInfos = New ToolKits.ReflectionAllMemberInfos

    AllMemberInfos.GetMembers obj
  
    For Each MemberInfos In AllMemberInfos.AllMemberInfos
    
        Select Case MemberInfos.MemberType

            Case INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT

                blnNewMember = Not BaseToolKit.Etc.CollectionContainsKey(Items, MemberInfos.MemberName)
      
                If blnNewMember Then
                    lngIndex = lngIndex + 1
        
                    Set aItem = New ClassEditInfos
                    aItem.Name = MemberInfos.MemberName
                    aItem.Alias = MemberInfos.MemberName
                    aItem.Value = MemberInfos.Value
                    aItem.SortIndex = lngIndex
        
                    Items.Add aItem, aItem.Name
                Else
                    Set aItem = Items.Item(MemberInfos.MemberName)
                End If

                Select Case MemberInfos.MemberType

                    Case INVOKE_PROPERTYGET
                        aItem.Readable = True

                    Case INVOKE_PROPERTYPUT
                        aItem.Writeable = True
                End Select
      
            Case Else
        End Select

    Next MemberInfos

    Set mObj = obj

    Exit Sub
  
Fehler:
    ShowError "GetItems"

    Exit Sub

End Sub

Public Sub Sort(ByVal Sort As eceSort)
Attribute Sort.VB_Description = "Sortiert die Eigenschaftenansicht nach Eigenschaftennamen (ceItemName) oder nach dem SortIndex (ceSortIndex)."
Dim astrNewOrder() As String
Dim astrHelp() As String
Dim colTemp As Collection
Dim ClassEditInfos As ClassEditInfos
Dim lngIndex As Long
Dim i As Long

  '{Felder für alle Member reservieren}
  ReDim astrNewOrder(1 To Items.Count)
  
  '{Feld mit der jetzigen Reihenfolge belegen}
  For Each ClassEditInfos In Items
  
    i = i + 1
    Select Case Sort
    Case ceItemName
      astrNewOrder(i) = ClassEditInfos.Alias & "{#~#}" & i
    Case ceSortIndex
      astrNewOrder(i) = ClassEditInfos.SortIndex & "{#~#}" & i
    End Select
  
  Next ClassEditInfos
      
  '{Feld sortieren}
  QuickSort astrNewOrder, 1, Items.Count, Sort

  Set colTemp = New Collection
  
  '{Collection über Temp-Collection umschreiben}
  For i = LBound(astrNewOrder) To UBound(astrNewOrder)
    astrHelp = Split(astrNewOrder(i), "{#~#}")
    lngIndex = CLng(astrHelp(1))
    Set ClassEditInfos = Items.Item(lngIndex)
    
    colTemp.Add ClassEditInfos, ClassEditInfos.Name
  Next i

  Set Items = colTemp
  Set colTemp = Nothing

End Sub

Public Sub ShowItems()
Attribute ShowItems.VB_Description = "Zeigt die Eigenschaften in der Eigenschaftenansicht an."
Dim aItem As ClassEditInfos
Dim aListItem As MSComctlLib.ListItem
  
On Error GoTo Fehler
  
  lvwClassEdit.ListItems.Clear
  
  For Each aItem In Items
  
   
    If aItem.Visible Then
            
      Set aListItem = lvwClassEdit.ListItems.Add
      aListItem.Text = aItem.Alias
      aListItem.Key = aItem.Name
      
      Select Case True
      Case aItem.Readable And aItem.Writeable
        aListItem.SmallIcon = "FullAccess"
      Case aItem.Readable
        aListItem.SmallIcon = "Readable"
      Case aItem.Writeable
        aListItem.SmallIcon = "Writeable"
      End Select
      
      Set aListItem.Tag = aItem
      aListItem.ListSubItems.Add 1, "Value", "Wert"
      
      aListItem.ListSubItems("Value").Text = Format(aItem.ShownValue, aItem.Format)
            
      Select Case aItem.SpecialType
      Case eceiSPassword
        aListItem.ListSubItems("Value").ReportIcon = "Password"
      Case eceiSeMail
        aListItem.ListSubItems("Value").ReportIcon = "Mail"
      Case eceiSMemo
        aListItem.ListSubItems("Value").ReportIcon = "Memo"
      Case eceiSSelect
        aListItem.ListSubItems("Value").ReportIcon = "Select"
      Case eceiSURL
        aListItem.ListSubItems("Value").ReportIcon = "URL"
      Case eceiSPhoneNumber
        aListItem.ListSubItems("Value").ReportIcon = "PhoneNumber"
      Case Else
        Select Case Me.AutoDetectSpecialType
        Case True
          Select Case True
          Case ToolKitsModule.BaseToolKit.Communication.mail.CheckMailAdress(aItem.Value)
            aListItem.ListSubItems("Value").ReportIcon = "Mail"
          Case InStr(LCase(aItem.Value), "http://") > 0
            aListItem.ListSubItems("Value").ReportIcon = "URL"
          Case InStr(LCase(aItem.Value), "ftp://") > 0
            aListItem.ListSubItems("Value").ReportIcon = "URL"
          Case Else
            aListItem.ListSubItems("Value").ReportIcon = "Blank"
          End Select
          
        Case False
          aListItem.ListSubItems("Value").ReportIcon = "Blank"
        End Select
      End Select
          
    End If
  Next aItem
  
  With lvwClassEdit
    .View = lvwReport
    ToolKitsModule.BaseToolKit.Controls.ListView.ResizeColumns lvwClassEdit, , True, True
    .Height = UserControl.ScaleHeight
    txtInput.Visible = False
  
    Set lvwClassEdit.SelectedItem = .ListItems(1)

    cmdUp.Visible = True
    cmdUp.Enabled = False
  
    cmdDown.Visible = True
    cmdDown.Enabled = ToolKitsModule.BaseToolKit.Controls.ListView.VisibleItemsCount _
    (lvwClassEdit) < lvwClassEdit.ListItems.Count
  End With
  Exit Sub
  
Fehler:
  ShowError "ShowItems"
  Exit Sub
End Sub

Public Sub SetValue _
(ByVal strItemName As String _
, ByVal strValue As String _
, Optional ByVal strValueID As String = "")
Attribute SetValue.VB_Description = "Setzt per Code den Wert einer Eigenschaft."

Dim aItem As ClassEditInfos

  Set aItem = Items.Item(strItemName)
  PrepareUpdate aItem, strValue, strValueID, True
    
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "Löscht die Eigenschaftenansicht."
  lvwClassEdit.ListItems.Clear
  ToolKitsModule.BaseToolKit.Controls.ListView.ResizeColumns lvwClassEdit, , True, True
  lvwClassEdit.Height = 0
  txtInput.Visible = False
  cmdUp.Visible = False
  cmdDown.Visible = False
End Sub

Public Sub CallSaveMethod()
Attribute CallSaveMethod.VB_Description = "Ruft die Speichermethode des Objektes manuell auf."
Dim x As Variant
Dim colX  As Collection
    
On Error GoTo Fehler

  If Me.SaveMethodName = "" Then Exit Sub
    
  Set colX = New Collection
  For Each x In mvarSaveMethodParams
    colX.Add x
  Next x

  Select Case UBound(mvarSaveMethodParams) + 1
  Case 0
    CallByName mObj, mstrSaveMethodName, VbMethod
  Case 1
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1)
  Case 2
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2)
  Case 3
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3)
  Case 4
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3) _
    , colX(4)
  Case 5
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3) _
    , colX(4) _
    , colX(5)
  Case 6
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3) _
    , colX(4) _
    , colX(5) _
    , colX(6)
  Case 7
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3) _
    , colX(4) _
    , colX(5) _
    , colX(6) _
    , colX(7)
  Case 8
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3) _
    , colX(4) _
    , colX(5) _
    , colX(6) _
    , colX(7) _
    , colX(8)
  Case 9
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3) _
    , colX(4) _
    , colX(5) _
    , colX(6) _
    , colX(7) _
    , colX(8) _
    , colX(9)
  Case 10
    CallByName mObj, mstrSaveMethodName, VbMethod _
    , colX(1) _
    , colX(2) _
    , colX(3) _
    , colX(4) _
    , colX(5) _
    , colX(6) _
    , colX(7) _
    , colX(8) _
    , colX(9) _
    , colX(10)
  End Select
  Exit Sub
  
Fehler:
  ShowError "CallSaveMethod"
  Exit Sub
End Sub

Public Function CheckRequiredFields() As ClassEditInfos
Dim aItem As ClassEditInfos

On Error GoTo Fehler

  Set CheckRequiredFields = New ClassEditInfos
  For Each aItem In Me.Items
    If aItem.Required Then
      If Trim(aItem.Value) = "" Then
        Set CheckRequiredFields = aItem
        Exit For
      End If
    End If
  Next aItem
  Exit Function
  
Fehler:
  ShowError "CheckRequiredFields"
  Exit Function
End Function
'{Ende Öffentliche Klassen-Methoden}


Public Property Get ClipDialPath() As String
  ClipDialPath = mClipDialPath
End Property

Public Property Let ClipDialPath(ByVal Value As String)
  mClipDialPath = Value
End Property
