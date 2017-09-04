VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl PropertiesViewControl 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ScaleHeight     =   3600
   ScaleWidth      =   7095
   ToolboxBitmap   =   "PropertiesViewControl.ctx":0000
   Begin VB.CommandButton cmdMemo 
      Caption         =   "..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrColResize 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   2880
   End
   Begin VB.TextBox txtEingabeFokus 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboEingabeFokus 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "cboEingabeFokus"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid flexEigenschaften 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "PropertiesViewControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'{Moduleigenschaften}
Private mReader As PropertiesViewReader

Private mstrSeparator As String
Private mblnUpdateEqualValues As Boolean

'{Modulvariablen}
Private mProperties() As typProperties
Private mintWidthCol_1 As Integer
Private mintWidthCol_2 As Integer
Private mblnCancelUpdate As Boolean
Private mlngMinWidth As Long

'{Globale-Moduleigenschaft}
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Public Event CellUpdated(ByVal intFieldIndex As Integer)
Public Event BeforeCellUpdate(ByVal intFieldIndex As Integer _
, ByRef Cancel As Boolean)

Public Event EnterCell(ByVal intFieldIndex As Integer)
Public Event LeaveCell(ByVal intFieldIndex As Integer)

Public Property Get FocusText() As String
Dim oEingabeFokus As Object

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  FocusText = oEingabeFokus.Text
End Property

Public Property Let FocusText(ByVal strFocusText As String)
Dim oEingabeFokus As Object

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  oEingabeFokus.Text = strFocusText
End Property

Public Property Get MinWidth() As Long
  MinWidth = mlngMinWidth
End Property

Public Property Let MinWidth(ByVal lngMinWidth As Long)
  mlngMinWidth = lngMinWidth
End Property

'{Prüft ob ComboBox umgeladen werden muß}
Private Sub AskRefill()
  With mProperties(flexEigenschaften.Row)
    If cboEingabeFokus.Tag = .DataString Then Exit Sub
    cboEingabeFokus.Tag = .DataString
    FillComboBox .DataString
  End With
End Sub

'{Prüft ob eine Eingabe ein Datum ist}
Private Function CheckDate _
(ByRef strValue As String _
, Optional blnNullOK As Boolean = False) As Boolean
Dim ValueArea As Variant
  
On Error GoTo Fehler
  
  If blnNullOK Then
    If strValue = "" Then
      CheckDate = True
      Exit Function
    End If
  End If
  
  If mProperties(flexEigenschaften.Row).ValueArea <> "" Then
    ValueArea = Split(mProperties(flexEigenschaften.Row).ValueArea, "#")
    strValue = mReader.datRead( _
    PruefWert:=strValue, _
    Default:="False", _
    Min:=ValueArea(0), _
    Max:=ValueArea(1))
  Else
    strValue = mReader.datRead(PruefWert:=strValue, Default:="False")
  End If

  If strValue <> "False" Then
    CheckDate = True
  Else
    Controls(GetCurrentEingabeFokus).Text _
    = flexEigenschaften.Text
    Controls(GetCurrentEingabeFokus).SetFocus
    CheckDate = False
  End If
  Exit Function
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "CheckDate"
  Exit Function
End Function

'{Prüft ob eine Eingabe eine Fließkommazahl ist}
Private Function CheckReal(ByRef strValue As String) As Boolean
Dim ValueArea As Variant
Dim blnError As Boolean
  
On Error GoTo Fehler
  
  If mProperties(flexEigenschaften.Row).ValueArea <> "" Then
    ValueArea = Split(mProperties(flexEigenschaften.Row).ValueArea, "#")
    
    If UBound(ValueArea) = 2 Then
      strValue = mReader.dblRead( _
      PruefWert:=strValue, _
      Min:=ValueArea(0), _
      Max:=ValueArea(1), _
      Dez:=ValueArea(2), _
      blnErr:=blnError)
    Else
      strValue = mReader.dblRead( _
      PruefWert:=strValue, _
      Min:=ValueArea(0), _
      Max:=ValueArea(1), _
      blnErr:=blnError)
    End If
  Else
      strValue = mReader.dblRead( _
      PruefWert:=strValue, _
      blnErr:=blnError)
  End If

  If Not blnError Then
    CheckReal = True
  Else
    Controls(GetCurrentEingabeFokus).Text _
    = flexEigenschaften.Text
    Controls(GetCurrentEingabeFokus).SetFocus
    CheckReal = False
  End If
  Exit Function
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "CheckReal"
  Exit Function
End Function

'{Prüft ob eine Eingabe eine ganze Zahl ist}
Private Function CheckInteger(ByRef strValue As String) As Boolean
Dim ValueArea As Variant
Dim blnError As Boolean
  
  On Error GoTo Fehler
  
  If mProperties(flexEigenschaften.Row).ValueArea <> "" Then
    ValueArea = Split(mProperties(flexEigenschaften.Row).ValueArea, "#")
    
    strValue = mReader.lngRead( _
    PruefWert:=strValue, _
    Min:=ValueArea(0), _
    Max:=ValueArea(1), _
    blnErr:=blnError)
    
  Else
      strValue = mReader.lngRead( _
      PruefWert:=strValue, _
      blnErr:=blnError)
  End If

  If Not blnError Then
    CheckInteger = True
  Else
    Controls(GetCurrentEingabeFokus).Text _
    = flexEigenschaften.Text
    Controls(GetCurrentEingabeFokus).SetFocus
    CheckInteger = False
  End If
  Exit Function
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "CheckInteger"
  Exit Function
End Function

Public Sub Clear()
  ReDim mProperties(0)
  flexEigenschaften.Rows = 0
  txtEingabeFokus.Visible = False
  cboEingabeFokus.Visible = False
  cmdMemo.Visible = False
  cboEingabeFokus.Text = ""
End Sub

'{Kreiert einen String mit allen Feldeigenschaften}
Private Function CreatePropertiesString() As String
Dim i As Integer

On Error GoTo Fehler

  With flexEigenschaften
    CreatePropertiesString = .Rows - 1
    For i = 1 To .Rows - 1
      CreatePropertiesString = CreatePropertiesString _
      & Me.Separator & .TextMatrix(i, 1)
    Next i
  End With
  Exit Function
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "CreatePropertiesString"
  Exit Function
End Function

'{Füllt das FlexGrid mit Feldeigenschaften}
Private Sub FillFlexGrid()
Dim i As Integer
Dim strHeader As String
Dim strEntry As String
Dim strValue As String

On Error GoTo Fehler

  tmrColResize.Enabled = False
  
  Screen.MousePointer = 11

  txtEingabeFokus.Visible = False
  cboEingabeFokus.Visible = False
  
  With flexEigenschaften
    .Font.Size = 10
    .Rows = 0
    .Cols = 2
    
    .ColAlignment(-1) = flexAlignLeftCenter
    .RowHeight(-1) = 315
    strHeader = "Eigenschaft" & vbTab & "Wert" & vbTab
    .AddItem strHeader
    
    For i = LBound(mProperties) To UBound(mProperties)
      strValue = FormatCell(mProperties(i).value, i)
      strEntry = mProperties(i).Field & vbTab & strValue & vbTab
      .AddItem strEntry
    Next i

    ToolKitsModule.BaseToolKit.Controls.flexGrid.ResizeColumns PropertiesViewResizeDialog, flexEigenschaften
    
    If (Me.MinWidth > 0) _
    And (flexEigenschaften.ColWidth(1) < Me.MinWidth) Then
      flexEigenschaften.ColWidth(1) = Me.MinWidth
    End If

    CheckMaxWidth
    
    .FixedRows = 1
    .FixedCols = 1
        
    SetEingabeFokusPosition GetCurrentEingabeFokus
    
    mintWidthCol_1 = .ColWidth(0)
    mintWidthCol_2 = .ColWidth(1)
    tmrColResize.Enabled = True
  
  End With
  Screen.MousePointer = 0
  Exit Sub
  
Fehler:
  MsgBox "(" & Err.Number & ")" & Err.Description _
  , 16, Name & "#FillFlexGrid"
  Exit Sub
End Sub

Private Sub CheckMaxWidth()
  '{Überprüfen ob maximale Spaltenbreite überschritten}
  With flexEigenschaften
    If .ColWidth(0) + .ColWidth(1) > .Width Then
      If .Width - (.ColWidth(0) + 250) > 0 Then
        .ColWidth(1) = .Width - (.ColWidth(0) + 250)
      End If
    End If
  End With
End Sub

'{Umladen der ComboBox}
Private Sub FillComboBox(ByVal DataString As String)
Dim i As Integer
Dim astrItems() As String
Dim intStep As Integer

On Error GoTo Fehler

  Screen.MousePointer = 11

  With cboEingabeFokus
    .Clear
      
    If mProperties(flexEigenschaften.Row).PutID Then
      intStep = 2
    Else
      intStep = 1
    End If
    
    astrItems = Split(DataString, "#")
    For i = LBound(astrItems) To UBound(astrItems) Step intStep
      .AddItem astrItems(i)
      If intStep = 2 Then
        .ItemData(.NewIndex) = astrItems(i + 1)
      End If
    Next i
    
  End With
  Screen.MousePointer = 0
  Exit Sub
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "FillComboBox"
  Exit Sub
End Sub

Private Function FormatCell _
(ByVal strValue As String _
, ByVal CellIndex As Integer) As String

Dim strFormatMask As String
Dim strFormatType As String
      
  If InStr(LCase(mProperties(CellIndex).FormatType) _
  , "uformat") > 0 Then
    strFormatType = mProperties(CellIndex).FormatType
    strFormatMask = Mid(strFormatType, InStr(strFormatType, "#") + 1)
  End If
  
  Select Case Mid(LCase(mProperties(CellIndex).FormatType), 1, 4)
  Case "norm" '{normal}
    strValue = strValue
  Case "lcas" '{lcase}
    strValue = LCase(strValue)
  Case "ucas" '{uCase}
    strValue = UCase(strValue)
  Case "ufor" '{uformat}
    strValue = Format(strValue, strFormatMask)
  End Select
  
  FormatCell = strValue
  
End Function


Public Sub PrintPropertiesView _
(ByVal LeftMargin As Single _
, ByVal TopMargin As Single _
, ByVal RightMargin As Single _
, ByVal BottomMargin As Single _
, Optional ByVal PrintTitle As String = "" _
, Optional ByVal PrintDate As String = "")

  ToolKitsModule.BaseToolKit.Controls.flexGrid.PrintData flexEigenschaften _
  , LeftMargin, TopMargin, RightMargin, BottomMargin _
  , PrintTitle, PrintDate
End Sub


'{Aktualisiert selektierte Zelle}
Private Sub SetCellValue(ByVal strValue As String)
  With flexEigenschaften
    .Text = strValue
  End With
  ToolKitsModule.BaseToolKit.Controls.flexGrid.ResizeColumns PropertiesViewResizeDialog, flexEigenschaften, 1
  
  If (Me.MinWidth > 0) _
  And (flexEigenschaften.ColWidth(1) < Me.MinWidth) Then
    flexEigenschaften.ColWidth(1) = Me.MinWidth
  End If
  
  CheckMaxWidth
End Sub

'{Positioniert den Eingabefokus und macht ihn schreibbereit}
Private Sub SetEingabeFokusPosition(ByVal strEingabeFokus As String)

On Error GoTo Fehler

  With Controls(strEingabeFokus)
    txtEingabeFokus.Visible = False
    cboEingabeFokus.Visible = False
    
    If LCase(mProperties(flexEigenschaften.Row).EingabeFokus) _
    = "textbox" Then
      .MaxLength = mProperties(flexEigenschaften.Row).MaxLen
      .Text = flexEigenschaften.Text
    Else
      AskRefill
      flexEigenschaften.Tag = ToolKitsModule.BaseToolKit.Controls.ComboBox.ItemInComboBox(cboEingabeFokus, flexEigenschaften.Text)
      cboEingabeFokus.ListIndex = flexEigenschaften.Tag
    End If
    
    
    .Top = flexEigenschaften.Top + flexEigenschaften.CellTop
    .Left = flexEigenschaften.Left + flexEigenschaften.CellLeft
    .Width = flexEigenschaften.CellWidth
        
    .Visible = True
    
    If CBool(mProperties(flexEigenschaften.Row).ReadOnly) Then
      PropertiesViewMemoDialog.txtMemo.Locked = True
      .Locked = True
    Else
      PropertiesViewMemoDialog.txtMemo.Locked = False
      .Locked = False
    End If
    
    If LCase(mProperties(flexEigenschaften.Row).ValueType) _
    = "memo" Then
      cmdMemo.Visible = True
      cmdMemo.Top = txtEingabeFokus.Top + 20
      cmdMemo.Left = txtEingabeFokus.Left _
      + txtEingabeFokus.Width - cmdMemo.Width
      .Width = flexEigenschaften.CellWidth - (cmdMemo.Width + 20)
    Else
      cmdMemo.Visible = False
      .Width = flexEigenschaften.CellWidth
    End If
  
    If LCase(mProperties(flexEigenschaften.Row).ValueType) _
    = "password" Then
      txtEingabeFokus.PasswordChar = "*"
    Else
      txtEingabeFokus.PasswordChar = ""
    End If
    
  End With
  Exit Sub
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "SetEingabeFokus"
  Exit Sub
End Sub

'{Prüft Eingabewerte und aktualisiert DB und Zelle}
Private Sub UpdateCell(ByVal strValue As String)
Dim strOldVAlue As String
Dim blnUpdate As Boolean
Dim oEingabeFokus As Object
Dim lngIndex As Long

On Error GoTo Fehler

  With flexEigenschaften
  
    Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
    
    '{Wenn Eingabefokus eine Combo}
    If oEingabeFokus.Name = "cboEingabeFokus" Then
      '{Wenn Auswahl aus Box erfolgen mußte}
      If LCase(mProperties(.Row).EingabeFokus) = "combobox" Then
        lngIndex = ToolKitsModule.BaseToolKit.Controls.ComboBox.ItemInComboBox(cboEingabeFokus, cboEingabeFokus.Text)
        '{Wenn kein Wert aus Box selektiert}
        If lngIndex = -1 Then
          '{Wert zurücksetzen}
          oEingabeFokus.ListIndex = ToolKitsModule.BaseToolKit.Controls.ComboBox.ItemInComboBox(cboEingabeFokus, .Text)
          Exit Sub
        Else
          '{Ansonsonsten ListIndex richtig setzen}
          oEingabeFokus.ListIndex = lngIndex
        End If
      End If
    End If
    
    '{Wurde alter Eintrag geändert}
    If UpdateEqualValues Then
      blnUpdate = True
    Else
      blnUpdate = .Text <> strValue
    End If
    
    If blnUpdate Then
    
      '{Update ausführen Standard False setzen}
      mblnCancelUpdate = False
    
      Select Case LCase(mProperties(.Row).EingabeFokus)
      Case "textbox"
        strOldVAlue = .Text
      Case "combobox"
        strOldVAlue = flexEigenschaften.Tag
      Case "comboboxex"
        strOldVAlue = .Text
      End Select
    
      '{Neuen Wert in Zelle setzen (für PropertieString)}
      .Text = oEingabeFokus.Text
    
      '{Ereignis BeforeCellUpdate auslösen}
      RaiseEvent BeforeCellUpdate(.Row, mblnCancelUpdate)
    
      '{Alten Zellenwert zurückschreiben}
      Select Case LCase(mProperties(.Row).EingabeFokus)
      Case "textbox"
        .Text = strOldVAlue
      Case "combobox"
        .Text = cboEingabeFokus.List(strOldVAlue)
      Case "comboboxex"
        .Text = strOldVAlue
      End Select
          
      '{Wurde im Ereignis Cancel gesetzt}
      If mblnCancelUpdate Then
        '{Alten Eingabefokuswert zurücksetzen}
        Select Case LCase(mProperties(.Row).EingabeFokus)
        Case "textbox"
          oEingabeFokus.Text = strOldVAlue
        Case "combobox"
          oEingabeFokus.ListIndex = strOldVAlue
        Case "comboboxex"
          oEingabeFokus.Text = strOldVAlue
        End Select
        Exit Sub
      End If
        
      Select Case LCase(mProperties(.Row).EingabeFokus)
      Case "combobox"
        flexEigenschaften.Tag = oEingabeFokus.ListIndex
      End Select
  
      '{FormatType und ValueType prüfen}
      Select Case Trim(LCase(mProperties(.Row).ValueType))
      Case "string"
      '{Keine weiteren Prüfungen}
      Case "date"
        If Not CheckDate(strValue) Then Exit Sub
      Case "datenull"
        If Not CheckDate(strValue, True) Then Exit Sub
      Case "integer"
        If Not CheckInteger(strValue) Then Exit Sub
      Case "real"
        If Not CheckReal(strValue) Then Exit Sub
      End Select
    
      strValue = FormatCell(strValue, .Row)
    
      If Trim(LCase(mProperties(.Row).ValueType)) = "real" Then
        strValue = Replace(strValue, ",", ".")
      End If
      '{Ende FormatType und ValueType prüfen}
    
      '{Wenn Eingabefokus eine ComboBox und als Wert ein Index aktualisiert werden soll}
      If (mProperties(.Row).PutID) _
      And (LCase(mProperties(.Row).EingabeFokus) = "combobox") Then
        strValue = cboEingabeFokus.ItemData(cboEingabeFokus.ListIndex)
        strOldVAlue = cboEingabeFokus.Text
      End If
    
      '{Eingabefokus aktualisieren}
      If oEingabeFokus.Name = "txtEingabeFokus" Then
        oEingabeFokus.Text = strValue
        '.Text = strValue
      Else
        Select Case LCase(mProperties(.Row).EingabeFokus)
        Case "combobox"
          If (mProperties(.Row).PutID) Then
            oEingabeFokus.ListIndex = ToolKitsModule.BaseToolKit.Controls.ComboBox.ItemInComboBox(cboEingabeFokus, strOldVAlue)
          Else
            oEingabeFokus.ListIndex = ToolKitsModule.BaseToolKit.Controls.ComboBox.ItemInComboBox(cboEingabeFokus, strValue)
          End If
        Case "comboboxex"
          oEingabeFokus.Text = strValue
        End Select
      End If
      
      '{Zelle aktualisieren}
      If oEingabeFokus.Name = "txtEingabeFokus" Then
        SetCellValue strValue
      Else
        If (mProperties(.Row).PutID) Then
          SetCellValue cboEingabeFokus.Text
        Else
          SetCellValue strValue
        End If
      End If
    
      '{Ereignis CellUpdated auslösen}
      RaiseEvent CellUpdated(.Row)
    End If
    On Error Resume Next
    .SetFocus
    On Error GoTo Fehler
  End With
  Exit Sub
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "UpdatCell"
  Exit Sub
End Sub

Private Sub cboEingabeFokus_Click()
  UpdateCell cboEingabeFokus.Text
End Sub

'{Wertet Tastatureingabe aus}
Private Sub cboEingabeFokus_KeyDown(KeyCode As Integer, Shift As Integer)
  
  RaiseEvent KeyDown(KeyCode, Shift)

  Select Case KeyCode
  Case 13 'Enter
    UpdateCell cboEingabeFokus.Text
  Case 27 'ESC
    cboEingabeFokus.ListIndex = ToolKitsModule.BaseToolKit.Controls.ComboBox.ItemInComboBox(cboEingabeFokus, flexEigenschaften.Text)
  End Select
End Sub

Private Sub cboEingabeFokus_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub


Private Sub cboEingabeFokus_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub cboEingabeFokus_Validate(Cancel As Boolean)
  UpdateCell cboEingabeFokus.Text
End Sub

Private Sub cmdMemo_Click()
  Dim f As PropertiesViewMemoDialog: Set f = New PropertiesViewMemoDialog
  f.Memo = txtEingabeFokus.Text
  f.Show 1
  txtEingabeFokus.Text = f.Memo
  UpdateCell txtEingabeFokus.Text
  Set f = Nothing
End Sub

'{Weist Eingabefokus den Fokus zu}
Private Sub flexEigenschaften_Click()
  On Error Resume Next
  Controls(GetCurrentEingabeFokus).SetFocus
End Sub

Private Sub flexEigenschaften_EnterCell()
  RaiseEvent EnterCell(flexEigenschaften.Row)
End Sub

'{Ermittelt den aktuellen Eingabefokus}
Private Function GetCurrentEingabeFokus() As String
  If LCase(mProperties(flexEigenschaften.Row).EingabeFokus) _
  = "textbox" Then
    GetCurrentEingabeFokus = "txtEingabeFokus"
  Else
    GetCurrentEingabeFokus = "cboEingabeFokus"
  End If
End Function

'{Wertet Tastatureingabe aus}
Private Sub flexEigenschaften_KeyDown(KeyCode As Integer, Shift As Integer)
  
  RaiseEvent KeyDown(KeyCode, Shift)

  Select Case KeyCode
  Case 13 'Enter
    Controls(GetCurrentEingabeFokus).SetFocus
  End Select
End Sub

Private Sub flexEigenschaften_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub


Private Sub flexEigenschaften_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub


'{Versucht Wert zu aktualisieren}
Private Sub flexEigenschaften_LeaveCell()
  With Controls(GetCurrentEingabeFokus)
    If flexEigenschaften.Text = .Text Then Exit Sub
    UpdateCell .Text
  End With
  RaiseEvent LeaveCell(flexEigenschaften.Row)
End Sub






Private Sub flexEigenschaften_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub flexEigenschaften_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub flexEigenschaften_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub


Private Sub flexEigenschaften_RowColChange()
  SetEingabeFokusPosition GetCurrentEingabeFokus
End Sub

Private Sub flexEigenschaften_Scroll()
  Controls(GetCurrentEingabeFokus).Visible = False
  cmdMemo.Visible = False
End Sub

Private Sub flexEigenschaften_SelChange()
  With flexEigenschaften
    If .RowSel <> .Row Then
      .Row = .RowSel
    End If
  End With
End Sub

Private Sub tmrColResize_Timer()
  If ((mintWidthCol_1 = flexEigenschaften.ColWidth(0))) _
  And ((mintWidthCol_2 = flexEigenschaften.ColWidth(1))) Then Exit Sub
  
  With Controls(GetCurrentEingabeFokus)
    .Top = flexEigenschaften.Top + flexEigenschaften.CellTop
    .Left = flexEigenschaften.Left + flexEigenschaften.CellLeft
    .Width = flexEigenschaften.CellWidth

    If mProperties(flexEigenschaften.Row).ValueType _
    = "memo" Then
      cmdMemo.Visible = True
      cmdMemo.Top = txtEingabeFokus.Top + 20
      cmdMemo.Left = txtEingabeFokus.Left _
      + txtEingabeFokus.Width - cmdMemo.Width
      .Width = flexEigenschaften.CellWidth - (cmdMemo.Width + 20)
    Else
      cmdMemo.Visible = False
      .Width = flexEigenschaften.CellWidth
    End If

    mintWidthCol_1 = flexEigenschaften.ColWidth(0)
    mintWidthCol_2 = flexEigenschaften.ColWidth(1)
  End With
End Sub


'{Markiert Text im Eingabefokus}
Private Sub txtEingabeFokus_GotFocus()
'*** Text markieren ***'
  With txtEingabeFokus
    .SelStart = 0
    .SelLength = Len(Trim(.Text))
  End With
'*** Ende Text markieren ***'
End Sub

'{Wertet Tastatureingabe aus}
Private Sub txtEingabeFokus_KeyDown(KeyCode As Integer, Shift As Integer)
  
  RaiseEvent KeyDown(KeyCode, Shift)

  Select Case KeyCode
  Case 13 'Enter
    UpdateCell txtEingabeFokus.Text
  Case 27 'ESC
    txtEingabeFokus.Text = flexEigenschaften.Text
  End Select
End Sub

'{Eigenschaft wieviele Felder vorhanden sind}
Public Property Get PropertyCount() As Integer
On Error GoTo NotDefined

  PropertyCount = UBound(mProperties)
  Exit Sub
  
NotDefined:
  PropertyCount = 0
End Property

'{Eigenschaft wieviele Felder vorhanden sind}
Public Property Let PropertyCount(ByVal intValue As Integer)
  
  '{Waren bereits Eigenschften angezeigt selektiertes Feld sichern}
  If flexEigenschaften.Rows > 0 Then
    With Controls(GetCurrentEingabeFokus)
      If flexEigenschaften.Text <> .Text Then
        UpdateCell .Text
      End If
    End With
  End If
  
  ReDim mProperties(1 To intValue)
End Property


'{Fügt Feldinformationen ein}
Public Sub AddField _
(ByVal PropertyNr As Integer _
, ByVal strField As String _
, ByVal strValue As String _
, Optional ByVal EingabeFokus As String = "" _
, Optional ByVal FormatType As String = "" _
, Optional ByVal DataString As String = "" _
, Optional ByVal blnPutID As String = "" _
, Optional ByVal ReadOnly As String = "" _
, Optional ByVal ValueType As String = "" _
, Optional ByVal MaxLen As String = "" _
, Optional ByVal ValueArea As String = "")

On Error GoTo Fehler

  With mProperties(PropertyNr)
    .Field = strField
    .value = strValue
    
    If EingabeFokus = "" Then
      .EingabeFokus = "TextBox"
    Else
      .EingabeFokus = EingabeFokus
    End If
    
    If FormatType = "" Then
      .FormatType = "Normal"
    Else
      .FormatType = FormatType
    End If
  
    If DataString = "" Then
      .DataString = ""
    Else
      .DataString = DataString
    End If
    
    If blnPutID = "" Then
      .PutID = True
    Else
      .PutID = blnPutID
    End If
    
    If ReadOnly = "" Then
      .ReadOnly = False
    Else
      .ReadOnly = ReadOnly
    End If
    
    If ValueType = "" Then
      .ValueType = "String"
    Else
      .ValueType = ValueType
    End If

    If MaxLen = "" Then
      .MaxLen = 0
    Else
      .MaxLen = MaxLen
    End If

    If ValueArea = "" Then
      .ValueArea = ""
    Else
      .ValueArea = ValueArea
    End If
  End With
  Exit Sub
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "AddField"
  Exit Sub
End Sub

Private Sub txtEingabeFokus_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtEingabeFokus_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtEingabeFokus_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub


Private Sub txtEingabeFokus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub


Private Sub txtEingabeFokus_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub


Private Sub txtEingabeFokus_Validate(Cancel As Boolean)
  UpdateCell txtEingabeFokus.Text
End Sub

'{Setzt Standardlistentrennzeichen}
Private Sub UserControl_Initialize()
  Set mReader = New PropertiesViewReader
  Me.Separator = "#"
End Sub

'{Paßt Größe des FlexGrids an}
Private Sub UserControl_Resize()
  With flexEigenschaften
    .Height = Height
    .Width = Width
  End With
End Sub

'{Öffne/aktualisiert Eigenschaftenansicht}
Public Sub OpenProperties()
On Error GoTo Fehler

  txtEingabeFokus.Text = ""
  cboEingabeFokus.Text = ""
  
  FillFlexGrid
  Exit Sub
  
Fehler:
  ToolKitsModule.BaseToolKit.ToolkitError.ShowError "OpenProperties"
  Exit Sub
End Sub

'{Schließt Verbindung zum SQL-Server}
Public Sub CloseProperties()
  With Controls(GetCurrentEingabeFokus)
    If flexEigenschaften.Text = .Text Then Exit Sub
    If Me.PropertyCount = 0 Then Exit Sub
    UpdateCell .Text
  End With
End Sub

'{Eigenschaft Listentrennzeichen}
Public Property Get Separator() As String
  Separator = mstrSeparator
End Property

'{Eigenschaft Listentrennzeichen}
Public Property Let Separator(ByVal strSeparator As String)
  mstrSeparator = strSeparator
End Property

Public Property Get SelText() As String
Dim oEingabeFokus As Control

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  SelText = oEingabeFokus.SelText
End Property

Public Property Let SelText(ByVal strSelText As String)
Dim oEingabeFokus As Control

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  oEingabeFokus.SelText = strSelText
End Property

Public Property Get SelLength() As Long
Dim oEingabeFokus As Control

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  SelLength = oEingabeFokus.SelLength
End Property

Public Property Let SelLength(ByVal lngSelLength As Long)
Dim oEingabeFokus As Control

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  oEingabeFokus.SelLength = lngSelLength
End Property

Public Property Get SelStart() As Long
Dim oEingabeFokus As Control

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  SelStart = oEingabeFokus.SelStart
End Property

Public Property Let SelStart(ByVal lngSelStart As Long)
Dim oEingabeFokus As Control

  Set oEingabeFokus = Controls(GetCurrentEingabeFokus)
  oEingabeFokus.SelStart = lngSelStart
End Property

'{Fordert PropertiesString an}
Public Property Get PropertiesString() As String
  PropertiesString = CreatePropertiesString
End Property

Public Property Get CellText(ByVal intIndex As Integer) As String
  CellText = flexEigenschaften.TextMatrix(intIndex, 1)
End Property

Public Sub SetCellText _
(ByVal intIndex As Integer _
, ByVal strValue As String)
  With flexEigenschaften
    '{flex aktualisieren}
    .TextMatrix(intIndex, 1) = strValue
    
    '.Row = intIndex
    '{event. Eingabefokus aktualisieren}
    If .Row = intIndex Then
      If Controls(GetCurrentEingabeFokus).Name = "txtEingabeFokus" Then
        Controls(GetCurrentEingabeFokus).Text = strValue
      Else
        Controls(GetCurrentEingabeFokus).ListIndex = ToolKitsModule.BaseToolKit.Controls.ComboBox.ItemDataInComboBox(cboEingabeFokus, strValue)
      End If
    End If
    ToolKitsModule.BaseToolKit.Controls.flexGrid.ResizeColumns PropertiesViewResizeDialog, flexEigenschaften, 1
    
    If (Me.MinWidth > 0) _
    And (flexEigenschaften.ColWidth(1) < Me.MinWidth) Then
      flexEigenschaften.ColWidth(1) = Me.MinWidth
    End If

    CheckMaxWidth
   End With
End Sub

Public Property Get Row() As Long
  Row = flexEigenschaften.Row
End Property

Public Property Let Row(ByVal lngRow As Long)
  flexEigenschaften.Row = lngRow
End Property

Public Property Get Text() As String
  Text = flexEigenschaften.Text
End Property

Public Property Get PropertyName(ByVal intIndex) As String
  PropertyName = flexEigenschaften.TextMatrix(intIndex, 0)
End Property

Public Sub SetPropertyName _
(ByVal intIndex As Integer _
, ByVal strValue As String)
  flexEigenschaften.TextMatrix(intIndex, 0) = strValue
  ToolKitsModule.BaseToolKit.Controls.flexGrid.ResizeColumns PropertiesViewResizeDialog, flexEigenschaften, 0
  
  If (Me.MinWidth > 0) _
  And (flexEigenschaften.ColWidth(1) < Me.MinWidth) Then
    flexEigenschaften.ColWidth(1) = Me.MinWidth
  End If

  CheckMaxWidth
End Sub

Public Property Get Count() As Long
  Count = flexEigenschaften.Rows
End Property

Public Property Get UpdateEqualValues() As Boolean
  UpdateEqualValues = mblnUpdateEqualValues
End Property

Public Property Let UpdateEqualValues(ByVal blnUpdateEqualValues As Boolean)
  mblnUpdateEqualValues = blnUpdateEqualValues
End Property

Private Sub UserControl_Terminate()
  Set mReader = Nothing
End Sub
