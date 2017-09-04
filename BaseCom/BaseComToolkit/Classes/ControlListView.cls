VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlListView
'    Project    : ToolKits
'
'    Description: Stellt ListView-Erweiterungen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum ListItemHitTestSuccessConstants
  lhsNone
  lhsListItem
  lhsListSubItem
End Enum

Public Enum ListItemHitTestPartConstants
  lhpNone
  lhpText
  lhpIcon
  lhpCheckBox
  lhpWhiteSpace
End Enum

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type LVHITTESTINFO
  pt As POINTAPI
  Flags As Long
  item As Long
  SubItem  As Long
End Type

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
Private Declare Function GetClientRect _
                Lib "user32" (ByVal hwnd As Long, _
                              lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Private Const LVNI_SELECTED = &H2
Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Private Const LVM_GETITEMRECT = (LVM_FIRST + 14)
Private Const LVHT_ONITEMICON      As Long = &H2
Private Const LVHT_ONITEMLABEL     As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM          As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
Private Const LVIR_BOUNDS = 0

Private Const lvwReport As Integer = 3



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Sub ListViewLastColumnFillSize(aListView As Object)
  Dim i      As Integer
  Dim nRect  As RECT
  Dim nWidth As Single

  With aListView

    For i = 1 To .ColumnHeaders.count - 1
      nWidth = nWidth + .ColumnHeaders(i).Width
    Next i

    GetClientRect .hwnd, nRect
    nRect.Right = (nRect.Right - 1) * Screen.TwipsPerPixelX

    If nRect.Right > nWidth Then
      .ColumnHeaders(.ColumnHeaders.count).Width = nRect.Right - nWidth
    End If

  End With

End Sub

Private Function HasIcon(ColumnHeader As Object) As Boolean
  Dim nIcon As Variant
  nIcon = ColumnHeader.Icon

  If nIcon <> "0" Then HasIcon = CBool(Len(nIcon))
End Function

Private Function LikeCompText(Str As String, match As String) As Boolean
  LikeCompText = CBool(Str Like match)
End Function



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ResizeColumns
' Description:       Aufrufbeispiel: ResizeColumns ListView1, mSpalte, True, True
'                    Der folgenden Hilfsprozedur übergeben Sie im ersten Parameter
'                    das betreffende ListView-Steuerelement.
'                    Optional können Sie im zweiten Parameter eine spezifische Spalte
'                    eine gültige Position einer Spalte, von 1 bis zur Anzahl der Spalten)
'                    angeben, deren Breite allein angepasst werden soll.
'                    Übergeben Sie hier einen ungültigen Wert, werden alle Spalten angepasst.
'                    Im nächsten optionalen Parameter IncludeHeaders geben Sie an, '
'                    ob die Breite der Texte der Spaltenköpfe mit berücksichtigt werden soll.
'                    Setzen Sie den letzten optionalen Parameter gleich True,
'                    füllt die letzte Spalte - soweit möglich -
'                    automatisch den verbleibenden Raum bis zum rechten Rand des
'                    ListView-Steuerelements bzw. bis zu einem gegebenenfalls sichtbaren
'                    vertikalen Rollbalken aus.
'                    Letzteren Effekt können Sie über die darauffolgende Prozedur
'                    ListViewLastColumnFillSize auch ohne Anpassung der
'                    übrigen Spaltenbreiten erreichen.
'                    Tip von: http://www.aboutvb.de/khw/artikel/khwlistviewadjustcolumnwidth.htm
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:10:39
'
' Parameters :       aListView (Object)
'                    Position (Integer)
'                    IncludeHeaders (Boolean)
'                    LastColumnFillSize (Boolean)
'--------------------------------------------------------------------------------
Public Sub ResizeColumns(ByRef aListView As Object, _
                          Optional ByVal Position As Integer, _
                          Optional ByVal IncludeHeaders As Boolean, _
                          Optional ByVal LastColumnFillSize As Boolean)
                          
  Dim i              As Integer
  Dim nListItem      As Object
  Dim nKey           As String
  Dim nColumn        As Integer
  Dim nPosition      As Integer
  Dim nSmallIconsSet As Boolean
  Dim nRect          As RECT
  Dim nWidth         As Single
  Const LVM_SETCOLUMNWIDTH = &H101E
  Const LVSCW_AUTOSIZE = -1&

  If aListView.ListItems.count = 0 Then Exit Sub

  With aListView
    '.View = lvwReport
    'If .ColumnHeaders.Count = 0 Then Exit Sub
    LockWindowUpdate .hwnd

    If IncludeHeaders Then
      If .SmallIcons Is Nothing Then
        Set .SmallIcons = .ColumnHeaderIcons
        nSmallIconsSet = True
      End If

      Select Case Position

        Case 1 To .ColumnHeaders.count
          nKey = CStr(Now)

          If HasIcon(.ColumnHeaders(1)) Then
            Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "           ")
          Else
            Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "  ")
          End If

          nPosition = .ColumnHeaders(1).Position

          If nPosition = Position Then nColumn = 0

          For i = 2 To .ColumnHeaders.count

            If HasIcon(.ColumnHeaders(i)) Then
              nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text & "         "
            Else
              nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text
            End If

            nPosition = .ColumnHeaders(i).Position

            If nPosition = Position Then nColumn = i - 1
          Next i

          SendMessage .hwnd, LVM_SETCOLUMNWIDTH, nColumn, LVSCW_AUTOSIZE

        Case Else
          nKey = CStr(Now)

          If HasIcon(.ColumnHeaders(1)) Then
            Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "           ")
          Else
            Set nListItem = .ListItems.Add(1, nKey, .ColumnHeaders(1).Text & "  ")
          End If

          SendMessage .hwnd, LVM_SETCOLUMNWIDTH, 0, LVSCW_AUTOSIZE

          For i = 2 To .ColumnHeaders.count

            If HasIcon(.ColumnHeaders(i)) Then
              nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text & "         "
            Else
              nListItem.ListSubItems.Add , , .ColumnHeaders(i).Text
            End If

            SendMessage .hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
          Next i

      End Select

      .ListItems.Remove nKey

      If nSmallIconsSet Then Set .SmallIcons = Nothing
    Else

      Select Case Position

        Case 1 To .ColumnHeaders.count
          nPosition = .ColumnHeaders(Position).Position

          If nPosition = Position Then
            SendMessage .hwnd, LVM_SETCOLUMNWIDTH, Position - 1, LVSCW_AUTOSIZE
          Else
            SendMessage .hwnd, LVM_SETCOLUMNWIDTH, nPosition - 1, LVSCW_AUTOSIZE
          End If

        Case Else

          For i = 0 To .ColumnHeaders.count
            SendMessage .hwnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE
          Next i

      End Select

    End If

    If LastColumnFillSize Then

      For i = 1 To .ColumnHeaders.count - 1
        nWidth = nWidth + .ColumnHeaders(i).Width
      Next i

      GetClientRect .hwnd, nRect
      nRect.Right = nRect.Right * Screen.TwipsPerPixelX

      If nRect.Right > nWidth Then
        .ColumnHeaders(.ColumnHeaders.count).Width = nRect.Right - nWidth
      End If
    End If

    .Refresh
  End With

  LockWindowUpdate 0&
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FillListViewByRecordsetWithoutClearing
' Description:       Befuellt ein Listview allein anhand eines Recordsets ohne das ListView vorher zu leeren
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       11.28.2016-16:33:01
'
' Parameters :       rs (Object)
'                    ListView (Object)
'--------------------------------------------------------------------------------
Public Sub FillListViewByRecordsetWithoutClearing(ByRef rs As Object, ByRef ListView As Object)

    '<EhHeader>
    On Error GoTo FillListView_Err

    '</EhHeader>

    Dim i              As Integer
    Dim iRsFieldsCount As Integer
    iRsFieldsCount = rs.Fields.count - 1

    With ListView
        .FullRowSelect = True
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideColumnHeaders = False

    End With

    Dim lvwItem As ListItem

    With rs

        Do While Not .EOF
            Set lvwItem = ListView.ListItems.Add(, , .Fields(0).value)

            For i = 1 To iRsFieldsCount
                Dim strValue As String
                strValue = Trim(.Fields(i).value & "")

                If Not IsNull(strValue) Then
                    lvwItem.SubItems(i) = strValue
                Else
                    lvwItem.SubItems(i) = ""
                End If

            Next i

            .MoveNext
        Loop

    End With

    If ListView.ListItems.count > 0 Then
        ResizeColumns ListView, , True
    End If
    
    '<EhFooter>
    Exit Sub

FillListView_Err:
    lvwItem.SubItems(i) = "!! DATENFEHLER !!"
    lvwItem.ListSubItems(i).ForeColor = vbRed

    Resume Next

    '</EhFooter>

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FillListViewByRecordset
' Description:       Befuellt ein Listview allein anhand eines Recordsets
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       10.30.2015-08:28:29
'
' Parameters :       rs (Object)
'                    listView (listView)
'--------------------------------------------------------------------------------
Public Sub FillListViewByRecordset(ByRef rs As Object, ByRef ListView As Object)

    '<EhHeader>
    On Error GoTo FillListView_Err

    '</EhHeader>
    
    Dim i              As Integer
    Dim iRsFieldsCount As Integer
    iRsFieldsCount = rs.Fields.count - 1
    
    With ListView
        .ListItems.Clear
        .FullRowSelect = True
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideColumnHeaders = False
        .ColumnHeaders.Clear

        For i = 0 To iRsFieldsCount
            .ColumnHeaders.Add , rs.Fields(i).name, rs.Fields(i).name
        Next i

    End With
    FillListViewByRecordsetWithoutClearing rs, ListView
    
    '<EhFooter>
    Exit Sub

FillListView_Err:
    Err.Raise vbObjectError, "Tnv.vm_frmVerlaengerungen.FillListView", "vm_frmVerlaengerungen component failure"

    Resume Next

    '</EhFooter>

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ListItemHitTest
' Description:       Liefert spezifische Informationen zum angeklickten ListItem.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:13:16
'
' Parameters :       ListView (Object)
'                    X (Single)
'                    y (Single)
'                    ListItem (Object)
'                    ListSubItem (Object)
'                    Part (ListItemHitTestPartConstants)
'                    SelectListItem (Boolean)
'--------------------------------------------------------------------------------
Public Function ListItemHitTest(ByRef ListView As Object, _
                                ByVal X As Single, _
                                ByVal Y As Single, _
                                ByRef ListItem As Object, _
                                ByRef ListSubItem As Object, _
                                ByRef Part As ListItemHitTestPartConstants, _
                                Optional ByVal SelectListItem As Boolean) As ListItemHitTestSuccessConstants
                                
  Dim nHITTESTINFO As LVHITTESTINFO
  Dim nListItem    As Object
  Dim l            As Long
  Dim nRect        As RECT
  Set ListItem = Nothing
  Set ListSubItem = Nothing
  Part = lhpNone

  With ListView

    If .View <> lvwReport Then Exit Function
    If .ListItems.count = 0 Then Exit Function

    With nHITTESTINFO
      .pt.X = (X \ Screen.TwipsPerPixelX)
      .pt.Y = (Y \ Screen.TwipsPerPixelY)
      .Flags = LVHT_ONITEM
    End With

    SendMessage .hwnd, LVM_SUBITEMHITTEST, 0, nHITTESTINFO
  End With

  Select Case nHITTESTINFO.item

    Case Is > -1
      Set nListItem = ListView.ListItems(nHITTESTINFO.item + 1)
    Case Else
      With ListView
        For l = .GetFirstVisible.Index To .ListItems.count
          Set nListItem = .ListItems(l)

          With nListItem
            Select Case Y
              Case .Top To .Top + .Height
                Exit For
              Case Else
                Set nListItem = Nothing
            End Select
          End With
        Next l
      End With

      If nListItem Is Nothing Then
        Exit Function
      Else
        nRect.Left = LVIR_BOUNDS
        SendMessage ListView.hwnd, LVM_GETITEMRECT, nListItem.Index - 1, nRect

        Select Case X \ Screen.TwipsPerPixelX
          Case nRect.Left To nRect.Right
            Part = lhpText
          Case Else
            Part = lhpWhiteSpace
        End Select
      End If
  End Select

  If nHITTESTINFO.SubItem > 0 Then
    Set ListSubItem = nListItem.ListSubItems(nHITTESTINFO.SubItem)
    ListItemHitTest = lhsListSubItem
  Else
    ListItemHitTest = lhsListItem
  End If

  Select Case nHITTESTINFO.Flags
    Case LVHT_ONITEMSTATEICON
      Part = lhpCheckBox
    Case LVHT_ONITEMICON
      If ListSubItem Is Nothing Then
        Part = lhpIcon
      Else
        If Len(ListSubItem.ReportIcon) Then
          Part = lhpIcon
        Else
          Part = lhpText
        End If
      End If
    Case LVHT_ONITEMLABEL
      Part = lhpText
  End Select

  If SelectListItem Then nListItem.Selected = True
  Set ListItem = nListItem
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetListViewSelectedItems
' Description:       Gibt ein Feld mit allen selektierten ListItem-Indices zurück
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:14:34
'
' Parameters :       aListView (Object)
'--------------------------------------------------------------------------------
Public Function GetListViewSelectedItems(ByRef aListView As Object) As Variant
  Dim lngCount   As Long
  Dim lngIndex   As Long
  Dim strIndizes As String
  Dim i          As Long

  With aListView
    lngIndex = -1
    lngCount = SendMessage(.hwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)

    For i = 1 To lngCount
      lngIndex = SendMessage(.hwnd, LVM_GETNEXTITEM, lngIndex, ByVal LVNI_SELECTED)

      If lngIndex <> -1 Then
        strIndizes = strIndizes & lngIndex + 1 & ","
      End If
    Next i

    strIndizes = Left$(strIndizes, Len(strIndizes) - 1)
    GetListViewSelectedItems = Split(strIndizes, ",")
  End With
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSelectedItemsCount
' Description:       Ermittelt die Anzahl der selektierten ListItems.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:15:37
'
' Parameters :       aListView (Object)
'--------------------------------------------------------------------------------
Public Function GetSelectedItemsCount(ByRef aListView As Object) As Long
  GetSelectedItemsCount = SendMessage(aListView.hwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       VisibleItemsCount
' Description:       Liefert die Anzahl der im ListView sehbaren Zeilen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:33:12
'
' Parameters :       aListView (Object)
'--------------------------------------------------------------------------------
Public Function VisibleItemsCount(ByRef aListView As Object)
  
  With aListView
    
    If aListView.View <> lvwReport Then Exit Function
    VisibleItemsCount = (.Height \ .ListItems(1).Height) - 1
  End With
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastVisibleListItem
' Description:       Liefret das letzte sichtbare ListItem .
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:33:59
'
' Parameters :       aListView (Object)
'--------------------------------------------------------------------------------
Public Function GetLastVisibleListItem(ByRef aListView As Object) As Object
  Dim lngLastIndex As Long

  With aListView
    If .View <> lvwReport Then Exit Function
    
    lngLastIndex = (.GetFirstVisible.Index + GetLastVisibleListItem(aListView) - 1)
    If lngLastIndex > .ListItems.count Then lngLastIndex = .ListItems.count
    Set GetLastVisibleListItem = .ListItems.item(lngLastIndex)
  End With

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SortByColumn
' Description:       Sortiert das ListView anhand von ColumnHeader.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:34:58
'
' Parameters :       aListView (Object)
'                    ColumnHeader (Object)
'                    listSortOrderConstant (Object)
'--------------------------------------------------------------------------------
Public Sub SortByColumn(ByVal aListView As Object, _
                        ByVal ColumnHeader As Object, _
                        ByVal listSortOrderConstant As Integer)
  With aListView
    .Sorted = True
    .SortKey = ColumnHeader.Index - 1
    .sortOrder = listSortOrderConstant
  End With
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetColumnIndex
' Description:       Liefert den Spaltenindex anhand der übergebenen X-Kooridate.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:36:18
'
' Parameters :       aListView (Object)
'                    X (Integer)
'--------------------------------------------------------------------------------
Public Function GetColumnIndex _
(ByVal aListView As Object, ByVal X As Integer) As Integer

  GetColumnIndex = -1
  
  With aListView
    Dim currentX As Integer: currentX = 0
    
    Dim i As Integer
    For i = 1 To .ColumnHeaders.count
      If (X <= currentX + .ColumnHeaders(i).Width) And (X >= currentX) Then
        GetColumnIndex = i
        Exit For
      End If
      currentX = currentX + .ColumnHeaders(i).Width
    Next i
  End With

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       Find
' Description:       Liefert die Indexe der ListItems, welche dem Suchmuster entsprechen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:37:21
'
' Parameters :       ListView (Object)
'                    searchString (String)
'                    FindIn (ListViewFindInConstants = lfBoth)
'--------------------------------------------------------------------------------
Public Function Find(ListView As Object, _
                     searchString As String, _
                     Optional ByVal FindIn As ListViewFindInConstants = lfBoth) As Collection
                     
  Dim nListItem    As Object
  Dim nListSubItem As Object
  Dim nFound       As Collection
  Set nFound = New Collection

  Select Case FindIn
    Case lfItems
      For Each nListItem In ListView.ListItems
        If LikeCompText(nListItem.Text, searchString) Then
          nFound.Add nListItem.Index
        End If
      Next
    Case lfSubItems
      For Each nListItem In ListView.ListItems
        For Each nListSubItem In nListItem.ListSubItems
          If LikeCompText(nListSubItem.Text, searchString) Then
            nFound.Add nListItem.Index
          End If
        Next
      Next
    Case lfBoth
      For Each nListItem In ListView.ListItems
        If LikeCompText(nListItem.Text, searchString) Then
          nFound.Add nListItem.Index
        End If

        For Each nListSubItem In nListItem.ListSubItems
          If LikeCompText(nListSubItem.Text, searchString) Then
            nFound.Add nListItem.Index
          End If
        Next
      Next
  End Select

  Set Find = nFound
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       TranformToFlexGridAndPrint
' Description:       Wandelt das übergebene ListView in ein FlexGrid
'                    um und druckt dessen Daten.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:39:44
'
' Parameters :       aForm (Object)
'                    aListView (Object)
'                    flexGrid (Object)
'                    LeftMargin (Single)
'                    TopMargin (Single)
'                    RightMargin (Single)
'                    BottomMargin (Single)
'                    PrintTitle (String = "")
'                    PrintDate (String = "")
'                    poOrientation (ePrinterOrientation = poPortrait)
'--------------------------------------------------------------------------------
Public Sub TranformToFlexGridAndPrint(ByVal aForm As Object, _
                                      ByRef aListView As Object, _
                                      ByRef flexGrid As Object, _
                                      ByVal LeftMargin As Single, _
                                      ByVal TopMargin As Single, _
                                      ByVal RightMargin As Single, _
                                      ByVal BottomMargin As Single, _
                                      Optional ByVal PrintTitle As String = "", _
                                      Optional ByVal PrintDate As String = "", _
                                      Optional ByVal poOrientation As ePrinterOrientation = poPortrait)
                                      
  Dim i        As Integer
  Dim j        As Integer
  Dim strEntry As String

  With aListView
    '{FlexGrid Initialisieren}
    flexGrid.rows = 0
    flexGrid.cols = .ColumnHeaders.count

    '{Header in FlexGrid einfügen}
    '{Spaltenköpfe auslesen}
    For i = 1 To .ColumnHeaders.count
      strEntry = strEntry & .ColumnHeaders(i).Text & vbTab
    Next i

    '{Spaltennamen einfügen}
    flexGrid.AddItem strEntry

    '{Ende Header in FlexGrid einfügen}
    For i = 1 To .ListItems.count
      strEntry = .ListItems(i).Text & vbTab

      For j = 2 To .ColumnHeaders.count
        strEntry = strEntry & .ListItems(i).ListSubItems(j - 1).Text & vbTab
      Next j

      flexGrid.AddItem strEntry
    Next i

    flexGrid.FixedRows = 1
  End With

  Dim flexGridClass As ControlFlexGrid: Set flexGridClass = New ControlFlexGrid
  flexGridClass.ResizeColumns aForm, flexGrid
  flexGridClass.PrintData flexGrid, LeftMargin, TopMargin, RightMargin, BottomMargin, PrintTitle, PrintDate, poOrientation
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SetAllCheckStates
' Description:       Setz im ListView die Checked-Eigenschaft aller ListItems auf den
'                    Wert von chekcState.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       26.01.2015-12:53:01
'
' Parameters :       aListView (Object)
'                    checkState (Boolean)
'--------------------------------------------------------------------------------
Public Sub SetAllCheckStates(ByVal aListView As Object, ByVal checkState As Boolean)
  
  With aListView
    Dim i As Integer
    For i = 1 To .ListItems.count
      .ListItems(i).Checked = checkState
    Next i
  End With
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetCheckedCount
' Description:       Liefert die Anzahl der ListItems mit dem
'                    Wert True ihrer Checked-Eigenschaft.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       26.01.2015-12:59:09
'
' Parameters :       aListView (Object)
'--------------------------------------------------------------------------------
Public Function GetCheckedCount(ByVal aListView As Object) As Integer
'
  With aListView
    GetCheckedCount = 0
  
    Dim i As Long
    For i = 1 To .ListItems.count
      GetCheckedCount = GetCheckedCount + IIf(.ListItems(i).Checked, 1, 0)
    Next i
  End With
End Function
