VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlTreeView
'    Project    : ToolKits
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Const TVHT_NOWHERE = &H1
Private Const TVHT_ONITEMICON = &H2
Private Const TVHT_ONITEMLABEL = &H4
Private Const TVHT_ONITEMINDENT = &H8
Private Const TVHT_ONITEMBUTTON = &H10
Private Const TVHT_ONITEMRIGHT = &H20
Private Const TVHT_ONITEMSTATEICON = &H40
Private Const TVHT_ABOVE = &H100
Private Const TVHT_BELOW = &H200
Private Const TVHT_TORIGHT = &H400
Private Const TVHT_TOLEFT = &H800
Private Const TVIS_CHECK      As Long = &H3000
Private Const TVIF_STATE      As Long = &H8
Private Const TV_FIRST        As Long = &H1100
Private Const TVM_GETITEM     As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM     As Long = (TV_FIRST + 13)
Private Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Private Const TVGN_CARET      As Long = &H9
Private Const HASCHECKBOX     As Long = 4098

Public Enum thtPartConstants
  thtClientNowhere = TVHT_NOWHERE
  thtItemIcon = TVHT_ONITEMICON
  thtItemLabel = TVHT_ONITEMLABEL
  thtItemIndent = TVHT_ONITEMINDENT
  thtItemButton = TVHT_ONITEMBUTTON
  thtItemRight = TVHT_ONITEMRIGHT
  thtItemCheck = TVHT_ONITEMSTATEICON
  thtClientAbove = TVHT_ABOVE
  thtClientBelow = TVHT_BELOW
  thtClientRight = TVHT_TORIGHT
  thtClientLeft = TVHT_TOLEFT
  thtClientLeftAbove = TVHT_TOLEFT Or TVHT_ABOVE
  thtClientLeftBelow = TVHT_TOLEFT Or TVHT_BELOW
  thtClientRightAbove = TVHT_TORIGHT Or TVHT_ABOVE
  thtClientRightBelow = TVHT_TORIGHT Or TVHT_BELOW
End Enum

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  X As Long
  y As Long
End Type

Private Type TVHITTESTINFO
  pt As POINTAPI
  Flags As Long
  hItem As Long
End Type

Private Type nodeItem
  Mask           As Long
  hItem          As Long
  State          As Long
  StateMask      As Long
  pszText        As String
  cchTextMax     As Long
  iImage         As Long
  iSelectedImage As Long
  cChildren      As Long
  lParam         As Long
End Type

Private Declare Function GetClientRect _
                Lib "user32" (ByVal hWnd As Long, _
                              lpRect As RECT) As Long
 
Private Declare Function GetScrollPos _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal nBar As Long) As Long
 
Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
  
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long
 
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Sub PrintTreeNode(ByVal aNode As Object, _
                          ByVal strTreeDeep As String, _
                          ByRef xPos As Integer, _
                          ByRef yPos As Integer, _
                          ByRef strTitle As String, _
                          ByRef strDate As String)

  Dim aChild        As Object
  Dim intChildCount As Integer

  With aNode

    If .Children > 0 Then
  
      CheckPrint aNode, strTreeDeep, xPos, xPos, strTitle, strDate
      Set aChild = .Child
      intChildCount = 1
      strTreeDeep = strTreeDeep & "   "
  
      While intChildCount <= aNode.Children
        PrintTreeNode aChild, strTreeDeep, xPos, yPos, strTitle, strDate
        Set aChild = aChild.Next
        intChildCount = intChildCount + 1
      Wend
    Else
      CheckPrint aNode, strTreeDeep, xPos, yPos, strTitle, strDate
    End If
  
  End With

End Sub

Private Sub CheckPrint(ByVal aNode As Object, _
                       ByVal strTreeDeep As String, _
                       ByRef xPos As Integer, _
                       ByRef yPos As Integer, _
                       ByRef strTitle As String, _
                       ByRef strDate As String)

  With printer

    If .CurrentY >= .ScaleHeight - yPos Then
      .NewPage
      PrintHeader xPos, yPos, strTitle, strDate
    End If

    .currentX = xPos
    printer.Print strTreeDeep & "- " & aNode.text
  End With

End Sub

Private Sub PrintHeader(ByRef xPos As Integer, _
                        ByRef yPos As Integer, _
                        ByRef strTitle As String, _
                        ByRef strDate As String, _
                        Optional ByVal blnInitialize As Boolean = False)

  Static intPageNumber As Integer

  If blnInitialize Then intPageNumber = 0
  
  intPageNumber = intPageNumber + 1

  With printer
    .FontBold = True
    .CurrentY = yPos
    .currentX = xPos
    printer.Print strTitle
    .currentX = xPos
    printer.Print "vom: " & strDate & " (Seite: " & intPageNumber & ")"
    .FontBold = False
    printer.Print ""
    printer.Print ""
  End With
  
End Sub



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ScrollOnDragOver
' Description:       Scrollt während eines Drag-Vorgangs, wenn der untere oder
'                    obere TreeView-Bereich mit der Maus betreten wird.
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:42:56
'
' Parameters :       aTreeView (Object)
'                    X (Single)
'                    y (Single)
'--------------------------------------------------------------------------------
Public Sub ScrollOnDragOver(ByRef aTreeView As Object, _
                            ByVal X As Single, _
                            ByVal y As Single)

  Const SB_LINEDOWN = 1
  Const SB_LINEUP = 0
  Const WM_VSCROLL = &H115

  With aTreeView

    Select Case True

      Case (y >= 0) And (y <= 50)
        SendMessageLong .hWnd, WM_VSCROLL, SB_LINEUP, 0

      Case (y >= .Height - 50) And (y <= .Height)
        SendMessageLong .hWnd, WM_VSCROLL, SB_LINEDOWN, 0
    End Select

  End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       HitTestExtended
' Description:       Stellt spezifische Informationen zum TreeNode zur Verfügung.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:45:10
'
' Parameters :       aTreeView (Object)
'                    X (Single)
'                    y (Single)
'                    Part (thtPartConstants)
'                    Button (Integer)
'                    Capture (Boolean)
'                    TestRow (Single = 1)
'--------------------------------------------------------------------------------
Public Function HitTestExtended(aTreeView As Object, _
                                ByVal X As Single, _
                                ByVal y As Single, _
                                Optional Part As thtPartConstants, _
                                Optional Button As Integer, _
                                Optional ByVal Capture As Boolean, _
                                Optional ByVal TestRow As Single = 1) As Object
                             
  Dim nHITTESTINFO As TVHITTESTINFO
  Dim nHitNode     As Object
  Dim nHit         As Boolean
  Dim nX           As Single
  Dim nRect        As RECT
    
  Const TVM_HITTEST = &H1111

  With aTreeView

    If Capture And CBool(Button) Then
      SetCapture .hWnd
    End If

    Set nHitNode = .HitTest(X, y)
  End With

  With nHITTESTINFO
    With .pt
      .X = X \ Screen.TwipsPerPixelX
      .y = y \ Screen.TwipsPerPixelX
    End With

    nHit = CBool(SendMessage(aTreeView.hWnd, TVM_HITTEST, 0, nHITTESTINFO))
    Part = .Flags
  End With

  If nHit Then
    If nHitNode Is Nothing Then
      If TestRow Then

        Select Case Part

          Case thtItemButton, thtItemIndent

            With aTreeView
              GetClientRect .hWnd, nRect

              For nX = X To nRect.Right * Screen.TwipsPerPixelX Step Abs(TestRow) * Screen.TwipsPerPixelX
                Set nHitNode = .HitTest(nX, y)

                If Not (nHitNode Is Nothing) Then
                  Set HitTestExtended = nHitNode
                  Exit For
                End If

              Next 'nX

            End With

          Case thtItemRight

            With aTreeView

              For nX = X To 0 Step Abs(TestRow) * -Screen.TwipsPerPixelX
                Set nHitNode = .HitTest(nX, y)

                If Not (nHitNode Is Nothing) Then
                  Set HitTestExtended = nHitNode
                  Exit For
                End If

              Next 'nX

            End With

        End Select

      End If

    Else
      Set HitTestExtended = nHitNode

      If (Part And TVHT_ONITEMSTATEICON) = TVHT_ONITEMSTATEICON Then
        Part = thtItemCheck
      ElseIf (Part And TVHT_ONITEMICON) = TVHT_ONITEMICON Then
        Part = thtItemIcon
      ElseIf (Part And TVHT_ONITEMLABEL) = TVHT_ONITEMLABEL Then
        Part = thtItemLabel
      End If
    End If
  End If

  If (Capture) And (Part >= TVHT_ABOVE) And (Button = 0) Then ReleaseCapture
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ScrollTo
' Description:       Scrollt im TreeView zur angegebenen Position.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:46:03
'
' Parameters :       aTreeView (Object)
'                    X (Single)
'                    y (Single)
'                    ScrollRegion (Long = 3)
'--------------------------------------------------------------------------------
Public Sub ScrollTo(ByVal aTreeView As Object, _
                    ByVal X As Single, _
                    ByVal y As Single, _
                    Optional ByVal ScrollRegion As Long = 3)

  Dim nRect         As RECT
  Dim nScrollRegion As Single
  Dim nScrollPos    As Long
  Dim nWnd          As Long
  
  Const SB_HORZ = 0
  Const SB_LINEDOWN = 1
  Const SB_LINELEFT = 0
  Const SB_LINERIGHT = 1
  Const SB_LINEUP = 0
  Const WM_HSCROLL = &H114
  Const WM_VSCROLL = &H115
  
  With aTreeView
    nWnd = .hWnd
    GetClientRect nWnd, nRect
    nRect.Bottom = nRect.Bottom * Screen.TwipsPerPixelY
    nScrollRegion = ScrollRegion * Screen.TwipsPerPixelY

    Select Case y

      Case 1 To nScrollRegion
        SendMessageLong nWnd, WM_VSCROLL, SB_LINEUP, 0

      Case Is >= nRect.Bottom - nScrollRegion
        SendMessageLong nWnd, WM_VSCROLL, SB_LINEDOWN, 0

      Case Else
        nRect.Right = nRect.Right * Screen.TwipsPerPixelX
        nScrollRegion = ScrollRegion * Screen.TwipsPerPixelX

        Select Case X

          Case 1 To nScrollRegion
            nScrollPos = GetScrollPos(nWnd, SB_HORZ)
            SendMessageLong nWnd, WM_HSCROLL, SB_LINELEFT, 0

            If nScrollPos <> GetScrollPos(nWnd, SB_HORZ) Then
              .Refresh
            End If

          Case Is >= nRect.Right - nScrollRegion
            nScrollPos = GetScrollPos(nWnd, SB_HORZ)
            SendMessageLong nWnd, WM_HSCROLL, SB_LINERIGHT, 0

            If nScrollPos <> GetScrollPos(nWnd, SB_HORZ) Then
              .Refresh
            End If

        End Select
    End Select

  End With

End Sub






'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       PrintTreeView
' Description:       Druckt die Treeview-Hierachie beginnend vom übergebenenen
'                    Node-Objekt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:47:12
'
' Parameters :       aNode (Object)
'                    xPos (Integer)
'                    yPos (Integer)
'                    intFontSize (Integer = 8)
'                    strTitle (String = "")
'                    strDate (String = "")
'--------------------------------------------------------------------------------
Public Sub PrintTreeView(ByRef aNode As Object, _
                         ByVal xPos As Integer, _
                         ByVal yPos As Integer, _
                         Optional intFontSize As Integer = 8, _
                         Optional ByVal strTitle As String = "", _
                         Optional ByVal strDate As String = "")

  With printer
    ' Maßeinheit "mm"
    .ScaleMode = 6
    '{Drucker initialisieren}
    printer.Print vbNullString
    .FontSize = intFontSize
    
    strTitle = IIf(strTitle = "", aNode.Root, strTitle)
    strDate = IIf(strDate = "", Format(Date, "yyyy-mm-dd"), Format(strDate, "yyyy-mm-dd"))
    
    PrintHeader xPos, yPos, strTitle, strDate, True
    PrintTreeNode aNode, "", xPos, yPos, strTitle, strDate
    .EndDoc
  End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SetNodeCheckBoxVisible
' Description:       Setzt die Sichtbarkeit der CheckBox des Node-Objektes (NodeKey).
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:48:26
'
' Parameters :       aTreeView (Object)
'                    strNodeKey (String)
'                    blnVisible (Boolean)
'--------------------------------------------------------------------------------
Public Sub SetNodeCheckBoxVisible(ByVal aTreeView As Object, _
                                  ByVal strNodeKey As String, _
                                  ByVal blnVisible As Boolean)

  On Error Resume Next
  
  aTreeView.Nodes.Item(strNodeKey).Selected = True
  Dim ret As Long: ret = SendMessage(aTreeView.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
  
  If ret > 0 Then
  
    Dim nodeItem As nodeItem

    With nodeItem
      .hItem = ret
      .Mask = TVIF_STATE
      .StateMask = TVIS_CHECK
      
      ' Status markierten Knotens ermitteln
      SendMessage aTreeView.hWnd, TVM_GETITEM, 0&, nodeItem
      .State = IIf(blnVisible, HASCHECKBOX, 0)
    End With
    
    ' Neuer Status übergeben
    SendMessage aTreeView.hWnd, TVM_SETITEM, 0&, nodeItem
  End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsNodeCheckBoxVisible
' Description:       Prüft, ob die CheckBox des Node-Objektes (NodeKey) sichtbar ist.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:50:00
'
' Parameters :       aTreeView (Object)
'                    strNodeKey (String)
'--------------------------------------------------------------------------------
Public Function IsNodeCheckBoxVisible(ByVal aTreeView As Object, _
                                      ByVal strNodeKey As String) As Boolean
  
  On Error Resume Next
  
  aTreeView.Nodes.Item(strNodeKey).Selected = True
  Dim ret As Long: ret = SendMessage(aTreeView.hWnd, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
  
  If ret > 0 Then
  
    Dim nodeItem As nodeItem

    With nodeItem
      .hItem = ret
      .Mask = TVIF_STATE
      .StateMask = TVIS_CHECK
      
      ' Status markierten Knotens ermitteln
      SendMessage aTreeView.hWnd, TVM_GETITEM, 0&, nodeItem
      
      'State-Werte
      '2      Node ohne Checkbox und ohne Childs
      '34     Node ohne Checkbox und mit Childs
      '4098   Node mit Checkbox unchecked und ohne Childs
      '8198   Node mit Checkbox checked und ohne Childs
      IsNodeCheckBoxVisible = (.State - 4000) > 0
    End With

  End If

End Function

