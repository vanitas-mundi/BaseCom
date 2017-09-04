VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlListBox
'    Project    : ToolKits
'
'    Description: Stellt ListBox-Erweiterungen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Declare Function SendMessage _
Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long _
, ByVal wParam As Long, lParam As Any) As Long



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ItemInListBox
' Description:       Prüft, ob sich item in der ListBox befindet.
'                    Liefert den ListIndex oder -1
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:28:14
'
' Parameters :       aListBox (Object)
'                    item (String)
'--------------------------------------------------------------------------------
Public Function ItemInListBox _
(ByVal aListBox As Object _
, ByVal item As String) As Long
  
  Const LB_FINDSTRINGEXACT = &H1A2
  ItemInListBox = SendMessage(aListBox.hWnd, LB_FINDSTRINGEXACT, ByVal -1&, ByVal item)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ItemDataInComboBox
' Description:       Prüft, ob item als ItemData existiert.
'                    Liefert den ListIndex oder -1
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:29:08
'
' Parameters :       combo (Object)
'                    itemData (Long)
'--------------------------------------------------------------------------------
Public Function ItemDataInListBox(ByVal aListBox As Object, ByVal itemData As Long) As Long

  ItemDataInListBox = -1

  Dim i As Long
  For i = 0 To aListBox.ListCount - 1
    If aListBox.itemData(i) = itemData Then
      ItemDataInListBox = i
      Exit Function
    End If
  Next i
End Function
