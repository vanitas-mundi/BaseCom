VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlUnicodeComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlUnicodeComboBox
'    Project    : ToolKits
'
'    Description: Stellt UnicodeComboBox-Erweiterungen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       PreSelectItemInComboBoxWhileTyping
' Description:       Wählt einen Eintrag während der Eingabe aus.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-14:16:58
'
' Parameters :       comboBox (Object)
'--------------------------------------------------------------------------------
Public Sub PreSelectItemInComboBoxWhileTyping(ByRef comboBox As Object)

  Static sTextOld As String
  Dim i           As Integer
  Dim sPos        As Integer
  
  If comboBox.text = "" Then comboBox.ListIndex = 0
  
  With comboBox
    ' nur wenn sich die Eingabe geändert hat
    
    If sTextOld <> .text Then
      For i = .ListIndex To .ListCount - 1
      
        If i >= 0 Then
        
          If LCase$(.List(i)) Like LCase$(.text + "*") Then
          
            sPos = .SelStart
            sTextOld = .text
            .text = .List(i)
            .SelStart = sPos
            .SelLength = Len(.text) - sPos + 1
            .ListIndex = i
            Exit For
          End If
        
        End If
      
      Next i
    
    End If
  End With
  
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetListIndexOfDropDownElement
' Description:       Ermittelt den ListIndex eines Elements in einer UnicodeCOmboBox.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-14:59:06
'
' Parameters :       text (String)
'                    cboBox (Object)
'                    bUseLikeOperator (Boolean)
'--------------------------------------------------------------------------------
Public Function GetListIndexOfDropDownElement(ByVal text As String, _
                                              ByRef cboBox As Object, _
                                              ByVal bUseLikeOperator As Boolean) As Integer
 
  Dim iUboundComboBox As Integer
  Dim i               As Integer
  
  iUboundComboBox = cboBox.ListCount - 1
  
  For i = 0 To iUboundComboBox
    Dim elementText As String: elementText = cboBox.List(i)
    
    If bUseLikeOperator Then
    
      If LCase$(elementText) Like LCase$(text & "*") Then
        GetListIndexOfDropDownElement = i
        Exit Function
      End If
      
    Else
      
      If LCase$(elementText) = LCase$(text) Then
        GetListIndexOfDropDownElement = i
        Exit Function
      End If
    
    End If
  Next i
  
  GetListIndexOfDropDownElement = -1
End Function

