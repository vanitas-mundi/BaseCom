VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlTextBox
'    Project    : ToolKits
'
'    Description: Stellt TextBox-Erweiterungen zur Verfügung.
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
' Procedure  :       SelectAllText
' Description:       Markiert den in der TextBox dargestellten Text.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-11:42:00
'
' Parameters :       aTextBox (Object)
'--------------------------------------------------------------------------------
Public Sub SelectAllText(ByVal aTextBox As Object)

  On Error Resume Next

  With aTextBox
    .SelStart = 0
    .SelLength = Len(Trim(.Text))
  End With

  On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CompleteDateWithDot
' Description:       Setzt bei Datumseingabe automatische Trennpunkte.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-11:42:07
'
' Parameters :       aTextBox (Object)
'--------------------------------------------------------------------------------
Public Sub CompleteDateWithDot(ByRef aTextBox As Object, ByVal KeyCode As Integer)

    If LenB(Trim$(aTextBox.Text)) = 0 Then Exit Sub
    If KeyCode <> 8 Then

        With aTextBox

            Select Case Len(.Text)

                Case 2
                    .Text = .Text & "."
                    .SelStart = Len(.Text)

                Case 5
                    .Text = .Text & "." & Mid$(Format$(Date, "dd.mm.yyyy"), 7, 4)
                    .SelStart = Len(.Text) - 4
                    .SelLength = Len(.Text)
            End Select

        End With

    End If

End Sub
