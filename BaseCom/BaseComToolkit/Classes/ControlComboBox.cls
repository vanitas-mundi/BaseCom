VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlComboBox
'    Project    : ToolKits
'
'    Description: Stellt ComboBox-Erweiterungen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
'Private Declare Function SendMessage _
'Lib "user32" Alias "SendMessageA" _
'(ByVal hwnd As Long, ByVal wMsg As Long _
', ByVal wParam As Long, lParam As Any) As Long

Private mConnectionString      As String


'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ConnectionString() As String
  ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal value As String)
  mConnectionString = value
End Property




'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ItemInComboBox
' Description:       Prüft, ob sich item in der ComboBox befindet.
'                    Liefert den ListIndex oder -1
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:10:57
'
' Parameters :       combo (Object)
'                    item (String)
'--------------------------------------------------------------------------------
Public Function ItemInComboBox(ByVal combo As Object, ByVal item As String) As Long
  'Findet keinen Leerstring
  'Const CB_FINDSTRINGEXACT = &H158
  'ItemInComboBox = SendMessage(combo.hwnd, CB_FINDSTRINGEXACT, ByVal -1&, ByVal item)
  

  ItemInComboBox = -1
  
  item = LCase$(item)
  
  Dim i As Long
  For i = 0 To combo.ListCount - 1
    If LCase$(combo.List(i)) = item Then
      ItemInComboBox = i
      Exit Function
    End If
  Next i
  
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ItemDataInComboBox
' Description:       Prüft, ob item als ItemData existiert.
'                    Liefert den ListIndex oder -1
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:11:49
'
' Parameters :       combo (Object)
'                    itemData (Long)
'--------------------------------------------------------------------------------
Public Function ItemDataInComboBox(ByVal combo As Object, ByVal itemData As Long) As Long
  
  ItemDataInComboBox = -1
  
  Dim i As Long
  For i = 0 To combo.ListCount - 1
    If combo.itemData(i) = itemData Then
      ItemDataInComboBox = i
      Exit Function
    End If
  Next i
End Function


'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       TextPropertyInCombo
' Description:       Prüft, ob sich der Text der Text-Eigenschaft in
'                    der ComboBox befindet.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-15:12:34
'
' Parameters :       combo (Object)
'--------------------------------------------------------------------------------
Public Function TextPropertyInCombo(ByRef combo As Object) As Boolean

  With combo
    Dim Index As Long: Index = ItemInComboBox(combo, .Text)
    .ListIndex = Index
    TextPropertyInCombo = (Index <> -1) Or (Trim(.Text) = "")
  End With
End Function



'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FillComboBoxFromT_Defaults
' Description:
'   Eine Combobox soll via generischer Funktion mit Werten aus "def.t_defaults" gefüllt werden.
'   Der existierende Inhalt der Combobox wird immer gelöscht.
'   Combobox wird gefüllt mit dem Inhalt der Spalte : "value".
'
' Created by :       Rainhard Frese
' Machine    :       VDI-IT-02-0009
' Date-Time  :       22.01.2016-12:04:16
'
' Parameters :       ComboBox (Object)
'                    Classname (String)
'                    Property (String)
'                    leererEintrag (Boolean)    <- Ob die Combobox auch einen leeren Eintrag haben soll.
'                                                   Der leere Eintrag wird immer als erstes hinzugefügt.
' Optionale Params:  Active (Boolean = True)
'                    IstGeloescht (Boolean = False) <- in der Tabelle ein DB-Timestamp: Wenn gefüllt, dann gilt das Objekt als gelöscht.
'--------------------------------------------------------------------------------
Public Sub FillComboBoxFromTableDefaults(ByRef combo As Object, Classname As String, Property As String, addEmptyItem As Boolean, _
                                        Optional Active As Boolean = True, Optional IsDeleted As Boolean = False)
    
    If combo Is Nothing Then Exit Sub
    Dim db  As Database: Set db = New Database
    db.ConnectionString = Me.ConnectionString
    
    Dim strSQL As String
    
    strSQL = "SELECT value From def.t_defaults " & vbCrLf & vbTab
    strSQL = strSQL & " WHERE " & vbCrLf & vbTab
    strSQL = strSQL & " classname = '" & Classname & "' " & vbCrLf & vbTab
    strSQL = strSQL & " AND Property = '" & Property & "'  AND " & vbCrLf & vbTab
    strSQL = strSQL & " Active = '" & IIf(Active, "Y", "N") & "' and " & vbCrLf & vbTab
    strSQL = strSQL & " IstGeloescht " & IIf(IsDeleted, "is not null", "is null") & vbCrLf
    strSQL = strSQL & " ORDER BY value asc ;"
                

    Dim rs As Object
    Set rs = db.ExecuteReaderConnected(strSQL)
    
    With combo
    
        .Clear
        If addEmptyItem Then
            .AddItem ""
        End If
        
        While Not rs.EOF
          .AddItem rs!value & ""
        rs.MoveNext
        Wend
    
    End With

    db.CloseRecordSetAndConnection rs
        
    Set db = Nothing
    
End Sub



'--------------------------------------------------------------------------------
' Project    :      ToolKits
' Procedure  :      FillComboWith_Y_N
' Description:      Eine Comboxbox soll mit den Werten: "Y" und "N" gefüllt werden.
'                   Der existierende Inhalt der Combobox wird immer gelöscht.
' Created by :      Rainhard Frese
' Machine    :      VDI-IT-02-0009
' Date-Time  :      22.01.2016-17:05:29
'
' Parameters :      combo (Object)
'                   leererEintrag (Boolean)    <- Ob die Combobox auch einen leeren Eintrag haben soll.
'                                                 Der leere Eintrag wird immer als erstes hinzugefügt.
'--------------------------------------------------------------------------------
Public Sub FillComboWithYesNo(ByRef combo As Object, addEmptyItem As Boolean)

    If combo Is Nothing Then Exit Sub
    
    With combo
        .Clear
        If addEmptyItem Then
            .AddItem ""
        End If
        .AddItem "Y"
        .AddItem "N"
        
    End With

End Sub



Public Sub PreSelectItemInComboBoxWhileTyping(ByRef comboBox As Object)

    If comboBox Is Nothing Then Exit Sub
    
    Static sTextOld As String
    Dim i           As Integer
    Dim sPos        As Integer
    Dim iListcount  As Integer

    If LenB(comboBox.Text) = 0 Then
        comboBox.ListIndex = 0
    End If

    iListcount = comboBox.ListCount - 1

    With comboBox

        ' nur wenn sich die Eingabe ge?ndert hat
        If sTextOld <> .Text Then
            Dim Text As String
            Text = .Text
            Dim Temp As String

            For i = 0 To iListcount
                Temp = Left$(.List(i), Len(Text))

                If LCase$(Temp) = LCase$(Text) Then
                    sPos = .SelStart
                    sTextOld = .Text
                    .Text = .List(i)
                    .SelStart = sPos
                    .SelLength = Len(.Text) - sPos + 1

                    Exit For

                End If

            Next i

        End If

    End With


End Sub


