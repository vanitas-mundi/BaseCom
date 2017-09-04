VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Convert
'    Project    : ToolKits
'
'    Description: Stellt Konvertierungsroutinen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum SortOrders
  Ascending = 0
  Descending = 1
End Enum

'---------------------- Konstruktor der Klasse ----------------------------------

'---------------------- Zugriffsmethoden der Klasse -----------------------------

'---------------------- Ereignismethoden der Klasse -----------------------------

'---------------------- Private Methoden der Klasse -----------------------------
Private Sub Quicksort(ByRef Data As Variant, ByVal sortOrder As SortOrders _
, ByVal firstIndex As Integer, ByVal lastIndex As Integer, Optional ByVal propertyName As String = "")

  If firstIndex < lastIndex Then
    Dim middleIndex As Integer
    middleIndex = SortSubset(Data, sortOrder, firstIndex, lastIndex, propertyName)
    Quicksort Data, sortOrder, firstIndex, middleIndex - 1, propertyName
    Quicksort Data, sortOrder, middleIndex + 1, lastIndex, propertyName
  End If
End Sub

Private Function GetCurrentValue(ByRef Data As Variant, ByVal Index As Integer _
, Optional ByVal propertyName As String = "") As Variant
    
  If propertyName = "" Then
    GetCurrentValue = Data(Index)
  Else
    GetCurrentValue = CallByName(Data(Index), propertyName, VbGet)
  End If
End Function

Private Sub TriangleExchange(ByRef Data As Variant, ByVal sourceIndex As Integer _
, ByVal destinationIndex As Integer, Optional ByVal propertyName As String = "")
  
  Dim Temp As Variant
  If propertyName = "" Then
    Temp = Data(sourceIndex)
    Data(sourceIndex) = Data(destinationIndex)
    Data(destinationIndex) = Temp
  Else
    Set Temp = Data(sourceIndex)
    Set Data(sourceIndex) = Data(destinationIndex)
    Set Data(destinationIndex) = Temp
  End If
End Sub

Private Function FirstCompareIsTrue _
(ByVal currentValue As Variant, ByVal pivot As Variant, ByVal sortOrder As SortOrders) As Boolean

  Select Case sortOrder
  Case Ascending
    FirstCompareIsTrue = (currentValue <= pivot)
  Case Descending
    FirstCompareIsTrue = (currentValue > pivot)
  Case Else
    FirstCompareIsTrue = False
  End Select
End Function

Private Function SecondCompareIsTrue _
(ByVal currentValue As Variant, ByVal pivot As Variant, ByVal sortOrder As SortOrders) As Boolean

  Select Case sortOrder
  Case Ascending
    SecondCompareIsTrue = (currentValue >= pivot)
  Case Descending
    SecondCompareIsTrue = (currentValue < pivot)
  Case Else
    SecondCompareIsTrue = False
  End Select
End Function

Private Function SortSubset(ByRef Data As Variant, ByVal sortOrder As SortOrders _
, ByVal firstIndex As Integer, ByVal lastIndex As Integer _
, Optional ByVal propertyName As String = "") As Integer

  Dim i As Integer: i = firstIndex
  Dim j As Integer: j = lastIndex - 1
  Dim pivot As Variant: pivot = GetCurrentValue(Data, lastIndex, propertyName)

  Do
    While FirstCompareIsTrue(GetCurrentValue(Data, i, propertyName), pivot, sortOrder) And i < lastIndex
      i = i + 1
    Wend
  
    While SecondCompareIsTrue(GetCurrentValue(Data, j, propertyName), pivot, sortOrder) And j > firstIndex
      j = j - 1
    Wend
  
    If i < j Then
      TriangleExchange Data, i, j, propertyName
    End If
  Loop While i < j 'solange i an j nicht vorbeigelaufen ist

  If SecondCompareIsTrue(GetCurrentValue(Data, i, propertyName), pivot, sortOrder) Then
    TriangleExchange Data, i, lastIndex, propertyName
  End If
  
  SortSubset = i
End Function


'---------------------- Öffentliche Methoden der Klasse -------------------------

'<summary>
'   Project     :   ToolKits
'   Procedure   :   BoolToYesNoString
'   Description :   Methode zur Umwandlung von Boolwerten in Yes/No, Y/N, Ja/Nein, J/N Angaben
'   Created by  :   sebastian.limke
'   Date-Time   :   7.22.2015-08:43:32'
'   Parameters  :   bValue (Boolean)
'              bLongWriting (Boolean)
'              strLanguage (String = "de")
'</summary>
Public Function BoolToYesNoString(ByVal bValue As Boolean, _
                                  ByVal bLongWriting As Boolean, _
                                  Optional ByVal strLanguage As String = "de") As String

    '<EhHeader>
    On Error GoTo BoolToYesNoString_Err

    '</EhHeader>

    If strLanguage = "en" Then
        If bLongWriting Then
            If bValue Then BoolToYesNoString = "Yes" Else BoolToYesNoString = "No"

            Exit Function

        Else

            If bValue Then BoolToYesNoString = "Y" Else BoolToYesNoString = "N"

            Exit Function

        End If

    Else

        If bLongWriting Then
            If bValue Then BoolToYesNoString = "Ja" Else BoolToYesNoString = "Nein"

            Exit Function

        Else

            If bValue Then BoolToYesNoString = "J" Else BoolToYesNoString = "N"

            Exit Function

        End If
    End If

    '<EhFooter>
    Exit Function

BoolToYesNoString_Err:
    Err.Raise vbObjectError, "ToolKits.Convert.BoolToYesNoString", Err.description

    '</EhFooter>

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CollectionToArray
' Description:       Wandelt die übergebene Collection col in ein Array um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:53:13
'
' Parameters :       col (Collection) As Variant()
'                    collectionContainsReferenceType - True wenn Collection instanzierte Klassen enthält.
'--------------------------------------------------------------------------------
Public Function CollectionToArray(ByVal col As Collection) As Variant()
'Public Function CollectionToArray(ByVal col As Collection _
', Optional ByVal collectionContainsReferenceType As Boolean = False) As Variant()
  Dim objArray() As Variant
  
  If col.count > 0 Then
    ReDim objArray(1 To col.count)
    
    Dim i As Long: i = 1
    Dim X As Variant
  
    For Each X In col
      objArray(i) = X
      i = i + 1
     Next X
  End If
  
  CollectionToArray = objArray
  
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CollectionToReferenceTypeArray
' Description:       Wandelt die übergebene Collection col in ein Array um, wenn die Collections mit instanzierten Objekten gefüllt ist.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       18.09.2015-12:53:13
'
' Parameters :       col (Collection) As Variant
'--------------------------------------------------------------------------------
Public Function CollectionToReferenceTypeArray(ByVal col As Collection) As Variant()
  Dim objArray() As Variant
  
  If col.count > 0 Then
    ReDim objArray(1 To col.count)
    
    Dim i As Integer: i = 1
    Dim X As Variant
  
    For Each X In col
      Set objArray(i) = X
      i = i + 1
     Next X
  End If
  
  CollectionToReferenceTypeArray = objArray
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ArrayToCollection
' Description:       Wandelt das übergebene Array sourceArray in eine Collection um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:53:50
'
' Parameters :       sourceArray (Variant)
'--------------------------------------------------------------------------------
Public Function ArrayToCollection(ByVal sourceArray As Variant) As Collection
  Set ArrayToCollection = New Collection

  Dim i As Long: i = 1
  For i = LBound(sourceArray) To UBound(sourceArray)
    ArrayToCollection.Add sourceArray(i)
  Next i
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       JoinCollection
' Description:       Wandelt die übergebene Collection in einen String um,
'                    die einzelnen Werte werden durch delimiter getrennt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:15
'
' Parameters :       col (Collection)
'                    delimeter (String)
'--------------------------------------------------------------------------------
Public Function JoinCollection _
(ByVal col As Collection, ByVal delimeter As String) As String
  JoinCollection = Join(CollectionToArray(col), delimeter)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SplitCollection
' Description:       Wandelt einen String in eine Collection um. Die einzelnen
'                    Elemente müssen im String durch delimiter getrennt sein.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:25
'
' Parameters :       value (String)
'                    delimeter (String)
'--------------------------------------------------------------------------------
Public Function SplitCollection _
(ByVal value As String, ByVal delimeter As String) As Collection
  Set SplitCollection = ArrayToCollection(Split(value, delimeter))
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       PointToComma
' Description:       Wandelt in einem String Punkte in Kommata um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:32
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function PointToComma(ByVal value As String) As String
  PointToComma = Replace(value, ".", ",")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CommaToPoint
' Description:       Wandelt in einem String in KommataPunkte um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:37
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function CommaToPoint(ByVal value As String) As String
  CommaToPoint = Replace(value, ",", ".")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FormatYMD
' Description:       Formatiert einen Datum-String in das Format JJJJ-MM-TT um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:42
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function FormatYMD(ByVal value As String) As String
  FormatYMD = Format(value, "yyyy-mm-dd")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FormatYMDHMS
' Description:       Formatiert einen Datum-String in das Format JJJJ-MM-TT HH:MM:SS um.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.04.2015-16:10:42
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function FormatYMDHMS(ByVal value As String) As String
  FormatYMDHMS = Format(value, "yyyy-mm-dd hh:mm:ss")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FormatDMY
' Description:       Formatiert einen Datum-String in das Format TT-MM-JJJJ um.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:46
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function FormatDMY(ByVal value As String) As String
  FormatDMY = Format(value, "dd.mm.yyyy")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MySqlReplaceEscape
' Description:       Maskiert in einem String MySQL geschütze Zeichen.
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:51
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function MySqlReplaceEscape(ByVal value As String) As String

  value = Replace(value, "\", "\\")
  value = Replace(value, "'", "\'")
  value = Replace(value, Chr(34), "\" & Chr(34))
  MySqlReplaceEscape = value
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       TimeSpanToTimeRay
' Description:       Liefert einen String welcher alle Minuten zwischen zwei
'                    Uhrzeiten enthält.
'                    Der Intervall beträgt eine Minute, kann aber über
'                    den optionalen Parameter geändert werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:54:58
'
' Parameters :       timeFrom (String)
'                    timeTo (String )
'                    intervalInMinutes (Integer = 1)
'--------------------------------------------------------------------------------
Public Function TimeSpanToTimeRay _
(ByVal timeFrom As String _
, ByVal timeTo As String _
, Optional intervalInMinutes As Integer = 1) As String

  Dim interval As String: interval = "00:" & Format(intervalInMinutes, "00")
  Dim result As Collection: Set result = New Collection

  Dim i As Date
  For i = timeFrom To timeTo Step interval
    result.Add "'" & Format(i, "hh:mm") & "'"
  Next i
  
  TimeSpanToTimeRay = JoinCollection(result, ", ")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DateSpanToDateRay
' Description:       Liefert einen String welcher alle Tage zwischen zwei
'                    Daten enthält.
'                    Der Intervall beträgt einen Tag, kann aber über
'                    den optionalen Parameter geändert werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:55:05
'
' Parameters :       dateFrom (String)
'                    dateTo (String )
'                    intervalInDays (Integer = 1)
'--------------------------------------------------------------------------------
Public Function DateSpanToDateRay _
(ByVal dateFrom As String _
, ByVal dateTo As String _
, Optional ByVal intervalInDays As Integer = 1) As String

  Dim result As Collection: Set result = New Collection
  
  Dim i As Date
  For i = dateFrom To dateTo Step intervalInDays
    result.Add "'" & Format(i, "yyyy-mm-dd") & "'"
  Next i
  
  DateSpanToDateRay = JoinCollection(result, ", ")
End Function






'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReverseCollection
' Description:       Liefert die übergebene Collection in umgekehrter Reihenfolge zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       18.09.2015-12:55:05
'
' Parameters :       col (Collection)
'--------------------------------------------------------------------------------
Public Function ReverseCollection(ByVal col As Collection) As Collection
  Dim Data As Collection: Set Data = New Collection
  
  Dim i As Integer
  For i = col.count To 1 Step -1
    Data.Add col.item(i)
  Next i
  
  Set ReverseCollection = Data
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReverseArray
' Description:       Liefert das übergebene Array in umgekehrter Reihenfolge zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       18.09.2015-12:55:05
'
' Parameters :       sourceArray (Variant)
'--------------------------------------------------------------------------------
Public Function ReverseArray(ByVal sourceArray As Variant) As Variant
  
  Dim l As Integer: l = LBound(sourceArray)
  Dim u As Integer: u = UBound(sourceArray)
  
  Dim Index As Integer
  Dim endIndex As Integer: endIndex = (u - l) \ 2

  For Index = l To endIndex
    Dim Temp As Variant
    Temp = sourceArray(l)
    sourceArray(l) = sourceArray(u)
    sourceArray(u) = Temp
    l = l + 1
    u = u - 1
  Next Index
  
  ReverseArray = sourceArray
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SortArray
' Description:       Sortiert das übergebene Array per Quicksort.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       18.09.2015-12:55:05
'
' Parameters :       data (Variant)
'                    sortOrder (SortOrders)
'                    propertyName (String) Die Eigenschaft des Objektes nach welcher sortiert wird
'--------------------------------------------------------------------------------
Public Function SortArray(ByVal Data As Variant _
, ByVal sortOrder As SortOrders, Optional ByVal propertyName As String = "") As Variant

  Dim firstIndex As Integer: firstIndex = LBound(Data)
  Dim lastIndex As Integer: lastIndex = UBound(Data)
  
  Quicksort Data, sortOrder, firstIndex, lastIndex, propertyName
  SortArray = Data
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SortCollection
' Description:       Sortiert die übergebene Collection per Quicksort.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       18.09.2015-12:55:05
'
' Parameters :       data (Collection)
'                    sortOrder (SortOrders)
'                    propertyName (String) Die Eigenschaft des Objektes nach welcher sortiert wird
'--------------------------------------------------------------------------------
Public Function SortCollection(ByVal Data As Collection _
, ByVal sortOrder As SortOrders, Optional ByVal propertyName As String = "") As Collection

  Dim firstIndex As Integer: firstIndex = 1
  Dim lastIndex As Integer: lastIndex = Data.count
  
  Dim ar As Variant
  If propertyName = "" Then
    ar = CollectionToArray(Data)
  Else
    ar = CollectionToReferenceTypeArray(Data)
  End If
  
  Quicksort ar, sortOrder, firstIndex, lastIndex, propertyName
  Set SortCollection = ArrayToCollection(ar)
End Function






