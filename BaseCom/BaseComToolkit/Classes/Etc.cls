VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Etc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Etc
'    Project    : ToolKits
'
'    Description: Stellt Hilfsfunktionalitäten zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum eldKindOfMaturity
  ldMonth = 0
  ldQuartal = 1
  ldSemester = 2
  ldHalfYear = 3
  ldYear = 4
End Enum

Private Declare Function GetModuleFileName Lib "kernel32" _
Alias "GetModuleFileNameA" _
(ByVal hModule As Long, _
ByVal lpFileName As String, _
ByVal nSize As Long) As Long



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DeleteLastArrayItemIfEmpty
' Description:       Ueberprueft, ob das letzte Arrayelement leer ist und loescht dieses ggfs.
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       10.29.2015-10:30:20
'
' Parameters :       strArray() (String) As String()
'--------------------------------------------------------------------------------
Public Function DeleteLastArrayItemIfEmpty(ByRef strArray() As String) As String()

    Dim iMax As Integer
    iMax = UBound(strArray)

    If iMax > 0 And Len(strArray(iMax)) = 0 Then
        ReDim Preserve strArray(iMax - 1)
    End If
    
    DeleteLastArrayItemIfEmpty = strArray

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsInRangeNumeric
' Description:       Prüft, ob eine Zahl in einem Wertebereich liegt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:45:34
'
' Parameters :       value (Long)
'                    min (Long)
'                    Max (Long)
'--------------------------------------------------------------------------------
Public Function IsInRangeNumeric(ByVal value As Long, _
                                 ByVal min As Long, _
                                 ByVal Max As Long) As Boolean

  IsInRangeNumeric = (value >= min) And (value <= Max)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsInRangeDate
' Description:       Prüft, ob ein Datum in einem Wertebereich liegt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:46:05
'
' Parameters :       value (Date)
'                    min (Date)
'                    Max (Date)
'--------------------------------------------------------------------------------
Public Function IsInRangeDate(ByVal value As Date, _
                              ByVal min As Date, _
                              ByVal Max As Date) As Boolean

  IsInRangeDate = (value >= min) And (value <= Max)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MinNumeric
' Description:       Liefert von zwei übergebenen Zahlen die kleinere.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:46:31
'
' Parameters :       value (Long)
'                    value2 (Long)
'--------------------------------------------------------------------------------
Public Function MinNumeric(ByVal value As Long, ByVal value2 As Long) As Long

  MinNumeric = IIf(value < value2, value, value2)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MaxNumeric
' Description:       Liefert von zwei übergebenen Zahlen die größere.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:47:25
'
' Parameters :       value (Long)
'                    value2 (Long)
'--------------------------------------------------------------------------------
Public Function MaxNumeric(ByVal value As Long, ByVal value2 As Long) As Long

  MaxNumeric = IIf(value > value2, value, value2)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MinDate
' Description:       Liefert von zwei übergebenen Daten die kleinere.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:47:46
'
' Parameters :       value (Date)
'                    value2 (Date)
'--------------------------------------------------------------------------------
Public Function MinDate(ByVal value As Date, ByVal value2 As Date) As Date

  MinDate = IIf(value < value2, value, value2)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MaxDate
' Description:       Liefert von zwei übergebenen Daten die größere.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:48:02
'
' Parameters :       value (Date)
'                    value2 (Date)
'--------------------------------------------------------------------------------
Public Function MaxDate(ByVal value As Date, ByVal value2 As Date) As Date

  MaxDate = IIf(value > value2, value, value2)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastDayOfMaturity
' Description:       Ruft anhand der Fälligkeitsart KindOfMaturity
'                    eine der nachfolgenden Funktionen auf und gibt
'                    deren Ergebnis zurück.
'                    Bei einer ungültigen Fälligkeitsart wird ???
'                    zurück gegeben.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:48:46
'
' Parameters :       validDate (String)
'                    KindOfMaturity (eldKindOfMaturity)
'                    blnDaysOnly (Boolean = False)
'--------------------------------------------------------------------------------
Public Function GetLastDayOfMaturity(ByVal validDate As String, _
                                     ByVal KindOfMaturity As eldKindOfMaturity, _
                                     Optional ByVal blnDaysOnly As Boolean = False) As String

  Select Case KindOfMaturity
    Case eldKindOfMaturity.ldMonth
      GetLastDayOfMaturity = GetLastDayOfMonth(validDate, blnDaysOnly)
    Case eldKindOfMaturity.ldQuartal
      GetLastDayOfMaturity = GetLastDayOfQuartal(validDate, blnDaysOnly)
    Case eldKindOfMaturity.ldSemester
      GetLastDayOfMaturity = GetLastDayOfSemester(validDate, blnDaysOnly)
    Case eldKindOfMaturity.ldHalfYear
      GetLastDayOfMaturity = GetLastDayOfHalfYear(validDate, blnDaysOnly)
    Case eldKindOfMaturity.ldYear
      GetLastDayOfMaturity = GetLastDayOfYear(validDate, blnDaysOnly)
    Case Else
      GetLastDayOfMaturity = "???"
  End Select

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastDayOfMonth
' Description:       Liefert den letzten Tag des Monats von ValidDate
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:49:55
'
' Parameters :       validDate (String)
'                    blnDaysOnly (Boolean = False)
'--------------------------------------------------------------------------------
Public Function GetLastDayOfMonth(ByVal validDate As String, _
                                  Optional ByVal blnDaysOnly As Boolean = False) As String

  On Error Resume Next
  GetLastDayOfMonth = DateSerial(Year(validDate), Month(validDate) + 1, 0)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastDayOfQuartal
' Description:       Liefert den letzten Tag des Quartals von ValidDate
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:50:15
'
' Parameters :       validDate (String)
'                    blnDaysOnly (Boolean = False)
'--------------------------------------------------------------------------------
Public Function GetLastDayOfQuartal(ByVal validDate As String, _
                                    Optional ByVal blnDaysOnly As Boolean = False) As String
                                    
  GetLastDayOfQuartal = DateSerial(Year(validDate), (DatePart("q", validDate) * 3) + 1, 0)
  If blnDaysOnly Then GetLastDayOfQuartal = Day(GetLastDayOfQuartal)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastDayOfHalfYear
' Description:       Liefert den letzten Tag eines Halbjahres von ValidDate.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:50:50
'
' Parameters :       validDate (String)
'                    blnDaysOnly (Boolean = False)
'--------------------------------------------------------------------------------
Public Function GetLastDayOfHalfYear(ByVal validDate As String, _
                                     Optional ByVal blnDaysOnly As Boolean = False) As String
                                     
  GetLastDayOfHalfYear = DateSerial(Year(validDate), (((Month(validDate) \ 7) + 1) * 6) + 1, 0)
  If blnDaysOnly Then GetLastDayOfHalfYear = Day(GetLastDayOfHalfYear)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastDayOfSemester
' Description:       Liefert den letzten Tag eines Semesters von ValidDate
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:51:28
'
' Parameters :       validDate (String)
'                    blnDaysOnly (Boolean = False)
'--------------------------------------------------------------------------------
Public Function GetLastDayOfSemester(ByVal validDate As String, _
                                     Optional ByVal blnDaysOnly As Boolean = False) As String

  Select Case Month(validDate)
    Case 1, 2
      GetLastDayOfSemester = (DateSerial(Year(validDate), 3, 0))
    Case 9, 10, 11, 12
      GetLastDayOfSemester = (DateSerial(Year(validDate) + 1, 3, 0))
    Case 3, 4, 5, 6, 7, 8
      GetLastDayOfSemester = (DateSerial(Year(validDate), 9, 0))
  End Select

  If blnDaysOnly Then GetLastDayOfSemester = Day(GetLastDayOfSemester)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastDayOfSemesterByInstitut
' Description:       Liefert den letzten Tag eines Semesters von ValidDate anhand des Instituts
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       11.20.2015-08:48:37
'
' Parameters :       validDate (String)
'                    strInstitut (String)
'                    blnDaysOnly (Boolean = False)
'--------------------------------------------------------------------------------
Public Function GetLastDayOfSemesterByInstitut(ByVal validDate As String, _
                                               ByVal strInstitut As String, _
                                               Optional ByVal blnDaysOnly As Boolean = False) As String
    Dim strLastDayOfSemester As String
    strLastDayOfSemester = GetLastDayOfSemester(validDate, blnDaysOnly)

    If strInstitut = "BA" Then
        strLastDayOfSemester = DateSerial(Year(strLastDayOfSemester), Month(strLastDayOfSemester) - 1, Day(strLastDayOfSemester))
    End If

    GetLastDayOfSemesterByInstitut = strLastDayOfSemester
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLastDayOfYear
' Description:       Liefert den letzten Tag eines Jahres von ValidDate.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:51:53
'
' Parameters :       validDate (String)
'                    blnDaysOnly (Boolean = False)
'--------------------------------------------------------------------------------
Public Function GetLastDayOfYear(ByVal validDate As String, _
                                 Optional ByVal blnDaysOnly As Boolean = False) As String
                                 
  GetLastDayOfYear = (DateSerial(Year(validDate) + 1, 1, 0))
  If blnDaysOnly Then GetLastDayOfYear = Day(GetLastDayOfYear)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetStartSemester
' Description:       Liefert das Startdatum anhand der Semesteranzahl
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:52:18
'
' Parameters :       lngSemester (Long)
'--------------------------------------------------------------------------------
Public Function GetStartSemester(ByVal lngSemester As Long) As String
  Dim validDate As String
  validDate = DateAdd("m", lngSemester * -6, Now)

  Select Case Month(validDate)
    Case 1, 2
      GetStartSemester = (DateSerial(Year(validDate), 3, 1))
    Case 9, 10, 11, 12
      GetStartSemester = (DateSerial(Year(validDate) + 1, 3, 1))
    Case 3, 4, 5, 6, 7, 8
      GetStartSemester = (DateSerial(Year(validDate), 9, 1))
  End Select

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSemester
' Description:       Liefert die Anzahl Semester anhand von ValidDate.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:52:43
'
' Parameters :       validDate (String)
'                    strBis (String)
'--------------------------------------------------------------------------------
Public Function GetSemester(ByVal validDate As String, strBis As String) As Long

        Dim i      As Date
        Dim lngSem As Long

        If Not IsDate(validDate) Then
                GetSemester = 0

                Exit Function
        Else
                For i = validDate To CDate(strBis)

                        If (Mid$(i, 1, 5) = "01.03") Or (Mid$(i, 1, 5) = "01.09") Then
                                lngSem = lngSem + 1
                        End If

                Next i

                GetSemester = lngSem
        End If

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       Concat
' Description:       Liefert die übergebenen Strings als einen Gesamtstring zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:53:35
'
' Parameters :       values() (Variant)
'--------------------------------------------------------------------------------
Public Function Concat(ParamArray values() As Variant) As String
  Concat = Join(values, "")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       StringFormat
' Description:       Tauscht im übergebenen String die Platzhalte {0}, {1} ...
'                    gegen die übergebenen Parameter.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:54:27
'
' Parameters :       s (String)
'                    values() (Variant)
'--------------------------------------------------------------------------------
Public Function StringFormat(ByVal s As String, ParamArray values() As Variant) As String
  
  Dim i As Integer
  For i = LBound(values) To UBound(values)
    s = Replace$(s, "{" & i & "}", values(i))
  Next i
  
  StringFormat = s
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetBOW
' Description:       Liefert das Datum des ersten Wochentages der Kalenderwoche intKW
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:55:44
'
' Parameters :       intYear (Integer)
'                    intKW (Integer)
'--------------------------------------------------------------------------------
Public Function GetBOW(ByVal intYear As Integer, ByVal intKW As Integer) As Date
  Dim dteFirstDay As Date
  Dim intDayDiff  As Integer
  Dim dteTemp     As Date
  
  dteFirstDay = DateSerial(intYear, 1, 1)

  Select Case Weekday(dteFirstDay)
    Case vbMonday
      intDayDiff = 0
    Case vbTuesday
      intDayDiff = -1
    Case vbWednesday
      intDayDiff = -2
    Case vbThursday
      intDayDiff = -3
    Case vbFriday
      intDayDiff = -4
    Case vbSaturday
      intDayDiff = -5
    Case vbSunday
      intDayDiff = -6
  End Select

  dteTemp = DateAdd("ww", intKW - 1, dteFirstDay)
  dteTemp = DateAdd("d", intDayDiff, dteTemp)
  GetBOW = IIf(dteTemp < dteFirstDay, dteFirstDay, dteTemp)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetEOW
' Description:       Liefert das Datum des letzten Wochentages der Kalenderwoche intKW
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:56:18
'
' Parameters :       intYear (Integer)
'                    intKW (Integer)
'--------------------------------------------------------------------------------
Public Function GetEOW(ByVal intYear As Integer, ByVal intKW As Integer) As Date
  Dim dteFirstDay As Date
  Dim dteLastDay  As Date
  Dim dteKWStart  As Date
  Dim intDayDiff  As Integer
  Dim dteTemp     As Date
  dteFirstDay = DateSerial(intYear, 1, 1)

  Select Case Weekday(dteFirstDay)
    Case vbMonday
      intDayDiff = 0
    Case vbTuesday
      intDayDiff = -1
    Case vbWednesday
      intDayDiff = -2
    Case vbThursday
      intDayDiff = -3
    Case vbFriday
      intDayDiff = -4
    Case vbSaturday
      intDayDiff = -5
    Case vbSunday
      intDayDiff = -6
  End Select

  dteTemp = DateAdd("ww", intKW - 1, dteFirstDay)
  dteKWStart = DateAdd("d", intDayDiff, dteTemp)
  dteTemp = DateAdd("d", 6, dteKWStart)
  dteLastDay = CDate("31.12." & intYear)
  GetEOW = IIf(dteTemp > dteLastDay, dteLastDay, dteTemp)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetQuartal
' Description:       Ermittelt das Quartal von ValidDate
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:56:39
'
' Parameters :       validDate (String)
'--------------------------------------------------------------------------------
Public Function GetQuartal(ByVal validDate As String) As Long

  Select Case Month(validDate)
    Case 1 To 3
      GetQuartal = 1
    Case 4 To 6
      GetQuartal = 2
    Case 7 To 9
      GetQuartal = 3
    Case 10 To 12
      GetQuartal = 4
  End Select
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetQuartalRange
' Description:       Ermittelt den Quartalzeitraum von ValidDate
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:57:02
'
' Parameters :       validDate (String)
'                    dteQuartalStart (Date)
'                    dteQuartalEnd (Date)
'--------------------------------------------------------------------------------
Public Function GetQuartalRange(ByVal validDate As String, _
                                ByRef dteQuartalStart As Date, _
                                ByRef dteQuartalEnd As Date) As Long
  Dim strYear As String
  strYear = Year(validDate)

  Select Case Month(validDate)
    Case 1 To 3
      dteQuartalStart = "01.01." & strYear
      dteQuartalEnd = "31.03." & strYear
      GetQuartalRange = 1
    Case 4 To 6
      dteQuartalStart = "01.04." & strYear
      dteQuartalEnd = "30.06." & strYear
      GetQuartalRange = 2
    Case 7 To 9
      dteQuartalStart = "01.07." & strYear
      dteQuartalEnd = "30.09." & strYear
      GetQuartalRange = 3
    Case 10 To 12
      dteQuartalStart = "01.10." & strYear
      dteQuartalEnd = "31.12." & strYear
      GetQuartalRange = 4
  End Select
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetSemesterCode
' Description:       Liefert das Semester-Kürzel anhand des übergebenen Datums.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:57:43
'
' Parameters :       validDate (Date)
'--------------------------------------------------------------------------------
Public Function GetSemesterCode(ByVal validDate As Date) As String
  Dim strYear As String
  Dim strCode As String

  Select Case Month(validDate)
    Case 1 To 2
      strYear = Year(validDate) - 1
      strCode = "WS"
    Case 3 To 8
      strYear = Year(validDate)
      strCode = "SS"
    Case 9 To 12
      strYear = Year(validDate)
      strCode = "WS"
  End Select

  GetSemesterCode = strYear & " " & strCode
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CollectionContainsKey
' Description:       Prüft, ob eine Collection einen bestimmten Key beinhaltet.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:58:04
'
' Parameters :       col (Collection)
'                    key (String)
'--------------------------------------------------------------------------------
Public Function CollectionContainsKey _
(ByVal col As Collection, ByVal key As String) As Boolean

  Dim firstCheckOk As Boolean
  Dim secondCheckOk As Boolean
  Dim item As Variant
  
  On Error Resume Next
  Set item = col.item(key)
  firstCheckOk = (Err.number = 0)
  On Error GoTo 0
  
  On Error Resume Next
  item = col.item(key)
  secondCheckOk = (Err.number = 0)
  On Error GoTo 0

  CollectionContainsKey = firstCheckOk Or secondCheckOk
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CollectionClear
' Description:       Entfernt alle Elemente aus einer Collection.
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       03.02.2015-15:37:11
'
' Parameters :       col (Collection)
'--------------------------------------------------------------------------------
Public Function CollectionClear(ByVal col As Collection) As Boolean

  Dim i As Integer
  For i = col.count To 1 Step -1
    col.Remove i
  Next i
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SetCollectionItem
' Description:       Erneuert in einer Collection das einem Key zugeordnete Element.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:58:39
'
' Parameters :       col (Collection)
'                    key (String)
'                    newItem (Variant)
'--------------------------------------------------------------------------------
Public Sub SetCollectionItem _
(ByVal col As Collection, ByVal key As String, ByVal newItem As Variant)

  If CollectionContainsKey(col, key) Then
    col.Remove key
    col.Add newItem, key
  End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ArrayCount
' Description:       Liefert die Anzahl in einem Array vorhandenen Elemente.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:59:23
'
' Parameters :       sourceArray (Variant)
'--------------------------------------------------------------------------------
Public Function ArrayCount(ByVal sourceArray As Variant) As Long

  On Error Resume Next
  Dim ret As Long: ret = UBound(sourceArray)
  If Err.number <> 0 Then
    ArrayCount = 0
  Else
    ArrayCount = (UBound(sourceArray) - LBound(sourceArray)) + 1
  End If
  On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       Contains
' Description:       Prüft, ob eine Zeichenfolge in einem String vorhanden ist.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:00:05
'
' Parameters :       source (String)
'                    containsString (String )
'                    ignoreCase (Boolean = False)
'--------------------------------------------------------------------------------
Public Function Contains _
(ByVal source As String _
, ByVal containsString As String _
, Optional ByVal ignoreCase As Boolean = False) As Boolean

  Contains = InStr(IIf(ignoreCase, LCase$(source), source) _
  , IIf(ignoreCase, LCase$(containsString), containsString)) > 0
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       TrimWhiteSpace
' Description:       Entfernt sämtlichen voran- und angehängten Whitespace aus
'                    einem String.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-16:00:37
'
' Parameters :       s (String)
'--------------------------------------------------------------------------------
Public Function TrimWhiteSpace(ByVal s As String) As String
  TrimWhiteSpace = ""
  Dim l As Long: l = Len(s)
  
  If l > 0 Then
    
    Dim posLeft As Long: posLeft = 1
    While (Asc(Mid$(s, posLeft, 1)) <= 32) And (posLeft < l)
      posLeft = posLeft + 1
    Wend
    
    Dim posRight As Long: posRight = l
    While (Asc(Mid$(s, posRight, 1)) <= 32) And (posRight > posLeft)
      posRight = posRight - 1
    Wend

    If Not ((l = posLeft) And (l = posRight)) Then
      TrimWhiteSpace = Mid$(s, posLeft, (posRight - posLeft) + 1)
    End If
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :
' Procedure  :       IsDevelopTime
' Description:       Überprüft, ob das Programm in der IDE ausgeführt wird.
' Created by :       Sebastian Limke
' Date-Time  :       5.19.2014-11:30:05'
'
' Parameters :       hInstanceApp (Long)
'--------------------------------------------------------------------------------
Public Function IsDevelopTime(ByVal hInstanceApp As Long) As Boolean
  Dim filename As String: filename = String(255, 0)
  Dim count As Long: count = GetModuleFileName(hInstanceApp, filename, 255)
  filename = Left(filename, count)
  filename = LCase(Mid$(filename, InStrRev(filename, "\") + 1))
  IsDevelopTime = (filename Like "vb*.exe")
End Function

'--------------------------------------------------------------------------------
' Project    :       Etc
' Procedure  :       IsTestVersion
' Description:       Ermittelt, ob es sich um die Version handelt, die auf den Testserver geht
' Created by :       Sebastian Limke
' Date-Time  :       10.20.2014-08:42:33'
'
' Parameters :       appPath (String)
'--------------------------------------------------------------------------------
Public Function IsTestVersion(ByVal appPath As String) As Boolean
  IsTestVersion = (appPath Like "*\_development\*")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       BetweenNumeric
' Description:       Prüft, ob value zwischen min und max liegt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       11.02.2015-14:53:33
'
' Parameters :       value (Long)
'                    min (Long)
'                    max (Long)
'--------------------------------------------------------------------------------
Public Function BetweenNumeric(ByVal value As Long, ByVal min As Long, ByVal Max As Long) As Boolean
  
  BetweenNumeric = (value >= min) And (value <= Max)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       BetweenDate
' Description:       Prüft, ob value zwischen min und max liegt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       11.02.2015-14:53:55
'
' Parameters :       value (Date)
'                    min (Date)
'                    max (Date)
'--------------------------------------------------------------------------------
Public Function BetweenDate(ByVal value As Date, ByVal min As Date, ByVal Max As Date) As Boolean

  BetweenDate = (value >= min) And (value <= Max)
End Function
