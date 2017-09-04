VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropertiesViewReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type typProperties
  Field As String           '{Orginärer Feldname }
  value As String
  EingabeFokus As String    '{Eingabe über TextBox,ComboBox,ComboBoxEx}
  
  'hack2
  
  DataString As String      '{Zum Füllen der Combobox erstes Feld ID zweites Wert}
  FormatType As String      '{Darstellung lCase, uCase, uFormat, "Normal" }
  PutID As Boolean          '{Soll ID-Wert geschrieben werden oder Wert (ComboBox)}
  ReadOnly As Boolean       '{Nur Lesezugriff}
  ValueType As String       '{[String*,password,Memo,Date,DateNull,!Time!,Integer,Real] Datentyp für Eingabeüberprüfung}
  MaxLen As String          '{Maximale Länge für Texteingabe in Textbox}
  ValueArea As String       '{Wertebereich von#bis#Nachkommastellen}
End Type

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long _
, ByVal wParam As Long, lParam As Any) As Long

Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_ERR = (-1)
Private Const CB_GETCURSEL = &H147

Public gstrMemo As String


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
'*  Letzte Änderung: 25.03.2000                            *'
'*                                                         *'
'*  Modul: datRead                                         *'
'*                                                         *'
'*  Version: 1.0                                           *'
'*                                                         *'
'*  Modulbeschreibung: datRead prüft, ob das Argument      *'
'*  <PruefWert> ein Datum ist. Sollte dies nicht der Fall  *'
'*  sein, liefert datRead einen Wert <Default>  zurück     *'
'*  oder beim Fehlen von <Default> das Systemdatum und gibt*'
'*  eine Fehlermeldung aus. Optional kann noch eine untere *'
'*  Datumsgrenze <Min> und eine obere Datumsgrenze <Max>   *'
'*  übergeben werden. Sollte bei Übergabe von <Min> und/   *'
'*  oder Max PruefWert nicht im Wertebereich befinden,     *'
'*  liefert datRead <Default> als Rückgabewert und gibt    *'
'*  eine Fehlermeldung aus. <Size> bestimmt das Ausgabe-   *'
'*  format ("Short Date", "Long Date"), bei fehlendem      *'
'*  Argument wird automatisch nach "Short Date" formatiert.*'
'*                                                         *'
'*  Probleme: keine bekannt                                *'
'*                                                         *'
'*  Bekannte Bugs: Keine                                   *'
'*                                                         *'
'*  Autor: Sascha Glinka alias Vanitas Mundi               *'
'*                                                         *'
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'

Public Function datRead(ByVal PruefWert As Variant _
, Optional ByVal Default As String, Optional ByVal Size As Variant _
, Optional ByVal Min As Variant _
, Optional ByVal Max As Variant) As String

On Error GoTo Fehler

'*** Zuweisen eines Standardformates ***'
  If IsMissing(Size) Then
    Size = "Short Date"
  End If
'*** Ende Zuweisen eines Standardformates ***'

'*** Festlegen eines defaultwertes ***'
  If IsMissing(Default) Then
    Default = Date
  End If
'*** Ende Festlegen eines defaultwertes ***'

'*** Ist Pruefwert ein Datum? ***'
  If Not IsDate(PruefWert) Then
    MsgBox "Eingabe ungültig - kein Datumsformat" _
    , 16, "Formatfehler"
    datRead = Default
    Exit Function
  End If
'*** Ende Ist Pruefwert ein Datum? ***'

'*** Wurde MinGrenze unterschritten? ***'
  If Not IsMissing(Min) Then
    If PruefWert < CDate(Min) Then
      MsgBox "Eingabe ungültig - Datumsgrenze von " _
      & Min & " unterschritten", 16, "Formatfehler"
      datRead = Default
    Exit Function
    End If
  End If
'*** Ende Wurde MinGrenze unterschritten? ***'
  
'*** Wurde MaxGrenze überschritten? ***'
  If Not IsMissing(Max) Then
    If PruefWert > CDate(Max) Then
      MsgBox "Eingabe ungültig - Datumsgrenze von " _
      & Max & " überschritten", 16, "Formatfehler"
      datRead = Default
    Exit Function
    End If
  End If
'*** Ende Wurde MaxGrenze überschritten? ***'
  
  
  datRead = Format(PruefWert, Size) '*** Rückgabe des PruefWertes
  
  Exit Function
  
'*** Bei Auftritt eines Fehlers ***'
Fehler:
  MsgBox "Datentyp-Diskrepanz der Eingangsargumente" _
  , 16, "Argumentfehler"
  datRead = Default
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
'*  Letzte Änderung: 25.03.2000                            *'
'*                                                         *'
'*  Modul: dblRead                                         *'
'*                                                         *'
'*  Version: 1.0                                           *'
'*                                                         *'
'*  Modulbeschreibung: dblRead prüft, ob das Argument      *'
'*  <PruefWert> numerisch ist. Sollte dies nicht der Fall  *'
'*  sein, liefert dblReadeine 0 zurück und gibt eine       *'
'*  Fehlermeldung aus. Optional kann noch eine untere      *'
'*  Wertgrenze <Min> und eine obere Wertgrenze <Max>       *'
'*  übergeben werden. Sollte bei Übergabe von <Min> und/   *'
'*  oder Max PruefWert nicht im Wertebereich befinden,     *'
'*  liefert dblRead eine 0 als Rückgabewert und gibt eine  *'
'*  Fehlermeldung aus. Der Parameter <Dez> legt die        *'
'*  zurückgegebenen Dezimalstellen fest.                   *'
'*                                                         *'
'*  Probleme: keine bekannt                                *'
'*                                                         *'
'*  Bekannte Bugs: Keine                                   *'
'*                                                         *'
'*  Autor: Sascha Glinka alias Vanitas Mundi               *'
'*                                                         *'
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'

Public Function dblRead _
(ByVal PruefWert _
, ByRef blnErr As Boolean _
, Optional ByVal Dez _
, Optional ByVal Min _
, Optional ByVal Max) As Double

Dim strFormatierung As String
Dim i As Byte

On Error GoTo Fehler

'*** Punkt durch Komma tauschen ***'
  If InStr(PruefWert, ".") > 0 Then
    PruefWert = Mid(PruefWert, 1, InStr(PruefWert, ".")) _
    & "," & Mid(PruefWert, InStr(PruefWert, ".") + 1)
  End If
'*** Ende Punkt durch Komma tauschen ***'

'*** Ist Pruefwert eine Zahl? ***'
  If Not IsNumeric(PruefWert) Then
    MsgBox "Eingabe ungültig - kein numerischer Wert" _
    , 16, "Formatfehler"
    dblRead = 0
    blnErr = True
    Exit Function
  End If
'*** Ende Ist Pruefwert eine Zahl? ***'

'*** Wurde MinGrenze unterschritten? ***'
  If Not IsMissing(Min) Then
    If PruefWert < CDbl(Min) Then
      MsgBox "Eingabe ungültig - untere Grenze von " _
      & Min & " unterschritten", 16, "Formatfehler"
      dblRead = 0
      blnErr = True
      Exit Function
    End If
  End If
'*** Ende Wurde MinGrenze unterschritten? ***'
  
'*** Wurde MaxGrenze überschritten? ***'
  If Not IsMissing(Max) Then
    If PruefWert > CDbl(Max) Then
      MsgBox "Eingabe ungültig - obere Grenze von " _
      & Max & " überschritten", 16, "Formatfehler"
      dblRead = 0
      blnErr = True
      Exit Function
    End If
  End If
'*** Ende Wurde MaxGrenze überschritten? ***'
  
  If Not IsMissing(Dez) Then
    '*** Formatierungsschlüssel festlegen ***'
    strFormatierung = "##,"
    For i = 1 To Dez
      strFormatierung = strFormatierung & "#"
    Next i
      strFormatierung = strFormatierung & "00."
    For i = 1 To Dez
      strFormatierung = strFormatierung & "0"
    Next i
    '*** Ende Formatierungsschlüssel festlegen ***'
    '*** Rückgabe des PruefWertes
    dblRead = CDbl(Format(Round(PruefWert, Dez), strFormatierung))
  Else
    dblRead = PruefWert
  End If
  Exit Function
  
'*** Bei Auftritt eines Fehlers ***'
Fehler:
  MsgBox "Datentyp-Diskrepanz der Eingangsargumente" _
  , 16, "Argumentfehler"
  dblRead = 0
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
'*  Letzte Änderung: 25.03.2000                            *'
'*                                                         *'
'*  Modul: lngRead                                         *'
'*                                                         *'
'*  Version: 1.0                                           *'
'*                                                         *'
'*  Modulbeschreibung: lngRead prüft, ob das Argument      *'
'*  <PruefWert> numerisch ist. Sollte dies nicht der Fall  *'
'*  sein, liefert lngRead eine 0 zurück und gibt eine      *'
'*  Fehlermeldung aus. Optional kann noch eine untere      *'
'*  Wertgrenze <Min> und eine obere Wertgrenze <Max>       *'
'*  übergeben werden. Sollte bei Übergabe von <Min> und/   *'
'*  oder Max PruefWert nicht im Wertebereich befinden,     *'
'*  liefert lngRead eine 0 als Rückgabewert und gibt eine  *'
'*  Fehlermeldung aus.                                     *'
'*                                                         *'
'*  Probleme: keine bekannt                                *'
'*                                                         *'
'*  Bekannte Bugs: Keine                                   *'
'*                                                         *'
'*  Autor: Sascha Glinka alias Vanitas Mundi               *'
'*                                                         *'
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'

Public Function lngRead(ByVal PruefWert As Variant _
, ByRef blnErr As Boolean _
, Optional ByVal Min As Variant _
, Optional ByVal Max As Variant) As Long

On Error GoTo Fehler

'*** Ist Pruefwert eine Zahl? ***'
  If Not IsNumeric(PruefWert) Then
    MsgBox "Eingabe ungültig - kein numerischer Wert" _
    , 16, "Formatfehler"
    lngRead = 0
    blnErr = True
    Exit Function
  End If
'*** Ende Ist Pruefwert eine Zahl? ***'

'*** Wurde MinGrenze unterschritten? ***'
  If Not IsMissing(Min) Then
    If PruefWert < CLng(Min) Then
      MsgBox "Eingabe ungültig - untere Grenze von " _
      & Min & " unterschritten", 16, "Formatfehler"
      lngRead = 0
      blnErr = True
    Exit Function
    End If
  End If
'*** Ende Wurde MinGrenze unterschritten? ***'
  
'*** Wurde MaxGrenze überschritten? ***'
  If Not IsMissing(Max) Then
    If PruefWert > CLng(Max) Then
      MsgBox "Eingabe ungültig - obere Grenze von " _
      & Max & " überschritten", 16, "Formatfehler"
      lngRead = 0
      blnErr = True
    Exit Function
    End If
  End If
'*** Ende Wurde MaxGrenze überschritten? ***'
  
  lngRead = CLng(PruefWert) '*** Rückgabe des PruefWertes
  
  Exit Function
  
'*** Bei Auftritt eines Fehlers ***'
Fehler:
  MsgBox "Datentyp-Diskrepanz der Eingangsargumente" _
  , 16, "Argumentfehler"
  lngRead = 0
End Function




