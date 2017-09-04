VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEditInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eceiSpecialType
  eceiSBlank = 0
  eceiSMemo = 1
  eceiSeMail = 2
  eceiSURL = 3
  eceiSSelect = 4
  eceiSPassword = 5
  eceiSPhoneNumber = 6
End Enum

Public Enum eValueType
  evtString = 0
  evtInteger = 1
  evtReal = 2
  evtDate = 3
  evtTime = 4
End Enum

Public Visible As Boolean
Attribute Visible.VB_VarDescription = "Blendet eine Eigenschaft in der Ansicht aus (false) oder wieder ein (true)."
Public Writeable As Boolean
Attribute Writeable.VB_VarDescription = "Schaltet den Schreibschutz der Eigenschaft ein (false) oder wieder aus (true)."
Public Readable As Boolean
Attribute Readable.VB_VarDescription = "Bestimmt ob eine Eigenschaft gelesen werden darf (true). In der Regel macht es keinen Sinn Readable auf false zu setzen."
Public Format As String
Attribute Format.VB_VarDescription = "Formatiert die Darstellung der Eigenschaft in der Ansicht (sämtliche Formatierungszeichenfolgen, welche die VB-Funktion Format unterstützt, sind kompatibel)."
Public Name As String
Attribute Name.VB_VarDescription = "Liefert den Namen der Eigenschaft, welcher z.B. benutzt werden kann, um per Items-Auflistung des Steuerelements, auf ein bestimmtes clsClassEsitInfos-Objekt zu verweisen."
Public Alias As String
Attribute Alias.VB_VarDescription = "Maskiert den Namen Objekt-Eigenschaft in der Ansicht. Wird die Alias-Eigenschaft nicht gesetzt, so wird der originäre Name verwendet."
Public ValueType As eValueType
Attribute ValueType.VB_VarDescription = "evtString = Beliebige Zeichenfolge\r\nevtInteger = Ganzzahl\r\nevtReal = Fließkommazahl\r\nevtDate = Datum\r\nevtTime = Uhrzeit\r\n"
Public ValueID As String
Attribute ValueID.VB_VarDescription = "Beinhaltet den ID-Wert, welcher benutzt wird, wenn ReturnID den Wert True besitzt"
Public SpecialType As eceiSpecialType
Attribute SpecialType.VB_VarDescription = "eceiSBlank = Schlichtes Eingabefeld.\r\neceiSMemo = Mehrzeiliges Eingabefeld.\r\neceiSeMail = Öffnet Fastmailer mit Eigenschaftswert als Adressat.\r\neceiSURL = Öffnet Eigenschaftswert als URL mit dem Standardbrowser.\r\neceiSSelect = Listet eine Auswahl(SelectD"
Public SelectData As Variant
Attribute SelectData.VB_VarDescription = "Stellt die Auswahlliste für den SpecialType eceiSSelect zur Verfügung. SelectData kann eine Collection sein oder aber ein String, wobei die einzelnen Optionen durch eine Raute getrennt sein müssen (Option1#Option2#...). Sollte statt der Option ein ID-Wer"
Public ReturnID As Boolean
Attribute ReturnID.VB_VarDescription = "Legt fest ob die Eigenschaft maskiert dargestellt und stellvertretend für den Value die ValueID gespeichert wird. Funktioniert nur mit dem SpecialType eceiSSelect und bei Übergabe von SelectData, sowie SelectOnlyFromSelectData gleich true."
Public SortIndex As Long
Attribute SortIndex.VB_VarDescription = "Setzt den Sortierungsindex für die Eigenschaft. Wenn keine alphanumerische Sortierung gewünscht ist, kann über die SortIndex-Eigenschaft, eine eigene Sortierung angegeben werden."
Public AllowNullString As Boolean
Attribute AllowNullString.VB_VarDescription = "Legt fest wie auf die Eingabe eines Nullstrings reagiert wird. Standardwert ist false, d.h. es werden keine Leereingaben akzeptiert."
Public Min As Variant
Attribute Min.VB_VarDescription = "Bestimmt die untere Grenze für die Werteingabe der Eigenschaft. Die Untergrenze kann bezüglich des Datentyps der Eigenschaft beliebige Werte aufnehmen."
Public Max As Variant
Attribute Max.VB_VarDescription = "Setzt eine Obergrenze für die Werteingabe der Eigenschaft. Die Obergrenze kann bezüglich des Datentyps der Eigenschaft beliebige Werte aufnehmen."
Public MaxLenght As Long
Attribute MaxLenght.VB_VarDescription = "Schränkt die Länge der Texteingabe auf MaxLenght Zeichen ein. Eine 0 repräsentiert beliebig viele Zeichen."
Public SelectOnlyFromSelectData As Boolean
Attribute SelectOnlyFromSelectData.VB_VarDescription = "Kann nur aktiviert werden mit dem SpecialType eceiSSelect und lässt den Benutzer entweder nur aus der Auswahl SelectData wählen (True) oder gewährt Freieingaben und eine Selektion aus SelectData."
Public Required As Boolean

Private mstrShownValue As String
Private mstrValue As String


Public Property Get Value() As String
Attribute Value.VB_Description = "Setzt den Wert der Eigenschaft."
  Value = mstrValue
End Property

Public Property Let Value(ByVal strValue As String)
  mstrValue = strValue
  If Me.SpecialType = eceiSPassword Then
    mstrShownValue = String(Len(strValue), "*")
  Else
    mstrShownValue = strValue
  End If
End Property

Public Property Get ShownValue() As String
Attribute ShownValue.VB_Description = "Beinhaltet den Wert der in der Ansicht angezeigt wird."
  ShownValue = mstrShownValue
End Property

Private Sub Class_Initialize()
  Me.SpecialType = eceiSBlank
  Me.Visible = True
  Me.Writeable = False
  Me.Readable = False
  Me.Format = ""
  Me.Name = ""
  Me.Alias = ""
  Me.ValueType = evtString
  mstrShownValue = ""
  Me.Value = ""
  Me.ValueID = ""
  Me.SelectData = ""
  Me.ReturnID = False
  Me.SortIndex = 0
  Me.AllowNullString = False
  Me.Min = ""
  Me.Max = ""
  Me.MaxLenght = 0
  Me.SelectOnlyFromSelectData = True
  Me.Required = False
End Sub



