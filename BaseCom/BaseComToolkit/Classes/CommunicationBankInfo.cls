VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommunicationBankInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : CommunicationBankInformation
'    Project    : ToolKits
'
'    Description: Stellt Methoden für Abfrage von Bankinformationen zur Verfügung.
'                 Benötigt einen gesetzten ConnectionString.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum BankInformationFields
  bankCode = 1
  BankName = 2
  Postcode = 3
  city = 4
  BankShortName = 5
  bic = 6
  CheckDigitCalculation = 7
End Enum

Public Enum BankInformationErrors
  NoError = 0
  UnknownError = 1
  InvalidLenght = 2
  InvalidCulture = 3
  InvalidCheckDigit = 4
  NotSupportedCulture = 5
  InvalidBicLenght = 6
End Enum

Private Const BANK_INFORMATION_STATEMENT As String = "SELECT {0} AS Value FROM def.t_bank_information WHERE {1} = '{2}' GROUP BY {0}"
Private mConnectionString         As String
Private mIbanLenghtCultures As Collection
Private mBankInformationError As BankInformationErrors



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mIbanLenghtCultures = New Collection
  InitializeIbanLenghtCultures
End Sub

Private Sub Class_Terminate()
  Set mIbanLenghtCultures = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ConnectionString() As String
  ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal value As String)
  mConnectionString = value
End Property

Public Property Get BankInformationError() As BankInformationErrors
  BankInformationError = mBankInformationError
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------

'---------------------- Private Methoden der Klasse -----------------------------
Private Sub InitializeIbanLenghtCultures()
  With mIbanLenghtCultures
    .Add 27, "EG"  'Ägypten
    .Add 28, "AL"  'Albanien
    .Add 24, "DZ"  'Algerien
    .Add 24, "AD"  'Andorra
    .Add 25, "AO"  'Angola
    .Add 28, "AZ"  'Aserbaidschan
    .Add 22, "BH"  'Bahrain
    .Add 16, "BE"  'Belgien
    .Add 28, "BJ"  'Benin
    .Add 20, "BA"  'Bosnien und Herzegowina
    .Add 29, "BR"  'Brasilien
    .Add 24, "VG"  'Britische Jungferninseln
    .Add 22, "BG"  'Bulgarien
    .Add 27, "BF"  'Burkina Faso
    .Add 16, "BI"  'Burundi
    .Add 21, "CR"  'Costa Rica
    .Add 28, "CI"  'Côte d'Ivoire (Elfenbeinküste)
    .Add 18, "DK"  'Dänemark
    .Add 22, "DE"  'Deutschland
    .Add 28, "DO"  'Dominikanische Republik
    .Add 20, "EE"  'Estland
    .Add 18, "FO"  'Färöer
    .Add 18, "FI"  'Finnland
    .Add 27, "FR"  'Frankreich, Französisch-Guayana, Französisch-Polynesien, Französische Süd- und Antarktisgebiete, Guadeloupe, Martinique, Réunion, Mayotte, Neukaledonien, Saint-Barthélemy, Saint-Martin, Saint-Pierre und Miquelon, Wallis und Futuna
    .Add 27, "GA"  'Gabun
    .Add 22, "GE"  'Georgien
    .Add 23, "GI"  'Gibraltar
    .Add 27, "GR"  'Griechenland
    .Add 18, "GL"  'Grönland
    .Add 28, "GT"  'Guatemala
    .Add 26, "IR"  'Iran
    .Add 22, "IE"  'Irland
    .Add 26, "IS"  'Island
    .Add 23, "IL"  'Israel
    .Add 27, "IT"  'Italien
    .Add 30, "JO"  'Jordanien
    .Add 27, "CM"  'Kamerun
    .Add 25, "CV"  'Kap Verde
    .Add 20, "KZ"  'Kasachstan
    .Add 29, "QA"  'Katar
    .Add 27, "CG"  'Kongo (Brazzaville)
    .Add 20, "XK"  'Kosovo
    .Add 21, "HR"  'Kroatien
    .Add 30, "KW"  'Kuwait
    .Add 21, "LV"  'Lettland
    .Add 28, "LB"  'Libanon
    .Add 21, "LI"  'Liechtenstein
    .Add 20, "LT"  'Litauen
    .Add 20, "LU"  'Luxemburg
    .Add 27, "MG"  'Madagaskar
    .Add 28, "ML"  'Mali
    .Add 31, "MT"  'Malta
    .Add 27, "MR"  'Mauretanien
    .Add 30, "MU"  'Mauritius
    .Add 19, "MK"  'Mazedonien
    .Add 24, "MD"  'Moldawien
    .Add 27, "MC"  'Monaco
    .Add 22, "ME"  'Montenegro
    .Add 25, "MZ"  'Mosambik
    .Add 18, "NL"  'Niederlande
    .Add 15, "NO"  'Norwegen
    .Add 20, "AT"  'Österreich
    .Add 23, "TL"  'Osttimor
    .Add 24, "PK"  'Pakistan
    .Add 29, "PS"  'Palästinensische Autonomiegebiete
    .Add 28, "PL"  'Polen
    .Add 25, "PT"  'Portugal
    .Add 24, "RO"  'Rumänien
    .Add 27, "SM"  'San Marino
    .Add 25, "ST"  'São Tomé und Príncipe
    .Add 24, "SA"  'Saudi-Arabien
    .Add 24, "SE"  'Schweden
    .Add 21, "CH"  'Schweiz
    .Add 28, "SN"  'Senegal
    .Add 22, "RS"  'Serbien
    .Add 24, "SK"  'Slowakei
    .Add 19, "SI"  'Slowenien
    .Add 24, "ES"  'Spanien
    .Add 24, "CZ"  'Tschechien
    .Add 24, "TN"  'Tunesien
    .Add 26, "TR"  'Türkei
    .Add 28, "HU"  'Ungarn
    .Add 23, "AE"  'Vereinigte Arabische Emirate
    .Add 22, "GB"  'Vereinigtes Königreich, Jersey, Guernsey, Isle of Man
    .Add 28, "CY"  'Zypern
    .Add 27, "CF"  'Zentralafrikanische Republik
  End With
End Sub

Private Function GetBankInformationFieldString(ByVal bankInformationField As BankInformationFields)

  Select Case bankInformationField
    Case BankInformationFields.bankCode
      GetBankInformationFieldString = "bank_code"

    Case BankInformationFields.BankName
      GetBankInformationFieldString = "name"

    Case BankInformationFields.BankShortName
      GetBankInformationFieldString = "short_name"

    Case BankInformationFields.bic
      GetBankInformationFieldString = "bic"

    Case BankInformationFields.CheckDigitCalculation
      GetBankInformationFieldString = "check_digit_calculation"
      
    Case BankInformationFields.city
      GetBankInformationFieldString = "city"
      
    Case BankInformationFields.Postcode
      GetBankInformationFieldString = "post_code"
      
    Case Else
      GetBankInformationFieldString = ""
  End Select

End Function

Private Function GetBankInformationFieldGermanString(ByVal bankInformationField As BankInformationFields)

  Select Case bankInformationField

    Case BankInformationFields.bankCode
      GetBankInformationFieldGermanString = "Bankleitzahl"

    Case BankInformationFields.BankName
      GetBankInformationFieldGermanString = "Bankname"

    Case BankInformationFields.BankShortName
      GetBankInformationFieldGermanString = "Bank-Kurzbezeichnung"

    Case BankInformationFields.bic
      GetBankInformationFieldGermanString = "BIC"

    Case BankInformationFields.CheckDigitCalculation
      GetBankInformationFieldGermanString = "Prüfziffer-Berechnungs-Methode"
      
    Case BankInformationFields.city
      GetBankInformationFieldGermanString = "Stadt"
      
    Case BankInformationFields.Postcode
      GetBankInformationFieldGermanString = "Postleitzahl"
      
    Case Else
      GetBankInformationFieldGermanString = ""
  End Select
End Function

Private Function GetValueBase(ByVal getField As BankInformationFields, _
                              ByVal getBy As BankInformationFields, _
                              ByVal value As String) As String
                              
  mBankInformationError = UnknownError

  Dim etc As etc: Set etc = New etc
  Dim s   As String: s = etc.StringFormat(BANK_INFORMATION_STATEMENT _
  , GetBankInformationFieldString(getField), GetBankInformationFieldString(getBy), value)

  Dim db  As Database: Set db = New Database
  db.ConnectionString = Me.ConnectionString
  Dim sb As DialogSelectEntry: Set sb = New DialogSelectEntry
  sb.SetDataSource db.ExecuteReader(s)
  sb.ReturnSingleEntryAutomatically = True
  
  Dim title As String: title = etc.StringFormat("Bitte {0} wählen ...", GetBankInformationFieldGermanString(getField))
  
  sb.SelectEntry , title, False, True
  
  GetValueBase = sb.ValueEntry
  
  Set etc = Nothing
  Set sb = Nothing
  Set db = Nothing
  
  mBankInformationError = NoError
End Function

Private Function IbanCharToNumber(ByVal c As String) As Integer
  'Die beiden Buchstaben des Länderkürzels sowie weitere etwa in der Kontonummer enthaltene Buchstaben
  'werden durch ihre Position im lateinischen Alphabet + 9 ersetzt (A = 10, B = 11, …, Z = 35).
  Dim number As Integer: number = Asc(UCase(c)) - 55
  '(A = 10, B = 11, …, Z = 35).
  IbanCharToNumber = IIf(number >= 10 And number <= 35, number, c)
End Function

Private Function NumberToIbanChar(ByVal n As Integer) As String
  If (n >= 10) And (n <= 35) Then
    NumberToIbanChar = Chr(n + 55)
  Else
    NumberToIbanChar = CStr(n)
  End If
End Function

Private Function Modulo97Result(ByVal checksum As String) As Long
  
  Dim currentLen As Long: currentLen = 0
  Dim lenRest As Long: lenRest = 0
  Dim checksumMod97 As String: checksumMod97 = ""
  
  Do
    lenRest = 9 - currentLen
    Dim value As String: value = Mid$(checksum, 1, lenRest)
    checksumMod97 = Val(checksumMod97 & value) Mod 97
    currentLen = Len(checksumMod97)
  
    checksum = Mid$(checksum, Len(value) + 1)
  Loop Until value = ""

  Modulo97Result = checksumMod97
End Function



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetBankCode
' Description:       Ermittelt die BLZ anhand des angegebenen Feldes.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (BankInformationFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetBankCode(ByVal getBy As BankInformationFields, ByVal value As String) As String
  
  GetBankCode = GetValueBase(bankCode, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetBankName
' Description:       Ermittelt den Banknamen anhand des angegebenen Feldes.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (BankInformationFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetBankName(ByVal getBy As BankInformationFields, ByVal value As String) As String
  
  GetBankName = GetValueBase(BankName, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetPostcode
' Description:       Ermittelt die Banken-PLZ anhand des angegebenen Feldes.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (BankInformationFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetPostcode(ByVal getBy As BankInformationFields, ByVal value As String) As String
  
  GetPostcode = GetValueBase(Postcode, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetCity
' Description:       Ermittelt die Banken-Stadt anhand des angegebenen Feldes.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (BankInformationFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetCity(ByVal getBy As BankInformationFields, ByVal value As String) As String
  
  GetCity = GetValueBase(city, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetBankShortName
' Description:       Ermittelt den Banken-Kurznamen anhand des angegebenen Feldes.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (BankInformationFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetBankShortName(ByVal getBy As BankInformationFields, ByVal value As String) As String
  
  GetBankShortName = GetValueBase(BankShortName, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetBic
' Description:       Ermittelt die BIC anhand des angegebenen Feldes.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (BankInformationFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetBic(ByVal getBy As BankInformationFields, ByVal value As String) As String
  
  GetBic = GetValueBase(bic, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetCheckDigitCalculation
' Description:       Ermittelt die Prüfzahl-Berechnungs-Methode anhand des angegebenen Feldes.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (BankInformationFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetCheckDigitCalculation(ByVal getBy As BankInformationFields, ByVal value As String) As String
  
  GetCheckDigitCalculation = GetValueBase(CheckDigitCalculation, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsValidIBAN
' Description:       Prüft die übergebene IBAN auf Gültigkeit.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       10.03.2015-17:48:16
'
' Parameters :       iban (String)
'--------------------------------------------------------------------------------
Public Function IsValidIBAN(ByVal iban As String) As Boolean
  
  mBankInformationError = UnknownError
  
  IsValidIBAN = True
  
  iban = UCase(Trim$(iban))
  
  If iban = "" Then
    IsValidIBAN = False
    mBankInformationError = InvalidCulture
  Else
    Dim culture As String: culture = Mid$(iban, 1, 2)
    
    Dim etc As etc: Set etc = New etc
    If etc.CollectionContainsKey(mIbanLenghtCultures, culture) Then
      Dim ibanLen As Integer: ibanLen = mIbanLenghtCultures.item(culture)
      
      If Not Len(iban) = ibanLen Then
        IsValidIBAN = False
        mBankInformationError = InvalidLenght
      End If
    Else
      IsValidIBAN = False
      mBankInformationError = InvalidCulture
    End If
  End If
  
  If IsValidIBAN Then mBankInformationError = NoError
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsValidCheckDigit
'
' Description:       Prüft die Prüfziffer der übergebenen IBAN auf Gültigkeit.
'                    Im Rahmen von IPI und ECBS wurde das Prüfziffernverfahren
'                    auf den Standard ISO 7064 mod 97-10 festgelegt
'                    und von den nationalen Normungsgremien der Banken übernommen.
'                    Die Prüfsumme liegt im Bereich „02“ bis „98“, durch abweichende
'                    Berechnungsmethoden bei verschiedenen Finanzinstituten kamen aber
'                    auch „00“, „01“ und „99“ vor.
'                    Die Korrekturen sind weitgehend abgeschlossen. Jedoch kann nicht
'                    ausgeschlossen werden, dass weiterhin IBANs mit „falschen“ Prüfziffern existieren.
'
'                    Zur Validierung der Prüfsumme wird zunächst eine Zahl erstellt.
'                    Diese setzt sich aus BBAN(in Deutschland z. B. 18 Stellen) + Länderkürzel
'                    kodiert (2 Stellen, siehe Punkt 2) + Prüfsumme zusammen.
'
'                    Die beiden Buchstaben des Länderkürzels sowie weitere etwa in der
'                    Kontonummer enthaltene Buchstaben werden durch ihre Position im lateinischen
'                    Alphabet + 9 ersetzt (A = 10, B = 11, …, Z = 35).
'
'                    Nun wird der Rest berechnet, der sich beim Teilen der Zahl durch 97 ergibt (Modulo 97).
'                    Das Ergebnis muss 1 sein, ansonsten ist die IBAN falsch.
'
'                    Beispiel:
'
'                    IBAN:      DE68 2105 0170 0012 3456 78
'                    Umstellung:     2105 0170 0012 3456 78DE 68
'                    Modulo:         210501700012345678131468 mod 97 = 1
'
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       10.03.2015-17:48:16
'
' Parameters :       iban (String)
'--------------------------------------------------------------------------------
Public Function IsValidCheckDigit(ByVal iban As String) As Boolean

  mBankInformationError = UnknownError
  
  iban = Mid$(iban, 5) + Mid$(iban, 1, 4)
  
  Dim checksum As String
  
  Dim i As Integer
  For i = 1 To Len(iban)
    checksum = checksum & IbanCharToNumber(Mid$(iban, i, 1))
  Next i
  
  IsValidCheckDigit = (Modulo97Result(checksum) = 1)
  mBankInformationError = IIf(IsValidCheckDigit, NoError, InvalidCheckDigit)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsValidBic
' Description:       Prüft die angegebene Bic auf formale Gültigkeit.
'
'                    Der BIC oder SWIFT-Code besteht aus acht oder elf alphanumerischen
'                    Zeichen und ist wie folgt untergliedert:
'
'                    4-stelliger Bankcode
'                    2-stelliger Ländercode
'                    2-stellige Codierung des Ortes
'                    3-stellige Kennzeichnung der Filiale (optional)
'
'                    Durch diesen Code kann innerhalb des neuen europäischen Zahlungsraumes,
'                    also in allen EU-Mitgliedsstaaten jedes Kreditinstitut einwandfrei
'                    identifiziert werden.
'
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       12.03.2015-11:46:20
'
' Parameters :       bic (String)
'--------------------------------------------------------------------------------
Public Function IsValidBic(ByVal bic As String) As Boolean

  mBankInformationError = UnknownError
  
  bic = Trim$(bic)
  IsValidBic = ((Len(bic) = 8) Or (Len(bic) = 11))

  mBankInformationError = IIf(IsValidBic, NoError, InvalidBicLenght)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CalcCheckDigit
'
' Description:       Berechnung der IBAN-Prüfsumme
'
'                    Durch Umkehrung der Validierungsmethode ist es möglich,
'                    bezüglich der Prüfsummenbildung korrekt validierende IBANs
'                    zu erzeugen. Die ECBS weist ausdrücklich darauf hin, dass
'                    nur die Banken selbst korrekte IBANs herausgeben.
'                    Benutzt man eine selbsterrechnete IBAN, so riskiert man im
'                    schlimmsten Fall den Verlust seiner Überweisung. Bereits
'                    die ISO 13616 legt fest, dass ausschließlich Finanzinstitute
'                    eine IBAN generieren dürfen. Hintergrund ist, dass die
'                    bisherige Kontonummer nicht immer rechtsbündig in die IBAN
'                    eingesetzt wird, sondern gelegentlich um zwei Stellen
'                    versetzt – die letzten beiden Ziffern entsprechen dann
'                    einem bisherigen Unterkonto.
'
'                    Ein möglicher Algorithmus zur Berechnung der Prüfziffern ist:
'
'                    Setze die beiden Prüfziffern auf 00
'                    (die IBAN beginnt dann z. B. mit DE00 für Deutschland).
'
'                    Stelle die vier ersten Stellen an das Ende der IBAN.
'
'                    Ersetze alle Buchstaben durch Zahlen, wobei A = 10, B = 11, …, Z = 35.
'
'                    Berechne den ganzzahligen Rest, der bei Division durch 97 bleibt.
'
'                    Subtrahiere den Rest von 98, das Ergebnis sind die beiden Prüfziffern.
'                    Falls das Ergebnis einstellig ist, wird es mit einer führenden Null ergänzt.
' Beispiel:
'
'                    IBAN:      DE00 2105 0170 0012 3456 78
'                    Umstellung:     2105 0170 0012 3456 78DE 00
'                    Modulo:         210501700012345678131400 mod 97 = 30
'                    Subtraktion:    98 - 30 = 68
'
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       10.03.2015-17:48:16
'
' Parameters :       iban (String)
'--------------------------------------------------------------------------------
Public Function CalcCheckDigit(ByVal iban As String) As String

  mBankInformationError = UnknownError
  
  'Setze die beiden Prüfziffern auf 00 (die IBAN beginnt dann z. B. mit DE00 für Deutschland).
  'Stelle die vier ersten Stellen an das Ende der IBAN.
  iban = Mid$(iban, 5) + Mid$(iban, 1, 2) & "00"
  
  Dim checksum As String
  
  'Ersetze alle Buchstaben durch Zahlen, wobei A = 10, B = 11, …, Z = 35.
  Dim i As Integer
  For i = 1 To Len(iban)
    checksum = checksum & IbanCharToNumber(Mid$(iban, i, 1))
  Next i
  
  'Berechne den ganzzahligen Rest, der bei Division durch 97 bleibt.
  CalcCheckDigit = Format(98 - Modulo97Result(checksum), "00")
  
  mBankInformationError = IIf(CLng(CalcCheckDigit) >= 2 And CLng(CalcCheckDigit) <= 98, NoError, InvalidCheckDigit)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CalcIBAN
' Description:       Berechnet eine IBAN anhand von Länderkennzeichen, BLZ und KontoNr.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       11.03.2015-12:12:38
'
' Parameters :       bankCulture (String)
'                    bankCode (String)
'                    accountNumber (String)
'--------------------------------------------------------------------------------
Public Function CalcIBAN(ByVal bankCulture As String, ByVal bankCode As String, ByVal accountNumber As String) As String

  mBankInformationError = UnknownError
  
  bankCulture = UCase$(bankCulture)
  
  If bankCulture = "DE" Then
    CalcIBAN = bankCulture & "00" & Format(bankCode, "00000000") & Format(accountNumber, "0000000000")
    Dim checkDigit As Long: checkDigit = CalcCheckDigit(CalcIBAN)
    CalcIBAN = bankCulture & checkDigit & Format(bankCode, "00000000") & Format(accountNumber, "0000000000")
    mBankInformationError = NoError
  Else
    CalcIBAN = ""
    mBankInformationError = NotSupportedCulture
  End If
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ShowBankInformationError
' Description:       Liefert eine MessageBox mit dem letzen aufgetretenen Fehler.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       11.03.2015-12:13:44
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub ShowBankInformationError()

  Dim msg As String: msg = ""
  Dim title As String: title = ""
  Select Case mBankInformationError
  Case BankInformationErrors.UnknownError
    msg = "Ein unbekannter IBAN-Fehler ist aufgetreten!"
    title = "Unbekannter Fehler"
  Case BankInformationErrors.InvalidCulture
    msg = "Das angegebene Länderkennzeichen der IBAN ist ungültig!"
    title = "Unbekanntes Länderkennzeichen"
  Case BankInformationErrors.InvalidLenght
    msg = "IBAN-Länge für das angegebene Länderkennzeichen ist ungültig!"
    title = "Falsche IBAN-Länge"
  Case BankInformationErrors.InvalidCheckDigit
    msg = "Es trat ein Fehler in der Prüfziffer auf!"
    title = "Prüfziffernfehler"
  Case BankInformationErrors.NotSupportedCulture
    msg = "Die IBAN-Berechnung unterstützt nur das Länderkürzel DE für Deutschland!"
    title = "Nicht unterstütztes Länderkennzeichen"
  Case BankInformationErrors.InvalidBicLenght
    msg = "Länge für die angegebene BIC ist ungültig!"
    title = "Falsche BIC-Länge"
  Case Else
    Exit Sub
  End Select
  
  MsgBox msg, vbCritical, title
End Sub
