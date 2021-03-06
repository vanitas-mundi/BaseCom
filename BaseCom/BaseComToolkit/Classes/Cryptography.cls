VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cryptography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Cryptography
'    Project    : ToolKits
'
'    Description: Stellt Verschl�sselungsroutinen zur Verf�gung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- �ffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetModulo10Code
' Description:       Liefert einen Modulo10-Code.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:03:33
'
' Parameters :       value (Long)
'--------------------------------------------------------------------------------
Public Function GetModulo10Code(ByVal value As Long) As Long
  'Modulo-10-Verfahren
  Dim i              As Long
  Dim lngSumUngerade As Long
  Dim lngSumGerade   As Long
  Dim lngSumGesamt   As Long
  Dim lngGewichtung  As Long
  
  'Beispiel:
  '4 7 2 0 5 6 0 X                     X = unbekannte Pr�fziffer
  '1. Addition der ungeraden Positionen: 4+2+5+0 = 11
  For i = 1 To Len(CStr(value)) Step 2
    lngSumUngerade = lngSumUngerade + Mid$(value, i, 1)
  Next i
  
  '2. Gewichtung mit Faktor 3:  11 * 3 = 33
  lngGewichtung = 3
  
  '3. Addition der geraden Positionen: 7+0+6 = 13
  For i = 2 To Len(CStr(value)) Step 2
    lngSumGerade = lngSumGerade + Mid$(value, i, 1)
  Next i
  
  '4. Addition aus Schritt 2 und 3: 33 + 13 = 46
  lngSumGesamt = (lngSumUngerade * lngGewichtung) + lngSumGerade
  '5. Bilden der Differenz zum n�chsten Vielfachen von 10.
  GetModulo10Code = Right$((10 - (lngSumGesamt Mod 10)), 1)
  'Die Differenz ist die Pr�fziffer.
  'In diesem Fall hat die Pr�fziffer den Wert 4.
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsModulo10CheckSumOK
' Description:       Pr�ft, ob es sich um einen g�ltigen Modulo10-Code handelt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:03:39
'
' Parameters :       value (Long)
'--------------------------------------------------------------------------------
Public Function IsModulo10CheckSumOK(ByVal value As Long) As Boolean
  Dim lngChecksum As Long
  '{Pr�fziffer extrahieren}
  lngChecksum = Mid$(value, Len(CStr(value)))
  '{Pr�fziffer abschneiden}
  value = Mid$(value, 1, Len(CStr(value)) - 1)
  '{Pr�fziffer validieren}
  IsModulo10CheckSumOK = lngChecksum = GetModulo10Code(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       EncryptString
' Description:       Diese Function wendet eine Ringverschl�sselung an,
'                    von jedem Zeichen wird die Ordinalzahl ermittelt
'                    und mit 128 addiert. Danach mu� gepr�ft werden ob
'                    die Zeichensatzgrenze von 256 �berschritten wurde,
'                    wenn die neu ermittelte Ordinalzahl gr��er ist,
'                    mu� 256 davon abgezogen werden. Aus der neuen
'                    Odinalzahl wird ein neues ASCII-Zeichen erzeugt.
'                    Wenn der Algorithmus zweimal angewendet wird,
'                    dann verschl�sselt er bei erster Anwendung und
'                    entschl�sselt bei zweiter Anwendung.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:03:43
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function EncryptString(ByVal value As String) As String
  
  Dim i As Long
  For i = 1 To Len(value)
    Dim strLetter As String: strLetter = Mid(value, i, 1)
    Dim lngOrd As Long: lngOrd = Asc(strLetter) + 128
    
    If lngOrd > 256 Then lngOrd = lngOrd - 256
    strLetter = Chr(lngOrd)
    EncryptString = EncryptString & strLetter
  Next i
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DecryptString
' Description:       Entschl�sselt einen mit EncryptString verschl�sselten String.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:04:00
'
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function DecryptString(ByVal value As String) As String
  DecryptString = EncryptString(value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       EncryptTextFile
' Description:       Verschl�sselt eine komplette Textdatei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:04:04
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Sub EncryptTextFile(ByVal filename As String)
  Dim io As FileIO: Set io = New FileIO
  Dim s As String: s = io.ReadAllText(filename)
  io.WriteAllText filename, EncryptString(s), False
  Set io = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DecryptTextFile
' Description:       Entschl�sselt eine komplette Textdatei.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-13:04:08
'
' Parameters :       filename (String)
'--------------------------------------------------------------------------------
Public Sub DecryptTextFile(ByVal filename As String)
  EncryptTextFile filename
End Sub
