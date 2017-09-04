VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommunicationSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------------------------
'    Component  : CommunicationSMS
'    Project    : ToolKits
'
'    Description: Bibliothek zu versenden von SMS-Nachrichten
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Public Type SmsInfoType
  Data As String
  Debugging As DebugType
  Enc As EncType
  MaxParts As Integer
  Opr As String
  Orig As String
  Project As String
  RegisteredDelivery As RegisteredDeliveryType
  MessageType As SMSType
End Type

Public Enum RegisteredDeliveryType
  WithOut = 0
  [With] = 1
End Enum

Public Enum EncType
  iso = 0
  usc2 = 1
  gsm = 2
  UTF8 = 3
End Enum

Public Enum DebugType
  Send = 0
  DebugMode = 1
End Enum

Public Enum SMSType
  SMS = 1
End Enum

'{0} user -{1} password
Private Const SEND_SMS_URL As String = "https://{0}:{1}@sms.openit.de/put.php?"
Private Const USER_NAME    As String = "fh-essen"
Private Const USER_PWD     As String = "pt33ls19"

Private Declare Sub InternetCloseHandle Lib "wininet.dll" ( _
    ByVal hInet As Long)
Private Declare Function InternetOpenA Lib "wininet.dll" ( _
    ByVal sAgent As String, ByVal lAccessType As Long, _
    ByVal sProxyName As String, ByVal sProxyBypass As String, _
    ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrlA Lib "wininet.dll" ( _
    ByVal hOpen As Long, ByVal sUrl As String, _
    ByVal sHeaders As String, ByVal lLength As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Sub InternetReadFile Lib "wininet.dll" ( _
    ByVal hFile As Long, ByVal sBuffer As String, _
    ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long)

'Enumeration für Internet:
Public Enum InternetOpenType
  IOTPreconfig = 0
  IOTDirect = 1
  IOTProxy = 3
End Enum



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Function BuildBaseURL() As String
  Dim etc As etc: Set etc = New etc
  BuildBaseURL = etc.StringFormat(SEND_SMS_URL, USER_NAME, USER_PWD)
  Set etc = Nothing
End Function

Private Function BuildURL(ByVal dest As String, ByRef smsInfo As SmsInfoType) As String
  
  With smsInfo
    Dim col As Collection: Set col = New Collection
    col.Add BuildBaseURL
    col.Add "orig=" & .Orig
    col.Add "&dest=" & dest
    col.Add "&type=" & .MessageType
    col.Add "&project=" & UrlEscapeString4SMS(.Project)
    col.Add "&data=" & UrlEscapeString4SMS(.Data)
    col.Add "&debug=" & .Debugging
    col.Add "&enc=" & GetEncString(.Enc) ''"iso", gsm "utf-8"  und "ucs2"
    col.Add "&maxparts=" & .MaxParts
    col.Add "&opr=" & .Opr
    col.Add "&registered_delivery=" & .RegisteredDelivery
    Dim conv As Convert: Set conv = New Convert
    BuildURL = conv.JoinCollection(col, "")
    Set conv = Nothing
  End With
End Function

Private Function GetEncString(ByVal value As EncType) As String
  Select Case value
  Case EncType.iso
    GetEncString = "iso"
  Case EncType.gsm
    GetEncString = "gsm"
  Case EncType.usc2
    GetEncString = "usc2"
  Case EncType.UTF8
    GetEncString = "utf-8"
  End Select
End Function


Private Function ApiOpenUrl(ByVal url As String, Optional ByVal OpenType As InternetOpenType = IOTPreconfig) As String
  
  Const INET_RELOAD = &H80000000
  Dim hInet As Long
  Dim hURL As Long
  Dim Buffer As String * 2048
  Dim bytes As Long
  
  'Inet-Connection öffnen:
  hInet = InternetOpenA("VB-Tec:INET", OpenType, vbNullString, vbNullString, 0)
  hURL = InternetOpenUrlA(hInet, url, vbNullString, 0, INET_RELOAD, 0)
  
  'Daten sammeln:
  Do
    InternetReadFile hURL, Buffer, Len(Buffer), bytes
    If bytes = 0 Then Exit Do
    ApiOpenUrl = ApiOpenUrl & Left$(Buffer, bytes)
  Loop
  
  'Inet-Connection schließen:
  InternetCloseHandle hURL
  InternetCloseHandle hInet
End Function



'--------------------------------------------------------------------------------
' Project    :      ToolKits
' Procedure  :      UrlEscape
' Description:      Entfernt kritische Zeichen asu TExt, der als Url-Parameter übergeben werden soll,
'                   und ersetzt sie mit einem '_'.
'                   Z.B.: & ? usw.
' Date-Time  :       11.03.2016-17:16:21
' Parameters :       value (String)
'--------------------------------------------------------------------------------
Public Function UrlEscapeString4SMS(value As String) As String
        '<EhHeader>
        On Error GoTo UrlEscapeWithUnderscore_Err
        '</EhHeader>

'        value = Replace$(value, "!", "%21")
'        value = Replace$(value, """", "%22")
'        value = Replace$(value, "#", " ")   '%23")
'        value = Replace$(value, "$", "%24")
'        value = Replace$(value, "%", "%25")
        value = Replace$(value, "&", " ")
'        value = Replace$(value, "'", "%27")
'        value = Replace$(value, "/", " ") '%2C")
        value = Replace$(value, "+", " ")
'        value = Replace$(value, ":", "%3A")
'        value = Replace$(value, "=", "%3D")
'        value = Replace$(value, "?", "_")   '%3F")
'        value = Replace$(value, "\", "%5C")
        
        UrlEscapeString4SMS = value
        '<EhFooter>
        Exit Function

UrlEscapeWithUnderscore_Err:
        Err.Raise vbObjectError + 100, _
                  "ToolKits.CommunicationSMS.UrlEscapeWithUnderscore", _
                  "CommunicationSMS component failure"
        '</EhFooter>
End Function


'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CreateSmsInfo
' Description:       Liefert eine Struktur mit Info-Felder für den SMS-Versand.
'                    Das Feld Data sollte mit der Nachricht gefüllt werden, die
'                    restlichen Felder sind optional.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-15:46:04
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function CreateSmsInfo() As SmsInfoType
  Dim smsInfo As SmsInfoType

  With smsInfo
    .Data = ""
    .Project = ""
    .Orig = "0049201810040"
    .Debugging = Send
    .Enc = UTF8
    .MaxParts = 1
    .Opr = ""
    .RegisteredDelivery = 0
    .MessageType = SMS
  End With
  
  CreateSmsInfo = smsInfo
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SendSms
' Description:       Sendet eine einzelne SMS an die Nummer (dest).
'                    Format: 0049201810040
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-15:48:31
'
' Parameters :       dest (String)
'                    smsInfo (SmsInfoType)
'--------------------------------------------------------------------------------
Public Sub SendSms(ByVal dest As String, ByRef smsInfo As SmsInfoType)
  
  Dim url  As String: url = BuildURL(dest, smsInfo)
  ApiOpenUrl url
End Sub


'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SendSmsByDestCollection
' Description:       Sendet eine SMS an alle übergebenen Nummern (dests).
'                    Format: 0049201810040
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-15:49:32
'
' Parameters :       dests (Collection)
'                    smsInfo (SmsInfoType)
'--------------------------------------------------------------------------------
Public Sub SendSmsByDestCollection(ByVal dests As Collection, _
                                   ByRef smsInfo As SmsInfoType)

  Dim X As Variant
  For Each X In dests
    Dim dest As String: dest = CStr(X)
    Me.SendSms dest, smsInfo
  Next X

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SendSmsByDestArray
' Description:       Sendet eine SMS an alle übergebenen Nummern (dests).
'                    Format: 0049201810040
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-15:49:32
'
' Parameters :       dests (Collection)
'                    smsInfo (SmsInfoType)
'--------------------------------------------------------------------------------
Public Sub SendSmsByDestArray(ByVal dests As Variant, ByRef smsInfo As SmsInfoType)

  Dim i As Integer

  For i = LBound(dests) To UBound(dests)
    Dim dest As String: dest = dests(i)
    Me.SendSms dest, smsInfo
  Next i
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SendSmsWithResult
' Description:       Sendet eine einzelne SMS an die Nummer (dest) und
'                    liefert die Serverantwort.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       02.03.2015-17:09:52
'
' Parameters :       dest (String)
'                    smsInfo (SmsInfoType)
'--------------------------------------------------------------------------------
Public Function SendSmsWithResult(ByVal dest As String, ByRef smsInfo As SmsInfoType) As String
  
  Dim url  As String: url = BuildURL(dest, smsInfo)
  SendSmsWithResult = ApiOpenUrl(url)
End Function

