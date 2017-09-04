VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebServiceMailing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : WebServiceMailing
'    Project    : ToolKits
'
'    Description: Stellt Zugrif auf den Mailing-WebService zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Const WEB_SERVICE_URL = "http://appserver01.bcw-intern.local/Services/MailService/MailServiceLibrary.MailService.svc?singleWsdl"



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SendMail
' Description:       Versendet eine E-Mail über den Mailing-WebService.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:35:36
'
' Parameters :       fromAddress (String)
'                    toAddresses (String)
'                    ccAddresses (Variant)
'                    bccAddresses (Variant)
'                    Subject (String)
'                    body (String)
'--------------------------------------------------------------------------------
Public Sub SendMail(ByVal fromAddress As String, _
                    ByVal toAddresses As String, _
                    ByVal ccAddresses, _
                    ByVal bccAddresses, _
                    ByVal Subject As String, _
                    ByVal body As String)
                    
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  soap.SendMail fromAddress, toAddresses, ccAddresses, bccAddresses, Subject, body, ""
  Set soap = Nothing
End Sub

