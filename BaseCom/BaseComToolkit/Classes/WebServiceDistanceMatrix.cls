VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebServiceDistanceMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : WebServiceDistanceMatrix
'    Project    : ToolKits
'
'    Description: Stellt Zugrif auf Google-Maps-Api zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Const WEB_SERVICE_URL = "http://appserver01.bcw-intern.local/Services/DistanceMatrixService/BCW.WebServices.DistanceMatrixServiceLibrary.WebService.DistanceMatrixService.svc?singleWsdl"



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
Public Function GetDistanceMatrix(ByVal originStreet As String, ByVal originHouseNumber As String _
    , ByVal originPostCode As String, ByVal originCity As String, ByVal originCountry As String _
    , ByVal destinationStreet As String, ByVal destinationHouseNumber As String _
    , ByVal destinationPostCode As String, ByVal destinationCity As String _
    , ByVal destinationCountry As String, ByVal languageCultureCode As String) As String

                    
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetDistanceMatrix = soap.GetDistanceMatrixVb6(originStreet, originHouseNumber, originPostCode _
  , originCity, originCountry, destinationStreet, destinationHouseNumber _
  , destinationPostCode, destinationCity, destinationCountry, languageCultureCode)
  Set soap = Nothing
End Function



