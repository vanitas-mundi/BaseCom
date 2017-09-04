VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ELO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ELO
'    Project    : ToolKits
'
'    Description: Stellt Routinen für den ELO-Zugriff zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Const WEB_SERVICE_URL = "http://appserver01.bcw-intern.local/Services/EloService/EloServiceLibrary.EloService.svc?singleWsdl"
Private Const strEloConsolePath As String = "M:\verwaltungsprogramme\tools\EloConsole\EloConsole.exe"
Private Const strEloTempFolder  As String = "C:\temp\elo\"

'---------------------- Konstruktor der Klasse ----------------------------------
'---------------------- Zugriffsmethoden der Klasse -----------------------------
'---------------------- Ereignismethoden der Klasse -----------------------------
'---------------------- Private Methoden der Klasse -----------------------------
'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CreateTaggingValuesTemplateByStatement
' Description:       Erzeugt die Verschlagwortung anhand einer ID
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       10.25.2016-08:45:05
'
' Parameters :       strEnumName (String)
'                    strId (String)
'                    strAcl (String)
'                    strUserName (String)
'--------------------------------------------------------------------------------
Public Function CreateTaggingValuesTemplateByStatement(ByVal strEnumName As String, _
                                                       ByVal strId As String, _
                                                       ByVal strAcl As String, _
                                                       ByVal strUserName As String) As String

    '<EhHeader>
    On Error GoTo CreateTaggingValuesTemplateByStatement_Err

    '</EhHeader>

    Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
    soap.MSSoapInit WEB_SERVICE_URL
    CreateTaggingValuesTemplateByStatement = soap.CreateTaggingValuesTemplateByStatement(strEnumName, strId, strAcl, strUserName)
    Set soap = Nothing

    '<EhFooter>
    Exit Function

CreateTaggingValuesTemplateByStatement_Err:
    Err.Raise vbObjectError, "ToolKits.ELO.CreateTaggingValuesTemplateByStatement", "ELO component failure"

    '</EhFooter>

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CreateEloXmlCommandXml
' Description:       Ermittlung des Command-abhaengigen ELO-Xmls.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:39:00
'
' Parameters :       strEloCommand (String)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Public Function CreateEloXmlCommandXml(ByVal strEloCommand As String, _
                                       ParamArray parameters() As Variant) As Object
    
  Dim io                            As fileIO: Set io = New fileIO
  Dim xmlDom                        As Object 'MSXML2.DOMDocument
  Dim strEloCommandTemplateFileName As String
        
  Set xmlDom = CreateObject("MSXML2.DOMDocument") 'New MSXML2.DOMDocument
    
  strEloCommandTemplateFileName = strEloCommand & "Template.Xml"
    
  'Ueberpruefen, ob der Dateipfad schon existiert
  If Not io.DirectoryExists(strEloTempFolder) Then
    io.CreateDirectory strEloTempFolder
  End If
    
  'Template fuer das jeweilige Commando von der Elo-Console erstellen lassen und einlesen
  Shell strEloConsolePath & " /t:" & strEloCommand & " /x:" & strEloTempFolder & strEloCommandTemplateFileName
    
  xmlDom.Load strEloTempFolder & strEloCommandTemplateFileName
  Set xmlDom = ReplaceEloParameterValuesInDom(xmlDom, parameters) 'strPara1, strPara2, strPara3, strPara4, strPara5, strPara6, strPara7, strPara8, strPara9)
    
  Set io = Nothing
    
  Set CreateEloXmlCommandXml = xmlDom
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ReplaceEloParameterValuesInDom
' Description:       Fuegt die gewuenschten Parameter in das DOM ein
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:39:48
'
' Parameters :       xmlDom (Object)
'                    parameters() (Variant)
'--------------------------------------------------------------------------------
Private Function ReplaceEloParameterValuesInDom(ByVal xmlDom As Object, _
                                                ParamArray parameters() As Variant) As Object
                                             
  Dim etc   As etc: Set etc = New etc
  Dim count As Integer: count = etc.ArrayCount(parameters)
    
  If count = 1 Then
    count = etc.ArrayCount(parameters(0))

    If count > 0 Then
      
      Dim params As Variant: params = parameters(0)
      
      Dim i      As Integer

      For i = LBound(params) To UBound(params)
        xmlDom.SelectSingleNode("//param" & (i + 1)).Text = params(i)
      Next i

      Set ReplaceEloParameterValuesInDom = xmlDom
    End If
  End If

  Set etc = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SendXmlCommandToEloConsole
' Description:       Setzt das Kommando zur Elo-Konsole ab.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-15:40:35
'
' Parameters :       strEloInFilePath (String)
'--------------------------------------------------------------------------------
Public Function SendXmlCommandToEloConsole(ByVal strEloInFilePath As String)
    
  Dim strEloOutFilePath As String
  strEloOutFilePath = strEloTempFolder & "eloOut.xml"
  Shell strEloConsolePath & " " & strEloInFilePath & " " & strEloOutFilePath
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       UploadFileToElo
' Description:       Verschiebt eine Datei mit Verschlagwortungsinformationen zum EloWebService,
'                    zum anschließenden Elo-Upload.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       06.10.2015-13:35:36
'
' Parameters :       fileName (String)
'                    tagginEnumName (String)
'                    taggingValues (String)
'--------------------------------------------------------------------------------
Public Function UploadFileToElo(ByVal PersonId As Long _
                    , ByVal filename As String _
                    , ByVal tagginEnumName As String _
                    , ByVal taggingValues As String) As String
                    
  On Error GoTo errLabel
  
  Dim io As fileIO: Set io = New fileIO
  Dim fileContent As Variant: fileContent = io.GetByteArrayFromFile(filename)
                    
  If Not Me.CheckTaggingValues(tagginEnumName, taggingValues) Then
    UploadFileToElo = "Error in taggingValues!"
    Exit Function
  End If
                    
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  soap.UploadFileToElo PersonId, fileContent, filename, tagginEnumName, taggingValues
  
  Set soap = Nothing
  Set io = Nothing
  
  UploadFileToElo = "OK"
  Exit Function
  
errLabel:
  UploadFileToElo = Err.description
  Exit Function
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetTaggingEnumNames
' Description:       Liefert die Namen aller verfügbaren TaggingEnums.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       06.10.2015-13:35:36
'--------------------------------------------------------------------------------
Public Function GetTaggingEnumNames() As String()
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetTaggingEnumNames = soap.GetTaggingEnumNames
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetPersonFolderIdByPersonId
' Description:       Ermittelt die FolderID anhand einer PersonenID
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       11.10.2015-14:43:43
'
' Parameters :       lPersonId (Long)
'--------------------------------------------------------------------------------
Public Function GetPersonFolderIdByPersonId(ByVal lPersonId As Long) As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetPersonFolderIdByPersonId = soap.GetPersonFolderIdByPersonId(lPersonId)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetParticipantDocumentEloFileName
' Description:       Legt den Namen einer Teilnehmerdatei (z.B. Imma, Exma, Rechnung) fest
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       11.10.2015-14:48:56
'
' Parameters :       lPersonID (Long)
'                    lMatrikelNr (Long)
'                    strDocType (String)
'                    strSemester (String)
'--------------------------------------------------------------------------------
Public Function GetParticipantDocumentEloFileName(ByVal lPersonId As Long, ByVal lMatriculationNumber As Long, ByVal strDocType As String, ByVal strAdditionalText As String) As String
    GetParticipantDocumentEloFileName = lPersonId & "_" & lMatriculationNumber & "_" & strDocType & "_" & strAdditionalText & ".pdf"
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetTaggingEnumValueNames
' Description:       Liefert aller verfügbaren Verschlagwortungsfelder einer taggingEnum.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       06.10.2015-13:35:36
'
' Parameters :       enumName (String)
'--------------------------------------------------------------------------------
Public Function GetTaggingEnumValueNames(ByVal enumName As String) As String()
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetTaggingEnumValueNames = soap.GetTaggingEnumValueNames(enumName)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetTaggingEnumNames
' Description:       Liefert true, wenn der Aufbauf von taggingValues dem Format der angegebenen taggingEnum entspricht.
'                    Verschlagwortungsfeld und Wert werden durch ":" getrennt, Key-Wert-Paare duch ";".
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       06.10.2015-13:35:36
'--------------------------------------------------------------------------------
Public Function CheckTaggingValues(ByVal enumName As String, ByVal taggingValues As String) As Boolean
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  CheckTaggingValues = soap.CheckTaggingValues(enumName, taggingValues)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CreateTaggingValuesTemplate
' Description:       Liefert ein Template für einen TaggingEnumString für das angegebene taggingEnum-Format.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       06.10.2015-13:35:36
'
' Parameters :       enumName (String)
'--------------------------------------------------------------------------------
Public Function CreateTaggingValuesTemplate(ByVal enumName As String) As String

  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  CreateTaggingValuesTemplate = soap.CreateTaggingValuesTemplate(enumName)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetAclStringHelp
' Description:       Liefert die Notation für die Übergabe der Tagging-Info Acl.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       25.01.2016-13:35:36
'--------------------------------------------------------------------------------
Public Function GetAclStringHelp() As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetAclStringHelp = soap.GetAclStringHelp
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExistDocument
' Description:       Prueft, ob ein Dokument bereits in Elo existiert und liefert dessen Id oder 0.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       25.01.2016-13:35:36
'
' Parameters :       maskName (String) Name der Verschlagwortungsmaske
'                    documentType (String) Wert des Indexfeldes Dokumententyp
'                    searchId (String) Wert des Indexfeldes Such-Id.
'--------------------------------------------------------------------------------
Public Function ExistDocument(ByVal maskName As String, ByVal documentType As String, ByVal searchId As String) As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  ExistDocument = soap.ExistDocument(maskName, documentType, searchId)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExistsDocumentByIndexField
' Description:       Prueft, ob ein Dokument bereits in Elo existiert und liefert dessen Id oder 0.
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       10.28.2016-08:53:43
'
' Parameters :       maskName (String)
'                    strIndexFieldName (String)
'                    strIndexFieldValue (String)
'--------------------------------------------------------------------------------
Public Function ExistsDocumentByIndexField(ByVal maskName As String, ByVal strIndexFieldName As String, ByVal strIndexFieldValue As String) As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  ExistsDocumentByIndexField = soap.ExistsDocumentByIndexField(maskName, strIndexFieldName, strIndexFieldValue)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ExistsDocumentByIndexFields
' Description:       Prueft, ob ein Dokument bereits in Elo existiert und liefert dessen Id oder 0.
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       10.28.2016-09:00:10
'
' Parameters :       maskName (String)
'                    indexFieldNames (Variant)
'                    indexFieldValues (Variant)
'--------------------------------------------------------------------------------
Public Function ExistsDocumentByIndexFields(ByVal maskName As String, ByVal indexFieldNames As Variant, ByVal indexFieldValues As Variant) As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  ExistsDocumentByIndexFields = soap.ExistsDocumentByIndexFields(maskName, indexFieldNames, indexFieldValues)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FindSordIdsByIndexField
' Description:       Prueft, ob ein Dokument bereits in Elo existiert und liefert dessen Id oder 0.
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       10.28.2016-09:00:22
'
' Parameters :       maskName (String)
'                    indexFieldName (String)
'                    indexFieldValue (String)
'--------------------------------------------------------------------------------
Public Function FindSordIdsByIndexField(ByVal maskName As String, ByVal indexFieldName As String, ByVal indexFieldValue As String) As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  FindSordIdsByIndexField = soap.FindSordIdsByIndexField(maskName, indexFieldName, indexFieldValue)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       FindSordIdsByIndexFields
' Description:       Prueft, ob ein Dokument bereits in Elo existiert und liefert dessen Id oder 0.
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       10.28.2016-09:00:27
'
' Parameters :       maskName (String)
'                    indexFieldNames (Variant)
'                    indexFieldValues (Variant)
'--------------------------------------------------------------------------------
Public Function FindSordIdsByIndexFields(ByVal maskName As String, ByVal indexFieldNames As Variant, ByVal indexFieldValues As Variant) As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  FindSordIdsByIndexFields = soap.FindSordIdsByIndexFields(maskName, indexFieldNames, indexFieldValues)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetDocumentVersionUrlById
' Description:       Liefert die Elo-Url für den Dokumenten-Download.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       25.01.2016-13:35:36
'
' Parameters :       id (String) Id des Sord (Elo-Dokuments).
'--------------------------------------------------------------------------------
Public Function GetDocumentVersionUrlById(id As String) As String
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetDocumentVersionUrlById = soap.GetDocumentVersionUrlById(id)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetDocumentVersionByteArrayById
' Description:       Liefert das hinterlegte Dokument als Byte-Array.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       25.01.2016-13:35:36
'
' Parameters :       id (String) Id des Sord (Elo-Dokuments).
'--------------------------------------------------------------------------------
Public Function GetDocumentVersionByteArrayById(id As String) As Variant
  
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetDocumentVersionByteArrayById = soap.GetDocumentVersionByteArrayById(id)
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       WriteDocumentVersionByteArrayById
' Description:       Speichert ein ELO-Dokument als eine lokale Datei im Temp-Verzeichnis
' Created by :       Sebastian Limke
' Machine    :       VDI-IT-02-0007
' Date-Time  :       1.29.2016-11:39:42
'
' Parameters :       id (String)
'--------------------------------------------------------------------------------
Public Function WriteDocumentVersionByteArrayById(id As String) As String

    '<EhHeader>
    On Error GoTo WriteDocumentVersionByteArrayById_Err

    '</EhHeader>

    Dim bSuccessfull As Boolean
    Dim strFilePath  As String
    Dim fileIO       As fileIO
    Dim byteArray()  As Byte
    byteArray = GetDocumentVersionByteArrayById(id)
    strFilePath = fileIO.GetTempFileName("pdf")
    bSuccessfull = fileIO.WriteAllBytes(byteArray, strFilePath)
    Set fileIO = Nothing
    WriteDocumentVersionByteArrayById = strFilePath

    '<EhFooter>
    Exit Function

WriteDocumentVersionByteArrayById_Err:
    WriteDocumentVersionByteArrayById = ""
    Err.Raise vbObjectError, "ToolKits.ELO.WriteDocumentVersionByteArrayById", "ELO component failure"

    '</EhFooter>

End Function
