VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileXml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : FileXml
'    Project    : ToolKits
'
'    Description: Stellt XML-Funktionalität zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum XmlEncoding
  UTF8 = 1
  Ansi = 2
  Ascii = 3
End Enum



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
End Sub

'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Function GetEncodingString(ByVal encoding As XmlEncoding) As String

  Select Case encoding
    Case XmlEncoding.UTF8
      GetEncodingString = "UTF-8"
    Case Else
      Err.Raise ErrorNumbers.NotImplemented, "FileXml.GetEncodingString"
  End Select

End Function



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CreateProcessingInstruction
' Description:       Liefert ein MSXML2.DOMDocument zurück.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       15.01.2015-10:14:31
'
' Parameters :       xmlDoc (MSXML2.DOMDocument)
'                    strXml (String)
'                    strVersion (String)
'                    strEncoding (String = " UTF-8")
'--------------------------------------------------------------------------------
Public Function CreateProcessingInstruction(ByRef xmlDoc As Object, _
                                            strXml As String, _
                                            strVersion As String, _
                                            Optional encoding As XmlEncoding = UTF8) As Object
    
  Dim strEncoding              As String: strEncoding = GetEncodingString(encoding)
    
  Dim xmlProcessingInstruction As Object 'MSXML2.IXMLDOMProcessingInstruction
  Set xmlProcessingInstruction = CreateObject("MSXML2.IXMLDOMProcessingInstruction")
  Set xmlProcessingInstruction = xmlDoc.CreateProcessingInstruction(strXml, strVersion & " " & strEncoding)
  xmlDoc.appendChild xmlProcessingInstruction
    
  Set xmlProcessingInstruction = Nothing
    
  Set CreateProcessingInstruction = xmlDoc
End Function

'--------------------------------------------------------------------------------
' Project    :       Dozentenverwaltung
' Procedure  :       CreateXmlAttribut
' Description:       Erzeugt ein Attribut fuer den uebergebenen Node
' Created by :       Sebastian Limke
' Date-Time  :       10.7.2014-10:40:21'
' Parameters :       xmlDoc (MSXML2.DOMDocument)
'                    node (IXMLDOMElement)
'                    strAttribName (String)
'                    strValue (String)
'--------------------------------------------------------------------------------
Public Function CreateXmlAttribut(ByRef xmlDoc As Object, _
                                  ByRef Node As Object, _
                                  strAttribName As String, _
                                  strValue As String) As Object 'IXMLDOMElement

  Dim attrib As Object 'IXMLDOMAttribute

  Set attrib = xmlDoc.createAttribute(strAttribName)
  attrib.value = strValue
  Node.setAttributeNode attrib
        
  Set attrib = Nothing
    
  Set CreateXmlAttribut = Node
End Function

'--------------------------------------------------------------------------------
' Project    :       Dozentenverwaltung
' Procedure  :       SaveDomAsXmlFile
' Description:       Speichert den Inhalt eines DOMs
' Created by :       Sebastian Limke
' Date-Time  :       11.28.2014-08:43:54'
' Parameters :       xmlDom (MSXML2.DOMDocument)
'                    strFilePath (String)
'--------------------------------------------------------------------------------
Public Sub SaveDomAsXmlFile(ByRef xmlDom As Object, _
                            ByVal strFilePath As String, _
                            ByVal strFileName As String)
    
  Dim io As FileIO: Set io = New FileIO
    
  If Not io.DirectoryExists(strFilePath) Then
    io.CreateDirectory strFilePath
  End If

  xmlDom.Save strFilePath & strFileName
    
  Set io = Nothing
End Sub

