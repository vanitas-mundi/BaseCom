VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommunicationPostCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : CommunicationPostCode
'    Project    : ToolKits
'
'    Description: Stellt Methoden für Abfrage von: PLZ, Stadt, Bundesland und
'                 Vorwahl zur Verfügung. Benötigt einen gesetzten ConnectionString.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum PostCodeFields
  Postcode = 1
  Phonecode = 2
  city = 3
  federalState = 4
End Enum

Private Const POST_CODE_STATEMENT As String = "SELECT {0} AS Value FROM def.t_postcodes where {1} = '{2}' GROUP BY {0}"

Private mConnectionString         As String

'---------------------- Konstruktor der Klasse ----------------------------------

'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ConnectionString() As String
  ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal value As String)
  mConnectionString = value
End Property

'---------------------- Ereignismethoden der Klasse -----------------------------

'---------------------- Private Methoden der Klasse -----------------------------
Private Function GetPostCodeFieldString(ByVal postCodeField As PostCodeFields)

  Select Case postCodeField

    Case PostCodeFields.city
      GetPostCodeFieldString = "City"

    Case PostCodeFields.federalState
      GetPostCodeFieldString = "FederalState"

    Case PostCodeFields.Phonecode
      GetPostCodeFieldString = "Phonecode"

    Case PostCodeFields.Postcode
      GetPostCodeFieldString = "Postcode"

    Case Else
      GetPostCodeFieldString = ""
  End Select

End Function

Private Function GetPostCodeFieldGermanString(ByVal postCodeField As PostCodeFields)

  Select Case postCodeField

    Case PostCodeFields.city
      GetPostCodeFieldGermanString = "Stadt"

    Case PostCodeFields.federalState
      GetPostCodeFieldGermanString = "Bundesland"

    Case PostCodeFields.Phonecode
      GetPostCodeFieldGermanString = "Vorwhal"

    Case PostCodeFields.Postcode
      GetPostCodeFieldGermanString = "Postleitzahl"

    Case Else
      GetPostCodeFieldGermanString = ""
  End Select

End Function

Private Function GetValueBase(ByVal getField As PostCodeFields, _
                              ByVal getBy As PostCodeFields, _
                              ByVal value As String) As String

  Dim etc As etc: Set etc = New etc
  Dim s   As String: s = etc.StringFormat(POST_CODE_STATEMENT, GetPostCodeFieldString(getField), GetPostCodeFieldString(getBy), value)

  Dim db  As Database: Set db = New Database
  db.ConnectionString = Me.ConnectionString
  Dim sb As DialogSelectEntry: Set sb = New DialogSelectEntry
  sb.SetDataSource db.ExecuteReader(s)
  sb.ReturnSingleEntryAutomatically = True
  
  Dim title As String: title = etc.StringFormat("Bitte {0} wählen ...", GetPostCodeFieldGermanString(getField))
  
  sb.SelectEntry , title, False, True
  
  GetValueBase = sb.ValueEntry
  
  Set etc = Nothing
  Set sb = Nothing
  Set db = Nothing
End Function

'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetCity
' Description:       Ermittelt die Stadt anhand von PLZ, Vorwahl oder Bundesland.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:16
'
' Parameters :       getBy (PostCodeFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetCity(ByVal getBy As PostCodeFields, ByVal value As String) As String
  
  GetCity = GetValueBase(city, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetFederalState
' Description:       Ermittelt das Bundesland anhand von PLZ, Vorwahl oder Stadt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:48:55
'
' Parameters :       getBy (PostCodeFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetFederalState(ByVal getBy As PostCodeFields, _
                                ByVal value As String) As String
  
  GetFederalState = GetValueBase(federalState, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetPhoneCode
' Description:       Ermittelt die Vorwahl anhand von PLZ, Bundesland oder Stadt.
' Created by :       Sascha GLinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:49:19
'
' Parameters :       getBy (PostCodeFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetPhoneCode(ByVal getBy As PostCodeFields, _
                             ByVal value As String) As String
  
  GetPhoneCode = GetValueBase(Phonecode, getBy, value)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetPostCode
' Description:       Ermittelt die PLZ anhand von Vorwahl, Bundesland oder Stadt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-17:50:01
'
' Parameters :       getBy (PostCodeFields)
'                    value (String)
'--------------------------------------------------------------------------------
Public Function GetPostCode(ByVal getBy As PostCodeFields, _
                            ByVal value As String) As String

  GetPostCode = GetValueBase(Postcode, getBy, value)
End Function

