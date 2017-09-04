VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolkitError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ToolkitError
'    Project    : ToolKits
'
'    Description: Stellt Standardfehler zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mErrorMessages As Collection
Private mEtc           As etc

Public Enum ErrorNumbers
  None = 0
  ParameterInputCanceledError = -666001
  ParameterInvalidInputFormat = -666002
  ParameterHasUnresolveableFormat = -666003
  SqlGroupFileStatementNameAlreadyExists = -666004
  SqlGroupFileStatementNotExists = -666005
  NotImplemented = -666006
  IsObsolete = -666007
End Enum



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mEtc = New etc
  Set mErrorMessages = New Collection
  
  mErrorMessages.Add "Parameterabfrage wurde abgebrochen!", "key" & ErrorNumbers.ParameterInputCanceledError
  
  mErrorMessages.Add "Format/Datentyp bei Parameterabfrage ungültig!", "key" & ErrorNumbers.ParameterInvalidInputFormat

  mErrorMessages.Add "Parameter kann aufgrund seines Formats nicht aufgelöst werden!", "key" & ErrorNumbers.ParameterHasUnresolveableFormat

  mErrorMessages.Add "Statementname bereits vorhanden!", "key" & ErrorNumbers.SqlGroupFileStatementNameAlreadyExists

  mErrorMessages.Add "Statementname existiert nicht!", "key" & ErrorNumbers.SqlGroupFileStatementNotExists

  mErrorMessages.Add "Routine wurde nicht implementiert!", "key" & ErrorNumbers.NotImplemented

  mErrorMessages.Add "Routine ist obsolet!", "key" & ErrorNumbers.IsObsolete
End Sub

Private Sub Class_Terminate()
  Set mEtc = Nothing
  Set mErrorMessages = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetErrorMessage
' Description:       Liefert den Error-Text zu einer Error-Number.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:20:11
'
' Parameters :       number (ErrorNumbers)
'--------------------------------------------------------------------------------
Public Function GetErrorMessage(ByVal number As ErrorNumbers) As String
  
  Dim key As String: key = "key" & number
  GetErrorMessage = IIf(mEtc.CollectionContainsKey(mErrorMessages, key), mErrorMessages.item(key), "")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ShowToolkitError
' Description:       Zeigt den aufgetretenen Fehler in einer MsgBox an.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:21:05
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub ShowError _
(Optional ByVal source As String = "" _
, Optional ByVal description As String = "")

  Screen.MousePointer = 0
  
  Dim number As Long: number = Err.number
  Dim msg As String: msg = Err.description
  Dim title  As String: title = "Error [" _
  & IIf(source = "", IIf(Err.source = "", "", Err.source), source) & "]"
  
  Dim key     As String: key = "key" & number
  Dim message As String: message = "[" & number & "]" & vbCrLf
  If mEtc.CollectionContainsKey(mErrorMessages, key) Then
    message = message & GetErrorMessage(number)
  Else
    message = message & msg
  End If
  
  message = message & IIf(description = "", "", vbCrLf & vbCrLf & description)
  
  MsgBox message, VbMsgBoxStyle.vbCritical, title
  Err.number = 0
End Sub


