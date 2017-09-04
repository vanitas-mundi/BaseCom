VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : RegEx
'    Project    : ToolKits
'
'    Description: Stellt Reguläre Ausdrücke zur Verfügung.
'                 http://www.regular-expressions.info/vbscript.html
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mRegEx As Object



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mRegEx = CreateObject("VBScript.RegExp")
End Sub

Private Sub Class_Terminate()
  Set mRegEx = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsPatternMatched
' Description:       Prüft, ob das übergebene Pattern in der Zeichenfolge
'                    expressionString enthalten ist und liefert true oder false.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.07.2016
'
' Parameters :       pattern (String)
'                    expressionString (String)
'                    ignoreCase (Boolean)
'--------------------------------------------------------------------------------
Public Function IsPatternMatched(ByVal pattern As String _
, ByVal expressionString As String _
, Optional ByVal ignoreCase As Boolean = True) As Boolean

  mRegEx.Global = False
  mRegEx.ignoreCase = ignoreCase
  mRegEx.pattern = pattern
  IsPatternMatched = mRegEx.Test(expressionString)
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetMatchesByPattern
' Description:       Liefert das Auftreten des Pattern als Match-Collection zurück.
'                    Das Match-Object besitzt folgende Eingenschaften:
'                    FirstIndex, Length, Value, SubMatches
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.07.2016
'
' Parameters :       pattern (String)
'                    expressionString (String)
'                    ignoreCase (Boolean)
'                    returnFirstOneOnly (Boolean)
'--------------------------------------------------------------------------------
Public Function GetMatchesByPattern(ByVal pattern As String _
, ByVal expressionString As String _
, Optional ByVal ignoreCase As Boolean = True _
, Optional ByVal returnFirstOneOnly As Boolean = False) As Collection

  mRegEx.Global = Not returnFirstOneOnly
  mRegEx.ignoreCase = ignoreCase
  mRegEx.pattern = pattern
  
  Dim matches As Object: Set matches = mRegEx.Execute(expressionString) 'MatchCollection
  Dim match As Object 'Match
  
  Dim result As Collection: Set result = New Collection
  
  For Each match In matches
    result.Add match
  Next match
  
  Set GetMatchesByPattern = result ' conv.CollectionToArray(result)
  
  Set matches = Nothing
  Set match = Nothing
  Set result = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsPatternMatched
' Description:       Ersetzt das übergebene Pattern in der Zeichenfolge
'                    expressionString mit der Zeichenfolge replaceString
'                    und liefert die ersetzte Zeichenfolge zurück.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.07.2016
'
' Parameters :       pattern (String)
'                    expressionString (String)
'                    replaceString (String)
'                    ignoreCase (Boolean)
'                    replaceFirstOneOnly (Boolean)
'--------------------------------------------------------------------------------
Public Function ReplacePattern(ByVal pattern As String _
, ByVal expressionString As String, ByVal replaceString As String _
, Optional ByVal ignoreCase As Boolean = True _
, Optional ByVal replaceFirstOneOnly As Boolean = False) As String

  mRegEx.Global = Not replaceFirstOneOnly
  mRegEx.ignoreCase = ignoreCase
  mRegEx.pattern = pattern
  ReplacePattern = mRegEx.Replace(expressionString, replaceString)
End Function

