VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommunicationAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : CommunicationAddress
'    Project    : ToolKits
'
'    Description: Stellt Address-Methoden zur Verfuegung
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Private mStrException As String

'---------------------- Konstruktor der Klasse ----------------------------------
'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get Exception() As String
    Exception = mStrException
End Property

'---------------------- Ereignismethoden der Klasse -----------------------------

'--------------------------------------------------------------------------------
' Project    :       Tnv
' Procedure  :       IsValidFormattedAddress
' Description:       Ueberprueft, ob die formatierte Anschrift gueltig ist
' Created by :       Project Administrator
' Machine    :       VDI-IT-02-0007
' Date-Time  :       9.30.2015-15:54:26
'
' Parameters :       strStrasse (String)
'                    iHausNr (Integer)
'--------------------------------------------------------------------------------
Public Function IsValidFormattedAddress(ByVal strFormatierteAnschrift As String) As Boolean

    '<EhHeader>
    On Error GoTo IsValidFormattedAddress_Err

    '</EhHeader>

    Dim mc     As Object
    Dim regexp As Object
    Set regexp = CreateObject("VBScript.RegExp")
    regexp.Pattern = "\n"
    regexp.Global = True
    Set mc = regexp.Execute(strFormatierteAnschrift)

    If mc.count > 5 Then
        IsValidFormattedAddress = False
        mStrException = "Die formatierte Anschrift ist für den Ausdruck zu lang, bitte einkürzen (weniger Zeilen)!"
    Else
        IsValidFormattedAddress = True
    End If

    Set regexp = Nothing
    Set mc = Nothing

    '<EhFooter>
    Exit Function

IsValidFormattedAddress_Err:
    Err.Raise vbObjectError, "ToolKits.CommunicationAddress.IsValidFormattedAddress", "CommunicationAddress component failure"

    '</EhFooter>

End Function

'--------------------------------------------------------------------------------
' Project    :       Tnv
' Procedure  :       IsValidStreet
' Description:       Ueberprueft, ob die Strasse ohne Hausnummer ist
' Created by :       Project Administrator
' Machine    :       VDI-IT-02-0007
' Date-Time  :       9.30.2015-15:54:26
'
' Parameters :       strStrasse (String)
'                    iHausNr (Integer)
'--------------------------------------------------------------------------------
Public Function IsValidStreet(ByVal strStrasse As String, _
                                         ByVal iHausNr As String) As Boolean

    '<EhHeader>
    On Error GoTo IsValidStreet_Err

    '</EhHeader>
    If LenB(Trim$(iHausNr)) = 0 Then
        IsValidStreet = True

        Exit Function

    End If

    Dim regexp As Object
    Set regexp = CreateObject("VBScript.RegExp")
    regexp.Pattern = iHausNr

    If regexp.Test(strStrasse) Then
        mStrException = "Die Hausnummer ist Bestandteil der Straße. Bitte entfernen Sie die Hausnummer aus dem Feld 'Straße'!"
        IsValidStreet = False
    Else
        mStrException = ""
        IsValidStreet = True
    End If

    Set regexp = Nothing

    '<EhFooter>
    Exit Function

IsValidStreet_Err:
    Err.Raise vbObjectError, "ToolKits.CommunicationAddress.IsValidStreet", "CommunicationAddress component failure"

    '</EhFooter>

End Function

'---------------------- Private Methoden der Klasse -----------------------------
