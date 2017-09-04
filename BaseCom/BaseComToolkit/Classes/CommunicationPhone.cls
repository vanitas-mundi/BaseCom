VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommunicationPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : CommunicationPhone
'    Project    : ToolKits
'
'    Description: Stellt Methoden für die CISCO-Treiber-Kommunikation zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CallNumber
' Description:       Wählt die angegebene Nummer über den CISCO-Treiber.
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-15:54:49
'
' Parameters :       phoneNumber (String)
'                    clipDialPath (String)
'--------------------------------------------------------------------------------
Public Sub CallNumber(ByVal phoneNumber As String, ByVal clipDialPath As String)
    
  Dim regex As Object: Set regex = CreateObject("VBScript.RegExp") 'RegExp
  regex.Pattern = "\D"
  regex.Global = True
  phoneNumber = regex.Replace(phoneNumber, "")
    
  Dim io       As FileIO: Set io = New FileIO
  Dim filename As String: filename = io.CombinePath(clipDialPath, "ClipDial.exe")
  shell filename & " " & phoneNumber
    
  Set regex = Nothing
  Set io = Nothing
End Sub

