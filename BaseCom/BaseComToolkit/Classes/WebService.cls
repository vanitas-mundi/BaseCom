VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : WebService
'    Project    : ToolKits
'
'    Description: Stellt Zugriff auf WebServices zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mAuthentication As WebServiceAuthentication
Private mMailing        As WebServiceMailing
Private mDistanceMatrix As WebServiceDistanceMatrix


'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mAuthentication = New WebServiceAuthentication
  Set mMailing = New WebServiceMailing
  Set mDistanceMatrix = New WebServiceDistanceMatrix
End Sub

Private Sub Class_Terminate()
  Set mAuthentication = Nothing
  Set mMailing = Nothing
  Set mDistanceMatrix = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get Authentication() As WebServiceAuthentication
  Set Authentication = mAuthentication
End Property

Public Property Get Mailing() As WebServiceMailing
  Set Mailing = mMailing
End Property

Public Property Get DistanceMatrix() As WebServiceDistanceMatrix
  Set DistanceMatrix = mDistanceMatrix
End Property


'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------

