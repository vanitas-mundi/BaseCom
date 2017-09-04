VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reflection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Reflection
'    Project    : ToolKits
'
'    Description: Stellt Routinen für den Zugriff auf Klasseninformationen zur Verfügung
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mAllMemberInfos As ReflectionAllMemberInfos
Private mMemberInfos As ReflectionMemberInfos



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mAllMemberInfos = New ReflectionAllMemberInfos
  Set mMemberInfos = New ReflectionMemberInfos
End Sub

Private Sub Class_Terminate()
  Set mAllMemberInfos = Nothing
  Set mMemberInfos = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------


