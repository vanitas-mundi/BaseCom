VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : FileSystem
'    Project    : ToolKits
'
'    Description: Stellt FileSystem-Routinen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mIni As FileIni
Private mIO As FileIO
Private mSqlGroupFile As FileSqlGroupFile
Private mXml As FileXml



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mIni = New FileIni
  Set mIO = New FileIO
  Set mSqlGroupFile = New FileSqlGroupFile
  Set mXml = New FileXml
End Sub

Private Sub Class_Terminate()
  Set mIni = Nothing
  Set mIO = Nothing
  Set mSqlGroupFile = Nothing
  Set mXml = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ini() As FileIni
  Set ini = mIni
End Property

Public Property Get io() As FileIO
  Set io = mIO
End Property

Public Property Get SqlGroupFile() As FileSqlGroupFile
  Set SqlGroupFile = mSqlGroupFile
End Property

Public Property Get Xml() As FileXml
  Set Xml = mXml
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------


