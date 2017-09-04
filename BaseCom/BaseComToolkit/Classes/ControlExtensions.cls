VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlExtensions
'    Project    : ToolKits
'
'    Description: Setllt Zugriff auf Control-Erweiteungen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Private mComboBox As ControlComboBox
Private mFlexGrid As ControlFlexGrid
Private mFormular As ControlForm
Private mListBox As ControlListBox
Private mListView As ControlListView
Private mTextBox As ControlTextBox
Private mTreeView As ControlTreeView
Private mUnicodeComboBox As ControlUnicodeComboBox



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mComboBox = New ControlComboBox
  Set mFlexGrid = New ControlFlexGrid
  Set mFormular = New ControlForm
  Set mListBox = New ControlListBox
  Set mListView = New ControlListView
  Set mTextBox = New ControlTextBox
  Set mTreeView = New ControlTreeView
  Set mUnicodeComboBox = New ControlUnicodeComboBox
End Sub

Private Sub Class_Terminate()
  Set mComboBox = Nothing
  Set mFlexGrid = Nothing
  Set mFormular = Nothing
  Set mListBox = Nothing
  Set mListView = Nothing
  Set mTextBox = Nothing
  Set mTreeView = Nothing
  Set mUnicodeComboBox = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ComboBox() As ControlComboBox
  Set ComboBox = mComboBox
End Property

Public Property Get flexGrid() As ControlFlexGrid
  Set flexGrid = mFlexGrid
End Property

Public Property Get Formular() As ControlForm
  Set Formular = mFormular
End Property

Public Property Get ListBox() As ControlListBox
  Set ListBox = mListBox
End Property

Public Property Get ListView() As ControlListView
  Set ListView = mListView
End Property

Public Property Get textBox() As ControlTextBox
  Set textBox = mTextBox
End Property

Public Property Get TreeView() As ControlTreeView
  Set TreeView = mTreeView
End Property

Public Property Get UnicodeComboBox() As ControlUnicodeComboBox
  Set UnicodeComboBox = mUnicodeComboBox
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------

