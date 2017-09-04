VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DialogSelectEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : DialogSelectEntry
'    Project    : ToolKits
'
'    Description: Stellt einen Dialog zur Auswahl von Elementen zur Verfügung
'
'    Modified   :
'--------------------------------------------------------------------------------



'---------------------- Eigenschaften der Klasse --------------------------------
Private mSelectEntryDialog As SelectEntryDiaolg



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Initialize
End Sub

Private Sub Class_Terminate()
  Unload mSelectEntryDialog
  Set mSelectEntryDialog = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get PreChecked() As Collection
  Set PreChecked = mSelectEntryDialog.PreChecked
End Property

Public Property Let PreChecked(ByRef colPreChecked As Collection)
  mSelectEntryDialog.PreChecked = colPreChecked
End Property

Public Property Get IsMarkDemarkAllVisible() As Boolean
  IsMarkDemarkAllVisible = mSelectEntryDialog.PreSelect
End Property

Public Property Let IsMarkDemarkAllVisible(ByVal blnIsMarkDemarkAllVisible As Boolean)
  mSelectEntryDialog.IsMarkDemarkAllVisible = blnIsMarkDemarkAllVisible
End Property

Public Property Get PreSelect() As Boolean
  PreSelect = mSelectEntryDialog.PreSelect
End Property

Public Property Let PreSelect(ByVal blnPreSelect As Boolean)
  mSelectEntryDialog.PreSelect = blnPreSelect
End Property

Public Property Get ValueEntry() As String
  ValueEntry = mSelectEntryDialog.ValueEntry
End Property

Public Property Let ValueEntry(ByVal strValueEntry As String)
  mSelectEntryDialog.ValueEntry = strValueEntry
End Property

Public Property Get ValueID() As String
  ValueID = mSelectEntryDialog.ValueID
End Property

Public Property Let ValueID(ByVal strValueID As String)
  mSelectEntryDialog.ValueID = strValueID
End Property

Public Property Get ValueEntries() As Collection
  Set ValueEntries = mSelectEntryDialog.ValueEntries
End Property

Public Property Get ValueIDs() As Collection
  Set ValueIDs = mSelectEntryDialog.ValueIDs
End Property

Public Property Get SortInFolder() As Boolean
  SortInFolder = mSelectEntryDialog.SortInFolder
End Property

Public Property Let SortInFolder(ByVal blnSortInFolder As Boolean)
  mSelectEntryDialog.SortInFolder = blnSortInFolder
End Property

Public Property Get UseFirstLetter() As Boolean
  UseFirstLetter = mSelectEntryDialog.UseFirstLetter
End Property

Public Property Let UseFirstLetter(ByVal blnUseFirstLetter As Boolean)
  mSelectEntryDialog.UseFirstLetter = blnUseFirstLetter
End Property

Public Property Get DataString() As String
  DataString = mSelectEntryDialog.DataString
End Property

Public Property Let DataString(ByVal strDataString As String)
  mSelectEntryDialog.DataString = strDataString
End Property

Public Property Get DataCollection() As Collection
  Set DataCollection = mSelectEntryDialog.DataCollection
End Property

Public Property Let DataCollection(ByVal colDataCollection As Collection)
  mSelectEntryDialog.DataCollection = colDataCollection
End Property

Public Property Get ReturnID() As Boolean
  ReturnID = mSelectEntryDialog.ReturnID
End Property

Public Property Let ReturnID(ByVal blnReturnID As Boolean)
  mSelectEntryDialog.ReturnID = blnReturnID
End Property

Public Property Get Sorted() As Boolean
  Sorted = mSelectEntryDialog.Sorted
End Property

Public Property Let Sorted(ByVal blnSorted As Boolean)
  mSelectEntryDialog.Sorted = blnSorted
End Property

Public Property Get Header() As String
  Header = mSelectEntryDialog.Header
End Property

Public Property Let Header(ByVal strHeader As String)
  mSelectEntryDialog.Header = strHeader
End Property

Public Property Get Delimiter() As String
  Delimiter = mSelectEntryDialog.Delimiter
End Property

Public Property Let Delimiter(ByVal strDelimiter As String)
  mSelectEntryDialog.Delimiter = strDelimiter
End Property

Public Property Get PreSelectValue() As String
  PreSelectValue = mSelectEntryDialog.PreSelectValue
End Property

Public Property Let PreSelectValue(ByVal strPreSelectValue As String)
  mSelectEntryDialog.PreSelectValue = strPreSelectValue
End Property

Public Property Get ReturnSingleEntryAutomatically() As Boolean
  ReturnSingleEntryAutomatically = mSelectEntryDialog.ReturnSingleEntryAutomatically
End Property

Public Property Let ReturnSingleEntryAutomatically(ByVal blnReturnSingleEntryAutomatically As Boolean)
  mSelectEntryDialog.ReturnSingleEntryAutomatically = blnReturnSingleEntryAutomatically
End Property

Public Property Get MultiSelect() As Boolean
  MultiSelect = mSelectEntryDialog.MultiSelect
End Property

Public Property Let MultiSelect(ByVal blnMultiSelect As Boolean)
  mSelectEntryDialog.MultiSelect = blnMultiSelect
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Sub Initialize()
  Set mSelectEntryDialog = New SelectEntryDiaolg
  Me.ReturnID = False
  Me.Sorted = True
  Me.SortInFolder = False
  Me.Header = "Bitte wählen ..."
  Me.Delimiter = "#"
  Me.PreSelectValue = ""
  Me.ReturnSingleEntryAutomatically = False
  Me.MultiSelect = False
  Me.PreSelect = False
  Me.UseFirstLetter = True
  
  mSelectEntryDialog.SelectEntry Me.DataCollection, , , , , , , True
End Sub

Private Function TransIntoCollection(ByVal strDataString As String, _
                                     ByVal blnReturnID As Boolean, _
                                     Optional ByVal blnSortInFolder As Boolean = False) As Collection
  Set TransIntoCollection = mSelectEntryDialog.TransIntoCollection(strDataString, blnReturnID, blnSortInFolder)
End Function



'---------------------- Öffentliche Methoden der Klasse -------------------------
Public Sub SelectEntry(Optional ByVal SelectData, _
                       Optional ByVal strHeader, _
                       Optional ByVal blnReturnID, _
                       Optional ByVal blnSorted, _
                       Optional ByVal blnSortInFolder, _
                       Optional ByVal strPreSelectValue, _
                       Optional ByVal blnReturnSingleEntryAutomatically)
  
  If IsMissing(blnReturnID) Then blnReturnID = Me.ReturnID
  
  If IsMissing(strHeader) Then strHeader = Me.Header
  If IsMissing(blnSorted) Then blnSorted = Me.Sorted
  If IsMissing(blnSortInFolder) Then blnSortInFolder = Me.SortInFolder
  If IsMissing(strPreSelectValue) Then strPreSelectValue = Me.PreSelectValue
  If IsMissing(blnReturnSingleEntryAutomatically) Then blnReturnSingleEntryAutomatically = Me.ReturnSingleEntryAutomatically
  
  If IsMissing(SelectData) Then
    Set SelectData = Me.DataCollection
  Else

    Select Case TypeName(SelectData)

      Case "Collection"

      Case "String"
        Set SelectData = TransIntoCollection(SelectData, blnReturnID, blnSortInFolder)
    End Select

  End If
  
  mSelectEntryDialog.SelectEntry SelectData, strHeader, blnReturnID, blnSorted, blnSortInFolder, strPreSelectValue, blnReturnSingleEntryAutomatically
End Sub

Public Sub SetDataSource(ByRef rs As Object)
  mSelectEntryDialog.SetDataSource rs
End Sub

Public Sub SetIniFileSource(ByVal strFileName As String, ByVal strSection As String)
  mSelectEntryDialog.SetIniFileSource strFileName, strSection
End Sub

Public Sub Reset()
  Initialize
End Sub
