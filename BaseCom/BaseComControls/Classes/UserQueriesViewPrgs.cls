VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserQueriesViewPrgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'{Eigenschaften der Klasse}
Private mstrUserFID As String
'{Ende Eigenschaften der Klasse}


'{Auflistung für Klausuren}
Private mcolPrgs As Collection


'{---------- Konstruktor und Destrunktor der Klasse ----------}
Private Sub Class_Initialize()
  '{Konstruktor der Klasse}
  Set mcolPrgs = New Collection
  Me.UserFID = -1
End Sub

Private Sub Class_Terminate()
  '{Destruktor der Klasse}
  Set mcolPrgs = Nothing
End Sub
'{---------- Ende Konstruktor und Destrunktor der Klasse ----------}


'{---------- Zugriffsmethoden der Eigenschaften der Klasse ----------}
Public Property Get UserFID() As String
  UserFID = mstrUserFID
End Property

Public Property Let UserFID(ByVal strUserFID As String)
  mstrUserFID = strUserFID
End Property
'{---------- Ende Zugriffsmethoden der Eigenschaften der Klasse ----------}


'{---------- Private Methoden der Klasse ----------}
Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub
'{---------- Ende Private Methoden der Klasse ----------}


'{---------- Öffentliche Methoden der Klasse ----------}
Public Function GetPrgs() As Collection
On Error GoTo errLabel:

  Set GetPrgs = mcolPrgs
  Exit Function

errLabel:
  ShowError "GetPrgs"
  Exit Function
End Function

Public Function GetPrgsDB() As Long

On Error GoTo errLabel:

  GetPrgsDB = 0
  Set mcolPrgs = New Collection

  Dim x As Variant
  For Each x In baseToolKit.WebService.Authentication.GetAppNames
    mcolPrgs.Add CStr(x), CStr(x)
  Next x
  Exit Function

errLabel:
  GetPrgsDB = -1
  ShowError "GetPrgsDB"
  Exit Function
End Function
'{---------- Ende Öffentliche Methoden der Klasse ----------}



