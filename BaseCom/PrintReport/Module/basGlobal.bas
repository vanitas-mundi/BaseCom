Attribute VB_Name = "basGlobal"
Option Explicit

Public ODBC As ODBCHandling.clsODBCHandling

Public Enum preOutput
  prScreen = 0
  prPrinter = 1
End Enum

Public Sub ShowError(ByVal strMessage As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strMessage
End Sub

