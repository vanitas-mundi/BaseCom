Attribute VB_Name = "ToolKitsModule"
Option Explicit

Private mBaseToolKit As ToolKits.BaseToolKitVb6

Public Property Get BaseToolKit() As ToolKits.BaseToolKitVb6
  Set BaseToolKit = mBaseToolKit
End Property

Public Property Let BaseToolKit(ByVal Value As ToolKits.BaseToolKitVb6)
  Set mBaseToolKit = Value
End Property

