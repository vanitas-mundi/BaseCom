VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReflectionMemberInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum InvokeKinds
    INVOKE_UNKNOWN = 0
    INVOKE_FUNC = 1
    INVOKE_PROPERTYGET = 2
    INVOKE_PROPERTYPUT = 4
    INVOKE_PROPERTYPUTREF = 8
    INVOKE_EVENTFUNC = 16
    INVOKE_CONST = 32
End Enum

Public MemberName As String
Public MemberID As String
Public HelpString As String
Public value As String
Public valueType As String
Public ParameterCount As String
Public MemberType As InvokeKinds

Public Property Get MemberTypeString() As String

  Select Case Me.MemberType
  Case INVOKE_CONST
    MemberTypeString = "CONST"
  Case INVOKE_EVENTFUNC
    MemberTypeString = "EVENTFUNC"
  Case INVOKE_FUNC
    MemberTypeString = "FUNC"
  Case INVOKE_PROPERTYGET
    MemberTypeString = "PROPERTYGET"
  Case INVOKE_PROPERTYPUT
    MemberTypeString = "PROPERTYPUT"
  Case INVOKE_PROPERTYPUTREF
    MemberTypeString = "PROPERTYPUTREF"
  Case INVOKE_UNKNOWN
    MemberTypeString = "UNKNOWN"
  Case Else
    MemberTypeString = "UNKNOWN"
  End Select
End Property

