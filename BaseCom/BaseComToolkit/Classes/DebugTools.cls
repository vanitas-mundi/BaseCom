VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DebugTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DataTypeEnumToString(ByVal value As Long) As String
  Select Case value
  Case 20
    DataTypeEnumToString = "adBigInt"
  Case 128
    DataTypeEnumToString = "adBinary"
  Case 11
    DataTypeEnumToString = "adBoolean"
  Case 8
    DataTypeEnumToString = "adBSTR"
  Case 136
    DataTypeEnumToString = "adChapter"
  Case 129
    DataTypeEnumToString = "adChar"
  Case 6
    DataTypeEnumToString = "adCurrency"
  Case 7
    DataTypeEnumToString = "adDate"
  Case 133
    DataTypeEnumToString = "adDBDate"
  Case 134
    DataTypeEnumToString = "adDBTime"
  Case 135
    DataTypeEnumToString = "adDBTimeStamp"
  Case 14
    DataTypeEnumToString = "adDecimal"
  Case 5
    DataTypeEnumToString = "adDouble"
  Case 0
    DataTypeEnumToString = "adEmpty"
  Case 10
    DataTypeEnumToString = "adError"
  Case 64
    DataTypeEnumToString = "adFileTime"
  Case 72
    DataTypeEnumToString = "adGUID"
  Case 9
    DataTypeEnumToString = "adIDispatch"
  Case 3
    DataTypeEnumToString = "adInteger"
  Case 13
    DataTypeEnumToString = "adIUnknown"
  Case 205
    DataTypeEnumToString = "adLongVarBinary"
  Case 201
    DataTypeEnumToString = "adLongVarChar"
  Case 203
    DataTypeEnumToString = "adLongVarWChar"
  Case 131
    DataTypeEnumToString = "adNumeric"
  Case 138
    DataTypeEnumToString = "adPropVariant"
  Case 4
    DataTypeEnumToString = "adSingle"
  Case 2
    DataTypeEnumToString = "adSmallInt"
  Case 16
    DataTypeEnumToString = "adTinyInt"
  Case 21
    DataTypeEnumToString = "adUnsignedBigInt"
  Case 19
    DataTypeEnumToString = "adUnsignedInt"
  Case 18
    DataTypeEnumToString = "adUnsignedSmallInt"
  Case 17
    DataTypeEnumToString = "adUnsignedTinyInt"
  Case 132
    DataTypeEnumToString = "adUserDefined"
  Case 204
    DataTypeEnumToString = "adVarBinary"
  Case 200
    DataTypeEnumToString = "adVarChar"
  Case 12
    DataTypeEnumToString = "adVariant"
  Case 139
    DataTypeEnumToString = "adVarNumeric"
  Case 202
    DataTypeEnumToString = "adVarWChar"
  Case 130
    DataTypeEnumToString = "adWChar"
  Case Else
    DataTypeEnumToString = "#####"
  End Select
End Function

