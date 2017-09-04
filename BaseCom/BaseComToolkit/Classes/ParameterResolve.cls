VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParameterResolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit

'---------------------- Eigenschaften der Klasse --------------------------------
Private mConv As Convert
Private mError As ToolkitError
Private mEtc As etc
Private mIni As FileIni



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mConv = New Convert
  Set mError = New ToolkitError
  Set mEtc = New etc
  Set mIni = New FileIni
End Sub

Private Sub Class_Terminate()
  Set mConv = New Convert
  Set mError = Nothing
  Set mEtc = Nothing
  Set mIni = Nothing
End Sub

Private Function MaskParameterValueParameters(ByVal value As String) As String
  value = Replace$(value, "[", "~pb;")
  value = Replace$(value, "]", "~pe;")
  MaskParameterValueParameters = value
End Function

Private Function UnMaskParameterValueParameters(ByVal value As String) As String
  value = Replace$(value, "~pb;", "[")
  value = Replace$(value, "~pe;", "]")
  UnMaskParameterValueParameters = value
End Function

Private Function ReplaceMissingParameters(ByVal statement As String) As String

  Dim parameters() As String: parameters = Split(statement, "[")

  Dim i As Long
  For i = LBound(parameters) To UBound(parameters)
    Dim pos As Long: pos = InStr(parameters(i), "]") - 1
    If pos > 0 Then
      Dim paramName As String: paramName = Mid$(parameters(i), 1, pos)
      Dim value As String: value = GetMissingParameterValue(paramName)
      value = MaskParameterValueParameters(value)
      statement = Replace(statement, "[" & paramName & "]", value)
    End If
  Next i

  ReplaceMissingParameters = statement
End Function

Private Function GetMissingParameterValue(ByVal parameter As String) As String

  Dim isConstant As Boolean: isConstant = Mid(parameter, 1, 1) = "?"
  Dim Temp As String: Temp = IIf(isConstant, Mid$(parameter, 2), parameter)
  Dim paramParts() As String: paramParts = Split(Temp, ">")
  Dim paramPartsCount As Long: paramPartsCount = mEtc.ArrayCount(paramParts)
  
  Dim paramName As String: paramName = ""
  Dim paramType As String: paramType = "default"
  Dim paramDefaultValue As String: paramDefaultValue = ""
  
  Select Case True
  Case (paramPartsCount = 1) And (Not isConstant)
    paramName = paramParts(0)
  Case (paramPartsCount = 1) And (isConstant)
    If mEtc.Contains(paramParts(0), "date", True) Then
      paramName = "Datum"
      paramType = "date"
      paramDefaultValue = "CURRENT_DATE()"
    Else
      paramName = "Uhrzeit"
      paramType = "time"
      paramDefaultValue = "CURRENT_TIME()"
    End If
  Case (paramPartsCount = 2) And (isConstant)
    paramName = paramParts(1)
    paramType = LCase(paramParts(0))
  Case (paramPartsCount = 3) And (isConstant)
    paramName = paramParts(1)
    paramType = LCase(paramParts(0))
    paramDefaultValue = paramParts(2)
  Case (paramPartsCount = 4) And (isConstant)
    paramName = paramParts(3)
    paramType = "default"
    paramDefaultValue = mIni.ReadEntryValue(paramParts(1), paramParts(2), paramParts(3))
  Case (paramPartsCount = 5) And (isConstant)
    paramName = paramParts(4)
    paramType = "default"
    paramDefaultValue = mIni.ReadEntryValue(paramParts(1), paramParts(2), paramParts(3))
  Case Else
    Err.Raise ParameterHasUnresolveableFormat, "FileSqlGroupFile.GetMissingParameterValue" _
    , mError.GetErrorMessage(ParameterHasUnresolveableFormat)
  End Select
  
  GetMissingParameterValue = GetMissingParameterValueByType(paramName, paramType, paramDefaultValue)
End Function

Private Function GetDateParameterValue _
(ByVal parameterName As String _
, Optional ByVal DefaultValue As String = "") As String

  Dim value As String
  value = InputBox("Bitte Wert für " & parameterName & " eingeben!" _
  , parameterName, mConv.FormatYMD(DefaultValue))
  
  If StrPtr(value) = 0 Then Err.Raise ParameterInputCanceledError
  If Not IsDate(value) Then Err.Raise ParameterInvalidInputFormat
  
  GetDateParameterValue = mConv.FormatYMD(value)
End Function

Private Function GetMissingParameterValueByType _
(ByVal paramName As String _
, ByVal paramType As String _
, ByVal paramDefaultValue As String) As String

  Dim value As String: value = ""

  Dim prompt As String: prompt = "Bitte Wert für " & paramName & " eingeben!"
  value = InputBox(prompt, paramName, paramDefaultValue)

  Dim occurredError As ErrorNumbers: occurredError = None

  Select Case True
  Case StrPtr(value) = 0 'Abbrechen wurde gedrückt
    occurredError = ParameterInputCanceledError
  Case paramType = "default" 'sonst würde ins else gesprungen werden
    value = value
  Case paramType = "date"
    If IsDate(value) Then
      value = Format(value, "yyyy-mm-dd")
    Else
      occurredError = ParameterInvalidInputFormat
    End If
  Case paramType = "time"
    If IsDate(value) Then
      value = Format(value, "hh:mm")
    Else
      occurredError = ParameterInvalidInputFormat
    End If
  Case paramType = "real"
    value = Replace$(value, ",", ".")
    If Not IsNumeric(value) Then
      occurredError = ParameterInvalidInputFormat
    End If
  Case paramType = "integer"
    value = Replace$(value, ",", ".")
    If (Not IsNumeric(value)) Or (InStr(value, ".") > 0) Then
      occurredError = ParameterInvalidInputFormat
    End If
  Case Else
    occurredError = ParameterHasUnresolveableFormat
  End Select

  If occurredError = None Then
    GetMissingParameterValueByType = value
  Else
    Err.Raise occurredError, "FileSqlGroupFile.GetMissingParameterValueByType" _
    , mError.GetErrorMessage(occurredError)
  End If
End Function

Public Function ReplaceParametersWithoutMissingParameters _
(ByVal statement As String, ByVal parameters As Variant) As String

  If mEtc.ArrayCount(parameters) > 0 Then
    Dim X As Variant
    For Each X In parameters
      Dim param As String: param = CStr(X)
      Dim pos As Integer: pos = InStr(param, ":")
  
      If pos > 0 Then
        Dim name As String: name = Mid$(param, 1, pos - 1)
        Dim value As String: value = MaskParameterValueParameters(Mid$(param, pos + 1))
        statement = Replace$(statement, "[" & name & "]", value)
      End If
    Next X
  End If
  
  ReplaceParametersWithoutMissingParameters = UnMaskParameterValueParameters(statement)
End Function

Public Function ReplaceParameters _
(ByVal statement As String, ByVal parameters As Variant) As String

  If mEtc.ArrayCount(parameters) > 0 Then
    Dim X As Variant
    For Each X In parameters
      Dim param As String: param = CStr(X)
      Dim pos As Integer: pos = InStr(param, ":")
  
      If pos > 0 Then
        Dim name As String: name = Mid$(param, 1, pos - 1)
        Dim value As String: value = MaskParameterValueParameters(Mid$(param, pos + 1))
        statement = Replace$(statement, "[" & name & "]", value)
      End If
    Next X
  End If
  
  statement = ReplaceMissingParameters(ByVal statement)
  
  ReplaceParameters = UnMaskParameterValueParameters(statement)
End Function


