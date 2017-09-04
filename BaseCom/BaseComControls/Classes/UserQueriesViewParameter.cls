VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserQueriesViewParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'{Modulvariablen}
Private mastrOperatoren(0 To 15) As String
Private mastrParameterTypes(0 To 3) As String


'{Eigenschaften der Klasse}
Private mstrParameterID As String
Private mstrQueryFID As String
Private mstrOwnerID As String
Private mstrName As String
Private mstrField As String
Private mstrValue As String
Private mOperator As eParameterOperator
Private mstrOperatorText As String
Private mParameterType As eParameterType
Private mstrParameterTypeText As String
Private mstrComment As String
'{Ende Eigenschaften der Klasse}

Public Enum eParameterOperator
  pveNoParameter = 0      'Kein Parameter
  pveEqual = 1            'entspricht
  pveNotEqual = 2         'entspricht nicht
  pveMemberOf = 3         'ist einer von
  pveNotMemberOf = 4      'ist nicht einer von
  pveLesser = 5           'ist kleiner als
  pveLesserEqual = 6      'ist kleiner oder gleich
  pveGreater = 7          'ist größer als
  pveGreaterEqual = 8     'ist größer oder gleich
  pveBetween = 9          'liegt zwischen
  pveNotBetween = 10      'liegt nicht zwischen
  pveLike = 11            'wie
  pveIntervalMinute = 12  'Minutenintervall
  pveIntervalHour = 13    'Stundenintervall
  
  pveIsNULL = 14          'Ist NULL
  pveIsNotNULL = 15       'Ist nicht NULL
End Enum

Public Enum eParameterType
  pteText = 0           'Text
  pteReal = 1           'Dezimalzahl
  pteInteger = 2        'Ganzzahl
  pteDate = 3           'Datum
End Enum

            
'{---------- Konstruktor und Destrunktor der Klasse ----------}
Private Sub Class_Initialize()
  '{Konstruktor der Klasse}
  Me.ParameterID = "-1"
  Me.QueryFID = "-1"
  Me.OwnerID = "-1"
  Me.Name = "NeuerParameter"
  Me.Field = ""
  Me.value = ""
  Me.Operator = pveNoParameter
  Me.ParameterType = pteText
  Me.Comment = ""
  
  mastrOperatoren(0) = "Kein Parameter"
  mastrOperatoren(1) = "entspricht"
  mastrOperatoren(2) = "entspricht nicht"
  mastrOperatoren(3) = "ist einer von"
  mastrOperatoren(4) = "ist nicht einer von"
  mastrOperatoren(5) = "ist kleiner als"
  mastrOperatoren(6) = "ist kleiner oder gleich"
  mastrOperatoren(7) = "ist größer als"
  mastrOperatoren(8) = "ist größer oder gleich"
  mastrOperatoren(9) = "liegt zwischen"
  mastrOperatoren(10) = "liegt nicht zwischen"
  mastrOperatoren(11) = "wie"
  mastrOperatoren(12) = "Minutenintervall"
  mastrOperatoren(13) = "Stundenintervall"
  mastrOperatoren(14) = "ist NULL"
  mastrOperatoren(15) = "ist nicht NULL"
  

  mastrParameterTypes(0) = "Text"
  mastrParameterTypes(1) = "Dezimalzahl"
  mastrParameterTypes(2) = "Ganzzahl"
  mastrParameterTypes(3) = "Datum"
End Sub

Private Sub Class_Terminate()
  '{Destruktor der Klasse}
End Sub
'{---------- Ende Konstruktor und Destrunktor der Klasse ----------}


'{---------- Zugriffsmethoden der Eigenschaften der Klasse ----------}
Public Property Get ParameterID() As String
  ParameterID = mstrParameterID
End Property

Public Property Let ParameterID(ByVal strParameterID As String)
  mstrParameterID = strParameterID
End Property

Public Property Get QueryFID() As String
  QueryFID = mstrQueryFID
End Property

Public Property Let QueryFID(ByVal strQueryFID As String)
  mstrQueryFID = strQueryFID
End Property

Public Property Get OwnerID() As String
  OwnerID = mstrOwnerID
End Property

Public Property Let OwnerID(ByVal strOwnerID As String)
  mstrOwnerID = strOwnerID
End Property

Public Property Get Name() As String
  Name = mstrName
End Property

Public Property Let Name(ByVal strName As String)
  mstrName = strName
End Property

Public Property Get Field() As String
  Field = mstrField
End Property

Public Property Let Field(ByVal strField As String)
  mstrField = strField
End Property

Public Property Get value() As String
  value = mstrValue
End Property

Public Property Let value(ByVal strValue As String)
  
  strValue = Replace(strValue, "\,", "<escape name=comma>")
  
  Dim astrValues() As String: astrValues = Split(strValue, ",")

  Dim i As Integer
  For i = LBound(astrValues) To UBound(astrValues)

    astrValues(i) = Trim(astrValues(i))

    If LCase(astrValues(i)) <> "#null#" Then
  
      Select Case Me.ParameterType
      Case eParameterType.pteText
      Case eParameterType.pteInteger
        If Not IsNumeric(astrValues(i)) Then
          MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
          Exit Property
        End If
      Case eParameterType.pteReal
        If Not IsNumeric(astrValues(i)) Then
          MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
          Exit Property
        End If
        
        astrValues(i) = BaseToolKit.FileSystem.SqlGroupFile.CommaToPoint(astrValues(i))
      Case eParameterType.pteDate
        If Not IsDate(astrValues(i)) Then
          MsgBox "Ungültige Eingabe!", 16, "Formatfehler"
          Exit Property
        End If
        
        astrValues(i) = BaseToolKit.FileSystem.SqlGroupFile.FormatYMD(astrValues(i))
      End Select
    Else
      astrValues(i) = "#NULL#"
    End If
    
  Next i
  
  mstrValue = BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Join(astrValues, ","))
End Property

Public Property Get Operator() As eParameterOperator
  Operator = mOperator
End Property

Public Property Let Operator(ByVal Operator As eParameterOperator)
  mOperator = Operator
  mstrOperatorText = mastrOperatoren(Operator)
End Property

Public Property Get OperatorText() As String
  OperatorText = mstrOperatorText
End Property

Public Property Get ParameterType() As eParameterType
  ParameterType = mParameterType
End Property

Public Property Let ParameterType(ByVal ParameterType As eParameterType)
Dim astrValues() As String
Dim i As Integer


  astrValues = Split(Me.value, ",")
  For i = LBound(astrValues) To UBound(astrValues)

    If Me.value <> "" Then
      Select Case ParameterType
      Case eParameterType.pteText
      Case eParameterType.pteReal
        If Not IsNumeric(astrValues(i)) Then
          MsgBox "Parameterwert wird zurückgesetzt!", 48, "Formatfehler"
          Me.value = ""
        End If
      Case eParameterType.pteInteger
        If Not IsNumeric(astrValues(i)) Then
          MsgBox "Parameterwert wird zurückgesetzt!", 48, "Formatfehler"
          Me.value = ""
        End If
      Case eParameterType.pteDate
        If Not IsDate(astrValues(i)) Then
          MsgBox "Parameterwert wird zurückgesetzt!", 48, "Formatfehler"
          Me.value = ""
        End If
      End Select
    End If
    
  Next i
  
  mParameterType = ParameterType
  mstrParameterTypeText = mastrParameterTypes(ParameterType)
End Property

Public Property Get ParameterTypeText() As String
  ParameterTypeText = mstrParameterTypeText
End Property

Public Property Get Comment() As String
  Comment = mstrComment
End Property

Public Property Let Comment(ByVal strComment As String)
  mstrComment = strComment
End Property
'{---------- Ende Zugriffsmethoden der Eigenschaften der Klasse ----------}


'{---------- Private Methoden der Klasse ----------}
Private Sub ShowError(ByVal strPlace As String)
  Screen.MousePointer = 0
  MsgBox "(" & Err.Number & ") " & Err.Description, 16, strPlace
End Sub
'{---------- Ende Private Methoden der Klasse ----------}


'{---------- Öffentliche Methoden der Klasse ----------}
Public Function GetParameter(ByVal strParameterID) As Long

On Error GoTo errLabel:

  GetParameter = 0
  
  Dim strSQL As String
  strSQL = strSQL & "SELECT" & vbCrLf & vbTab
  strSQL = strSQL & "p.ParameterID," & vbCrLf & vbTab
  strSQL = strSQL & "p.QueryFID," & vbCrLf & vbTab
  strSQL = strSQL & "p.OwnerFID," & vbCrLf & vbTab
  strSQL = strSQL & "p.Name," & vbCrLf & vbTab
  strSQL = strSQL & "p.Field," & vbCrLf & vbTab
  strSQL = strSQL & "p.Value," & vbCrLf & vbTab
  strSQL = strSQL & "p.Operator," & vbCrLf & vbTab
  strSQL = strSQL & "p.ParameterType," & vbCrLf
  strSQL = strSQL & "p.Comment" & vbCrLf
  strSQL = strSQL & "FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_parameter p" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "p.ParameterID = '" & strParameterID & "'"
  
  Dim rs As Object: Set rs = BaseToolKit.Database.ExecuteReaderConnected(strSQL)
  
  If Not rs.EOF Then
    Me.ParameterID = rs!ParameterID & ""
    Me.QueryFID = rs!QueryFID & ""
    Me.OwnerID = rs!OwnerFID & ""
    Me.Name = rs!Name & ""
    Me.Field = rs!Field & ""
    Me.value = rs!value & ""
    Me.Operator = rs!Operator & ""
    Me.ParameterType = rs!ParameterType & ""
    Me.Comment = rs!Comment & ""
  End If
  
  BaseToolKit.Database.CloseRecordSet rs
  Exit Function

errLabel:
  GetParameter = -1
  ShowError "GetParameter"
  Exit Function
End Function

Public Function SaveParameter() As Long

On Error GoTo errLabel:

  SaveParameter = 0

  If Me.ParameterID = "-1" Then '{Neuanlegen}

    Dim strSQL As String
    strSQL = strSQL & "INSERT INTO" & vbCrLf & vbTab
    strSQL = strSQL & "queries.t_parameter" & vbCrLf & vbTab
    strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "QueryFID," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "OwnerFID," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Name," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Field," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Value," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Operator," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "ParameterType," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "Comment" & vbCrLf & vbTab
    strSQL = strSQL & ")" & vbCrLf
    strSQL = strSQL & "VALUES" & vbCrLf & vbTab
    strSQL = strSQL & "(" & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.QueryFID & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.OwnerID & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Name) & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.Field & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.value & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.Operator & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & Me.ParameterType & "'," & vbCrLf & vbTab & vbTab
    strSQL = strSQL & "'" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Comment) & "'" & vbCrLf & vbTab
    strSQL = strSQL & ")"
    
    Me.ParameterID = BaseToolKit.Database.ExecuteNonQuery(strSQL)
  Else '{Aktualisieren}
    strSQL = strSQL & "UPDATE" & vbCrLf & vbTab
    strSQL = strSQL & "queries.t_parameter" & vbCrLf
    strSQL = strSQL & "SET" & vbCrLf & vbTab
    strSQL = strSQL & "QueryFID = '" & Me.QueryFID & "'," & vbCrLf & vbTab
    strSQL = strSQL & "OwnerFID = '" & Me.OwnerID & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Name = '" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Name) & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Field = '" & Me.Field & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Value = '" & Me.value & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Operator = '" & Me.Operator & "'," & vbCrLf & vbTab
    strSQL = strSQL & "ParameterType = '" & Me.ParameterType & "'," & vbCrLf & vbTab
    strSQL = strSQL & "Comment = '" & BaseToolKit.FileSystem.SqlGroupFile.ReplaceEscape(Me.Comment) & "'" & vbCrLf
    strSQL = strSQL & "WHERE" & vbCrLf & vbTab
    strSQL = strSQL & "ParameterID = '" & Me.ParameterID & "'"
    BaseToolKit.Database.ExecuteNonQuery strSQL
  End If
  Exit Function

errLabel:
  SaveParameter = -1
  ShowError "SaveParameter"
  Exit Function
End Function

Public Function DeleteParameter() As Long

On Error GoTo errLabel:

  DeleteParameter = 0
  
  Dim strSQL As String
  strSQL = strSQL & "DELETE FROM" & vbCrLf & vbTab
  strSQL = strSQL & "queries.t_parameter" & vbCrLf
  strSQL = strSQL & "WHERE" & vbCrLf & vbTab
  strSQL = strSQL & "ParameterID = '" & Me.ParameterID & "'"
  BaseToolKit.Database.ExecuteNonQuery strSQL
  Exit Function

errLabel:
  DeleteParameter = -1
  ShowError "DeleteParameter"
  Exit Function
End Function

Public Function GetOperatorTextString() As String
  GetOperatorTextString = GetOperatorTextString & "#Kein Parameter#0"
  GetOperatorTextString = GetOperatorTextString & "#entspricht#1"
  GetOperatorTextString = GetOperatorTextString & "#entspricht nicht#2"
  GetOperatorTextString = GetOperatorTextString & "#ist einer von#3"
  GetOperatorTextString = GetOperatorTextString & "#ist nicht einer von#4"
  GetOperatorTextString = GetOperatorTextString & "#ist kleiner als#5"
  GetOperatorTextString = GetOperatorTextString & "#ist kleiner oder gleich#6"
  GetOperatorTextString = GetOperatorTextString & "#ist größer als#7"
  GetOperatorTextString = GetOperatorTextString & "#ist größer oder gleich#8"
  GetOperatorTextString = GetOperatorTextString & "#liegt zwischen#9"
  GetOperatorTextString = GetOperatorTextString & "#liegt nicht zwischen#10"
  GetOperatorTextString = GetOperatorTextString & "#wie#11"
  GetOperatorTextString = GetOperatorTextString & "#Minutenintervall#12"
  GetOperatorTextString = GetOperatorTextString & "#Stundenintervall#13"
  GetOperatorTextString = GetOperatorTextString & "#ist NULL#14"
  GetOperatorTextString = GetOperatorTextString & "#ist nicht NULL#15"
End Function

Public Function GetParameterTypeTextString() As String
  GetParameterTypeTextString = GetParameterTypeTextString & "#Text#0"
  GetParameterTypeTextString = GetParameterTypeTextString & "#Dezimalzahl#1"
  GetParameterTypeTextString = GetParameterTypeTextString & "#Ganzzahl#2"
  GetParameterTypeTextString = GetParameterTypeTextString & "#Datum#3"
End Function
'{---------- Ende Öffentliche Methoden der Klasse ----------}

