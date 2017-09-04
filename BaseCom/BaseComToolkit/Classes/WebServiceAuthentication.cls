VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebServiceAuthentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : WebServiceAuthentication
'    Project    : ToolKits
'
'    Description: Stellt Kommunikation mit dem Autorisierungs-WebService zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private Const WEB_SERVICE_URL = "http://appserver01.bcw-intern.local/Services/GrantService/GrantServiceLibrary.GrantService.svc?singleWsdl"

Private mConnectionString As String

Private mGrants      As Collection
Private mPersonId    As Long
Private mUserName    As String
Private mLastName    As String
Private mFirstName   As String
Private mMailAddress As String
Private mDirectDial  As String
Private mPhoneNumber As String


'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mGrants = New Collection
End Sub

Private Sub Class_Terminate()
  Set mGrants = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ConnectionString() As String
  ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal value As String)
  mConnectionString = value
End Property

Public Property Get PersonId() As Long
  PersonId = mPersonId
End Property

Public Property Get userName() As String
  userName = mUserName
End Property

Public Property Get LastName() As String
  LastName = mLastName
End Property

Public Property Get FirstName() As String
  FirstName = mFirstName
End Property

Public Property Get FullName() As String
  FullName = Trim(Me.FirstName & " " & Me.LastName)
End Property

Public Property Get FullNameReverse() As String
  FullNameReverse = Trim(Me.LastName & ", " & Me.FirstName)
End Property

Public Property Get mailAddress() As String
  mailAddress = mMailAddress
End Property

Public Property Get DirectDial() As String
  DirectDial = mDirectDial
End Property

Public Property Get phoneNumber() As String
  phoneNumber = mPhoneNumber
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ShowLoginDialog
' Description:       Öffnet den Anmeldedialog.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:26:35
'
' Parameters :       appName (String)
'--------------------------------------------------------------------------------
Public Sub ShowLoginDialog(ByVal AppName As String)
  Dim f As WebServiceLoginDialog: Set f = New WebServiceLoginDialog
  f.Authentication = Me
  f.AppName = AppName
  f.Show 1

  If f.Cancel Then
    Set mGrants = New Collection
    mGrants.Add "n", "LoginOk"
    mGrants.Add "CancelPressed", "LoginResult"
    mGrants.Add "n", "Execute"
  End If
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SelectEmployeeAndLogin
' Description:       Führt eine Anmeldung mit dem ausgewählten User durch.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       24.09.2015-13:26:35
'
' Parameters :       appName (String)
'--------------------------------------------------------------------------------
Public Sub SelectEmployeeAndLogin(ByVal AppName As String)

  Dim statement As Collection: Set statement = New Collection
  statement.Add "SELECT"
  statement.Add vbTab & "PersonenFID AS ID,"
  statement.Add vbTab & "Name AS Value"
  statement.Add "FROM"
  statement.Add vbTab & "tele_v.v_active_employees"
  statement.Add "ORDER BY"
  statement.Add vbTab & "Name"
  Dim conv As Convert: Set conv = New Convert

  Dim db As Database: Set db = New Database
  db.ConnectionString = Me.ConnectionString
  Dim rs As Object: Set rs = db.ExecuteReaderConnected(conv.JoinCollection(statement, vbCrLf))
  
  Dim dialogs As dialog: Set dialogs = New dialog
  dialogs.SelectEntry.Reset
  dialogs.SelectEntry.SetDataSource rs
  dialogs.SelectEntry.SelectEntry , , True, True
  db.CloseRecordSet rs
  
  If dialogs.SelectEntry.ValueEntry = "" Then
    Set mGrants = New Collection
    mGrants.Add "n", "LoginOk"
    mGrants.Add "CancelPressed", "LoginResult"
    mGrants.Add "n", "Execute"
  Else
    Set statement = New Collection
    statement.Add "SELECT"
    statement.Add vbTab & "LoginName AS UserName"
    statement.Add "FROM"
    statement.Add vbTab & "datapool.t_personen"
    statement.Add "WHERE"
    statement.Add vbTab & "_rowid = " & dialogs.SelectEntry.ValueID
    Dim userName As String: userName = db.ExecuteScalar(conv.JoinCollection(statement, vbCrLf))
    Me.LoadGrants AppName, userName
  End If
 
  Set conv = Nothing
  Set db = Nothing
  Set rs = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetGrantString
' Description:       Liefert den GrantString vom WebService.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:27:19
'
' Parameters :       AppName (String)
'                    userName (String)
'                    passWord (String = "")
'--------------------------------------------------------------------------------
Public Function GetGrantString(ByVal AppName As String, _
                               ByVal userName As String, _
                               Optional ByVal passWord As String = "") As String
                    
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  
  If passWord = "" Then
    GetGrantString = soap.GetGrantStringSingleSignOn(AppName, userName)
  Else
    GetGrantString = soap.GetGrantString(AppName, userName, passWord)
  End If
  
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetPersonId
' Description:       Liefert die im AD hinterlegte Personen-Id anhand des Usernames.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:27:50
'
' Parameters :       userName (String)
'--------------------------------------------------------------------------------
Public Function GetPersonId(ByVal userName As String) As String
                    
  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetPersonId = soap.GetUserProperty(userName, "employeeId")
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetGrants
' Description:       Liefert eine Collection mit den Grant-Namen und -Werten einer App.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:28:36
'
' Parameters :       AppName (String)
'                    userName (String)
'                    passWord (String)
'--------------------------------------------------------------------------------
Private Function GetGrants(ByVal AppName As String, _
                           ByVal userName As String, _
                           Optional ByVal passWord = "") As Collection

  Set GetGrants = New Collection
  
  Dim grantString As String: grantString = Me.GetGrantString(AppName, userName, passWord)

  Dim X           As Variant

  For Each X In Split(grantString, ",")
    Dim parts() As String: parts = Split(X, ":")
    Dim name    As String: name = LCase(parts(0))
    Dim value   As String: value = parts(1)
    GetGrants.Add value, name
  Next X

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetAppNames
' Description:       Liefert alle im AD hinterlegten AppNames.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:29:31
'
' Parameters :       ) (String()
'--------------------------------------------------------------------------------
Public Function GetAppNames() As String()

  Dim soap As Object: Set soap = CreateObject("MSSOAP.SoapClient30")
  soap.MSSoapInit WEB_SERVICE_URL
  GetAppNames = soap.GetAppNames
  Set soap = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetGrantNames
' Description:       Liefert alle Berechtigungsnamen einer App.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:29:56
'
' Parameters :       AppName (String)
'--------------------------------------------------------------------------------
Public Function GetGrantNames(ByVal AppName As String) As Collection

  Set GetGrantNames = New Collection

  Dim grantString As String: grantString = Me.GetGrantString(AppName, "Administrator")

  Dim X           As Variant

  For Each X In Split(grantString, ",")
    Dim parts() As String: parts = Split(X, ":")
    Dim name    As String: name = parts(0)
    GetGrantNames.Add name
  Next X

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetGrantedInstitutes
' Description:       Liefert gesetzte Institutsberechtigungen anhand des UserNames.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:30:42
'
' Parameters :       userName (String)
'--------------------------------------------------------------------------------
Public Function GetGrantedInstitutes(ByVal userName As String) As Collection
  
  Set GetGrantedInstitutes = New Collection
  
  Dim grantString         As String: grantString = GetGrantString("granted_institute", userName)
  Dim grantedInstitutes() As String: grantedInstitutes = Split(grantString, ",")
  
  Dim X                   As Variant

  For Each X In grantedInstitutes
    Dim parts() As String: parts = Split(X, ":")
    Dim name    As String: name = parts(0)
    Dim value   As String: value = LCase(parts(1))

    If value = "y" Then GetGrantedInstitutes.Add name
  Next X

End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       LoadGrants
' Description:       Lädt Berechtigunen und Zusatzinformationen eines Users.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:31:51
'
' Parameters :       ConnectionString (String)
'                    AppName (String)
'                    userName (String)
'                    passWord (String)
'--------------------------------------------------------------------------------
Public Sub LoadGrants(ByVal AppName As String, _
                      ByVal userName As String, _
                      Optional ByVal passWord = "")

  mPersonId = 0
  mUserName = ""
  mLastName = ""
  mFirstName = ""
  mMailAddress = ""
  mDirectDial = ""

  Set mGrants = GetGrants(AppName, userName, passWord)
  
  If Me.IsLoginSuccessfull Then
    mPersonId = Me.GetPersonId(userName)
    mUserName = userName
    
    Dim db As Database: Set db = New Database
    db.ConnectionString = Me.ConnectionString
    
    Dim statement As Collection: Set statement = New Collection
    statement.Add "SELECT"
    statement.Add vbTab & "Nachname AS LastName,"
    statement.Add vbTab & "Vorname AS FirstName,"
    statement.Add vbTab & "datapool.GetCommunication(_rowid,'E-Mail (dienstlich)') AS MailAddress,"
    statement.Add vbTab & "datapool.GetCommunication(_rowid,'Durchwahl (dienstlich)') AS DirectDial,"
    statement.Add vbTab & "datapool.GetCommunication(_rowid,'Telefon (dienstlich)') AS PhoneNumber"
    statement.Add "FROM"
    statement.Add vbTab & "datapool.t_personen"
    statement.Add "WHERE"
    statement.Add vbTab & "_rowid = " & Me.PersonId
    Dim conv As Convert: Set conv = New Convert
    
    Dim rs As Object: Set rs = db.ExecuteReaderConnected(conv.JoinCollection(statement, vbCrLf))
    While Not rs.EOF
      mLastName = rs!LastName & ""
      mFirstName = rs!FirstName & ""
      mMailAddress = rs!mailAddress & ""
      mDirectDial = rs!DirectDial & ""
      mPhoneNumber = rs!phoneNumber & ""
      rs.MoveNext
    Wend
    db.CloseRecordSet rs
    
    Set db = Nothing
    Set conv = Nothing
  End If

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CanExecute
' Description:       Prüft, ob Execute-Berechtigung vorhanden ist.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:32:26
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function CanExecute() As Boolean
  CanExecute = IsGranted("Execute")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsGranted
' Description:       Prüft, ob die angegebene Berechtigung gewährt wurde.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:32:52
'
' Parameters :       grantName (String)
'--------------------------------------------------------------------------------
Public Function IsGranted(ByVal grantName As String) As Boolean
  IsGranted = False

  Dim etc As etc: Set etc = New etc
  grantName = LCase$(grantName)
  
  If etc.CollectionContainsKey(mGrants, grantName) Then
    If LCase(mGrants.item(grantName)) = "y" Then IsGranted = True
  End If

  Set etc = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       IsLoginSuccessfull
' Description:       Prüft, ob der Anmeldeversuch erfolgreich verlaufen ist.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:33:27
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function IsLoginSuccessfull() As Boolean

  IsLoginSuccessfull = Me.IsGranted("LoginOk") And Me.IsGranted("Execute")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetLoginResult
' Description:       Liefert das Ergebnis des Login-Versuchs als String.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-13:34:00
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function GetLoginResult() As String

  GetLoginResult = "GrantsNotLoaded"

  Dim etc       As etc: Set etc = New etc
  Dim grantName As String: grantName = LCase$("LoginResult")
  
  If etc.CollectionContainsKey(mGrants, grantName) Then
    GetLoginResult = mGrants.item(grantName)
  End If

  If (GetLoginResult = "Successful") And (Not Me.IsGranted("Execute")) Then
    GetLoginResult = "ExecuteGrantMissing"
  End If

  Set etc = Nothing
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetWindowsUserName
' Description:       Liefert den an Windows angemeldeten Benutzernamen.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       26.01.2015-16:20:22
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function GetWindowsUserName() As String
  GetWindowsUserName = Environ("USERNAME")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetDomainName
' Description:       Liefert den Domänennamen.
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       26.01.2015-16:21:32
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function GetDomainName() As String
  GetDomainName = Environ("USERDOMAIN")
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       InstituteGranted
' Description:       Prüft, ob angemeldeter Benutzer Zugriff auf das angegebene
'                    Institut besitzt.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       11.02.2015-10:34:42
'
' Parameters :       institute (String)
'--------------------------------------------------------------------------------
Public Function InstituteGranted(ByVal institute As String) As Boolean

  InstituteGranted = False
  institute = LCase$(institute)
  
  Dim X As Variant
  For Each X In Me.GetGrantedInstitutes(Me.userName)
    If LCase(CStr(X)) = institute Then
      InstituteGranted = True
      Exit Function
    End If
  Next X
End Function

