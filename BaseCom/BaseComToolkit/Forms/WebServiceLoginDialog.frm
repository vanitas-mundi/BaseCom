VERSION 5.00
Begin VB.Form WebServiceLoginDialog 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Anmeldung"
   ClientHeight    =   2445
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5175
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1444.587
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   4859.045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkDontShowLoginDialog 
      Caption         =   "Anmeldedialog nicht anzeigen"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1997
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cboAuthentication 
      Height          =   315
      ItemData        =   "WebServiceLoginDialog.frx":0000
      Left            =   1440
      List            =   "WebServiceLoginDialog.frx":0002
      Style           =   2  'Dropdown-Liste
      TabIndex        =   4
      Top             =   195
      Width           =   3615
   End
   Begin VB.Frame fraAuthentication 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   4935
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1305
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   720
         Width           =   3510
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblPassword 
         Caption         =   "&Kennwort:"
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label lblUserName 
         Caption         =   "&Benutzername:"
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   390
      Left            =   3919
      TabIndex        =   3
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label lblAuthentication 
      Caption         =   "Authentifizierung:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "WebServiceLoginDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'{-------- Enums der Klasse -------------}
Public Enum AuthenticationTypes
  Windows = 0
  WindowsWithoutPwd = 1
End Enum
'{-------- Ende Enums der Klasse -------------}

'Private Const PAUSE_KEY_CODE = 19
Private Const CONTROL_KEY_CODE = 17
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'{-------- Eigenschaften der Klasse -------------}
Private mAuthentication As WebServiceAuthentication
Private mAppName As String
Private mCancel As Boolean

Private mstrLastUserName As String
'{-------- Ende Eigenschaften der Klasse -------------}

'{-------------- Zugrifsmethoden der Klasseneigenschaften -----------------}
Public Property Get Cancel() As Boolean
  Cancel = mCancel
End Property

Public Property Get Authentication() As WebServiceAuthentication
  Set Authentication = mAuthentication
End Property

Public Property Let Authentication(ByVal value As WebServiceAuthentication)
  Set mAuthentication = value
End Property

Public Property Get AppName() As String
  AppName = mAppName
End Property

Public Property Let AppName(ByVal value As String)
  mAppName = value
End Property

Public Property Get auth() As AuthenticationTypes
  auth = cboAuthentication.ListIndex
End Property
'{-------------- Ende Zugrifsmethoden der Klasseneigenschaften -----------------}



'{---------------- Ereignis-Methoden der Klasse -----------------}
Private Sub Form_Load()
  Initialize
End Sub

Private Sub Form_Activate()
  CheckSingleSignOn
End Sub

Private Sub Form_Unload(Cancel As Integer)
  UnloadForm
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
  Case 13 'Enter
    CheckLogin
  Case 27 'Escape
    CancelDialog
  End Select
End Sub
  
Private Sub cboAuthentication_Click()
  ChangeAuthentication
End Sub

Private Sub cmdOK_Click()
  CheckLogin
End Sub

Private Sub cmdCancel_Click()
  CancelDialog
End Sub
'{---------------- Ende Ereignis-Methoden der Klasse -----------------}

'{---------------- Private Methoden der Klasse -----------------}
Private Sub Initialize()
  Me.tag = Me.AppName & "-Anmeldung"
  Me.Caption = Me.tag
  
  With cboAuthentication
    .AddItem "Windows Authentifizierung"
    .AddItem "Windows Authentifizierung ohne Kennwort"
    .ListIndex = GetSetting(app.EXEName, Me.name _
    , "Authentication", AuthenticationTypes.WindowsWithoutPwd)
  End With

  txtUserName.Text = GetSetting(app.EXEName, Me.name, "UserName", Me.Authentication.GetWindowsUserName)
  mstrLastUserName = txtUserName.Text

  chkDontShowLoginDialog.value = GetSetting(app.EXEName, Me.name, "DontShowLoginDialog", 0)
End Sub

Private Sub UnloadForm()
  SaveSetting app.EXEName, Me.name, "Authentication", cboAuthentication.ListIndex

  If Trim(txtUserName.Text) = "" Then
    SaveSetting app.EXEName, Me.name, "UserName", Me.Authentication.GetWindowsUserName
  Else
    SaveSetting app.EXEName, Me.name, "UserName", txtUserName.Text
  End If
  SaveSetting app.EXEName, Me.name, "DontShowLoginDialog", chkDontShowLoginDialog.value
End Sub

Private Function PauseKeyPressed() As Boolean
  Dim result As Integer: result = GetAsyncKeyState(CONTROL_KEY_CODE)
  PauseKeyPressed = (result = -32767) Or (result = 1)
End Function

Private Sub CheckSingleSignOn()

  If (Me.auth = WindowsWithoutPwd) And (chkDontShowLoginDialog.value) And (Not PauseKeyPressed) Then
    Me.Authentication.LoadGrants Me.AppName, txtUserName.Text
    Unload Me
  End If
End Sub

Private Sub ChangeAuthentication()

  Dim enabled As Boolean: enabled = (Me.auth = AuthenticationTypes.Windows)
  chkDontShowLoginDialog.Visible = Not enabled

  Select Case Me.auth
  Case AuthenticationTypes.Windows
    txtUserName.Text = mstrLastUserName
  Case AuthenticationTypes.WindowsWithoutPwd
    mstrLastUserName = txtUserName.Text
    txtUserName.Text = Me.Authentication.GetWindowsUserName
  End Select

  lblUserName.enabled = enabled
  txtUserName.enabled = enabled
  lblPassword.enabled = enabled
  txtPassword.enabled = enabled
End Sub

Private Sub CheckLogin()
  If Trim(txtUserName.Text) = "" Then
    Me.Caption = Me.tag & " [UserName missing]"
    Exit Sub
  End If
  
  If (Trim(txtPassword.Text) = "") And (Me.auth = Windows) Then
    Me.Caption = Me.tag & " [PassWord missing]"
    Exit Sub
  End If
  
  Me.Authentication.LoadGrants Me.AppName, txtUserName.Text, txtPassword.Text
  
  If Me.Authentication.IsLoginSuccessfull Then
    If Me.Authentication.CanExecute Then
      mCancel = False
      Unload Me
    Else
      Me.Caption = Me.tag & " [NoExecuteGranted]"
    End If
  Else
    Me.Caption = Me.tag & " [" & Me.Authentication.GetLoginResult & "]"
  End If
End Sub

Private Sub CancelDialog()
  mCancel = True
  Unload Me
End Sub
'{---------------- Ende Private Methoden der Klasse -----------------}
