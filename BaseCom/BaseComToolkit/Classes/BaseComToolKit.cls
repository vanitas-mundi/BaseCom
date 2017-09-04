VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseToolKitVb6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : BaseToolKitVb6
'    Project    : ToolKits
'
'    Description: Stellt den Zugriff auf die gesamte ToolBox bereit.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Private mConnectionString As String

Private mCommunication As Communication
Private mControls      As ControlExtensions
Private mConvert       As Convert
Private mCryptography  As Cryptography
Private mDatabase      As Database
Private mDialog        As dialog
Private mElo           As Elo
Private mEtc           As etc
Private mFileSystem    As FileSystem
Private mReflection     As Reflection
Private mSettings      As Settings
Private mToolkitError  As ToolkitError
Private mWebService    As WebService
Private mWin32Api      As win32API
Private mDebugTools    As DebugTools
Private mRegEx         As RegEx


'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mCommunication = New Communication
  Set mControls = New ControlExtensions
  Set mConvert = New Convert
  Set mCryptography = New Cryptography
  Set mDatabase = New Database
  Set mDialog = New dialog
  Set mElo = New Elo
  Set mEtc = New etc
  Set mFileSystem = New FileSystem
  Set mReflection = New Reflection
  Set mSettings = New Settings
  Set mToolkitError = New ToolkitError
  Set mWebService = New WebService
  Set mWin32Api = New win32API
  Set mDebugTools = New DebugTools
  Set mRegEx = New RegEx
End Sub

Private Sub Class_Terminate()
  Set mCommunication = Nothing
  Set mControls = Nothing
  Set mConvert = Nothing
  Set mCryptography = Nothing
  Set mDatabase = Nothing
  Set mDialog = Nothing
  Set mElo = Nothing
  Set mEtc = Nothing
  Set mFileSystem = Nothing
  Set mReflection = Nothing
  Set mSettings = Nothing
  Set mToolkitError = Nothing
  Set mWebService = Nothing
  Set mWin32Api = Nothing
  Set mDebugTools = Nothing
  Set mRegEx = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get ConnectionString() As String
  ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal value As String)
  mConnectionString = value
  Me.WebService.Authentication.ConnectionString = value
  Me.Communication.BankInformation.ConnectionString = value
  Me.Communication.Postcode.ConnectionString = value
  Me.Controls.comboBox.ConnectionString = value
  Me.Database.ConnectionString = value
  Me.FileSystem.SqlGroupFile.DatabaseIntern = Me.Database
End Property

Public Property Get Communication() As Communication
  Set Communication = mCommunication
End Property

Public Property Get Controls() As ControlExtensions
  Set Controls = mControls
End Property

Public Property Get Convert() As Convert
  Set Convert = mConvert
End Property

Public Property Get Cryptography() As Cryptography
  Set Cryptography = mCryptography
End Property

Public Property Get Database() As Database
  Set Database = mDatabase
End Property

Public Property Get dialog() As dialog
  Set dialog = mDialog
End Property

Public Property Get Elo() As Elo
  Set Elo = mElo
End Property

Public Property Get etc() As etc
  Set etc = mEtc
End Property

Public Property Get FileSystem() As FileSystem
  Set FileSystem = mFileSystem
End Property

Public Property Get Reflection() As Reflection
  Set Reflection = mReflection
End Property

Public Property Get Settings() As Settings
  Set Settings = mSettings
End Property

Public Property Get ToolkitError() As ToolkitError
  Set ToolkitError = mToolkitError
End Property

Public Property Get WebService() As WebService
  Set WebService = mWebService
End Property

Public Property Get win32API() As win32API
  Set win32API = mWin32Api
End Property

Public Property Get DebugTools() As DebugTools
  Set DebugTools = mDebugTools
End Property

Public Property Get RegEx() As RegEx
  Set RegEx = mRegEx
End Property


'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
