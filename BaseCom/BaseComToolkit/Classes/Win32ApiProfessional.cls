VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiProfessional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mSysDialogs As Win32ApiDialogs
Private mSysDialogsEx As Win32ApiDialogsEx
Private mSysDrives As Win32ApiSysDrives
Private mSysEventLog As Win32ApiEventLog
Private mSysFiles As Win32ApiSysFiles
Private mSysFolders As Win32ApiSysFolders
Private mSysFonts As Win32ApiSysFonts
Private mSysInfo As Win32ApiSysInfo
Private mSysKeyboard As Win32ApiKeyboard
Private mSysNLS As Win32ApiSysNLS
Private mSysPrinter As Win32ApiSysPrinter
Private mSysRegistry As Win32ApiRegistry
Private mSysShell As Win32ApiSysShell
Private mSysWindow As Win32ApiSysWindow

Public Property Get SysDialogs() As Win32ApiDialogs
  Set SysDialogs = mSysDialogs
End Property

Public Property Get SysDialogsEx() As Win32ApiDialogsEx
  Set SysDialogsEx = mSysDialogsEx
End Property

Public Property Get SysDrives() As Win32ApiSysDrives
  Set SysDrives = mSysDrives
End Property

Public Property Get SysEventLog() As Win32ApiEventLog
  Set SysEventLog = mSysEventLog
End Property

Public Property Get SysFiles() As Win32ApiSysFiles
  Set SysFiles = mSysFiles
End Property

Public Property Get SysFolders() As Win32ApiSysFolders
  Set SysFolders = mSysFolders
End Property

Public Property Get SysFonts() As Win32ApiSysFonts
  Set SysFonts = mSysFonts
End Property

Public Property Get SysInfo() As Win32ApiSysInfo
  Set SysInfo = mSysInfo
End Property

Public Property Get SysKeyboard() As Win32ApiKeyboard
  Set SysKeyboard = mSysKeyboard
End Property

Public Property Get SysNLS() As Win32ApiSysNLS
  Set SysNLS = mSysNLS
End Property

Public Property Get SysPrinter() As Win32ApiSysPrinter
  Set SysPrinter = mSysPrinter
End Property

Public Property Get SysRegistry() As Win32ApiRegistry
  Set SysRegistry = mSysRegistry
End Property

Public Property Get SysShell() As Win32ApiSysShell
  Set SysShell = mSysShell
End Property

Public Property Get SysWindow() As Win32ApiSysWindow
  Set SysWindow = mSysWindow
End Property

Private Sub Class_Initialize()
  Set mSysDialogs = New Win32ApiDialogs
  Set mSysDialogsEx = New Win32ApiDialogsEx
  Set mSysDrives = New Win32ApiSysDrives
  Set mSysEventLog = New Win32ApiEventLog
  Set mSysFiles = New Win32ApiSysFiles
  Set mSysFolders = New Win32ApiSysFolders
  Set mSysFonts = New Win32ApiSysFonts
  Set mSysInfo = New Win32ApiSysInfo
  Set mSysKeyboard = New Win32ApiKeyboard
  Set mSysNLS = New Win32ApiSysNLS
  Set mSysPrinter = New Win32ApiSysPrinter
  Set mSysRegistry = New Win32ApiRegistry
  Set mSysShell = New Win32ApiSysShell
  Set mSysWindow = New Win32ApiSysWindow
End Sub

Private Sub Class_Terminate()
  Set mSysDialogs = Nothing
  Set mSysDialogsEx = Nothing
  Set mSysDrives = Nothing
  Set mSysEventLog = Nothing
  Set mSysFiles = Nothing
  Set mSysFolders = Nothing
  Set mSysFonts = Nothing
  Set mSysInfo = Nothing
  Set mSysKeyboard = Nothing
  Set mSysNLS = Nothing
  Set mSysPrinter = Nothing
  Set mSysRegistry = Nothing
  Set mSysShell = Nothing
  Set mSysWindow = Nothing
End Sub


