VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Communication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : Communication
'    Project    : ToolKits
'
'    Description: Stellt Methoden für die Kommunikation zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private mBankInformation As CommunicationBankInfo
Private mMail As CommunicationMail
Private mPhone As CommunicationPhone
Private mPostcode As CommunicationPostCode
Private mSms As CommunicationSMS
Private mAddress As CommunicationAddress



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub Class_Initialize()
  Set mBankInformation = New CommunicationBankInfo
  Set mMail = New CommunicationMail
  Set mPhone = New CommunicationPhone
  Set mPostcode = New CommunicationPostCode
  Set mSms = New CommunicationSMS
  Set mAddress = New CommunicationAddress
End Sub

Private Sub Class_Terminate()
  Set mBankInformation = Nothing
  Set mMail = Nothing
  Set mPhone = Nothing
  Set mPostcode = Nothing
  Set mSms = Nothing
  Set mAddress = Nothing
End Sub



'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get BankInformation() As CommunicationBankInfo
  Set BankInformation = mBankInformation
End Property

Public Property Get mail() As CommunicationMail
  Set mail = mMail
End Property

Public Property Get Phone() As CommunicationPhone
  Set Phone = mPhone
End Property

Public Property Get Postcode() As CommunicationPostCode
  Set Postcode = mPostcode
End Property

Public Property Get SMS() As CommunicationSMS
  Set SMS = mSms
End Property

Public Property Get Address() As CommunicationAddress
  Set Address = mAddress
End Property

'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------




'---------------------- Öffentliche Methoden der Klasse -------------------------





