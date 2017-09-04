VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtSQL 
      Height          =   7815
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCommand1_Click()
  Dim cs As String
  cs = "DRIVER={MYSQL ODBC 5.1 Driver};SERVER=SQL01;UID=apps;PWD=bcw;PORT=3306;"
  
  Dim con As New ADODB.Connection
  con.ConnectionString = cs
  con.Open
   
  Dim rs As ADODB.Recordset
  Set rs = con.Execute(txtSQL.Text) '"SELECT Nachname AS Name, Vorname, Geburtsdatum, 120 AS Groesse FROM datapool.t_personen WHERE _rowid = 7")
   
  Dim printReport As New PrintReport10.clsPrintReport
  printReport.ReportFileName = "C:\_projects\VB6\PrintReport10\TestApp\teilnehmerliste.rpt" ' "h:\fuck.rpt"
  printReport.SetDataSource rs
  printReport.ShowReport 1
  
  con.Close
  MsgBox "fertig"
End Sub
