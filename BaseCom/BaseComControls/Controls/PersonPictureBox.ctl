VERSION 5.00
Begin VB.UserControl PersonPictureBox 
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1440
   ScaleWidth      =   3405
   ToolboxBitmap   =   "PersonPictureBox.ctx":0000
   Begin VB.PictureBox removeImagePictureBox 
      BorderStyle     =   0  'Kein
      Height          =   250
      Left            =   1320
      Picture         =   "PersonPictureBox.ctx":0312
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   250
   End
   Begin VB.PictureBox imagePictureBox 
      BorderStyle     =   0  'Kein
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblSaveNewImage 
      BackStyle       =   0  'Transparent
      Caption         =   "Neues Bild speichern"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1725
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "PersonPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : PersonPictureBox
'    Project    : BcwControls
'
'    Description: UserControl zur Darstellung und Änderung eines Personenbildes
'
'    Modified   : 05.12.2014 by Sebastian Limke
'--------------------------------------------------------------------------------



'---------------------- Eigenschaften der Klasse --------------------------------
Public Enum PictureFields
  Picture = 0
  PictureHQ = 1
End Enum

Private Const MAX_IMAGE_SIZE_KB = 20

Private mPersonId As Long
Private mPictureField As PictureFields



'---------------------- Konstruktor der Klasse ----------------------------------
Private Sub UserControl_Initialize()
End Sub

'---------------------- Zugriffsmethoden der Klasse -----------------------------
Public Property Get PictureFieldString() As String
  Select Case Me.PictureField
  Case PictureFields.Picture
    PictureFieldString = "Picture"
  Case PictureFields.PictureHQ
    PictureFieldString = "PictureHQ"
  Case Else
    PictureFieldString = ""
  End Select
End Property

Public Property Get PictureField() As PictureFields
  PictureField = mPictureField
End Property

Public Property Let PictureField(ByVal Value As PictureFields)
  mPictureField = Value
End Property

Public Property Get PersonId() As Long
  PersonId = mPersonId
End Property

Public Property Let PersonId(ByVal Value As Long)
  mPersonId = Value
End Property

Public Property Get SaveButtonVisible() As Boolean
  SaveButtonVisible = lblSaveNewImage.Visible
End Property

Public Property Let SaveButtonVisible(ByVal Value As Boolean)
  lblSaveNewImage.Visible = Value
  ResizeMe
End Property

Public Property Get RemoveImageButtonVisible() As Boolean
  RemoveImageButtonVisible = removeImagePictureBox.Visible
End Property

Public Property Let RemoveImageButtonVisible(ByVal Value As Boolean)
  removeImagePictureBox.Visible = Value
End Property



'---------------------- Ereignismethoden der Klasse -----------------------------
Private Sub UserControl_Resize()
  ResizeMe
End Sub

Private Sub lblSaveNewImage_Click()
  SaveNewPersonPicture
End Sub

Private Sub removeImagePictureBox_Click()
  RemoveImage
End Sub



'---------------------- Private Methoden der Klasse -----------------------------
'--------------------------------------------------------------------------------
' Project    :       BcwControls
' Procedure  :       LoadPersonPicture
' Description:       Laedt das Bild der Person
' Created by :       Project Administrator
' Machine    :       Sebastian Limke
' Date-Time  :       16.01.2015-11:59:14
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub LoadPersonPicture()
  
  '<EhHeader>
  On Error GoTo LoadPersonPicture_Err
  '</EhHeader>
  
  Set imagePictureBox.Picture = ToolKitsModule.BaseToolKit.Database.GetImageFromDatabase _
  ("datapool", "t_photos", "PersonenFID", Me.PictureFieldString, Me.PersonId)
  
  '<EhFooter>
  Exit Sub
  
LoadPersonPicture_Err:
  Err.Raise vbObjectError + 100, "BcwControls.PersonPictureBox.LoadPersonPicture", Err.Description
  '</EhFooter>
End Sub

'--------------------------------------------------------------------------------
' Project    :       BcwControls
' Procedure  :       SaveNewPersonPicture
' Description:       Speichert ein neues Personenbild
' Created by :       Sebastian Limke
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-12:03:52
'
' Parameters :
'--------------------------------------------------------------------------------
Private Sub SaveNewPersonPicture()
  
  '<EhHeader>
  On Error GoTo SaveNewPersonPicture
  '</EhHeader>
  
  If Me.PersonId <= 0 Then Exit Sub
  
  Dim fileName As String: fileName = ToolKitsModule.BaseToolKit.Dialog.GetOpenFileName _
  ("Bitte wählen Sie eine Bilddatei aus:", "", "", True, "jpg", "bmp", "gif", "png")
  
  If fileName = "" Then Exit Sub
  
  If (ToolKitsModule.BaseToolKit.FileSystem.io.GetFileSize(fileName, KBytes) > MAX_IMAGE_SIZE_KB) _
  And (Me.PictureField = Picture) Then
    ToolKitsModule.BaseToolKit.ToolkitError.ShowError _
    "Bild speichern", "Bilddatei darf nicht größer als 20KB sein!"
    Exit Sub
  End If
  
  Dim s As String
  With ToolKitsModule.BaseToolKit.Database
    s = "SELECT IF(COUNT(1) = 0, false, true) AS PhotoExists FROM datapool.t_photos WHERE PersonenFID = " & Me.PersonId
    If CBool(.ExecuteScalar(s)) Then
      .UpdateImageInDatabase "datapool", "t_photos" _
      , "PersonenFID", Me.PictureFieldString, Me.PersonId, fileName
    Else
      Dim id As Long: id = .InsertImageIntoDatabase _
      ("datapool", "t_photos", Me.PictureFieldString, fileName)
      s = "UPDATE datapool.t_photos SET PersonenFID = " & Me.PersonId & " WHERE _rowid = " & id
      .ExecuteNonQuery s
    End If
  End With
  
  ReloadPicture
 
  '<EhFooter>
  Exit Sub
  
SaveNewPersonPicture:
  Err.Raise vbObjectError + 100, "BcwControls.PersonPictureBox.SaveNewPersonPicture", Err.Description
  '</EhFooter>
End Sub

Private Sub ResizeMe()
  removeImagePictureBox.Top = 0
  removeImagePictureBox.Left = UserControl.Width - removeImagePictureBox.Width

  imagePictureBox.Left = 0
  imagePictureBox.Top = 0
  imagePictureBox.Width = UserControl.Width
  
  If lblSaveNewImage.Visible Then
    lblSaveNewImage.Left = 0
    lblSaveNewImage.Width = UserControl.Width
    lblSaveNewImage.Top = UserControl.Height - lblSaveNewImage.Height
    imagePictureBox.Height = UserControl.Height - (lblSaveNewImage.Height + 40)
  Else
    imagePictureBox.Height = UserControl.Height
  End If
  
End Sub



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       BcwControls
' Procedure  :       Reset
' Description:       Loescht das UserPicture aus der Anzeige
' Created by :       Project Administrator
' Machine    :       Sebastian Limke
' Date-Time  :       16.01.2015-12:01:34
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub Reset()
  Set imagePictureBox.Picture = Nothing
End Sub

'--------------------------------------------------------------------------------
' Project    :       BcwControls
' Procedure  :       ReloadPicture
' Description:       Laedt das PersonPicture erneut in die PictureBox
' Created by :       Sebastian Limke
' Date-Time  :       16.01.2015-12:02:31
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub ReloadPicture()
  LoadPersonPicture
End Sub

'--------------------------------------------------------------------------------
' Project    :       BcwControls
' Procedure  :       RemoveImage
' Description:       Löscht das Personenbild aus der Datenbank
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       16.01.2015-12:42:03
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub RemoveImage()
  Dim prompt As String: prompt = "Soll das Bild wirklich aus der Datenbank gelöscht werden?"
  Dim title As String: title = "Bild löschen"
  
  If MsgBox(prompt, 36, title) = vbYes Then
    Dim s As String: s = "UPDATE datapool.t_photos SET " _
    & Me.PictureFieldString & " = NULL WHERE PersonenFID = " & Me.PersonId
    ToolKitsModule.BaseToolKit.Database.ExecuteNonQuery s
    Reset
  End If
End Sub
