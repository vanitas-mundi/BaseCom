VERSION 5.00
Begin VB.UserControl SplitterWidth 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   MousePointer    =   9  'Größenänderung W O
   ScaleHeight     =   810
   ScaleWidth      =   1620
   ToolboxBitmap   =   "SplitterWidth.ctx":0000
   Begin VB.Image Image 
      Height          =   585
      Left            =   120
      Picture         =   "SplitterWidth.ctx":0532
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "SplitterWidth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'{---------Eigenschaften der Klasse--------------}

Private msngLastX As Single
Private mblnMoving As Boolean
Private mblnParentNotForm As Boolean
Private mlngSplitterColor As Long
Private mlngSplitterWidth As Long
Private mlngMinWidthLeftControl As Long
Private mlngMinWidthRightControl As Long

Private mcolLeftControls As Collection
Private mcolRightControls As Collection
Private mstrAppExeName As String

'{---------Ende Eigenschaften der Klasse--------------}


'{-------------Konstruktor und Destruktor der Klasse------------}

Private Sub UserControl_Initialize()
  Me.SplitterColor = &H808080
  Me.SplitterWidth = 75
  Me.MinWidthLeftControl = 1000
  Me.MinWidthRightControl = 1000
  Me.ParentNotForm = False
  Me.AppExeName = App.EXEName
  Set mcolLeftControls = New Collection
  Set mcolRightControls = New Collection
End Sub

Private Sub UserControl_Terminate()
  Set mcolLeftControls = Nothing
  Set mcolRightControls = Nothing
End Sub

'{-------------Konstruktor und Destruktor der Klasse------------}


'{-------------Zugriffsmethoden der Klasseneigenschaften----------}

Public Property Get SplitterColor() As Long
Attribute SplitterColor.VB_ProcData.VB_Invoke_Property = "StandardColor"
Attribute SplitterColor.VB_MemberFlags = "400"
  SplitterColor = mlngSplitterColor
End Property

Public Property Let SplitterColor(ByVal lngSplitterColor As Long)
  mlngSplitterColor = lngSplitterColor
End Property

Public Property Get SplitterWidth() As Long
Attribute SplitterWidth.VB_MemberFlags = "400"
  SplitterWidth = mlngSplitterWidth
End Property

Public Property Let SplitterWidth(ByVal lngSplitterWidth As Long)
  mlngSplitterWidth = lngSplitterWidth
  UserControl.Width = lngSplitterWidth
End Property

Public Property Get MinWidthLeftControl() As Long
Attribute MinWidthLeftControl.VB_MemberFlags = "400"
  MinWidthLeftControl = mlngMinWidthLeftControl
End Property

Public Property Let MinWidthLeftControl(ByVal lngMinWidthLeftControl As Long)
  mlngMinWidthLeftControl = lngMinWidthLeftControl
End Property

Public Property Get MinWidthRightControl() As Long
Attribute MinWidthRightControl.VB_MemberFlags = "400"
  MinWidthRightControl = mlngMinWidthRightControl
End Property

Public Property Let MinWidthRightControl(ByVal lngMinWidthRightControl As Long)
  mlngMinWidthRightControl = lngMinWidthRightControl
End Property

Public Property Get ParentNotForm() As Boolean
Attribute ParentNotForm.VB_MemberFlags = "400"
  ParentNotForm = mblnParentNotForm
End Property

Public Property Let ParentNotForm(ByVal blnParentNotForm As Boolean)
  mblnParentNotForm = blnParentNotForm
End Property

Public Property Get AppExeName() As String
Attribute AppExeName.VB_MemberFlags = "400"
  AppExeName = mstrAppExeName
End Property

Public Property Let AppExeName(ByVal strAppExeName As String)
  mstrAppExeName = strAppExeName
End Property

Public Property Get RightControls() As Collection
  Set RightControls = mcolRightControls
End Property

Public Property Get LeftControls() As Collection
  Set LeftControls = mcolLeftControls
End Property

'{-------------Ende Zugriffsmethoden der Klasseneigenschaften----------}


'{------------Private KLassenmethoden-------------------}

Private Sub UserControl_Resize()
  UserControl.Width = Me.SplitterWidth
End Sub

Private Sub SizeControls(ByVal sngDelta As Single)
Dim obj As Object
Dim sngSurfaceSize As Single

On Error Resume Next
    
  sngSurfaceSize = (Me.RightControls(1).Left + Me.RightControls(1).Width) - (Me.LeftControls(1).Left)
    
  For Each obj In Me.LeftControls
    obj.Width = obj.Width + sngDelta
    SaveSurfacePart obj.Name, (obj.Width * 100) / sngSurfaceSize
  Next obj
                
  For Each obj In Me.RightControls
    obj.Width = obj.Width - (sngDelta)
    SaveSurfacePart obj.Name, (obj.Width * 100) / sngSurfaceSize
        
    If (TypeOf obj.Container Is Form) Or (Me.ParentNotForm) Then
      obj.Left = (obj.Left + sngDelta)
    End If
  Next obj
    
End Sub

Private Sub SaveSurfacePart(ByVal strObjName As String, ByVal sngValue As Single)
  SaveSetting Me.AppExeName, UserControl.Parent.Name, "SurfacePartWidth" & strObjName, sngValue
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 1 Then Exit Sub
  BackColor = Me.SplitterColor
  UserControl.Extender.ZOrder
  msngLastX = UserControl.Extender.Left
  mblnMoving = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 1 Then Exit Sub
  SizeControls UserControl.Extender.Left - msngLastX
  BackColor = &H8000000F
  mblnMoving = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngPos As Single

  If Button <> 1 Then Exit Sub

  If Not mblnMoving Then Exit Sub
  
  sngPos = X + UserControl.Extender.Left
    
  With UserControl.Extender
    Select Case True
    Case sngPos < Me.MinWidthLeftControl + Me.LeftControls(1).Left
      .Left = Me.MinWidthLeftControl + Me.LeftControls(1).Left
    Case sngPos > (Me.RightControls(1).Left + Me.RightControls(1).Width) - Me.MinWidthRightControl
      .Left = (Me.RightControls(1).Left + Me.RightControls(1).Width) - Me.MinWidthRightControl
    Case Else
      .Left = sngPos
    End Select
  End With
  
End Sub

'{------------Ende Private KLassenmethoden-------------------}


'{------------Public KLassenmethoden-------------------}

Public Sub SetStoredProportions()
Dim sngValue As Single
Dim obj As Object
Dim sngSurfaceSize As Single
Dim sngLeft As Single

'On Error Resume Next

  sngSurfaceSize = (Me.RightControls(1).Left + Me.RightControls(1).Width) - (Me.LeftControls(1).Left)

  For Each obj In Me.LeftControls
    sngValue = GetSetting(Me.AppExeName, UserControl.Parent.Name, "SurfacePart" & obj.Name, 0)
    obj.Width = (sngValue * sngSurfaceSize) / 100
  Next obj

  sngLeft = Me.LeftControls(1).Left + Me.LeftControls(1).Width
  UserControl.Extender.Left = sngLeft
  sngLeft = sngLeft + UserControl.Width
  
  For Each obj In Me.RightControls
    sngValue = GetSetting(Me.AppExeName, UserControl.Parent.Name, "SurfacePart" & obj.Name, 0)
    obj.Width = (sngValue * sngSurfaceSize) / 100
    
    If (TypeOf obj.Container Is Form) Or (Me.ParentNotForm) Then
      obj.Left = sngLeft
    End If
  Next obj

End Sub

'{------------Ende Public KLassenmethoden-------------------}

