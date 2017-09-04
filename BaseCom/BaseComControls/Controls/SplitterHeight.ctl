VERSION 5.00
Begin VB.UserControl SplitterHeight 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   MousePointer    =   7  'Größenänderung N S
   ScaleHeight     =   810
   ScaleWidth      =   1620
   ToolboxBitmap   =   "SplitterHeight.ctx":0000
   Begin VB.Image Image 
      Height          =   585
      Left            =   120
      Picture         =   "SplitterHeight.ctx":0532
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "SplitterHeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'{---------Eigenschaften der Klasse--------------}

Private msngLastY As Single
Private mblnMoving As Boolean
Private mblnParentNotForm As Boolean
Private mlngSplitterColor As Long
Private mlngSplitterHeight As Long
Private mlngMinHeightTopControl As Long
Private mlngMinHeightBottomControl As Long

Private mcolTopControls As Collection
Private mcolBottomControls As Collection
Private mstrAppExeName As String

'{---------Ende Eigenschaften der Klasse--------------}


'{-------------Konstruktor und Destruktor der Klasse------------}

Private Sub UserControl_Initialize()
  Me.SplitterColor = &H808080
  Me.SplitterHeight = 90
  Me.MinHeightTopControl = 1000
  Me.MinHeightBottomControl = 1000
  Me.ParentNotForm = False
  Me.AppExeName = App.EXEName
  Set mcolTopControls = New Collection
  Set mcolBottomControls = New Collection
End Sub

Private Sub UserControl_Terminate()
  Set mcolTopControls = Nothing
  Set mcolBottomControls = Nothing
End Sub

'{-------------Konstruktor und Destruktor der Klasse------------}


'{-------------Zugriffsmethoden der Klasseneigenschaften----------}

Public Property Get SplitterColor() As Long
Attribute SplitterColor.VB_MemberFlags = "400"
  SplitterColor = mlngSplitterColor
End Property

Public Property Let SplitterColor(ByVal lngSplitterColor As Long)
  mlngSplitterColor = lngSplitterColor
End Property

Public Property Get SplitterHeight() As Long
Attribute SplitterHeight.VB_MemberFlags = "400"
  SplitterHeight = mlngSplitterHeight
End Property

Public Property Let SplitterHeight(ByVal lngSplitterHeight As Long)
  mlngSplitterHeight = lngSplitterHeight
  UserControl.Width = lngSplitterHeight
End Property

Public Property Get MinHeightTopControl() As Long
Attribute MinHeightTopControl.VB_MemberFlags = "400"
  MinHeightTopControl = mlngMinHeightTopControl
End Property

Public Property Let MinHeightTopControl(ByVal lngMinHeightTopControl As Long)
  mlngMinHeightTopControl = lngMinHeightTopControl
End Property

Public Property Get MinHeightBottomControl() As Long
Attribute MinHeightBottomControl.VB_MemberFlags = "400"
  MinHeightBottomControl = mlngMinHeightBottomControl
End Property

Public Property Let MinHeightBottomControl(ByVal lngMinHeightBottomControl As Long)
  mlngMinHeightBottomControl = lngMinHeightBottomControl
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

Public Property Get BottomControls() As Collection
  Set BottomControls = mcolBottomControls
End Property

Public Property Get TopControls() As Collection
  Set TopControls = mcolTopControls
End Property

'{-------------Ende Zugriffsmethoden der Klasseneigenschaften----------}


'{------------Private KLassenmethoden-------------------}

Private Sub UserControl_Resize()
  UserControl.Height = Me.SplitterHeight
End Sub

Private Sub SizeControls(ByVal sngDelta As Single)
Dim obj As Object
Dim sngSurfaceSize As Single

On Error Resume Next
    
  sngSurfaceSize = (Me.BottomControls(1).Top + Me.BottomControls(1).Height) - (Me.TopControls(1).Top)
    
  For Each obj In Me.TopControls
    obj.Height = obj.Height + sngDelta
    SaveSurfacePart obj.Name, (obj.Height * 100) / sngSurfaceSize
  Next obj
                
  For Each obj In Me.BottomControls
    obj.Height = obj.Height - sngDelta
    SaveSurfacePart obj.Name, (obj.Height * 100) / sngSurfaceSize
        
    If (TypeOf obj.Container Is Form) Or (Me.ParentNotForm) Then
      obj.Top = (obj.Top + sngDelta)
    End If
  Next obj
    
End Sub

Private Sub SaveSurfacePart(ByVal strObjName As String, ByVal sngValue As Single)
  SaveSetting Me.AppExeName, UserControl.Parent.Name, "SurfacePartHeight" & strObjName, sngValue
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 1 Then Exit Sub
  BackColor = Me.SplitterColor
  UserControl.Extender.ZOrder
  msngLastY = UserControl.Extender.Top
  mblnMoving = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 1 Then Exit Sub
  SizeControls UserControl.Extender.Top - msngLastY
  BackColor = &H8000000F
  mblnMoving = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngPos As Single

  If Button <> 1 Then Exit Sub

  If Not mblnMoving Then Exit Sub
  
  sngPos = Y + UserControl.Extender.Top
    
  With UserControl.Extender
    Select Case True
    Case sngPos < Me.MinHeightTopControl + Me.TopControls(1).Top
      .Top = Me.MinHeightTopControl + Me.TopControls(1).Top
    Case sngPos > (Me.BottomControls(1).Top + Me.BottomControls(1).Height) - Me.MinHeightBottomControl
      .Top = (Me.BottomControls(1).Top + Me.BottomControls(1).Height) - Me.MinHeightBottomControl
    Case Else
      .Top = sngPos
    End Select
  End With
  
End Sub

'{------------Ende Private KLassenmethoden-------------------}


'{------------Public KLassenmethoden-------------------}

Public Sub SetStoredProportions()
Dim sngValue As Single
Dim obj As Object
Dim sngSurfaceSize As Single
Dim sngTop As Single

  sngSurfaceSize = (Me.BottomControls(1).Top + Me.BottomControls(1).Height) - (Me.TopControls(1).Top)

  For Each obj In Me.TopControls
    sngValue = GetSetting(Me.AppExeName, UserControl.Parent.Name, "SurfacePart" & obj.Name, 0)
    obj.Height = (sngValue * sngSurfaceSize) / 100
  Next obj

  sngTop = Me.TopControls(1).Top + Me.TopControls(1).Height
  UserControl.Extender.Top = sngTop
  sngTop = sngTop + UserControl.Height
  
  For Each obj In Me.BottomControls
    sngValue = GetSetting(Me.AppExeName, UserControl.Parent.Name, "SurfacePart" & obj.Name, 0)
    If (TypeOf obj.Container Is Form) Or (Me.ParentNotForm) Then
      obj.Top = sngTop
    End If
    
    obj.Height = (sngValue * sngSurfaceSize) / 100
    
  Next obj

End Sub

'{------------Ende Public KLassenmethoden-------------------}


