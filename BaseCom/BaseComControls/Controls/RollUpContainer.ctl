VERSION 5.00
Begin VB.UserControl RollUpContainer 
   Alignable       =   -1  'True
   BackColor       =   &H80000014&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "RollUpContainer.ctx":0000
   Begin VB.Timer tmrStateChange 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3105
      Top             =   1905
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2055
      Top             =   1920
   End
   Begin VB.PictureBox picBack 
      Align           =   1  'Oben ausrichten
      Height          =   1320
      Left            =   0
      ScaleHeight     =   1260
      ScaleWidth      =   4740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4800
      Begin VB.CheckBox chkHeader 
         Height          =   345
         Left            =   645
         Style           =   1  'Grafisch
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   315
         Width           =   1125
      End
      Begin VB.CheckBox chkButton 
         Height          =   300
         Left            =   2340
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Grafisch
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   255
         UseMaskColor    =   -1  'True
         Width           =   645
      End
   End
End
Attribute VB_Name = "RollUpContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum RollUpStateBitmapConstants
    rusUp = 101
    rusDown = 102
End Enum

Private mNoResizeEvent As Boolean

Public Event BeforeStateChange(ByVal OldState As RollUpStateConstants, Cancel As Boolean)
Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Public Event HeaderClick()
Public Event HeaderMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event HeaderMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event HeaderMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Public Event Resize()
Public Event StateChanged(ByVal State As RollUpStateConstants)

Public Enum RollUpBorderStyleContants
    ruBorderStyleNone
    ruBorderStyleSunken
End Enum

Public Enum RollUpButtonPositionConstants
    ruButtonRight
    ruButtonLeft
End Enum

Public Enum RollUpStateConstants
    ruRollDown = False
    ruRollUp = True
End Enum

Private pBorderStyle As RollUpBorderStyleContants
Private pButtonPosition As RollUpButtonPositionConstants
Private pHeaderHeight As Single
Private pHeaderVisible As Boolean
Private pState As RollUpStateConstants

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute BackColor.VB_UserMemId = -501
    BackColor = picBack.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picBack.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As RollUpBorderStyleContants
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = pBorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As RollUpBorderStyleContants)
    Select Case New_BorderStyle
        Case pBorderStyle
            Exit Property
        Case ruBorderStyleNone, ruBorderStyleSunken
            pBorderStyle = New_BorderStyle
        Case Else
            Err.Raise 380
    End Select
    UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

Public Property Get ButtonPosition() As RollUpButtonPositionConstants
Attribute ButtonPosition.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    ButtonPosition = pButtonPosition
End Property

Public Property Let ButtonPosition(ByVal New_ButtonPosition As RollUpButtonPositionConstants)
    Select Case New_ButtonPosition
        Case pButtonPosition
            Exit Property
        Case ruButtonRight, ruButtonLeft
        Case Else
            Err.Raise 380
    End Select
    pButtonPosition = New_ButtonPosition
    UserControl_Resize
    PropertyChanged "ButtonPosition"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    chkHeader.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get HeaderBackColor() As OLE_COLOR
Attribute HeaderBackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    HeaderBackColor = chkHeader.BackColor
End Property

Public Property Let HeaderBackColor(ByVal New_HeaderBackColor As OLE_COLOR)
    chkHeader.BackColor = New_HeaderBackColor
    PropertyChanged "HeaderBackColor"
End Property

Public Property Get HeaderCaption() As String
Attribute HeaderCaption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute HeaderCaption.VB_UserMemId = -518
    HeaderCaption = chkHeader.Caption
End Property

Public Property Let HeaderCaption(New_HeaderCaption As String)
    chkHeader.Caption = New_HeaderCaption
    PropertyChanged "HeaderCaption"
End Property

Public Property Get HeaderFont() As Font
Attribute HeaderFont.VB_ProcData.VB_Invoke_Property = ";Schriftart"
Attribute HeaderFont.VB_UserMemId = -512
    Set HeaderFont = chkHeader.Font
End Property

Public Property Let HeaderFont(New_HeaderFont As Font)
    zSetHeaderFont New_HeaderFont
End Property

Public Property Set HeaderFont(New_HeaderFont As Font)
    zSetHeaderFont New_HeaderFont
End Property

Private Sub zSetHeaderFont(New_HeaderFont As Font)
    Set chkHeader.Font = New_HeaderFont
    PropertyChanged "HeaderFont"
End Sub

Public Property Get HeaderForeColor() As OLE_COLOR
Attribute HeaderForeColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
Attribute HeaderForeColor.VB_UserMemId = -513
    HeaderForeColor = chkHeader.ForeColor
End Property

Public Property Let HeaderForeColor(ByVal New_HeaderForeColor As OLE_COLOR)
    chkHeader.ForeColor = New_HeaderForeColor
    PropertyChanged "HeaderForeColor"
End Property

Public Property Get HeaderHeight() As Single
Attribute HeaderHeight.VB_ProcData.VB_Invoke_Property = ";Maﬂstab"
    HeaderHeight = pHeaderHeight
End Property

Public Property Let HeaderHeight(ByVal New_HeaderHeight As Single)
    Select Case New_HeaderHeight
        Case Is > 0
            pHeaderHeight = New_HeaderHeight
            UserControl_Resize
        Case Else
            Err.Raise 380
    End Select
    PropertyChanged "HeaderHeight"
End Property

Public Property Get HeaderLeft() As Single
Attribute HeaderLeft.VB_ProcData.VB_Invoke_Property = ";Maﬂstab"
    HeaderLeft = chkHeader.Left
End Property

Public Property Get HeaderVisible() As Boolean
Attribute HeaderVisible.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    HeaderVisible = pHeaderVisible
End Property

Public Property Let HeaderVisible(ByVal New_HeaderVisible As Boolean)
    pHeaderVisible = New_HeaderVisible
    chkHeader.Visible = pHeaderVisible
    PropertyChanged "HeaderVisible"
End Property

Public Property Get HeaderWidth() As Single
Attribute HeaderWidth.VB_ProcData.VB_Invoke_Property = ";Maﬂstab"
    HeaderWidth = chkHeader.Width
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Set Picture = picBack.Picture
End Property

Public Property Let Picture(New_Picture As Picture)
    zSetPicture New_Picture
End Property

Public Property Set Picture(New_Picture As Picture)
    zSetPicture New_Picture
End Property

Private Sub zSetPicture(New_Picture As Picture)
    Set picBack.Picture = New_Picture
    PropertyChanged "Picture"
End Sub

Public Property Get State() As RollUpStateConstants
Attribute State.VB_ProcData.VB_Invoke_Property = ";Verhalten"
Attribute State.VB_MemberFlags = "200"
    State = pState
End Property

Public Property Let State(ByVal New_State As RollUpStateConstants)
    Dim nCancel As Boolean
    
    Select Case New_State
        Case pState
            Exit Property
        Case ruRollDown, ruRollUp
            RaiseEvent BeforeStateChange(pState, nCancel)
            If Not nCancel Then
                mNoResizeEvent = True
                pState = New_State
                UserControl_Resize
                tmrStateChange.Enabled = True
                mNoResizeEvent = False
            End If
        Case Else
            Err.Raise 380
    End Select
    PropertyChanged "State"
End Property

Public Property Get RollUpSize() As Single
Attribute RollUpSize.VB_ProcData.VB_Invoke_Property = ";Position"
    RollUpSize = picBack.Height
End Property

Public Property Get ScaleHeight() As Single
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ScaleWidth() As Single
    ScaleWidth = UserControl.ScaleWidth
End Property

Private Sub chkButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkButton.Value = vbChecked
    picBack.SetFocus
    Me.State = Not Me.State
End Sub

Private Sub chkButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkButton.Value = vbUnchecked
End Sub

Private Sub chkHeader_Click()
    RaiseEvent HeaderClick
End Sub

Private Sub chkHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack.SetFocus
    RaiseEvent HeaderMouseDown(Button, Shift, X, Y)
End Sub

Private Sub chkHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent HeaderMouseMove(Button, Shift, X, Y)
End Sub

Private Sub chkHeader_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent HeaderMouseUp(Button, Shift, X, Y)
    RaiseEvent HeaderClick
End Sub

Private Sub picBack_Click()
    RaiseEvent Click
End Sub

Private Sub picBack_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picBack_GotFocus()
    Dim nControl As Control
    Dim nTabStop As Boolean
    Dim nSetControl As Control
    Dim nLowestTabIndex As Integer
    
    With UserControl
        On Error Resume Next
        nLowestTabIndex = .ParentControls.Count
        For Each nControl In .ContainedControls
            nTabStop = False
            Err.Clear
            With nControl
                nTabStop = .TabStop
                If Err.Number = 0 Then
                    If .Visible And .Enabled Then
                        If nTabStop Then
                            If .TabIndex < nLowestTabIndex Then
                                Set nSetControl = nControl
                            End If
                        End If
                    End If
                End If
            End With
        Next
    End With
    If Not (nSetControl Is Nothing) Then
        nSetControl.SetFocus
    End If
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub tmrResize_Timer()
    tmrResize.Enabled = False
    UserControl_Resize
End Sub

Private Sub tmrStateChange_Timer()
    tmrStateChange.Enabled = False
    RaiseEvent StateChanged(pState)
End Sub

Private Sub UserControl_Initialize()
    pBorderStyle = ruBorderStyleSunken
    pButtonPosition = ruButtonRight
    pHeaderVisible = True
    pHeaderHeight = 20 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_InitProperties()
    chkHeader.Caption = Ambient.DisplayName
    tmrResize.Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        picBack.BackColor = .ReadProperty("BackColor", vbButtonFace)
        pBorderStyle = .ReadProperty("BorderStyle", ruBorderStyleSunken)
        Me.Enabled = .ReadProperty("Enabled", True)
        chkHeader.BackColor = .ReadProperty("HeaderBackColor", vbButtonFace)
        pButtonPosition = .ReadProperty("ButtonPosition", ruButtonRight)
        chkHeader.Caption = .ReadProperty("HeaderCaption", Ambient.DisplayName)
        Set chkHeader.Font = .ReadProperty("HeaderFont", Ambient.Font)
        chkHeader.ForeColor = .ReadProperty("HeaderForeColor", vbWindowText)
        pHeaderHeight = .ReadProperty("HeaderHeight", chkButton.Height)
        pHeaderVisible = .ReadProperty("HeaderVisible", True)
        Set picBack.Picture = .ReadProperty("Picture", Nothing)
        pState = .ReadProperty("State", False)
    End With
    tmrResize.Enabled = True
End Sub

Private Sub UserControl_Resize()
    Dim nBitMap As RollUpStateBitmapConstants
    
    With UserControl
        picBack.Visible = False
        Select Case pState
            Case ruRollUp
                nBitMap = rusUp
                .BackStyle = 0
                Select Case pBorderStyle
                    Case ruBorderStyleNone
                        With picBack
                            .BorderStyle = 0
                            .Height = pHeaderHeight
                        End With
                    Case ruBorderStyleSunken
                        With picBack
                            .BorderStyle = 1
                            .Height = .Height - .ScaleHeight + pHeaderHeight
                        End With
                End Select
            Case ruRollDown
                nBitMap = rusDown
                .BackStyle = 1
                picBack.BorderStyle = pBorderStyle
                picBack.Height = .ScaleHeight
        End Select
        With chkButton
            Select Case pButtonPosition
                Case ruButtonRight
                    .Move picBack.ScaleWidth - pHeaderHeight, 0, pHeaderHeight, pHeaderHeight
                    chkHeader.Move 0, 0, picBack.ScaleWidth - pHeaderHeight, pHeaderHeight
                Case ruButtonLeft
                    .Move 0, 0, pHeaderHeight, pHeaderHeight
                    chkHeader.Move pHeaderHeight, 0, picBack.ScaleWidth - pHeaderHeight, pHeaderHeight
            End Select
            Set .Picture = LoadResPicture(nBitMap, vbResBitmap)
        End With
        chkHeader.Visible = pHeaderVisible
        picBack.Visible = True
        If Not mNoResizeEvent Then
            RaiseEvent Resize
        End If
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", picBack.BackColor, vbButtonFace
        .WriteProperty "BorderStyle", pBorderStyle, ruBorderStyleSunken
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "HeaderBackColor", chkHeader.BackColor, vbButtonFace
        .WriteProperty "ButtonPosition", pButtonPosition, ruButtonRight
        .WriteProperty "HeaderCaption", chkHeader.Caption, Ambient.DisplayName
        .WriteProperty "HeaderFont", chkHeader.Font, Ambient.Font
        .WriteProperty "HeaderForeColor", chkHeader.ForeColor, vbWindowText
        .WriteProperty "HeaderHeight", pHeaderHeight, chkButton.Height
        .WriteProperty "HeaderVisible", pHeaderVisible, True
        .WriteProperty "Picture", picBack.Picture, Nothing
        .WriteProperty "State", pState, False
    End With
End Sub

