VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------------------------------
'    Component  : ControlForm
'    Project    : ToolKits
'
'    Description: Stellt Formular-Erweiterungen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------


'---------------------- Eigenschaften der Klasse --------------------------------
Private WHook&

Private Const GWL_EXSTYLE = -20

Private Const SWP_DRAWFRAME = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

Private Const WN_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetParent _
                Lib "user32" (ByVal hWndChild As Long, _
                              ByVal hWndNewParent As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook&) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Private mInitBoxStyle As Long
Private mInitLeft     As Integer
Private mInitTop      As Integer
Private mInitWidth    As Integer
Private mInitHeight   As Integer



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------
Private Sub SetControlStyle(ByRef aForm As Object, style As Long, aControl As Object)
  
  If style& Then
    SetWindowLong aControl.hWnd, GWL_STYLE, style
    SetWindowPos aControl.hWnd, aForm.hWnd, 0, 0, 0, 0, SWP_FLAGS
  End If

End Sub



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       AddControlToFormTitle
' Description:       Verlagert einen CommandButton in die Titelleiste eines Formulars.
'                    Spätestens beim Entladen der Form muss RemoveControlFromFormTitle
'                    aufgerufen werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:02:00
'
' Parameters :       aForm (Object)
'                    aControl (Object)
'                    lngPos (Long)
'--------------------------------------------------------------------------------
Public Sub AddControlToFormTitle(ByRef aForm As Object, _
                                 ByRef aControl As Object, _
                                 ByVal lngPos As Long)
  aForm.Hide
  WHook = ControlFormHelperModule.SetWindowsHookExModule(aForm, aControl, lngPos)
  SetWindowLong aControl.hWnd, GWL_EXSTYLE, &H80
  SetParent aControl.hWnd, GetParent(aForm.hWnd)
  aForm.Show
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       RemoveControlFromFormTitle
' Description:      Entfernt eine Verlagerung eines CommandButton aus der Titelleiste eines Formulars.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:02:14
'
' Parameters :       aForm (Object)
'                    aControl (Object)
'--------------------------------------------------------------------------------
Public Sub RemoveControlFromFormTitle(ByRef aForm As Object, ByRef aControl As Object)
  UnhookWindowsHookEx WHook
  SetParent aControl.hWnd, aForm.hWnd
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ActivateMoveControl
' Description:       Aktiviert automatische Control-Verschiebung innerhalb eines Formulars.
'                    Es muss zwingend DeActivateMoveControl aufgerufen werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:02:37
'
' Parameters :       aForm (Object)
'                    aControl (Object)
'--------------------------------------------------------------------------------
Public Sub ActivateMoveControl(ByRef aForm As Object, ByRef aControl As Object)
  
  mInitBoxStyle = GetWindowLong(aControl.hWnd, GWL_STYLE)
  mInitLeft = aControl.Left
  mInitTop = aControl.Top
  mInitWidth = aControl.Width
  mInitHeight = aControl.Height
  
  SetControlStyle aForm, mInitBoxStyle, aControl
    
  Dim style As Long
  style& = GetWindowLong(aControl.hWnd, GWL_STYLE)
  style& = style& Or WS_THICKFRAME
  SetControlStyle aForm, style&, aControl
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       DeActivateMoveControl
' Description:       Deaktiviert automatische Control-Verschiebung innerhalb eines Formulars.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:02:47
'
' Parameters :       aForm (Object)
'                    aControl (Object)
'--------------------------------------------------------------------------------
Public Sub DeActivateMoveControl(ByRef aForm As Object, ByRef aControl As Object)
  SetControlStyle aForm, mInitBoxStyle, aControl
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       MoveControlInForm
' Description:       Verschiebt das angegebne Control auf dem Formular.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-12:02:52
'
' Parameters :       aControl (Object)
'--------------------------------------------------------------------------------
Public Sub MoveControlInForm(ByRef aControl As Object)
  ReleaseCapture
  SendMessage aControl.hWnd, WN_NCLBUTTONDOWN, HTCAPTION, 0&
  aControl.Refresh
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       AddControl
' Description:       Lädt dynamisch ein Control und zeigt es auf dem angegebenen
'                    Formular an.
'                    Beispiel:
'                    Private WithEvents mDynControl As VBControlExtender
'
'                    Private Sub mDynControl_ObjectEvent(Info As EventInfo)
'                      If Info.name = "Click" Then
'                        MsgBox "Click"
'                      End If
'                    End Sub
'
'                    ...
'                    Set mDynControl = AddControl(Me, "MSFlexGridLib.MSFlexGrid", "flexGrid1")
'
' Created by :       Project Administrator
' Machine    :       VDI-EDV-0003
' Date-Time  :       20.01.2015-14:01:46
'
' Parameters :       aForm (Object)
'                    progId (String)
'                    controlName (String)
'--------------------------------------------------------------------------------
Public Function AddControl(ByVal aForm As Object, ByVal progId As String, ByVal controlName As String) As Object

  ' Add the license for the treeview to the license collection.
  ' If the license is already in the collection you will get
  ' the run-time error number 732.
  Licenses.Add progId
  
  ' Dynamically add a TreeView control to the form.
  ' If you want the control to be added to a different
  ' container such as a Frame or PictureBox, you use the third
  ' parameter of the Controls.Add to specify the container.
  
  Set AddControl = aForm.Controls.Add(progId, controlName, aForm)
  
  ' set the location and size of the control.
  AddControl.Move 1, 1, 2500, 3500
  
  ' Make the control visible.
  AddControl.Visible = True
  AddControl.ZOrder
End Function


Public Sub GetWindowPos(ByVal appName As String, ByVal aForm As Object)

  With aForm
    .Top = GetSetting(appName, .name, "Top", .Top)
    .Left = GetSetting(appName, .name, "Left", .Left)
    .Height = GetSetting(appName, .name, "Height", .Height)
    .Width = GetSetting(appName, .name, "Width", .Width)
    .WindowState = GetSetting(appName, .name, "WindowState", .WindowState)
  End With
End Sub

Public Sub PutWindowPos(ByVal appName As String, ByVal aForm As Object)
  With aForm
    SaveSetting appName, .name, "Top", .Top
    SaveSetting appName, .name, "Left", .Left
    SaveSetting appName, .name, "Height", .Height
    SaveSetting appName, .name, "Width", .Width
    SaveSetting appName, .name, "WindowState", .WindowState
  End With
End Sub

