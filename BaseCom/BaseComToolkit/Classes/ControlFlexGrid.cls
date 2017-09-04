VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

'--------------------------------------------------------------------------------
'    Component  : ControlFlexGrid
'    Project    : ToolKits
'
'    Description: Stellt FlexGrid-Erweiterungen zur Verfügung.
'
'    Modified   :
'--------------------------------------------------------------------------------

'---------------------- Eigenschaften der Klasse --------------------------------
Private Const WM_USER = &H400
Private Const VP_FORMATRANGE = WM_USER + 125
Private Const VP_YESIDO = 456654

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Type RECT
  Left    As Long
  Top     As Long
  Right   As Long
  Bottom  As Long
End Type

Private Type TFormatRange
  hDC         As Long
  hdcTarget   As Long
  rc          As RECT
  rcPage      As RECT
End Type

Public Enum ePrinterOrientation
  poPortrait = 1
  poLandscape = 2
End Enum

Public Enum ListViewFindInConstants
  lfItems = 1
  lfSubItems = 2
  lfBoth = 3
End Enum



'---------------------- Konstruktor der Klasse ----------------------------------



'---------------------- Zugriffsmethoden der Klasse -----------------------------



'---------------------- Ereignismethoden der Klasse -----------------------------



'---------------------- Private Methoden der Klasse -----------------------------



'---------------------- Öffentliche Methoden der Klasse -------------------------
'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       ResizeColumns
' Description:       Resize FlexGrid Paßt die Spalten eines FlexGrids an den Inhalt an.
'                    Um die Funktion TextWidth benutzen zu können muss zudem
'                    ein Formular übergeben werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-11:54:40
'
' Parameters :       aForm (Object)
'                    flexGrid (Object)
'                    intCol (Integer = -1)
'--------------------------------------------------------------------------------
Public Sub ResizeColumns(ByVal aForm As Object, _
                         ByRef flexGrid As Object, _
                         Optional ByVal intCol As Integer = -1)
                          
  Dim r           As Integer
  Dim c           As Integer
  Dim Max         As Long
  Dim X           As Long
  Dim intStartCol As Integer
  Dim intEndCol   As Integer

  With flexGrid
    aForm.Font.name = .Font.name
    aForm.Font.Size = .Font.Size
    aForm.Font.Bold = .Font.Bold
    aForm.Font.Italic = .Font.Italic
    aForm.Font.Strikethrough = .Font.Strikethrough
    aForm.Font.Underline = .Font.Underline

    '{Alle Cols anpassen}
    If intCol = -1 Then
      intStartCol = 0
      intEndCol = .cols - 1
    Else '{gezielt eine Spalte anpassen}
      intStartCol = intCol
      intEndCol = intCol
    End If

    For c = intStartCol To intEndCol '.Cols - 1
      For r = 0 To .rows - 1
        X = aForm.TextWidth(.TextMatrix(r, c))

        If X > Max Then Max = X
      Next r

      .ColWidth(c) = Max + 150
      Max = 0
    Next c

  End With
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       PrintData
' Description:       Druckt den sichtbaren Inhalt eines FlexGrids aus.
'                    Spalten welche nicht gedruckt werden sollen werden auf Width = 0 gesetzt.
'                    Übergeben wird die Referenz zu einem FlexGrid; mit den Margin-Parametern
'                    können die Seitenränder in mm angegeben werden.
'                    Optional kann noch ein Berichtstitel und ein Druckdatum übergeben werden.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-11:56:04
'
' Parameters :       flexGrid (Object)
'                    LeftMargin (Single)
'                    TopMargin (Single)
'                    RightMargin (Single)
'                    BottomMargin (Single)
'                    PrintTitle (String = "")
'                    PrintDate (String = "")
'                    poOrientation (ePrinterOrientation = poPortrait)
'--------------------------------------------------------------------------------
Public Sub PrintData(flexGrid As Object, _
                     ByVal LeftMargin As Single, _
                     ByVal TopMargin As Single, _
                     ByVal RightMargin As Single, _
                     ByVal BottomMargin As Single, _
                     Optional ByVal PrintTitle As String = "", _
                     Optional ByVal PrintDate As String = "", _
                     Optional ByVal poOrientation As ePrinterOrientation = poPortrait)
                  
  Dim intCol        As Integer
  Dim intCol2       As Integer
  Dim intStartCol   As Integer
  Dim intEndCol     As Integer
  Dim tRange        As TFormatRange
  Dim blnReturn     As Long
  Dim alngCopyCW()  As Long
  Dim intPageNumber As Integer
  Dim X             As Long

  If flexGrid.rows = 0 Then Exit Sub
  If flexGrid.cols = 0 Then Exit Sub
  '{Seiten-Ausrichtung festlegen}
  printer.Orientation = poOrientation
  '{Selektierte Spalten ermitteln}
  intCol = flexGrid.col
  intCol2 = flexGrid.ColSel

  '{Ende Selektierte Spalten ermitteln}
  '{Spalten für den Druck bestimmen}
  Select Case True

    Case intCol = intCol2 '{Nur eine Spalte selektiert}

      If MsgBox("Soll die gesamte Ansicht gedruckt werden?", 36, "Ansicht drucken") = 6 Then '{Wenn gesamtes Grid gedruckt werden soll}
        intStartCol = 0
        intEndCol = flexGrid.cols - 1
      Else '{Wenn nur eine Spalte gedruckt werden soll}
        intStartCol = intCol2
        intEndCol = intCol
      End If

    Case intCol > intCol2 '{Es wurde nach links hin slektiert}
      intStartCol = intCol2
      intEndCol = intCol

    Case intCol2 > intCol '{Es wurde nach rechts hin slektiert}
      intStartCol = intCol
      intEndCol = intCol2
  End Select

  '{Ende Spalten für den Druck bestimmen}
  '{Selektierung für Ausdruck aufheben}
  flexGrid.ColSel = flexGrid.col
  flexGrid.RowSel = flexGrid.Row
  '{Ende Selektierung für Ausdruck aufheben}
  '{Erste Seitenzahl festlegen}
  intPageNumber = 1
  '{Feld für ColWidths dimensionieren}
  ReDim alngCopyCW(0 To flexGrid.cols - 1)
  '{Automatisches Neuzeichnen der Grids ausschalten}
  flexGrid.Redraw = False

  '{Alle Spalten die nicht gedruckt werden sollen ColWidth auf 0 setzen}
  '{Alten Wert im Array zwischenspeichern}
  For X = 0 To intStartCol - 1
    alngCopyCW(X) = flexGrid.ColWidth(X)
    flexGrid.ColWidth(X) = 0
  Next X

  For X = intEndCol + 1 To flexGrid.cols - 1
    alngCopyCW(X) = flexGrid.ColWidth(X)
    flexGrid.ColWidth(X) = 0
  Next X

  '{Ende Alle Spalten die nicht gedruckt werden sollen ColWidth auf 0 setzen}
  '{mit wParam <> 0 kann überprüft werden}
  '{ob das Control OPP unterstützt, wenn ja wird}
  '{456654 (VP_YESIDO) zurückgeliefert}
  blnReturn = SendMessage(flexGrid.hwnd, VP_FORMATRANGE, 1, 0)

  If blnReturn = VP_YESIDO Then
    '{Struktur mit Formatierungsinformationen füllen}
    printer.ScaleMode = vbPixels

    With tRange
      .hDC = printer.hDC
      '{Höhe und Breite einer Seite (in Pixel)}
      .rcPage.Right = printer.ScaleWidth
      .rcPage.Bottom = printer.ScaleHeight
      '{Lage und Abmessungen des Bereichs auf den}
      '{gedruckt werden soll (in Pixel)}
      .rc.Left = printer.ScaleX(LeftMargin, vbMillimeters)
      .rc.Top = printer.ScaleY(TopMargin, vbMillimeters)
      .rc.Right = .rcPage.Right - printer.ScaleX(RightMargin, vbMillimeters)
      .rc.Bottom = .rcPage.Bottom - printer.ScaleY(BottomMargin, vbMillimeters)
    End With

    '{Drucker initialisieren}
    printer.Print vbNullString

    '{Seite(n) drucken}
    Do
      printer.currentX = printer.ScaleX(LeftMargin, vbMillimeters)
      printer.CurrentY = printer.ScaleY(10, vbMillimeters)

      If PrintTitle <> "" Then
        printer.Print PrintTitle & " (Seite: " & intPageNumber & ")"
      End If

      printer.currentX = printer.ScaleX(LeftMargin, vbMillimeters)
      printer.CurrentY = printer.ScaleY(16, vbMillimeters)

      If PrintDate = "" Then
        printer.Print Format$(Now, "yyyy-mm-dd")
      Else
        printer.Print Format$(PrintDate, "yyyy-mm-dd")
      End If

      blnReturn = SendMessage(flexGrid.hwnd, VP_FORMATRANGE, 0, VarPtr(tRange))

      If blnReturn < 0 Then
        Exit Do
      Else
        intPageNumber = intPageNumber + 1
        printer.NewPage
      End If

    Loop

    printer.EndDoc
    'Reset
    blnReturn = SendMessage(flexGrid.hwnd, VP_FORMATRANGE, 0, 0)
  End If

  '{ColWidth derversteckten Spalten wieder rücksetzen}
  For X = 0 To intStartCol - 1
    flexGrid.ColWidth(X) = alngCopyCW(X)
  Next X

  For X = intEndCol + 1 To flexGrid.cols - 1
    flexGrid.ColWidth(X) = alngCopyCW(X)
  Next X

  '{ColWidth derversteckten Spalten wieder rücksetzen}
  '{Automatisches neuzeichnen einschalten}
  flexGrid.Redraw = True
  'PrintFlexGrid(MSFlexGrid1, 20, 25, 20, 20, "Titel " , date)
End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       CopyToClipBoard
' Description:       Kopiert die selektierten FlexGrid-Daten in die Zwischenablage.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-11:58:34
'
' Parameters :       flexGrid (Object)
'--------------------------------------------------------------------------------
Public Sub CopyToClipBoard(ByRef flexGrid As Object)
  Dim strView As String
  Dim X       As Long
  Dim Y       As Long
  Dim z1      As Long
  Dim z2      As Long
  Dim s1      As Long
  Dim s2      As Long

  With flexGrid

    If .Row > .RowSel Then
      z1 = .RowSel
      z2 = .Row
    Else
      z1 = .Row
      z2 = .RowSel
    End If

    If .col > .ColSel Then
      s1 = .ColSel
      s2 = .col
    Else
      s1 = .col
      s2 = .ColSel
    End If

    For Y = z1 To z2
      For X = s1 To s2
        strView = strView & .TextMatrix(Y, X) & vbTab '";"
      Next X

      strView = Mid$(strView, 1, Len(strView) - 1)
      strView = strView & vbCrLf
    Next Y

    Clipboard.Clear
    Clipboard.SetText strView
  End With

End Sub

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       GetData
' Description:       Liefert die FLexGrid-Daten als String.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       19.01.2015-11:59:54
'
' Parameters :       flexGrid (Object)
'--------------------------------------------------------------------------------
Public Function GetData(ByRef flexGrid As Object) As String
  Dim c       As Long
  Dim r       As Long
  Dim strRow  As String
  Dim strData As String
  Dim strCell As String

  With flexGrid

    For r = 0 To .rows - 1
      strRow = ""

      For c = 0 To .cols - 1
        strCell = Replace(.TextMatrix(r, c), Chr$(10), "")
        strCell = Replace(strCell, Chr$(13), " ")
        strRow = strRow & strCell & vbTab
      Next c

      strData = strData & strRow & vbCrLf
    Next r

  End With

  GetData = strData
  Exit Function
End Function

'--------------------------------------------------------------------------------
' Project    :       ToolKits
' Procedure  :       SelectAll
' Description:       Markiert das komplette FlexGrid.
' Created by :       Sascha Glinka
' Machine    :       VDI-EDV-0003
' Date-Time  :       11.03.2015-16:11:06
'
' Parameters :       flexGrid (Object)
'--------------------------------------------------------------------------------
Public Sub SelectAll(ByRef flexGrid As Object)
  
  With flexGrid
    .col = 0            ' start selection in this column
    .Row = 0            ' start selection in this row
    .ColSel = .cols - 1 ' end selection in this column
    .RowSel = .rows - 1 ' end selection in this row
  End With
End Sub
