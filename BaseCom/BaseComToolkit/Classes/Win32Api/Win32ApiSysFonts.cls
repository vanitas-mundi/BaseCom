VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cSysInfo    As New Win32ApiSysInfo
Private cSysReg     As New Win32ApiRegistry
Private cSysWnd     As New Win32ApiSysWindow
Private strFonts()  As String    'verfügbare Fonts
Private objControl  As Object    'Steuerelement zur Manipulation von Texten
Private cSysPrinter As New Win32ApiSysPrinter
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Const DEFAULT_CHARSET = 1

'Struktur zum Beschreiben eines logischen Fonts
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * LF_FACESIZE
End Type

'Texteigenschaften für GetTextMetrics
Private Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type
    
Private Type ENUMLOGFONTEX
  elfLogFont As LOGFONT
  elfFullName(LF_FULLFACESIZE) As Byte
  elfStyle(LF_FACESIZE) As Byte
  elfScript(LF_FACESIZE) As Byte
End Type

'Verwendete APIs
Private Declare Function GetTextFace _
                Lib "gdi32" _
                Alias "GetTextFaceA" (ByVal hDC As Long, _
                                      ByVal nCount As Long, _
                                      ByVal lpFaceName As String) As Long
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextMetrics _
                Lib "gdi32" _
                Alias "GetTextMetricsA" (ByVal hDC As Long, _
                                         lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nMapMode As Long) As Long
Private Declare Function CreateFontIndirect _
                Lib "gdi32" _
                Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EnumFontFamiliesEx _
                Lib "gdi32" _
                Alias "EnumFontFamiliesExA" (ByVal hDC As Long, _
                                             lpLogFont As LOGFONT, _
                                             ByVal lpEnumFontProc As Long, _
                                             ByVal lParam As Long, _
                                             ByVal dw As Long) As Long

Private Sub Class_Initialize()
'  ReDim strFonts(0)
'  Call EnumFonts
End Sub

Private Sub EnumFonts()
'  '====================================================================================
'  '2000-04-14, Klaus Prinz
'  '   EnumFontFamiliesEx für System und Default Printer aufrufen
'  '====================================================================================
'  Dim hDC As Long                 'Device Context des Desktops
'  Dim LF  As LOGFONT               'LOGFONT-Struktur
'  On Error GoTo errHandler
'  'LOGFONT parametrieren
'  LF.lfCharSet = DEFAULT_CHARSET
'  'Device Context für System ermitteln und Fonts enumerieren
'  hDC = cSysWnd.SystemHDC
'  EnumFontFamiliesEx hDC, LF, AddressOf EnumFontCallBack, 0&, 0&
'  'Device Context für Default-Printer ermitteln und Fonts enumerieren
'  hDC = cSysPrinter.hDC
'  EnumFontFamiliesEx hDC, LF, AddressOf EnumFontCallBack, 0&, 0&
'  Exit Sub
'
'errHandler:
'  Err.Raise Err.number, "clsSysFonts:EnumFonts"
  
  Err.Raise ErrorNumbers.IsObsolete, "Win32ApiSysFonts.EnumFonts", "NotImplemented"
End Sub

Private Function GetDeviceFontName(ByVal hDC As Long) As String
  '====================================================================================
  '2000-04-14, Klaus Prinz
  '   Rückgabe des Namens des angeforderten SystemFonts
  '------------------------------------------------------------------------------------
  'Argumente:
  '   hDC:    'Handle des Device Context
  '====================================================================================
  Dim lpFaceName As String * 100  'String für Rückgabe aus GetFaceText
  Dim nSize      As Long               'Nettolänge des Fonts
  On Error GoTo errHandler
  'Font-Namen ermitteln
  nSize = GetTextFace(hDC, 100, lpFaceName)
  'Font auf evt. Mapping prüfen ...
  GetDeviceFontName = GetMappedFont(Left(lpFaceName, nSize))
  Exit Function

errHandler:
  Err.Raise Err.number, "clsSysFonts:GetDeviceFontName"
End Function

Private Function GetMappedFont(ByVal strFont As String) As String
  '====================================================================================
  '2001-03-13, Klaus Prinz
  '   Rückgabe des gemappedten Fonts, sofern vorhanden
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strFont:    'Name des zu suchenden Fonts
  '====================================================================================
  Const strMapKey As String = "Software\Microsoft\Windows NT\CurrentVersion\FontSubstitutes"
  On Error GoTo errHandler
  GetMappedFont = strFont

  If cSysReg.HasKey(enumHKEY_LOCAL_MACHINE, strMapKey, strFont) Then
    strFont = cSysReg.GetRegistryValue(enumHKEY_LOCAL_MACHINE, strMapKey, strFont)
  Else
    GetMappedFont = strFont
  End If

  Exit Function

errHandler:
  Err.Raise Err.number, "clsSysFonts:GetMappedFont"
End Function

Private Function GetDeviceFontSize(ByVal hDC As Long) As Single
  '====================================================================================
  '2000-04-14, Klaus Prinz
  '   Rückgabe der Größe des SytemFonts in Punkt
  '------------------------------------------------------------------------------------
  'Argumente:
  '   hDC:    'Objekt-Handle
  '====================================================================================
  Dim udtTM          As TEXTMETRIC         'TextMetric-Type
  Dim lngPrevMapMode As Long      'vorangehender MapMode
  Const MM_TWIPS = 6
  On Error GoTo errHandler
  lngPrevMapMode = SetMapMode(hDC, MM_TWIPS)
  GetTextMetrics hDC, udtTM
  lngPrevMapMode = SetMapMode(hDC, lngPrevMapMode)
  GetDeviceFontSize = udtTM.tmHeight / 20
  Exit Function

errHandler:
  Err.Raise Err.number, "clsSysFonts:GetDeviceFontSize"
End Function

'Dieser Font wird von Windows in Menüs und Dialogen verwendet.
Public Property Get DefaultGUIFontName() As String
  DefaultGUIFontName = GetDeviceFontName(cSysWnd.SystemHDC)
End Property

Public Property Get DefaultGUIFontSize() As Single
  DefaultGUIFontSize = GetDeviceFontSize(cSysWnd.SystemHDC)
End Property

Public Property Get FontCount() As Long
  FontCount = UBound(strFonts) + 1
End Property

Public Property Get FontName(ByVal Index As Long) As String
  FontName = strFonts(Index - 1)
End Property

Public Sub AddFont(ByVal strFontFace As String)
  '====================================================================================
  '2000-03-05, Klaus Prinz
  '   den übergebenen Font in das Array strFonts einsortieren
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strFontFace:    'Fontname
  '====================================================================================
  Dim iFont         As Long               'Zeiger in strFonts
  Dim lpLFX         As ENUMLOGFONTEX
  'Dim strFontFace As String
  Dim strFontString As String     'Array-Inhalt für InStr-Funktion
  On Error GoTo errHandler

  'strFontFace = ClearAPIString(StrConv(lpLFX.elfFullName, vbUnicode))
  If UBound(strFonts) = 0 And LenB(strFonts(0)) = 0 Then
    '1. Font
    strFonts(0) = strFontFace
  Else
    'ab 2. Font
    strFontString = ";" & Join(strFonts, ";") & ";"

    'Abbruch, wenn Font bereits existiert
    If InStr(1, strFontString, ";" & strFontFace & ";") > 0 Then Exit Sub
    'Array vergrößern
    ReDim Preserve strFonts(UBound(strFonts) + 1)

    'Fonts zur Einsortierung durchlaufen
    For iFont = UBound(strFonts) - 1 To 0 Step -1

      If strFonts(iFont) > strFontFace Then
        'Font nach oben schieben
        strFonts(iFont + 1) = strFonts(iFont)

        'prüfen, ob erster Font erreicht wurde
        If iFont = 0 Then
          'neuen Font reinschreiben
          strFonts(iFont) = strFontFace
          Exit For
        End If

      Else
        'neuen Font reinschreiben
        strFonts(iFont + 1) = strFontFace
        Exit For
      End If

    Next

  End If

  Exit Sub

errHandler:
  Err.Raise Err.number, "clsSysFonts:AddFont"
End Sub

'**************************************************************************************************
'Rotate Font
'**************************************************************************************************
Public Sub WriteVerticalText(aPictureBox As Object, _
                             ByVal strFontName As String, _
                             ByVal sngFontSize As Single, _
                             ByVal strText As String)
  '====================================================================================
  '2000-03-05, Klaus Prinz
  '   den übergebenen Font senkrecht in die PictureBox schreiben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   aPictureBox:         Picture-Steuerelement
  '   strFontName:    Fontname
  '   sngFontSize:    Fontgröße
  '   strText:        Text
  '====================================================================================
  Dim udtFont     As LOGFONT          'LOGFONT-Struktur
  Dim lngPrevFont As Long         'vorheriger Font
  Dim lngPrevX    As Long            'vorherige CurrentX-Position
  Dim lngPrevY    As Long            'vorherige CurrentY-Position
  Dim hFont       As Long               'Handle auf neuen Font
  Dim lngReturn   As Long           'Rückgabe aus Funktion
  On Error GoTo errHandler
  'Hintergrund löschen
  aPictureBox.Picture = LoadPicture()
  'aktuelle Koordinaten sichern
  lngPrevX = aPictureBox.currentX
  lngPrevY = aPictureBox.CurrentY
  'LOGFONT parametrieren
  udtFont.lfEscapement = 900
  udtFont.lfFaceName = strFontName & Chr(0)
  udtFont.lfHeight = (sngFontSize * -20) / 15 '15 = Twips per Pixel
  'logischen Font erzeugen und für Control auswählen
  hFont = CreateFontIndirect(udtFont)
  lngPrevFont = SelectObject(aPictureBox.hDC, hFont)
  'neue Koordinaten einstellen und Text ausgeben
  aPictureBox.currentX = 0
  aPictureBox.CurrentY = aPictureBox.ScaleHeight - 60
  aPictureBox.Print strText
  'vorherigen Font wieder einstellen und Font-Handle löschen
  lngReturn = SelectObject(aPictureBox.hDC, lngPrevFont)
  lngReturn = DeleteObject(hFont)
  'vorherige Koordinaten einstellen
  aPictureBox.currentX = lngPrevX
  aPictureBox.CurrentY = lngPrevY
  Exit Sub

errHandler:
  lngReturn = DeleteObject(hFont)
  Err.Raise Err.number, "clsSysFonts:WriteVerticalText"
End Sub

