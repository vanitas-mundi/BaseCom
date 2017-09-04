VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'© 2000 Klaus Prinz Software Consulting
'       Schnittstelle zu comdlg32-Dialogen und SHBrowseForFolder
'==================================================================================================
'Methoden:
'--------------------------------------------------------------------------------------------------
'   GetColor
'       Farbdialog erzeugen und Auswahl zurückgeben
'   GetOpenFileName
'       GetOpenFileName-Dialog erzeugen und Auswahl zurückgeben
'   GetSaveFileName
'       GetSaveFileName-Dialog erzeugen und Auswahl zurückgeben
'   GetPath
'       SHBrowseForFolder-Dialog erzeugen und Auswahl zurückgeben
'==================================================================================================
'Eigenschaften:
'--------------------------------------------------------------------------------------------------
'   CustomizedColor
'       Schreib- und Lesezugriff auf Array der benutzerdefinierten Farben
'==================================================================================================
'Private Methoden:
'--------------------------------------------------------------------------------------------------
'   ClearAPIString
'       Übergebenen String bis vor ersten Chr(0) abschneiden
'   ComposeFileFilter
'       Filterausdruck für OPENFILENAME.lpstrFilter zusammenstellen
'==================================================================================================

Private cSysReg           As New Win32ApiRegistry   'Instanz auf Registry-Klasse
Private lngCustColors(15) As Long       'Array für die 16 benutzerdefinierte Farben

'Enumeration zur Unterscheidung von GetOpen- und GetSaveFileName-Dialog
Private Enum enumSaveOpenFileName
  enumOpenFileName = 1
  enumSaveFileName = 2
End Enum

'Struktur für OpenFile und SaveFile
Private Type OpenFilename
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'Struktur für APIChooseColor
Private Type CHOOSECOLOR
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  RgbResult As Long
  lpCustColors As Long
  Flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'Konstanten für CHOOSECOLOR.flags
Private Const CC_ANYCOLOR = &H100
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CC_SOLIDCOLOR = &H80

'Konstanten für BROWSEINFO.ulFlags
Private Const BIF_RETURNONLYFSDIRS = 1  'nur Auswahl von Verzeichnissen ermöglichen
Private Const BIF_DONTGOBELOWDOMAIN = 2 'keine Domainenmitglieder zeigen

'Type zur Parametrierung der SHBrowseForFolder
Private Type BROWSEINFO
  hwndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As String
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

'Verwendete APIs
Private Declare Function APIChooseColor _
                Lib "comdlg32.dll" _
                Alias "ChooseColorA" (pCHOOSECOLOR As CHOOSECOLOR) As Long
Private Declare Function APIGetOpenFileName _
                Lib "comdlg32.dll" _
                Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
Private Declare Function APIGetSaveFileName _
                Lib "comdlg32.dll" _
                Alias "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList _
                Lib "shell32" (ByVal pidList As Long, _
                               ByVal lpBuffer As String) As Long

'**************************************************************************************************
'Datei Öffnen- und Speichern-Dialoge
'**************************************************************************************************
Public Function GetOpenFileName(ByVal initialPath As String, _
                                ByVal filename As String, _
                                ByVal FileTypes As String, _
                                Optional ByVal title As String, _
                                Optional ByVal hwndOwner As Long) As String
  GetOpenFileName = GetFileName(enumOpenFileName, initialPath, filename, FileTypes, title, hwndOwner)
End Function

Public Function GetSaveFileName(ByVal initialPath As String, _
                                ByVal filename As String, _
                                ByVal FileTypes As String, _
                                Optional ByVal title As String, _
                                Optional ByVal hwndOwner As Long) As String
  GetSaveFileName = GetFileName(enumSaveFileName, initialPath, filename, FileTypes, title, hwndOwner)
End Function

Private Function GetFileName(ByVal lngMode As enumSaveOpenFileName, _
                             ByVal strInitialPath As String, _
                             ByVal strFileName As String, _
                             ByVal strFileTypes As String, _
                             Optional ByVal strTitle As String, _
                             Optional ByVal hwndOwner As Long) As String
  '====================================================================================
  '2000-04-20, Klaus Prinz
  '   GetOpenFileName- oder GetSaveFileName-Dialog erzeugen und Auswahl zurückgeben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngMode:            Modus (Datei Öffnen oder Speichern)
  '   strInitialPath:     Startverzeichnis
  '   strFileName:        vorgeschlagener Dateiname
  '   strFileTypes:       Dateitypen, Semikolon-getrennt (z.B. "xls" oder "doc;txt")
  '   strTitle:           Dialogtitel, wenn "", dann wird "Öffnen" angezeigt
  '   hWndOwner:          Handle des Owner-Fensters
  '------------------------------------------------------------------------------------
  'Rückgabe:              vollständiger Name der ausgewählten Datei oder "" bei Abbruch
  '====================================================================================
  Dim OpenFile As OpenFilename    'Datenstruktur für APIGetOpenFileName
  On Error GoTo errHandler
  OpenFile.hwndOwner = hwndOwner
  OpenFile.lStructSize = Len(OpenFile)
  OpenFile.lpstrFilter = ComposeFileFilter(strFileTypes)
  OpenFile.nFilterIndex = 1
  OpenFile.lpstrFile = strFileName & String(260 - Len(strFileName), 0)
  OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
  OpenFile.lpstrFileTitle = OpenFile.lpstrFile
  OpenFile.nMaxFileTitle = OpenFile.nMaxFile
  OpenFile.lpstrInitialDir = strInitialPath

  If strTitle <> "" Then
    OpenFile.lpstrTitle = strTitle
  End If

  If lngMode = enumOpenFileName Then
    APIGetOpenFileName OpenFile
  Else
    APIGetSaveFileName OpenFile
  End If

  GetFileName = ClearAPIString(OpenFile.lpstrFile)
  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "cSysDialogs:GetOpenFileName"
  Err.Raise Err.number
End Function

Private Function ClearAPIString(ByVal strAPIString As String) As String
  '====================================================================================
  '2000-04-20, Klaus Prinz
  '   Übergebenen String bis vor ersten Chr(0) abschneiden
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strAPIString:   API-String
  '====================================================================================
  On Error GoTo errHandler

  If InStr(1, strAPIString, Chr(0)) > 0 Then
    ClearAPIString = Left(strAPIString, InStr(1, strAPIString, Chr(0)) - 1)
  Else
    ClearAPIString = strAPIString
  End If

  Exit Function

errHandler:
  Err.Raise Err.number, "cSysDialogs:ClearAPIString"
End Function

Private Function ComposeFileFilter(ByVal strExtensions As String) As String
  '====================================================================================
  '2000-04-20, Klaus Prinz
  '   Filterausdruck für OPENFILENAME.lpstrFilter zusammenstellen
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strExtension:   Dateierweiterungen, Semikolon-getrennt
  '====================================================================================
  Dim strExt() As String          'Array zur Aufnahme der übergebenen Dateierweiterungen
  Dim iExt     As Long                'Zeiger in strExt()
  On Error GoTo errHandler

  If strExtensions = "" Then
    ComposeFileFilter = ""
  Else
    strExt = Split(strExtensions, ";")

    For iExt = 0 To UBound(strExt)

      If InStr(1, strExt(iExt), ".") = 0 Then
        'nur Dateierweiterung, z. B. "xml"
        ComposeFileFilter = ComposeFileFilter & cSysReg.ExtensionToFileType(strExt(iExt)) & " (*." & strExt(iExt) & ")" & Chr(0) & "*." & strExt(iExt) & Chr(0)
      Else
        'Dateimaske, z. B. "Import*.xml"
        ComposeFileFilter = ComposeFileFilter & strExt(iExt) & Chr(0) & strExt(iExt) & Chr(0)
      End If

    Next

  End If

  If UBound(strExt) > 0 Then
    'mit einem weiteren NULL anschließen, wenn mehrere Einzelfilter
    ComposeFileFilter = ComposeFileFilter & Chr(0)
  End If

  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "cSysDialogs:ComposeFileFilter"
  Err.Raise Err.number
End Function

'**************************************************************************************************
'Verzeichnisdialog
'**************************************************************************************************
Public Function GetPath(Optional ByVal hwndOwner As Long) As String
  '====================================================================================
  '2000-02-24, Klaus Prinz
  '   Verzeichnis ermitteln und zurückgeben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   hWndOwner:   Handle des aufrufenden Fensters
  '------------------------------------------------------------------------------------
  'Rückgabe:       vollständiger Name des ausgewählten Verzeichnisse oder "" bei Abbruch
  '====================================================================================
  Dim lngIDList   As Long       'Zeiger auf das ausgesuchte Verzeichnis
  Dim strTitle    As String      'Titel des Dialogs
  Dim tBrowseInfo As BROWSEINFO
  On Error GoTo errHandler
  tBrowseInfo.hwndOwner = hwndOwner
  strTitle = "Bitte wählen Sie ein Verzeichnis aus"
  tBrowseInfo.lpszTitle = strTitle
  tBrowseInfo.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  lngIDList = SHBrowseForFolder(tBrowseInfo)

  If lngIDList <> 0 Then
    GetPath = Space(260)
    SHGetPathFromIDList lngIDList, GetPath
    GetPath = Left(GetPath, InStr(GetPath, vbNullChar) - 1)
  End If

  Exit Function

errHandler:
  Err.Raise Err.number + vbObjectError, "clsRegistry:GetPath"
End Function

'**************************************************************************************************
'Farbdialog
'**************************************************************************************************
Public Property Let CustomizedColor(ByVal Index As Long, ByVal NewColor As Long)

  If Index < 1 Or Index > 16 Then
    Err.Raise 9, "Der Index muss zwischen 1 und 16 liegen."
  End If

  lngCustColors(Index - 1) = NewColor
End Property

Public Property Get CustomizedColor(ByVal Index As Long) As Long

  If Index < 1 Or Index > 16 Then
    Err.Raise 9, "Der Index muss zwischen 1 und 16 liegen."
  End If

  CustomizedColor = lngCustColors(Index - 1)
End Property

Public Function GetColor(Optional InitColor As Long = -1, _
                         Optional ByVal hwndOwner As Long) As Long
  '====================================================================================
  '2000-02-24, Klaus Prinz
  '   Farbdialog anzeigen und ausgewählte Farbe zurückgeben
  '------------------------------------------------------------------------------------
  'Argumente:
  '   Initcolor:   Initalisierungsfarbe oder -1 für keine Initialisierung
  '   hWndOwner:      Handle des aufrufenden Fensters
  '------------------------------------------------------------------------------------
  'Rückgabe:          RGB-Farbe oder -1, wenn Dialog über Abbrechen geschlossen wurde
  '====================================================================================
  Dim tChooseColor As CHOOSECOLOR 'Datenstruktur für APIChooseColor
  Dim lngReturn    As Long           'Rückgabe aus Funktion
  On Error GoTo errHandler
  tChooseColor.lStructSize = Len(tChooseColor)
  tChooseColor.hwndOwner = hwndOwner
  tChooseColor.lpCustColors = VarPtr(lngCustColors(0))

  If InitColor = -1 Then
    tChooseColor.Flags = CC_SOLIDCOLOR + CC_FULLOPEN
  Else
    tChooseColor.Flags = CC_SOLIDCOLOR + CC_FULLOPEN + CC_RGBINIT
    tChooseColor.RgbResult = InitColor
  End If

  lngReturn = APIChooseColor(tChooseColor)

  If lngReturn = 0 Then
    GetColor = -1
  Else
    GetColor = tChooseColor.RgbResult
  End If

  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "cSysDialogs:GetColor"
  Err.Raise Err.number
End Function

