VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cSysInfo       As New Win32ApiSysInfo
Private cSysReg        As New Win32ApiRegistry
Private lngDC          As Long                   'Gerätekontext
Private nPrinters      As Long               'Anzahl der Drucker
Private strPrinter     As String            'aktueller Drucker
Private strPrinters()  As String         'Druckernamen, 0-basiert
Private strPaperBins() As String        'Papierzufuhrnamen, 0-basiert
Private intPaperBins() As Integer       'Papierzufuhrkonstanten, 0-basiert
'Konstanten für SendMessage
Private Const HWND_BROADCAST = &HFFFF
Private Const WM_WININICHANGE = &H1A
'Richtungskonstanten für DocumentProperties
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
' Constants for DMFIELDS
Private Const DM_ORIENTATION = &H1&         'Orientation
Private Const DM_PAPERSIZE = &H2&           'Papiergröße (siehe DMPAPER*)
Private Const DM_PAPERLENGTH = &H4&         'Papierlänge in Zehntel mm
Private Const DM_PAPERWIDTH = &H8&          'Papierbreite in Zehntel mm
Private Const DM_SCALE = &H10&              'Skalierung bezogen auf 100
Private Const DM_COPIES = &H100&            'Anzahl der Kopien
Private Const DM_DEFAULTSOURCE = &H200&     'Papierzufuhr
Private Const DM_PRINTQUALITY = &H400&      'Anzahl der DPI oder eine der DMRES*-Konstanten
Private Const DM_COLOR = &H800&             'DMCOLOR_MONOCHROME oder DMCOLOR_COLOR
Private Const DM_DUPLEX = &H1000&           'beidseitiges Drucken
Private Const DM_YRESOLUTION = &H2000&      'Anzahl der DPI in Y-Richtung
Private Const DM_TTOPTION = &H4000&         'Handling von TrueType-Schriften
Private Const DM_COLLATE = &H8000&          'Download bei Mehrfachkopien
Private Const DM_FORMNAME = &H1000000       'bei NT Schriftfamilienunterstützung
Private Const DM_ICMMETHOD = &H2000000      'Image-Color-Management Unterstützung (9x und 2000)
Private Const DM_ICMINTENT = &H4000000      'Art der ICM-Unterstützung (9x und 2000)
Private Const DM_MEDIATYPE = &H8000000      'Art des Druckmediums (9x und 2000)
Private Const DM_DITHERTYPE = &H10000000    'Art des Dithering (9x und 2000)
'Konstanten für Auflösung
Private Const DMRES_DRAFT = (-1)
Private Const DMRES_HIGH = (-4)
Private Const DMRES_LOW = (-2)
Private Const DMRES_MEDIUM = (-3)
'Farbkonstanten
Private Const DMCOLOR_MONOCHROME = 1
Private Const DMCOLOR_COLOR = 2
'Duplexdruck-Konstanten
Private Const DMDUP_SIMPLEX = 1     'nur eineitiger Druck
Private Const DMDUP_VERTICAL = 2    'beidseitig, Bindung an Schmalseite
Private Const DMDUP_HORIZONTAL = 3  'beidseitig, Bindung an Breitseite
'Download bei Mehfachkopien
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1
'TrueType-Konstanten
Private Const DMTT_BITMAP = 1           'TrueType-Fonts als Grafik drucken
Private Const DMTT_DOWNLOAD = 2         'TrueType-Fonts als Soft Fonts in Drucker laden
Private Const DMTT_SUBDEV = 3           'Geräte-Fonts für TrueType-Fonts ersetzen
Private Const DMTT_DOWNLOAD_OUTLINE = 4 'TrueType-Fonts als Outline Soft Fonts laden
Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12
'Paperzufuhrkonstanten
Private Const DMBIN_ONLYONE = 1         'nur ein Behälter vorhanden
Private Const DMBIN_UPPER = 1           'oberer Behälter
Private Const DMBIN_LOWER = 2           'unterer Behälter
Private Const DMBIN_MIDDLE = 3          'mittlerer Behälter
Private Const DMBIN_MANUAL = 4          'manuelle Blattzufuhr
Private Const DMBIN_ENVELOPE = 5        'Behälter für Umschläge
Private Const DMBIN_ENVMANUAL = 6       'manuelle Umschlagzufuhr
Private Const DMBIN_AUTO = 7            'automatische Blattzufuhr
Private Const DMBIN_TRACTOR = 8         'Traktoreinzug
Private Const DMBIN_SMALLFMT = 9        'Behälter für kleine Formate
Private Const DMBIN_LARGEFMT = 10       'Behälter für große Formate
Private Const DMBIN_LARGECAPACITY = 11  'Behälter für große Kapazitäten
Private Const DMBIN_CASSETTE = 14       'Papierkassette
Private Const DMBIN_USER = 256          'Grenze für treiberspezifische Werte

Public Enum enumPaperBin

  enumPaperBinUpper = 1
  enumPaperBinLower = 2
  enumPaperBinMiddle = 3
  enumPaperBinManual = 4
  enumPaperBinEnvelope = 5
  enumPaperBinEnvelopeManual = 6
  enumPaperBinAutomatic = 7

End Enum

'Papierausrichtung
Private Const DMORIENT_PORTRAIT = 1     'Hochformat
Private Const DMORIENT_LANDSCAPE = 2    'Querformat

Public Enum enumOrientation

  enumOrientationPortrait = DMORIENT_PORTRAIT
  enumOrientationLandscape = DMORIENT_LANDSCAPE

End Enum

'Papiergrößen, zöllig
Private Const DMPAPER_LETTER = 1                'Letter,                 8 1/2  x   11     in
Private Const DMPAPER_LEGAL = 5                 'Legal                   8 1/2  x   14     in
Private Const DMPAPER_10X14 = 16                '                       10      x   14     in
Private Const DMPAPER_11X17 = 17                '                       11      x   17     in
Private Const DMPAPER_EXECUTIVE = 7             'Executive               7 1/4  x   10 1/2 in
Private Const DMPAPER_FANFOLD_US = 39           'US Std Fanfold         14 7/8  x   11     in
Private Const DMPAPER_FANFOLD_STD_GERMAN = 40   'German Std Fanfold      8 1/2  x   12     in
Private Const DMPAPER_FANFOLD_LGL_GERMAN = 41   'German Legal Fanfold    8 1/2  x   13     in
Private Const DMPAPER_FOLIO = 14                'Folio                   8 1/2  x   13     in
Private Const DMPAPER_LEDGER = 4                'Ledger                 17      x   11     in
Private Const DMPAPER_LETTERSMALL = 2           'Letter Small            8 1/2  x   11     in
Private Const DMPAPER_NOTE = 18                 'Note                    8 1/2  x   11     in
Private Const DMPAPER_STATEMENT = 6             'Statement               5 1/2  x    8 1/2 in
Private Const DMPAPER_TABLOID = 3               'Tabloid                11      x   17     in
Private Const DMPAPER_CSHEET = 24               'C size sheet
Private Const DMPAPER_DSHEET = 25               'D size sheet
Private Const DMPAPER_ENV_9 = 19                'Envelope #9             3 7/8  x    8 7/8 in
Private Const DMPAPER_ENV_10 = 20               'Envelope #10            4 1/8  x    9 1/2 in
Private Const DMPAPER_ENV_11 = 21               'Envelope #11            4 1/2  x   10 3/8 in
Private Const DMPAPER_ENV_12 = 22               'Envelope #12            4 1/2  x   11, (auch 4, 4.25 und 4.75 x 11)
Private Const DMPAPER_ENV_14 = 23               'Envelope #14            5      x   11 1/2 in
Private Const DMPAPER_ENV_MONARCH = 37          'Envelope Monarch        3 7/8  x    7 1/2 in
Private Const DMPAPER_ENV_PERSONAL = 38         '6 3/4 Envelope          3 5/8  x    6 1/2 in
'Papiergrößen, metrisch
Private Const DMPAPER_A3 = 8                    'A3             297 x 420 mm
Private Const DMPAPER_A4 = 9                    'A4             210 x 297 mm
Private Const DMPAPER_A5 = 11                   'A5             148 x 210 mm
Private Const DMPAPER_B4 = 12                   'B4             250 x 354 mm
Private Const DMPAPER_B5 = 13                   'B5             182 x 257 mm
Private Const DMPAPER_QUARTO = 15               'Quarto         215 x 275 mm
Private Const DMPAPER_ENV_C3 = 29               'Envelope C3    324 x 458 mm
Private Const DMPAPER_ENV_C4 = 30               'Envelope C4    229 x 324 mm
Private Const DMPAPER_ENV_C5 = 28               'Envelope C5    162 x 229 mm
Private Const DMPAPER_ENV_C6 = 31               'Envelope C6    114 x 162 mm
Private Const DMPAPER_ENV_C65 = 32              'Envelope C65   114 x 229 mm
Private Const DMPAPER_ENV_B4 = 33               'Envelope B4    250 x 353 mm
Private Const DMPAPER_ENV_B5 = 34               'Envelope B5    176 x 250 mm
Private Const DMPAPER_ENV_B6 = 35               'Envelope B6    176 x 125 mm
Private Const DMPAPER_ENV_DL = 27               'Envelope DL    110 x 220 mm
Private Const DMPAPER_ENV_ITALY = 36            'Envelope       110 x 230 mm
Private Const PRINTER_ENUM_DEFAULT = 1
Private Const PRINTER_ENUM_LOCAL = 2
Private Const PRINTER_ENUM_CONNECTIONS = 4
Private Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Private Const PRINTER_ATTRIBUTE_LOCAL = &H40
Private Const PRINTER_ATTRIBUTE_NETWORK = &H10
'Konstanten für DesiredAccess member of PRINTER_DEFAULTS
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Type PRINTER_DEFAULTS

  pDatatype As String
  pDevMode As Long
  DesiredAccess As Long   'Recht

End Type

'Struktur zur Ermittlung der Drucker unter NT
Private Type PRINTER_INFO_4

  pPrinterName As String
  pServerName As String
  Attributes As Long

End Type

'Struktur zur Ermittlung der Drucker unter NT
Private Type PRINTER_INFO_5

  pPrinterName As String
  pPortName As String
  Attributes As Long
  DeviceNotSelectedTimeout As Long
  TransmissionRetryTimeout As Long

End Type

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

'DEVMODE-Struktur
Private Type DEVMODE

  dmDeviceName As String * CCHDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCHFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Long
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long

End Type

Private Type ACL

  AclRevision As Byte
  Sbz1 As Byte
  AclSize As Integer
  AceCount As Integer
  Sbz2 As Integer

End Type

Private Type SECURITY_DESCRIPTOR

  Revision As Byte
  Sbz1 As Byte
  Control As Long
  Owner As Long
  Group As Long
  Sacl As ACL
  Dacl As ACL

End Type

Private Type PRINTER_INFO_2

  pServerName As String         '\\NTDC
  pPrinterName As String        '\\NTDC\HP LaserJet 2100 Series PCL 6
  pShareName As String          'HPLJ2100
  pPortName As String           'LPT1:
  pDriverName As String         'HP LaserJet 2100 Series PCL 6
  pComment As String            '
  pLocation As String           '
  pDevMode As DEVMODE
  pSepFile As String            '
  pPrintProcessor As String     'winprint
  pDatatype As String           'RAW
  pParameters As String         '
  pSecurityDescriptor As SECURITY_DESCRIPTOR
  Attributes As Long
  Priority As Long
  DefaultPriority As Long
  StartTime As Long
  UntilTime As Long
  Status As Long
  cJobs As Long
  AveragePPM As Long

End Type

Private Declare Function EnumPrinters _
                Lib "winspool.drv" _
                Alias "EnumPrintersA" (ByVal Flags As Long, _
                                       ByVal name As String, _
                                       ByVal Level As Long, _
                                       pPrinterEnum As Long, _
                                       ByVal cdBuf As Long, _
                                       pcbNeeded As Long, _
                                       pcReturned As Long) As Long
Private Declare Function OpenPrinter _
                Lib "winspool.drv" _
                Alias "OpenPrinterA" (ByVal pPrinterName As String, _
                                      phPrinter As Long, _
                                      pDefault As Any) As Long 'PRINTER_DEFAULTS
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function GetPrinter _
                Lib "winspool.drv" _
                Alias "GetPrinterA" (ByVal hPrinter As Long, _
                                     ByVal Level As Long, _
                                     pPrinter As Any, _
                                     ByVal cbBuf As Long, _
                                     pcbNeeded As Long) As Long
Private Declare Function SetPrinter _
                Lib "winspool.drv" _
                Alias "SetPrinterA" (ByVal hPrinter As Long, _
                                     ByVal Level As Long, _
                                     pPrinter As Any, _
                                     ByVal Command As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (hpvDest As Any, _
                                       hpvSource As Any, _
                                       ByVal cbCopy As Long)
Private Declare Function lstrcat _
                Lib "kernel32" _
                Alias "lstrcatA" (ByVal lpString1 As String, _
                                  ByVal lpString2 As String) As Long
Private Declare Function PtrToStr _
                Lib "kernel32" _
                Alias "lstrcpyA" (ByVal retVal As String, _
                                  ByVal Ptr As Long) As Long
Private Declare Function PtrToPtr _
                Lib "kernel32" _
                Alias "lstrcpyA" (retVal As Any, _
                                  Ptr As Any) As Long
Private Declare Function StrLen _
                Lib "kernel32" _
                Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function GetDefaultPrinter _
                Lib "winspool.drv" _
                Alias "GetDefaultPrinterA" (ByVal pszBuffer As String, _
                                            pcchBuffer As Long) As Long
Private Declare Function SetDefaultPrinter _
                Lib "winspool.drv" _
                Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Long
Private Declare Function DeviceCapabilities _
                Lib "winspool.drv" _
                Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
                                             ByVal lpPort As String, _
                                             ByVal iIndex As Long, _
                                             lpOutput As Any, _
                                             ByVal dev As Long) As Long
Private Declare Function DocumentProperties _
                Lib "winspool.drv" _
                Alias "DocumentPropertiesA" (ByVal hwnd As Long, _
                                             ByVal hPrinter As Long, _
                                             ByVal pDeviceName As String, _
                                             pDevModeOutput As Any, _
                                             pDevModeInput As Any, _
                                             ByVal fMode As Long) As Long
Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Private Declare Function CreateDC _
                Lib "gdi32" _
                Alias "CreateDCA" (ByVal lpDriverName As String, _
                                   ByVal lpDeviceName As String, _
                                   ByVal lpOutput As String, _
                                   ByVal lpInitData As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Sub Class_Initialize()
  '====================================================================================
  '2000-09-12, Klaus Prinz
  '   Aufruf der EnumeratePrinters mit den Flags für lokale und Netzwerkdrucker
  '====================================================================================
  Const lngFlags As Long = PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL
  Dim strReturn  As String         'Rückgabe aus EnumeratePrinters

  If cSysInfo.WinNT Then
    'Aufruf mit PRINTER_INFO_4, 3 Elemente
    strReturn = EnumeratePrinters(lngFlags, 4, 3)
  Else
    'Aufruf mit PRINTER_INFO_5, 5 Elemente
    strReturn = EnumeratePrinters(lngFlags, 5, 5)
  End If

  'Rückgabe in Array umwandeln
  If LenB(strReturn) > 0 Then
    strPrinters = Split(strReturn, ";")
  End If

  strPrinter = Me.DefaultPrinter
End Sub

Private Sub Class_Terminate()

  '====================================================================================
  '2002-03-05, Klaus Prinz
  '   evt. vorhandenen Gerätekontext freigeben
  '====================================================================================
  If lngDC <> 0 Then
    DeleteDC lngDC
  End If

End Sub

Public Property Get hDC() As Long
  '====================================================================================
  '2002-03-05, Klaus Prinz
  '   Rückgabe des Device-Context des aktuellen Druckers
  '====================================================================================
  Dim strDriverName As String     'Name des Druckertreibers

  On Error GoTo errHandler

  If cSysInfo.WinNT Then
    strDriverName = "winspool"
  End If

  lngDC = CreateDC(strDriverName, strPrinter, 0, 0)
  hDC = lngDC

  Exit Property

errHandler:
  Err.Raise Err.number, "clsPrinter:Get hDC"
End Property

Public Property Get PrinterCount() As Long
  PrinterCount = nPrinters
End Property

Public Property Let printer(ByVal strNewValue As String)

  '====================================================================================
  '2002-02-21, Klaus Prinz
  '   Aufruf der EnumerateBins
  '====================================================================================
  On Error GoTo errHandler

  'evt. vorhandenen Gerätekontext freigeben
  If lngDC <> 0 Then
    DeleteDC lngDC
    lngDC = 0
  End If

  strPrinter = strNewValue
  EnumerateBins strPrinter, GetPrinter_2_Value(strPrinter, 3)

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Let Printer"
  Err.Raise Err.number
End Property

Public Property Get printerName(ByVal Index As Long) As String
  printerName = strPrinters(Index - 1)
End Property

Public Property Get ServerName(ByVal Index As Long) As String
  ServerName = GetPrinter_2_Value(strPrinters(Index - 1), 0)
End Property

Private Function GetPrinter_2_Value(ByVal strPrinter As String, _
                                    ByVal iIndex As Long) As Variant
  Dim lngResult   As Long           'Rückgabe aus Funktion
  Dim hPrinter    As Long            'Drücker-Handle
  Dim lngBuffer() As Long         'Daten-Array
  Dim lngRequSize As Long         'Buffer-Größe
  Dim strTemp     As String           'Zwischenvariable
  ReDim lngBuffer(0)
  lngResult = OpenPrinter(strPrinter, hPrinter, 0&)
  lngResult = GetPrinter(hPrinter, 2, lngBuffer(0), 0&, lngRequSize)
  ReDim lngBuffer(lngRequSize / 4)
  'Alternativ mit Buffer für Strings
  lngResult = GetPrinter(hPrinter, 2, lngBuffer(0), lngRequSize, lngRequSize)
  lngResult = ClosePrinter(hPrinter)
  strTemp = Space$(StrLen(lngBuffer(iIndex)))
  lngResult = PtrToStr(strTemp, lngBuffer(iIndex))
  GetPrinter_2_Value = strTemp
End Function

Private Function EnumeratePrinters(ByVal lngFlags As Long, _
                                   ByVal iLevel As Long, _
                                   ByVal nElements As Long) As String
  '====================================================================================
  '2000-09-12, Klaus Prinz
  '   Rückgabe aller Drucker, die den Flags entsprechen
  '------------------------------------------------------------------------------------
  'Argumente:
  '   lngFlags:   Flags, denen die enumerierten Drucker entsprechen müssen
  '   iLevel:     Struktur-Kennung (PRINTER_INFO_?)
  '   nElements:  Anzahl der Elemente der Struktur
  '====================================================================================
  Dim lngResult       As Long           'Rückgabe aus Funktion
  Dim lngRequiredSize As Long     'von der Funktion ermittelte Buffergröße
  Dim lngBuffer()     As Long         'Daten-Array
  Dim iPrinter        As Long            'Druckerzeiger
  Dim strTemp()       As String         'Zwischenvariable für Rückgabe

  On Error GoTo errHandler

  'Größe des Buffers bestimmen
  lngResult = EnumPrinters(lngFlags, vbNullString, iLevel, 0&, 0&, lngRequiredSize, nPrinters)
  'Buffer einstellen und Funktion erneut aufrufen
  ReDim lngBuffer(lngRequiredSize \ 4) As Long
  lngResult = EnumPrinters(lngFlags, vbNullString, iLevel, lngBuffer(0), lngRequiredSize, lngRequiredSize, nPrinters)

  If nPrinters > 0 Then
    ReDim strTemp(nPrinters - 1)

    For iPrinter = 0 To nPrinters - 1
      strTemp(iPrinter) = Space$(StrLen(lngBuffer(iPrinter * nElements)))
      lngResult = PtrToStr(strTemp(iPrinter), lngBuffer(iPrinter * nElements))
    Next

  End If

  EnumeratePrinters = Join(strTemp, ";")

  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:EnumeratePrinters"
  Err.Raise Err.number
End Function

Public Property Get DefaultPrinter() As String
  '====================================================================================
  '2000-09-12, Klaus Prinz
  '   Rückgabe des Default-Printers
  '====================================================================================
  Dim strValue  As String          'Default-Printer
  Dim nSize     As Long               'Bruttogröße der Rückgabe
  Dim lngResult As Long           'Rückgabe aus GetDefaultPrinter

  On Error GoTo errHandler

  If cSysInfo.Win95 Or cSysInfo.Win98 Then
    'in Win 95 und 98 Enumerieren mit Default-Kostante und Struktur 5
    DefaultPrinter = EnumeratePrinters(PRINTER_ENUM_DEFAULT, 5, 5)
  ElseIf cSysInfo.WinNT Then

    Select Case cSysInfo.NTVersion

      Case 4
        'in NT bis Version 4 in Registry nachsehen
        strValue = cSysReg.GetRegistryValue(enumHKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device")

        If InStr(1, strValue, ",") > 0 Then
          DefaultPrinter = Left$(strValue, InStr(1, strValue, ",") - 1)
        Else
          DefaultPrinter = strValue
        End If

      Case 5
        'in Windows 2000ff neue Funktion nutzen
        lngResult = GetDefaultPrinter(strValue, nSize)
        strValue = Space$(nSize)
        lngResult = GetDefaultPrinter(strValue, nSize)
        DefaultPrinter = Left$(strValue, nSize - 1)
    End Select

  End If

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Get DefaultPrinter"
  Err.Raise Err.number
End Property

Public Property Let DefaultPrinter(ByVal strNewValue As String)
  '====================================================================================
  '2000-09-12, Klaus Prinz
  '   Übergebenen Werte als Default-Printer einstellen
  '====================================================================================
  Dim hPrinter    As Long                'Drucker-Handle
  Dim lngRequSize As Long             'erforderliche Variablengröße
  Dim lngResult   As Long               'Rückgabe aus GetDefaultPrinter
  Dim udtDefaults As PRINTER_DEFAULTS 'Struktur mit Zugangsrechten
  Dim lngInfo()   As Long               'aktuelle PRINTER_INFO-Daten

  On Error GoTo errHandler

  If cSysInfo.Win95 Or cSysInfo.Win98 Then
    'Rechte einstellen
    udtDefaults.pDatatype = 0&
    udtDefaults.DesiredAccess = PRINTER_ALL_ACCESS
    'Drucker öffnen
    lngResult = OpenPrinter(strNewValue, hPrinter, udtDefaults)

    If lngResult = 0 Then
      Err.Raise 9999, , "Der Drucker " & strNewValue & " konnte nicht geöffnet werden."
    End If

    'Anzahl der benötigten Bytes bestimmen und Drucker auslesen
    lngResult = GetPrinter(hPrinter, 5, ByVal 0&, 0&, lngRequSize)
    ReDim lngInfo((lngRequSize \ 4))
    lngResult = GetPrinter(hPrinter, 5, lngInfo(0), lngRequSize, lngRequSize)
    lngInfo(2) = lngInfo(2) Or PRINTER_ATTRIBUTE_DEFAULT
    lngResult = SetPrinter(hPrinter, 5, lngInfo(0), 0&)

    If lngResult = 0 Then
      MsgBox "SetPrinter Failed."
    End If

    Call ClosePrinter(hPrinter)
  ElseIf cSysInfo.WinNT Then

    Select Case cSysInfo.NTVersion

      Case 4
        'in NT bis Version 4 in Registry eintragen
        cSysReg.SetRegistryValue enumHKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Windows", "Device", 1, strNewValue
        lngResult = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, 0)

      Case 5
        lngResult = SetDefaultPrinter(strNewValue & vbNullChar)
    End Select

  End If

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Let DefaultPrinter"
  Err.Raise Err.number
End Property

Public Function GetPrinterConnection(ByVal strPrinter As String, _
                                     ByVal lngLangID As Long) As String
  '====================================================================================
  '2000-09-12, Klaus Prinz
  '   Ermitteln des Net-Resource-Druckernamens zur Zuweisung per ActivePrinter
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strPrinter:     Druckername aus EnumPrinters
  '   lngLangID:      Language-ID
  '====================================================================================
  Dim strRes As String            'Ressource (z.B. Ne00: bei NT)

  On Error GoTo errHandler

  If cSysInfo.Win95 Or cSysInfo.Win98 Then
    'in Win 95 und 98 Enumerieren mit Default-Kostante und Struktur 5
    strPrinter = GetPrinter_2_Value(strPrinter, 4)
    strRes = cSysReg.GetRegistryValue(enumHKEY_LOCAL_MACHINE, "System\CurrentControlSet\Control\Print\Printers\" & strPrinter, "Port")
  ElseIf cSysInfo.WinNT Then
    strRes = cSysReg.GetRegistryValue(enumHKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", strPrinter)
    strRes = Right$(strRes, Len(strRes) - InStr(1, strRes, ","))
  End If

  Select Case lngLangID

    Case 1031
      GetPrinterConnection = strPrinter & " auf " & strRes

    Case 1033
      GetPrinterConnection = strPrinter & " on " & strRes
  End Select

  Exit Function

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:GetPrinterConnection"
  Err.Raise Err.number
End Function

'**************************************************************************************************
'Papierzufuhr
'**************************************************************************************************
Public Property Get PaperBin() As String
  '====================================================================================
  '2000-02-28, Klaus Prinz
  '   aktuelle Papierzufuhr zurückgeben
  '====================================================================================
  Dim pDevMode  As DEVMODE         'DEVMODE-Struktur zur Ermittlung des Offset
  Dim lngOffset As Long           'Adress-Offset in DEVMODE-Struktur
  Dim intBin    As Integer           'aktuelle Bin-Konstante
  Dim iBin      As Long                'Zeiger in strPaperBins()

  On Error GoTo errHandler

  'Offset ermitteln
  lngOffset = VarPtr(pDevMode.dmDefaultSource) - VarPtr(pDevMode.dmSpecVersion) + 32
  'aktuelle Bin ermitteln
  intBin = GetDevModeValue(strPrinter, lngOffset)

  'intBin mit intPaperBins()-Elementen vergleichen ...
  For iBin = 0 To UBound(strPaperBins)

    If intBin = intPaperBins(iBin) Then
      '... und dazugehörigen Namen zurückgeben
      PaperBin = strPaperBins(iBin)

      Exit For

    End If

  Next

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Get PaperBin"
  Err.Raise Err.number
End Property

Public Property Let PaperBin(ByVal strNewValue As String)
  '====================================================================================
  '2000-02-28, Klaus Prinz
  '   neue Papierzufuhr einstellen
  '====================================================================================
  Dim pDevMode  As DEVMODE         'DEVMODE-Struktur zur Ermittlung des Offset
  Dim lngOffset As Long           'Adress-Offset in DEVMODE-Struktur
  Dim iBin      As Long                'Zeiger in strPaperBins()

  On Error GoTo errHandler

  'Offset ermitteln
  lngOffset = VarPtr(pDevMode.dmDefaultSource) - VarPtr(pDevMode.dmSpecVersion) + 32

  'Bin-Konstante ermitteln ...
  For iBin = 0 To UBound(strPaperBins)

    If strNewValue = strPaperBins(iBin) Then
      '... und damit DEVMODE ändern
      SetDevModeValue strPrinter, lngOffset, intPaperBins(iBin)

      Exit For

    End If

  Next

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Let PaperBin"
  Err.Raise Err.number
End Property

Public Property Get PaperBinCount() As Long

  '====================================================================================
  '2000-02-28, Klaus Prinz
  '   Anzahl der Papierzufuhrbehältnisse ermitteln
  '====================================================================================
  On Error GoTo errHandler

  PaperBinCount = UBound(intPaperBins) + 1

  Exit Property

errHandler:

  If Err.number = 9 Then
    Err.description = "Das PaperBin-Array wurde noch nicht initialisiert."
  End If

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Get PaperBinCount"
  Err.Raise Err.number
End Property

Public Function GetPaperBinName(ByVal iBin As Long) As String

  '====================================================================================
  '2000-02-28, Klaus Prinz
  '   Name des Papierzufuhrbehältnisses mit Index iBin zurückgeben
  '====================================================================================
  On Error GoTo errHandler

  GetPaperBinName = strPaperBins(iBin - 1)

  Exit Function

errHandler:

  If Err.number = 9 Then
    Err.description = "Das PaperBin-Array wurde noch nicht initialisiert."
  End If

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:GetPaperBinName"
  Err.Raise Err.number
End Function

Private Sub EnumerateBins(strPrinter As String, ByVal strPort As String)
  '====================================================================================
  '2000-09-19, Klaus Prinz
  '   Paper-Bin-Arrays füllen
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strPrinter:     Druckernamen
  '   strPort:        Port
  '====================================================================================
  Dim nBins   As Long                'Anzahl der Bins des Druckers
  Dim iBin    As Long                 'Zeiger in nBins
  Dim strBins As String            'String mit allen verfügbaren Bin-Namen

  On Error GoTo errHandler

  'NULLs anhängen
  strPrinter = strPrinter & vbNullChar
  strPort = strPort & vbNullChar
  'Anzahl der Bins ermitteln
  nBins = DeviceCapabilities(strPrinter, strPort, DC_BINS, ByVal vbNullString, 0)
  'Arrays dimensionieren
  ReDim strPaperBins(nBins - 1)
  ReDim intPaperBins(nBins - 1)
  'String vorbelegen
  strBins = String$(24 * (nBins), 0)
  'Bin-Konstanten-Array füllen lassen
  nBins = DeviceCapabilities(strPrinter, strPort, DC_BINS, intPaperBins(0), 0)
  'Bin-Namen-Array füllen lassen
  nBins = DeviceCapabilities(strPrinter, strPort, DC_BINNAMES, ByVal strBins, 0) - 1

  For iBin = 0 To nBins
    'Bin-Namen in Namens-Array übertragen
    strPaperBins(iBin) = Mid$(strBins, 24 * iBin + 1, 24)
    strPaperBins(iBin) = Left$(strPaperBins(iBin), InStr(1, strPaperBins(iBin), Chr$(0)) - 1)
  Next

  Exit Sub

errHandler:
  Err.Raise Err.number, "clsPrinter:EnumerateBins"
End Sub

'**************************************************************************************************
'Orientation
'**************************************************************************************************
Public Property Get Orientation() As enumOrientation
  '====================================================================================
  '2000-02-28, Klaus Prinz
  '   aktuelle Papierausrichtung zurückgeben
  '====================================================================================
  Dim pDevMode  As DEVMODE         'DEVMODE-Struktur zur Ermittlung des Offset
  Dim lngOffset As Long           'Adress-Offset in DEVMODE-Struktur

  On Error GoTo errHandler

  'Offset ermitteln
  lngOffset = VarPtr(pDevMode.dmOrientation) - VarPtr(pDevMode.dmSpecVersion) + 32
  'aktuelle Ausrichtung ermitteln
  Orientation = GetDevModeValue(strPrinter, lngOffset)

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Get Orientation"
  Err.Raise Err.number
End Property

Public Property Let Orientation(ByVal lngNewValue As enumOrientation)
  '====================================================================================
  '2000-02-28, Klaus Prinz
  '   neue Papierausrichtung einstellen
  '====================================================================================
  Dim pDevMode  As DEVMODE         'DEVMODE-Struktur zur Ermittlung des Offset
  Dim lngOffset As Long           'Adress-Offset in DEVMODE-Struktur

  On Error GoTo errHandler

  'Offset ermitteln
  lngOffset = VarPtr(pDevMode.dmOrientation) - VarPtr(pDevMode.dmSpecVersion) + 32
  'neue Ausrichtung übergeben
  SetDevModeValue strPrinter, lngOffset, lngNewValue

  Exit Property

errHandler:

  If InStr(1, Err.source, ":") = 0 Then Err.source = "clsPrinter:Let Orientation"
  Err.Raise Err.number
End Property

'**************************************************************************************************
'lokale DEVMODE-Strukturen
'**************************************************************************************************
Private Function GetDevModeValue(ByVal strPrinter As String, _
                                 ByVal lngOffset As Long) As Integer
  '====================================================================================
  '2001-02-23, Klaus Prinz
  '   Wert aus DEVMODE-Abschnitt der PRINTER_INFO_2-Struktur auslesen
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strPrinter:     Druckernamen
  '   lngOffset:      Offset zu DEVMODE-Basisadresse
  '====================================================================================
  Dim nSize          As Long               'Puffergröße nach Ansicht der GetPrinter-Funktion
  Dim hPrinter       As Long            'Drucker-Handle
  Dim pd             As PRINTER_DEFAULTS      'Zugangsrechte
  Dim aDevMode(3046) As Byte      'Daten aus GetPrinter-Funktion
  Dim lngAddr(7)     As Long          'Adress-Array für die ersten 8 Elemente der PRINTER_INFO_2
  Dim lngResult      As Long           'Rückgabe aus Funktion

  On Error GoTo errHandler

  pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess
  lngResult = OpenPrinter(strPrinter, hPrinter, pd)
  lngResult = GetPrinter(hPrinter, 2, aDevMode(0), UBound(aDevMode), nSize)
  Call CopyMemory(lngAddr(0), aDevMode(0), 32)
  lngResult = PtrToPtr(GetDevModeValue, ByVal (lngAddr(7) + lngOffset))
  lngResult = ClosePrinter(hPrinter)

  Exit Function

errHandler:
  Err.Raise Err.number, "clsSysPrinter:GetDevModeValue"
End Function

Private Sub SetDevModeValue(ByVal strPrinter As String, _
                            ByVal lngOffset As Long, _
                            ByVal intValue As Integer)
  '====================================================================================
  '2001-02-23, Klaus Prinz
  '   Wert in DEVMODE-Abschnitt der PRINTER_INFO_2-Struktur ändern
  '------------------------------------------------------------------------------------
  'Argumente:
  '   strPrinter:     Druckernamen
  '   lngOffset:      Offset zu DEVMODE-Basisadresse
  '   intValue:       einzustellender Wert
  '====================================================================================
  Dim nSize          As Long               'Puffergröße nach Ansicht der GetPrinter-Funktion
  Dim hPrinter       As Long            'Drucker-Handle
  Dim pd             As PRINTER_DEFAULTS      'Zugangsrechte
  Dim aDevMode(3046) As Byte      'Daten aus GetPrinter-Funktion
  Dim lngAddr(7)     As Long          'Adress-Array für die ersten 8 Elemente der PRINTER_INFO_2
  Dim lngResult      As Long           'Rückgabe aus Funktion

  On Error GoTo errHandler

  pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess
  lngResult = OpenPrinter(strPrinter, hPrinter, pd)

  If lngResult > 0 Then
    'Schreiben nur, wenn Drucker auch geöffnet werden kann
    lngResult = GetPrinter(hPrinter, 2, aDevMode(0), UBound(aDevMode), nSize)
    Call CopyMemory(lngAddr(0), aDevMode(0), 32)
    lngResult = PtrToPtr(ByVal (lngAddr(7) + lngOffset), intValue)
    lngResult = SetPrinter(hPrinter, 2, aDevMode(0), 0)
    lngResult = ClosePrinter(hPrinter)
  End If

  Exit Sub

errHandler:
  Err.Raise Err.number, "clsSysPrinter:SetDevModeValue"
End Sub

