VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32ApiSysNLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'© 2000 Klaus Prinz EDV-Dienstleistungen
'       Diese Klasse dient als Schnittstelle zu allgemeinen Registry-Einträgen
'==================================================================================================
'Methoden:
'--------------------------------------------------------------------------------------------------
'   GetLocaleValue:     Private
'       Ermittelt die Locale-Einstellung der übergebenen LCType-Konstanten
'==================================================================================================
'Eigenschaften: (alle Read-only)
'--------------------------------------------------------------------------------------------------
'   CountryCode:                    CountryCode (049)
'   CountryEnglish:                 englischer Landesname (Germany)
'   CountryLocal:                   lokaler Landesname (Deutschland)
'   CurrencySymbolInternational:    internationales Währungssymbol (DEM)
'   CurrencySymbolLocal:            lokales Währungssymbol (DM)
'   DecimalSeparator:               Dezimaltrennzeichen (,)
'   LanguageEnglish:                englischer Sprachname (German)
'   LanguageNative:                 lokaler Sprachname (Deutsch)
'   ListDelimiter:                  Listentrennzeichen (;)
'   MonthName:                      lokaler Name des i-ten Monats
'   MonthNameArray:                 lokale Monatsnamen in Array
'   Systemfont:                     System-Font
'==================================================================================================

'Globale Konstanten
Private Const LOCALE_USER_DEFAULT = &H400
'Verwendete APIs Locale Settings
Private Declare Function GetLocaleInfo _
                Lib "kernel32" _
                Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                        ByVal LCType As Long, _
                                        ByVal lpLCData As String, _
                                        ByVal cchData As Long) As Long

Private Function GetLocaleValue(ByVal lngLCType As Long) As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   Ermittelt die Locale-Einstellung der übergebenen LCType-Konstanten
  '====================================================================================
  Dim nSize     As Long               'Länge des Rückgabe
  Dim strBuffer As String         'Rückgabe
  Dim lngDummy  As Long            'Dummy für Rückgabe
  On Error GoTo errHandler
  nSize = GetLocaleInfo(LOCALE_USER_DEFAULT, lngLCType, strBuffer, nSize)
  strBuffer = Space(nSize)
  lngDummy = GetLocaleInfo(LOCALE_USER_DEFAULT, lngLCType, strBuffer, nSize)
  GetLocaleValue = Left(strBuffer, nSize - 1)
  Exit Function

errHandler:
  Err.Raise vbObjectError + Err.number, "clsSystem:GetLocaleValue"
End Function

'**************************************************************************************************
'Eigenschaften
'**************************************************************************************************
Public Property Get CountryCode() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   CountryCode (049)
  '====================================================================================
  Const LOCALE_ICOUNTRY = &H5
  CountryCode = GetLocaleValue(LOCALE_ICOUNTRY)
End Property

Public Property Get CountryEnglish() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   englischer Landesname (Germany)
  '====================================================================================
  Const LOCALE_SENGCOUNTRY = &H1002
  CountryEnglish = GetLocaleValue(LOCALE_SENGCOUNTRY)
End Property

Public Property Get CountryLocal() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   lokaler Landesname (Deutschland)
  '====================================================================================
  Const LOCALE_SCOUNTRY = &H6
  CountryLocal = GetLocaleValue(LOCALE_SCOUNTRY)
End Property

Public Property Get CurrencySymbolInternational() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   internationales Währungssymbol (DEM)
  '====================================================================================
  Const LOCALE_SINTLSYMBOL = &H15
  CurrencySymbolInternational = GetLocaleValue(LOCALE_SINTLSYMBOL)
End Property

Public Property Get CurrencySymbolLocal() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   lokales Währungssymbol (DM)
  '====================================================================================
  Const LOCALE_SCURRENCY = &H14
  CurrencySymbolLocal = GetLocaleValue(LOCALE_SCURRENCY)
End Property

Public Property Get DecimalSeparator() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   Dezimaltrennzeichen (DM)
  '====================================================================================
  Const LOCALE_SDECIMAL = &HE
  DecimalSeparator = GetLocaleValue(LOCALE_SDECIMAL)
End Property

Public Property Get LanguageEnglish() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   englischer Sprachname (German)
  '====================================================================================
  Const LOCALE_SENGLANGUAGE = &H1001
  LanguageEnglish = GetLocaleValue(LOCALE_SENGLANGUAGE)
End Property

Public Property Get LanguageNative() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   lokaler Sprachname (Deutsch)
  '====================================================================================
  Const LOCALE_SNATIVELANGNAME = &H4
  LanguageNative = GetLocaleValue(LOCALE_SNATIVELANGNAME)
End Property

Public Property Get ListDelimiter() As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   Listentrennzeichen (;)
  '====================================================================================
  Const LOCALE_SLIST = &HC
  ListDelimiter = GetLocaleValue(LOCALE_SLIST)
End Property

Public Property Get MonthName(ByVal iMonth As Long) As String
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   lokaler Name des i-ten Monats
  '====================================================================================
  Dim lngMonthConst As Long
  lngMonthConst = 55 + iMonth
  MonthName = GetLocaleValue(lngMonthConst)
End Property

Public Property Get MonthNameArray() As String()
  '====================================================================================
  '2000-02-08, Klaus Prinz
  '   lokale Monatsnamen in Array
  '====================================================================================
  Dim strMonths(12) As String
  Dim iMonth        As Long

  For iMonth = 1 To 12
    strMonths(iMonth) = GetLocaleValue(55 + iMonth)
  Next

  MonthNameArray = strMonths
End Property

