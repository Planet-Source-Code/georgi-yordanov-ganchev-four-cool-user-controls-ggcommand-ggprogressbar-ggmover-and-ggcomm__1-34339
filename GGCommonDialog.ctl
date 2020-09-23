VERSION 5.00
Begin VB.UserControl GGCommonDialog 
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GGCommonDialog.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "GGCommonDialog.ctx":00F4
End
Attribute VB_Name = "GGCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FW_NORMAL = 400
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const FF_ROMAN = 16
Private Const CF_PRINTERFONTS = &H2
Private Const CF_SCREENFONTS = &H1
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const REGULAR_FONTTYPE = &H400
Private Const LF_FACESIZE = 32
Private Const CCHDEVICENAME = 132
Private Const CCHFORMNAME = 32
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const PD_PRINTSETUP = &H40
Private Const PD_DISABLEPRINTTOFILE = &H80000

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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
        lfFaceName As String * 31
End Type

Private Type ChooseFont
        lStructSize As Long
        hWndOwner As Long          '  caller's window handle
        hDC As Long                '  printer DC/IC or NULL
        lpLogFont As Long          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        Flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Type DEVMODE_TYPE
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
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type


Private CustomColors() As Byte

Public Enum ColorConstants
 'ChooseColor
 cdlCCFullOpen = &H2
 cdlCCHelpButton = &H8
 cdlCCPreventFullOpen = &H4
 cdlCCRGBInit = &H1
End Enum

Public Enum FileOpenConstants
 'GetOpenFileName and GetSaveFileName
 cdlOFNAllowMultiselect = &H200
 cdlOFNCreatePrompt = &H2000
 cdlOFNExplorer = &H80000
 cdlOFNExtensionDifferent = &H400
 cdlOFNFileMustExist = &H1000
 cdlOFNHelpButton = &H10
 cdlOFNHideReadOnly = &H4
 cdlOFNLongNames = &H200000
 cdlOFNNoChangeDir = &H8
 cdlOFNNoDereferenceLinks = &H100000
 cdlOFNNoLongNames = &H40000
 cdlOFNNoReadOnlyReturn = &H8000
 cdlOFNNoValidate = &H100
 cdlOFNOverwritePrompt = &H2
 cdlOFNPathMustExist = &H800
 cdlOFNReadOnly = &H1
 cdlOFNShareAware = &H4000
End Enum

Public Enum FontsConstants
 'ChooseFont
 cdlCFANSIOnly = &H400
 cdlCFApply = &H200
 cdlCFBoth = &H3
 cdlCFEffects = &H100
 cdlCFFixedPitchOnly = &H4000
 cdlCFForceFontExist = &H10000
 cdlCFHelpButton = &H4
 cdlCFLimitSize = &H2000
 cdlCFNoFaceSel = &H80000
 cdlCFNoSimulations = &H1000
 cdlCFNoSizeSel = &H200000
 cdlCFNoStyleSel = &H100000
 cdlCFNoVectorFonts = &H800
 cdlCFPrinterFonts = &H2
 cdlCFScalableOnly = &H20000
 cdlCFScreenFonts = &H1
 cdlCFTTOnly = &H40000
 cdlCFWYSIWYG = &H8000
End Enum

Public Enum HelpConstants
 'ShowHelp
 cdlHelpCommandHelp = &H102
 cdlHelpContents = &H3
 cdlHelpContext = &H1
 cdlHelpContextPopup = &H8
 cdlHelpForceFile = &H9
 cdlHelpHelpOnHelp = &H4
 cdlHelpIndex = &H3
 cdlHelpKey = &H101
 cdlHelpPartialKey = &H105
 cdlHelpQuit = &H2
 cdlHelpSetContents = &H5
 cdlHelpSetIndex = &H5
End Enum

Public Enum PrinterConstants
 'PrintDialog - PrinterConstants
 cdlPDAllPages = &H0
 cdlPDCollate = &H10
 cdlPDDisablePrintToFile = &H80000
 cdlPDHelpButton = &H800
 cdlPDHidePrintToFile = &H100000
 cdlPDNoPageNums = &H8
 cdlPDNoSelection = &H4
 cdlPDNoWarning = &H80
 cdlPDPageNums = &H2
 cdlPDPrintSetup = &H40
 cdlPDPrintToFile = &H20
 cdlPDReturnDC = &H100
 cdlPDReturnDefault = &H400
 cdlPDReturnIC = &H200
 cdlPDSelection = &H1
 cdlPDUseDevModeCopies = &H40000
End Enum

Public Enum PrinterOrientationConstants
 'PrintDialog - PrinterOrientationConstants
 cdlLandscape = &H2
 cdlPortrait = &H1
End Enum

Public Enum BrowseForFolderConstants
 'BrowseForFolder - BrowseForFolderConstants
 BIF_RETURNONLYFSDIRS = &H1     ' For finding a folder to start document searching
 BIF_DONTGOBELOWDOMAIN = &H2    ' For starting the Find Computer
 BIF_STATUSTEXT = &H4
 BIF_RETURNFSANCESTORS = &H8
 BIF_EDITBOX = &H10
 BIF_VALIDATE = &H20             ' insist on valid result (or CANCEL)
 BIF_BROWSEFORCOMPUTER = &H1000 ' Browsing for Computers.
 BIF_BROWSEFORPRINTER = &H2000   ' Browsing for Printers
 BIF_BROWSEINCLUDEFILES = &H4000 ' Browsing for Everything
End Enum

Public Enum BrowseForFolderSpecialFolder
  CSIDL_DESKTOP = &H0
  CSIDL_INTERNET = &H1
  CSIDL_PROGRAMS = &H2
  CSIDL_CONTROLS = &H3
  CSIDL_PRINTERS = &H4
  CSIDL_PERSONAL = &H5
  CSIDL_FAVORITES = &H6
  CSIDL_STARTUP = &H7
  CSIDL_RECENT = &H8
  CSIDL_SENDTO = &H9
  CSIDL_BITBUCKET = &HA
  CSIDL_STARTMENU = &HB
  CSIDL_DESKTOPDIRECTORY = &H10
  CSIDL_DRIVES = &H11
  CSIDL_NETWORK = &H12
  CSIDL_NETHOOD = &H13
  CSIDL_FONTS = &H14
  CSIDL_TEMPLATES = &H15
  CSIDL_COMMON_STARTMENU = &H16
  CSIDL_COMMON_PROGRAMS = &H17
  CSIDL_COMMON_STARTUP = &H18
  CSIDL_COMMON_DESKTOPDIRECTORY = &H19
  CSIDL_APPDATA = &H1A
  CSIDL_PRINTHOOD = &H1B
  CSIDL_ALTSTARTUP = &H1D          ' DBCS
  CSIDL_COMMON_ALTSTARTUP = &H1E   ' DBCS
  CSIDL_COMMON_FAVORITES = &H1F
  CSIDL_INTERNET_CACHE = &H20
  CSIDL_COOKIES = &H21
  CSIDL_HISTORY = &H22
End Enum


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'Default Property Values:
Const m_def_BrowseSpecialFolder = 0
Const m_def_FolderPath = ""
Const m_def_PrinterName = ""
Const m_def_PrinterDefaultSource = 0
'Const m_def_PrinterName = ""
Const m_def_HelpCommand = 0
Const m_def_HelpContext = 0
Const m_def_HelpFile = ""
Const m_def_HelpKey = ""
Const m_def_Action = 0
Const m_def_Orientation = 1
Const m_def_FromPage = 0
Const m_def_ToPage = 0
Const m_def_Min = 0
Const m_def_Max = 0
Const m_def_PrinterDefault = 0
Const m_def_hDC = 0
Const m_def_FontItalic = 0
Const m_def_FontCharSet = 0
Const m_def_FontBold = 0
Const m_def_FontName = ""
Const m_def_FontSize = 8
Const m_def_FontStrikeThru = 0
Const m_def_FontUnderLine = 0
Const m_def_FileOffset = 0
Const m_def_ExtensionOffset = 0
Const m_def_MaxFileSize = 260
Const m_def_FilterIndex = 0
'Const m_def_FilterIndex = 0
Const m_def_Copies = 1
Const m_def_FileTitle = ""
Const m_def_DialogTitle = ""
Const m_def_InitDir = ""
Const m_def_Filter = ""
Const m_def_FileName = ""
Const m_def_DefaultExt = ""
Const m_def_Flags = 0
Const m_def_CancelError = 0
Const m_def_Color = 0
'Property Variables:
Dim m_BrowseSpecialFolder As BrowseForFolderSpecialFolder
Dim m_FolderPath As String
Dim m_PrinterExtraCollate As Integer
Dim m_PrinterExtraColor As Integer
Dim m_PrinterExtraDisplayFlags As Long
Dim m_PrinterExtraDriverVersion As Integer
Dim m_PrinterExtraDuplex As Integer
Dim m_PrinterExtraFields As Long
Dim m_PrinterExtraPaperLength As Integer
Dim m_PrinterExtraPaperSize As Integer
Dim m_PrinterExtraPaperWidth As Integer
Dim m_PrinterExtraPrintQuality As Integer
Dim m_PrinterExtraScale As Integer
Dim m_PrinterExtraSpecVersion As Integer
Dim m_PrinterExtraTTOption As Integer
Dim m_PrinterExtraYResolution As Integer
Dim m_PrinterExtraExtra As String
Dim m_PrinterExtraDefault As Integer
Dim m_PrinterExtraDeviceOffset As Integer
Dim m_PrinterExtraDriverOffset As Integer
Dim m_PrinterExtraOutputOffset As Integer
Dim m_PrinterName As String
Dim m_PrinterDefaultSource As Long
'Dim m_PrinterName As String
Dim m_HelpCommand As Integer
Dim m_HelpContext As Long
Dim m_HelpFile As String
Dim m_HelpKey As String
Dim m_Action As Integer
Dim m_Orientation As PrinterOrientationConstants
Dim m_FromPage As Integer
Dim m_ToPage As Integer
Dim m_Min As Integer
Dim m_Max As Integer
Dim m_PrinterDefault As Boolean
Dim m_hDC As Long
Dim m_FontItalic As Boolean
Dim m_FontCharSet As Long
Dim m_FontBold As Boolean
Dim m_FontName As String
Dim m_FontSize As Single
Dim m_FontStrikeThru As Boolean
Dim m_FontUnderLine As Boolean
Dim m_FileOffset As Integer
Dim m_ExtensionOffset As Integer
Dim m_MaxFileSize As Integer
Dim m_FilterIndex As Integer
'Dim m_FilterIndex As Integer
Dim m_Copies As Integer
Dim m_FileTitle As String
Dim m_DialogTitle As String
Dim m_InitDir As String
Dim m_Filter As String
Dim m_FileName As String
Dim m_DefaultExt As String
Dim m_Flags As Long
Dim m_CancelError As Boolean
Dim m_Color As OLE_COLOR


Dim plngOwnerHwnd As Long

Public Sub About()
Attribute About.VB_UserMemId = -552

MsgBox "GGCommonDialog user control" & vbCrLf & _
       "Compatible with 'Microsoft Common Dialog Control 6.0 (COMDLG32.OCX)'" & vbCrLf & _
       "Coded by Georgi Yordanov Ganchev - Varna, BULGARIA... - http://georgi-ganchev.tripod.com" & vbCrLf & vbCrLf & _
       "...with special thanks to KPDTeam - API-Guide (http://www.allapi.net/)" & vbCrLf & vbCrLf & _
       "Send comments to GogoX@Lycos.com", _
       vbInformation, _
       "GGCommonDialog user control - Georgi Yordanov Ganchev ©2000"

End Sub


Public Property Get OwnerHwnd() As Long

OwnerHwnd = plngOwnerHwnd

End Property

Public Property Let OwnerHwnd(New_OwnerHwnd As Long)

plngOwnerHwnd = New_OwnerHwnd

End Property

Public Function ShowBrowseForFolder() As String
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'KPDTeam@Allapi.net
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

If plngOwnerHwnd = 0 Then plngOwnerHwnd = UserControl.Parent.hwnd

    With udtBI
        .pIDLRoot = m_BrowseSpecialFolder
        'Set the owner window
        .hWndOwner = plngOwnerHwnd ' UserControl.Parent.hwnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(m_DialogTitle, "")
        'Return only if the user selected a directory
        .ulFlags = m_Flags
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(m_MaxFileSize, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    Else
        If m_CancelError = True Then
        Err.Raise 32755, "GGCommonDialog user control", "Cancel was selected."
        m_FileName = ""
        End If
    End If

    m_FolderPath = sPath
    
End Function

Public Sub ShowHelp()

If plngOwnerHwnd = 0 Then plngOwnerHwnd = UserControl.Parent.hwnd

'I don't know how to handle help commands
'So if it is 'CommandHelp' then use HelpContext property
If HelpCommand = 1 Then
 WinHelp plngOwnerHwnd, m_HelpFile, m_HelpCommand, ByVal m_HelpContext
'Else use HelpKey property
Else
 WinHelp plngOwnerHwnd, m_HelpFile, m_HelpCommand, ByVal m_HelpKey
End If

End Sub

Public Sub ShowPrinter()
    
    '-> Code by Donald Grover
    'Modified by Georgi Jordanov Ganchev - GogoX@Mailcity.com
    
    On Error Resume Next
    
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
If plngOwnerHwnd = 0 Then plngOwnerHwnd = UserControl.Parent.hwnd

    PrintDlg.hWndOwner = plngOwnerHwnd  'UserControl.Parent.hwnd
    PrintDlg.nMinPage = m_Min
    PrintDlg.nMaxPage = m_Max
    PrintDlg.nFromPage = m_FromPage
    PrintDlg.nToPage = m_ToPage
    PrintDlg.Flags = m_Flags
    

    On Error Resume Next
  'Some bug I found in original code by Donald Grover - "Method _ShowPrinter of object _GGCommonDialog failed" and GPF
  'when set the following remarked lines
  'So don't set the current orientation and duplex setting
  '
  'Set the current orientation and duplex setting
  '  DevMode.dmDeviceName = Printer.DeviceName
     DevMode.dmSize = Len(DevMode)
  '  DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX Or cdlPDCollate
  '  DevMode.dmPaperWidth = Printer.Width
  '  DevMode.dmOrientation = Printer.Orientation
  '  DevMode.dmPaperSize = Printer.PaperSize
  '  DevMode.dmDuplex = Printer.Duplex
'    On Error GoTo 0



    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))

        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    If m_PrinterDefault = True Then
                    Set Printer = objPrinter
                    End If
                    'set printer toolbar name at this point
                End If
            Next
        End If
        m_Flags = PrintDlg.Flags
        m_Copies = PrintDlg.nCopies
        m_FromPage = PrintDlg.nFromPage
        m_ToPage = PrintDlg.nToPage
        m_Orientation = DevMode.dmOrientation
        m_hDC = PrintDlg.hDC
        m_PrinterName = Left$(DevMode.dmDeviceName, InStr(1, DevMode.dmDeviceName, Chr$(0)) - 1)
        m_PrinterDefaultSource = DevMode.dmDefaultSource

    'Set extra added properties
    'Some not work - PaperSize for example :(
        m_PrinterExtraCollate = DevMode.dmCollate
        m_PrinterExtraColor = DevMode.dmColor
        m_PrinterExtraDisplayFlags = DevMode.dmDisplayFlags
        m_PrinterExtraDriverVersion = DevMode.dmDriverVersion
        m_PrinterExtraDuplex = DevMode.dmDuplex
        m_PrinterExtraFields = DevMode.dmFields
        m_PrinterExtraPaperLength = DevMode.dmPaperLength
        m_PrinterExtraPaperSize = DevMode.dmPaperSize
        m_PrinterExtraPaperWidth = DevMode.dmPaperWidth
        m_PrinterExtraPrintQuality = DevMode.dmPrintQuality
        m_PrinterExtraScale = DevMode.dmScale
        m_PrinterExtraSpecVersion = DevMode.dmSpecVersion
        m_PrinterExtraTTOption = DevMode.dmTTOption
        m_PrinterExtraYResolution = DevMode.dmYResolution
        m_PrinterExtraExtra = DevName.extra
        m_PrinterExtraDefault = DevName.wDefault
        m_PrinterExtraDeviceOffset = DevName.wDeviceOffset
        m_PrinterExtraDriverOffset = DevName.wDriverOffset
        m_PrinterExtraOutputOffset = DevName.wOutputOffset

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        If m_PrinterDefault = True Then
        Printer.Copies = PrintDlg.nCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        End If

        On Error GoTo 0
    Else
        If m_CancelError = True Then
        Err.Raise 32755, "GGCommonDialog user control", "Cancel was selected."
        m_FileName = ""
        End If
        
    End If
    
        
        
        
End Sub



Public Sub ShowSave()
Dim SFName As OPENFILENAME
'    'Clear FileName property
'    m_FileName = ""
'    m_FileTitle = ""
'    m_FolderPath = ""
    'Set the structure size
    SFName.lStructSize = Len(SFName)
    'Set the owner window
If plngOwnerHwnd = 0 Then plngOwnerHwnd = UserControl.Parent.hwnd
    SFName.hWndOwner = plngOwnerHwnd ' UserControl.Parent.hwnd
    'Set the application's instance
    SFName.hInstance = App.hInstance
    'Set the filet
    SFName.lpstrFilter = m_Filter
    'Default extension
    SFName.lpstrDefExt = m_DefaultExt
    'Create a buffer
    SFName.lpstrFile = Space$(MaxFileSize)
    'Set the maximum number of chars
    SFName.nMaxFile = MaxFileSize + 1
    'Create a buffer
    SFName.lpstrFileTitle = Space$(MaxFileSize)
    'Set the maximum number of chars
    SFName.nMaxFileTitle = MaxFileSize + 1
    'Set the initial directory
    SFName.lpstrInitialDir = m_InitDir
    'Set the dialog title
    SFName.lpstrTitle = m_DialogTitle
    'no extra flags
    SFName.Flags = m_Flags

    'Show the 'Save File'-dialog
    If GetSaveFileName(SFName) Then
        'Set FileName property to Path+FileName without ChrS(0)
        m_FileName = Trim$(Left$(SFName.lpstrFile, InStr(1, SFName.lpstrFile, Chr$(0)) - 1))
        'Set other properties
        m_FileOffset = SFName.nFileOffset
        m_ExtensionOffset = SFName.nFileExtension
        m_FileTitle = Left$(SFName.lpstrFileTitle, InStr(1, SFName.lpstrFileTitle, Chr$(0)) - 1)
        m_FilterIndex = SFName.nFilterIndex
    Else
        If m_CancelError = True Then
        Err.Raise 32755, "GGCommonDialog user control", "Cancel was selected."
        m_FileName = ""
        End If

    End If
End Sub


Public Sub ShowOpen()
Dim OFName As OPENFILENAME
'    'Clear FileName property
'    m_FileName = ""
'    m_FileTitle = ""
'    m_FolderPath = ""
    'Structure size
    OFName.lStructSize = Len(OFName)
    'Owner window
If plngOwnerHwnd = 0 Then plngOwnerHwnd = UserControl.Parent.hwnd
    
    OFName.hWndOwner = plngOwnerHwnd ' UserControl.Parent.hwnd
    'Application's instance
    OFName.hInstance = App.hInstance
    'Filter
    OFName.lpstrFilter = m_Filter
    'Default extension
    OFName.lpstrDefExt = m_DefaultExt
    'Create a buffer
    OFName.lpstrFile = Space$(MaxFileSize)
    'Maximum number of chars per file
    OFName.nMaxFile = MaxFileSize + 1
    'Create a buffer
    OFName.lpstrFileTitle = Space$(MaxFileSize)
    'Maximum number of chars per file
    OFName.nMaxFileTitle = MaxFileSize + 1
    'Initial directory
    OFName.lpstrInitialDir = m_InitDir
    'Dialog title
    OFName.lpstrTitle = m_DialogTitle
    'Flags
    OFName.Flags = m_Flags

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        'Set FileName property to Path+FileName without ChrS(0)
        m_FileName = Trim$(Left$(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr$(0)) - 1))
        'Set other properties
        m_FileOffset = OFName.nFileOffset
        m_ExtensionOffset = OFName.nFileExtension
        m_FileTitle = Left$(OFName.lpstrFileTitle, InStr(1, OFName.lpstrFileTitle, Chr$(0)) - 1)
        m_FilterIndex = OFName.nFilterIndex
    Else
        If m_CancelError = True Then
        Err.Raise 32755, "GGCommonDialog user control", "Cancel was selected."
        m_FileName = ""
        End If
    End If
End Sub



Public Sub ShowFont()
    Dim cf As ChooseFont, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim retval As Long
    lfont.lfHeight = 0  ' determine default height
    lfont.lfWidth = 0  ' determine default width
    lfont.lfEscapement = 0  ' angle between baseline and escapement vector
    lfont.lfOrientation = 0  ' angle between baseline and orientation vector
    lfont.lfWeight = 400  ' normal weight i.e. not bold
    lfont.lfCharSet = 1  ' use default character set
    lfont.lfOutPrecision = 0  ' default precision mapping
    lfont.lfClipPrecision = 0  ' default clipping precision
    lfont.lfQuality = 0  ' default quality setting
    lfont.lfPitchAndFamily = 16  ' default pitch, proportional with serifs
    lfont.lfFaceName = m_FontName & vbNullChar  ' string must be null-terminated
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(&H42, Len(lfont))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    CopyMemory ByVal pMem, lfont, Len(lfont)  ' copy structure's contents into block
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
    cf.lStructSize = Len(cf)  ' size of structure

If plngOwnerHwnd = 0 Then plngOwnerHwnd = UserControl.Parent.hwnd

    cf.hWndOwner = plngOwnerHwnd ' UserControl.Parent.hwnd  ' window Form1 is opening this dialog box
    cf.hDC = 0 ' Printer.hDC  ' device context of default printer (using VB's mechanism)
    cf.lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
    cf.nSizeMin = m_Min   'Minimum font size
    cf.nSizeMax = m_Max   'Maximum font size
    'cf.iPointSize = m_FontSize * 10 ' 12 point font (in units of 1/10 point)
    cf.Flags = m_Flags ' CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_LIMITSIZE
    cf.rgbColors = RGB(0, 0, 0)  ' black
    cf.nFontType = &H400  ' regular font type i.e. not bold or anything
    ' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    retval = ChooseFont(cf)  ' open the dialog box
    If retval <> 0 Then  ' success
        CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        m_FontName = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
    m_Color = cf.rgbColors
    m_FontCharSet = lfont.lfCharSet
    m_FontItalic = CBool(lfont.lfItalic)
    m_FontBold = CBool(lfont.lfWeight > 400)
    m_FontStrikeThru = CBool(lfont.lfStrikeOut)
    m_FontUnderLine = CBool(lfont.lfUnderline)
    m_FontSize = CSng(Round((-lfont.lfHeight / GetDeviceCaps(GetDC(0), 90)) * 72, 0))
    m_hDC = cf.hDC
    Else
     If m_CancelError = True Then
     Err.Raise 32755, "GGCommonDialog user control", "Cancel was selected."
     End If

    End If
    
    ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    retval = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    retval = GlobalFree(hMem)  ' free the allocated memory
End Sub

Public Sub ShowColor()
    Dim cc As ChooseColor
    Dim Custcolor(16) As Long
  
If plngOwnerHwnd = 0 Then plngOwnerHwnd = UserControl.Parent.hwnd

    'Structure size
    cc.lStructSize = Len(cc)
    'Owner hwnd
    cc.hWndOwner = plngOwnerHwnd ' UserControl.Parent.hwnd
    'Application's instance
    cc.hInstance = App.hInstance
    'Custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'Flags
    cc.Flags = m_Flags

    'Show the 'Select Color'-dialog
    If ChooseColor(cc) <> 0 Then
        m_Color = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        If m_CancelError = True Then
        Err.Raise 32755, "GGCommonDialog user control", "Cancel was selected."
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
'ReDim custom colors array for Color dialog
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim I As Integer
    For I = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(I) = 0
    Next I

End Sub

Private Sub UserControl_Resize()
UserControl.Size 420, 420

End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Color = m_def_Color
    m_CancelError = m_def_CancelError
    m_Flags = m_def_Flags
    m_DefaultExt = m_def_DefaultExt
    m_FileName = m_def_FileName
    m_Filter = m_def_Filter
    m_InitDir = m_def_InitDir
    m_DialogTitle = m_def_DialogTitle
    m_FileTitle = m_def_FileTitle
    m_Copies = m_def_Copies
    m_FilterIndex = m_def_FilterIndex
    m_MaxFileSize = m_def_MaxFileSize
    m_FileOffset = m_def_FileOffset
    m_ExtensionOffset = m_def_ExtensionOffset
    m_FontBold = m_def_FontBold
    m_FontName = m_def_FontName
    m_FontSize = m_def_FontSize
    m_FontStrikeThru = m_def_FontStrikeThru
    m_FontUnderLine = m_def_FontUnderLine
    m_FontCharSet = m_def_FontCharSet
    m_FontItalic = m_def_FontItalic
    m_hDC = m_def_hDC
    m_PrinterDefault = m_def_PrinterDefault
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_FromPage = m_def_FromPage
    m_ToPage = m_def_ToPage
    m_Orientation = m_def_Orientation
    m_Action = m_def_Action
    m_HelpCommand = m_def_HelpCommand
    m_HelpContext = m_def_HelpContext
    m_HelpFile = m_def_HelpFile
    m_HelpKey = m_def_HelpKey
    m_PrinterName = m_def_PrinterName
    m_PrinterDefaultSource = m_def_PrinterDefaultSource
    m_FolderPath = m_def_FolderPath
    m_BrowseSpecialFolder = m_def_BrowseSpecialFolder
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Color = PropBag.ReadProperty("Color", m_def_Color)
    m_CancelError = PropBag.ReadProperty("CancelError", m_def_CancelError)
    m_Flags = PropBag.ReadProperty("Flags", m_def_Flags)
    m_DefaultExt = PropBag.ReadProperty("DefaultExt", m_def_DefaultExt)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    m_Filter = PropBag.ReadProperty("Filter", m_def_Filter)
    m_InitDir = PropBag.ReadProperty("InitDir", m_def_InitDir)
    m_DialogTitle = PropBag.ReadProperty("DialogTitle", m_def_DialogTitle)
    m_FileTitle = PropBag.ReadProperty("FileTitle", m_def_FileTitle)
    m_Copies = PropBag.ReadProperty("Copies", m_def_Copies)
    m_FilterIndex = PropBag.ReadProperty("FilterIndex", m_def_FilterIndex)
    m_MaxFileSize = PropBag.ReadProperty("MaxFileSize", m_def_MaxFileSize)
    m_FileOffset = PropBag.ReadProperty("FileOffset", m_def_FileOffset)
    m_ExtensionOffset = PropBag.ReadProperty("ExtensionOffset", m_def_ExtensionOffset)
    m_FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    m_FontSize = PropBag.ReadProperty("FontSize", m_def_FontSize)
    m_FontStrikeThru = PropBag.ReadProperty("FontStrikeThru", m_def_FontStrikeThru)
    m_FontUnderLine = PropBag.ReadProperty("FontUnderLine", m_def_FontUnderLine)
    m_FontCharSet = PropBag.ReadProperty("FontCharSet", m_def_FontCharSet)
    m_FontItalic = PropBag.ReadProperty("FontItalic", m_def_FontItalic)
    m_hDC = PropBag.ReadProperty("hDC", m_def_hDC)
    m_PrinterDefault = PropBag.ReadProperty("PrinterDefault", m_def_PrinterDefault)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_FromPage = PropBag.ReadProperty("FromPage", m_def_FromPage)
    m_ToPage = PropBag.ReadProperty("ToPage", m_def_ToPage)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_Action = PropBag.ReadProperty("Action", m_def_Action)
    m_HelpCommand = PropBag.ReadProperty("HelpCommand", m_def_HelpCommand)
    m_HelpContext = PropBag.ReadProperty("HelpContext", m_def_HelpContext)
    m_HelpFile = PropBag.ReadProperty("HelpFile", m_def_HelpFile)
    m_HelpKey = PropBag.ReadProperty("HelpKey", m_def_HelpKey)
    m_PrinterName = PropBag.ReadProperty("PrinterName", m_def_PrinterName)
    m_PrinterDefaultSource = PropBag.ReadProperty("PrinterDefaultSource", m_def_PrinterDefaultSource)
    m_FolderPath = PropBag.ReadProperty("FolderPath", m_def_FolderPath)
    m_BrowseSpecialFolder = PropBag.ReadProperty("BrowseSpecialFolder", m_def_BrowseSpecialFolder)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
    Call PropBag.WriteProperty("CancelError", m_CancelError, m_def_CancelError)
    Call PropBag.WriteProperty("Flags", m_Flags, m_def_Flags)
    Call PropBag.WriteProperty("DefaultExt", m_DefaultExt, m_def_DefaultExt)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Call PropBag.WriteProperty("Filter", m_Filter, m_def_Filter)
    Call PropBag.WriteProperty("InitDir", m_InitDir, m_def_InitDir)
    Call PropBag.WriteProperty("DialogTitle", m_DialogTitle, m_def_DialogTitle)
    Call PropBag.WriteProperty("FileTitle", m_FileTitle, m_def_FileTitle)
    Call PropBag.WriteProperty("Copies", m_Copies, m_def_Copies)
    Call PropBag.WriteProperty("FilterIndex", m_FilterIndex, m_def_FilterIndex)
    Call PropBag.WriteProperty("MaxFileSize", m_MaxFileSize, m_def_MaxFileSize)
    Call PropBag.WriteProperty("FileOffset", m_FileOffset, m_def_FileOffset)
    Call PropBag.WriteProperty("ExtensionOffset", m_ExtensionOffset, m_def_ExtensionOffset)
    Call PropBag.WriteProperty("FontBold", m_FontBold, m_def_FontBold)
    Call PropBag.WriteProperty("FontName", m_FontName, m_def_FontName)
    Call PropBag.WriteProperty("FontSize", m_FontSize, m_def_FontSize)
    Call PropBag.WriteProperty("FontStrikeThru", m_FontStrikeThru, m_def_FontStrikeThru)
    Call PropBag.WriteProperty("FontUnderLine", m_FontUnderLine, m_def_FontUnderLine)
    Call PropBag.WriteProperty("FontCharSet", m_FontCharSet, m_def_FontCharSet)
    Call PropBag.WriteProperty("FontItalic", m_FontItalic, m_def_FontItalic)
    Call PropBag.WriteProperty("hDC", m_hDC, m_def_hDC)
    Call PropBag.WriteProperty("PrinterDefault", m_PrinterDefault, m_def_PrinterDefault)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("FromPage", m_FromPage, m_def_FromPage)
    Call PropBag.WriteProperty("ToPage", m_ToPage, m_def_ToPage)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("Action", m_Action, m_def_Action)
    Call PropBag.WriteProperty("HelpCommand", m_HelpCommand, m_def_HelpCommand)
    Call PropBag.WriteProperty("HelpContext", m_HelpContext, m_def_HelpContext)
    Call PropBag.WriteProperty("HelpFile", m_HelpFile, m_def_HelpFile)
    Call PropBag.WriteProperty("HelpKey", m_HelpKey, m_def_HelpKey)
    Call PropBag.WriteProperty("PrinterName", m_PrinterName, m_def_PrinterName)
    Call PropBag.WriteProperty("PrinterDefaultSource", m_PrinterDefaultSource, m_def_PrinterDefaultSource)

    Call PropBag.WriteProperty("FolderPath", m_FolderPath, m_def_FolderPath)
    Call PropBag.WriteProperty("BrowseSpecialFolder", m_BrowseSpecialFolder, m_def_BrowseSpecialFolder)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get Color() As OLE_COLOR
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As OLE_COLOR)
    m_Color = New_Color
    PropertyChanged "Color"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CancelError() As Boolean
    CancelError = m_CancelError
End Property

Public Property Let CancelError(ByVal New_CancelError As Boolean)
    m_CancelError = New_CancelError
    PropertyChanged "CancelError"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Flags() As Long
    Flags = m_Flags
End Property

Public Property Let Flags(ByVal New_Flags As Long)
    m_Flags = New_Flags
    PropertyChanged "Flags"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DefaultExt() As String
    DefaultExt = m_DefaultExt
End Property

Public Property Let DefaultExt(ByVal New_DefaultExt As String)
    m_DefaultExt = New_DefaultExt
    PropertyChanged "DefaultExt"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Filter() As String
    Filter = m_Filter
End Property

Public Property Let Filter(ByVal New_Filter As String)
    New_Filter = Replace(New_Filter, "|", Chr$(0))
    m_Filter = New_Filter
    PropertyChanged "Filter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get InitDir() As String
    InitDir = m_InitDir
End Property

Public Property Let InitDir(ByVal New_InitDir As String)
    m_InitDir = New_InitDir
    PropertyChanged "InitDir"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get DialogTitle() As String
    DialogTitle = m_DialogTitle
End Property

Public Property Let DialogTitle(ByVal New_DialogTitle As String)
    m_DialogTitle = New_DialogTitle
    PropertyChanged "DialogTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get FileTitle() As String
Attribute FileTitle.VB_MemberFlags = "400"
    FileTitle = m_FileTitle
End Property

Public Property Let FileTitle(ByVal New_FileTitle As String)
    If Ambient.UserMode = False Then Err.Raise 387
    m_FileTitle = New_FileTitle
    PropertyChanged "FileTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get Copies() As Integer
    Copies = m_Copies
End Property

Public Property Let Copies(ByVal New_Copies As Integer)
    m_Copies = New_Copies
    PropertyChanged "Copies"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FilterIndex() As Integer
    FilterIndex = m_FilterIndex
End Property

Public Property Let FilterIndex(ByVal New_FilterIndex As Integer)
    m_FilterIndex = New_FilterIndex
    PropertyChanged "FilterIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,260
Public Property Get MaxFileSize() As Integer
    MaxFileSize = m_MaxFileSize
End Property

Public Property Let MaxFileSize(ByVal New_MaxFileSize As Integer)
    m_MaxFileSize = New_MaxFileSize
    PropertyChanged "MaxFileSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,2,0
Public Property Get FileOffset() As Integer
Attribute FileOffset.VB_MemberFlags = "400"
    FileOffset = m_FileOffset
End Property

Public Property Let FileOffset(ByVal New_FileOffset As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    m_FileOffset = New_FileOffset
    PropertyChanged "FileOffset"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,2,0
Public Property Get ExtensionOffset() As Integer
Attribute ExtensionOffset.VB_MemberFlags = "400"
    ExtensionOffset = m_ExtensionOffset
End Property

Public Property Let ExtensionOffset(ByVal New_ExtensionOffset As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    m_ExtensionOffset = New_ExtensionOffset
    PropertyChanged "ExtensionOffset"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    m_FontSize = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = m_FontStrikeThru
End Property

Public Property Let FontStrikeThru(ByVal New_FontStrikeThru As Boolean)
    m_FontStrikeThru = New_FontStrikeThru
    PropertyChanged "FontStrikeThru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontUnderLine() As Boolean
    FontUnderLine = m_FontUnderLine
End Property

Public Property Let FontUnderLine(ByVal New_FontUnderLine As Boolean)
    m_FontUnderLine = New_FontUnderLine
    PropertyChanged "FontUnderLine"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FontCharSet() As Long
    FontCharSet = m_FontCharSet
End Property

Public Property Let FontCharSet(ByVal New_FontCharSet As Long)
    m_FontCharSet = New_FontCharSet
    PropertyChanged "FontCharSet"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontItalic() As Boolean
    FontItalic = m_FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    m_FontItalic = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get hDC() As Long
Attribute hDC.VB_MemberFlags = "400"
    hDC = m_hDC
End Property

Public Property Let hDC(ByVal New_hDC As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_hDC = New_hDC
    PropertyChanged "hDC"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get PrinterDefault() As Boolean
    PrinterDefault = m_PrinterDefault
End Property

Public Property Let PrinterDefault(ByVal New_PrinterDefault As Boolean)
    m_PrinterDefault = New_PrinterDefault
    PropertyChanged "PrinterDefault"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Min() As Integer
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Integer)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get FromPage() As Integer
    FromPage = m_FromPage
End Property

Public Property Let FromPage(ByVal New_FromPage As Integer)
    m_FromPage = New_FromPage
    PropertyChanged "FromPage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ToPage() As Integer
    ToPage = m_ToPage
End Property

Public Property Let ToPage(ByVal New_ToPage As Integer)
    m_ToPage = New_ToPage
    PropertyChanged "ToPage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Orientation() As PrinterOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As PrinterOrientationConstants)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,2,0
Public Property Get Action() As Integer
Attribute Action.VB_MemberFlags = "400"
    Action = m_Action
End Property

Public Property Let Action(ByVal New_Action As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    If New_Action < 0 Or New_Action > 6 Then Err.Raise 380
    m_Action = New_Action
    Select Case New_Action
     Case 0
      'Do nothing
     Case 1
      ShowOpen
     Case 2
      ShowSave
     Case 3
      ShowColor
     Case 4
      ShowFont
     Case 5
      ShowPrinter
     Case 6
      ShowHelp
    End Select
    
    PropertyChanged "Action"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get HelpCommand() As Integer
    HelpCommand = m_HelpCommand
End Property

Public Property Let HelpCommand(ByVal New_HelpCommand As Integer)
    m_HelpCommand = New_HelpCommand
    PropertyChanged "HelpCommand"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get HelpContext() As Long
    HelpContext = m_HelpContext
End Property

Public Property Let HelpContext(ByVal New_HelpContext As Long)
    m_HelpContext = New_HelpContext
    PropertyChanged "HelpContext"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get HelpFile() As String
    HelpFile = m_HelpFile
End Property

Public Property Let HelpFile(ByVal New_HelpFile As String)
    m_HelpFile = New_HelpFile
    PropertyChanged "HelpFile"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get HelpKey() As Variant
    HelpKey = m_HelpKey
End Property

Public Property Let HelpKey(ByVal New_HelpKey As Variant)
    m_HelpKey = New_HelpKey
    PropertyChanged "HelpKey"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=13,0,0,
'Public Property Get PrinterName() As String
'    PrinterName = m_PrinterName
'End Property
'
'Public Property Let PrinterName(ByVal New_PrinterName As String)
'    m_PrinterName = New_PrinterName
'    PropertyChanged "PrinterName"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Get PrinterName() As String
Attribute PrinterName.VB_MemberFlags = "400"
    PrinterName = m_PrinterName
End Property

Public Property Let PrinterName(ByVal New_PrinterName As String)
    If Ambient.UserMode = False Then Err.Raise 387
    m_PrinterName = New_PrinterName
    PropertyChanged "PrinterName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get PrinterDefaultSource() As Long
Attribute PrinterDefaultSource.VB_MemberFlags = "400"
    PrinterDefaultSource = m_PrinterDefaultSource
End Property

Public Property Let PrinterDefaultSource(ByVal New_PrinterDefaultSource As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    m_PrinterDefaultSource = New_PrinterDefaultSource
    PropertyChanged "PrinterDefaultSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraCollate() As Integer
    PrinterExtraCollate = m_PrinterExtraCollate
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraColor() As Integer
    PrinterExtraColor = m_PrinterExtraColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraDisplayFlags() As Long
    PrinterExtraDisplayFlags = m_PrinterExtraDisplayFlags
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraDriverVersion() As Integer
    PrinterExtraDriverVersion = m_PrinterExtraDriverVersion
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraDuplex() As Integer
    PrinterExtraDuplex = m_PrinterExtraDuplex
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraFields() As Long
    PrinterExtraFields = m_PrinterExtraFields
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraPaperLength() As Integer
    PrinterExtraPaperLength = m_PrinterExtraPaperLength
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraPaperSize() As Integer
    PrinterExtraPaperSize = m_PrinterExtraPaperSize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraPaperWidth() As Integer
    PrinterExtraPaperWidth = m_PrinterExtraPaperWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraPrintQuality() As Integer
    PrinterExtraPrintQuality = m_PrinterExtraPrintQuality
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraScale() As Integer
    PrinterExtraScale = m_PrinterExtraScale
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraSpecVersion() As Integer
    PrinterExtraSpecVersion = m_PrinterExtraSpecVersion
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraTTOption() As Integer
    PrinterExtraTTOption = m_PrinterExtraTTOption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraYResolution() As Integer
    PrinterExtraYResolution = m_PrinterExtraYResolution
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraExtra() As String
    PrinterExtraExtra = m_PrinterExtraExtra
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraDefault() As Integer
    PrinterExtraDefault = m_PrinterExtraDefault
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraDeviceOffset() As Integer
    PrinterExtraDeviceOffset = m_PrinterExtraDeviceOffset
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraDriverOffset() As Integer
    PrinterExtraDriverOffset = m_PrinterExtraDriverOffset
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PrinterExtraOutputOffset() As Integer
    PrinterExtraOutputOffset = m_PrinterExtraOutputOffset
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FolderPath() As String
    FolderPath = m_FolderPath
End Property

Public Property Let FolderPath(ByVal New_FolderPath As String)
    m_FolderPath = New_FolderPath
    PropertyChanged "FolderPath"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BrowseSpecialFolder() As BrowseForFolderSpecialFolder
    BrowseSpecialFolder = m_BrowseSpecialFolder
End Property

Public Property Let BrowseSpecialFolder(ByVal New_BrowseSpecialFolder As BrowseForFolderSpecialFolder)
    m_BrowseSpecialFolder = New_BrowseSpecialFolder
    PropertyChanged "BrowseSpecialFolder"
End Property

