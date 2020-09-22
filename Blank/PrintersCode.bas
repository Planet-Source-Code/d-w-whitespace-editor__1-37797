Attribute VB_Name = "PrinterCode"
Option Explicit

Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Long) As Long

Public Const CCHDEVICENAME = 32
' size of a device name string
Public Const CCHFORMNAME = 32
' size of a form name string
Public Type DEVMODE
dmDeviceName As String * CCHDEVICENAME
'The name of the device.
dmSpecVersion As Integer
'The version number of the device's initialization information specification.
dmDriverVersion As Integer
'The version number of the device driver.
dmSize As Integer
'The size of the structure, in bytes.
dmDriverExtra As Integer
'The number of bytes of information trailing the structure in memory.
dmFields As Long
'A combination of  flags specifying which of the rest of the structure's members contain information about the device:
dmOrientation As Integer
'Contains DM_ORIENTATION information
dmPaperSize As Integer
'Contains DM_PAPERSIZE information
dmPaperLength As Integer
'Contains DM_PAPERLENGTH information
dmPaperWidth As Integer
'Contains DM_PAPERWIDTH information
dmScale As Integer
'Contains DM_SCALE information
dmCopies As Integer
'Contains DM_COPIES information
dmDefaultSource As Integer
'Contains DM_DEFAULTSOURCE information
dmPrintQuality As Integer
'Contains DM_PRINTQUALITY information
dmColor As Integer
'Contains DM_COLOR information
dmDuplex As Integer
'Contains DM_DUPLEX information
dmYResolution As Integer
'Contains DM_YRESOLUTION information
dmTTOption As Integer
'Contains DM_TTOPTION information
dmCollate As Integer
'Contains DM_COLLATE information
dmFormName As String * CCHFORMNAME
'Contains DM_FORMNAME information
dmUnusedPadding As Integer
'Reserved -- set to 0. This member merely takes up space to align other members in memory.
dmBitsPerPel As Integer
'The number of color bits used per pixel on the display device.
dmPelsWidth As Long
'The width of the display, measured in pixels.
dmPelsHeight As Long
'The height of the display, measured in pixels.
dmDisplayFlags As Long
'A combination flags specifying the device's display mode:
'DM_GRAYSCALE The display does not support color. (If this flag is omitted, assume color is supported.)
'DM_INTERLACED The display is interlaced.
dmDisplayFrequency As Long
'The display frequency of the display, measured in Hz.
End Type

' current version of specification
Public Const DM_SPECVERSION = &H320
' field selection bits
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_SCALE = &H10&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_PRINTQUALITY = &H400&
Public Const DM_COLOR = &H800&
Public Const DM_DUPLEX = &H1000&
Public Const DM_YRESOLUTION = &H2000&
Public Const DM_TTOPTION = &H4000&
Public Const DM_COLLATE As Long = &H8000
Public Const DM_FORMNAME As Long = &H10000
' orientation selections
Public Const DMORIENT_PORTRAIT = 1
Public Const DMORIENT_LANDSCAPE = 2
' paper selections
Public Const DMPAPER_LETTER = 1 ' Letter 8 1/2 x 11 in
Public Const DMPAPER_FIRST = DMPAPER_LETTER ' Letter 8 1/2 x 11 in
Public Const DMPAPER_LETTERSMALL = 2 ' Letter Small 8 1/2 x 11 in
Public Const DMPAPER_TABLOID = 3 ' Tabloid 11 x 17 in
Public Const DMPAPER_LEDGER = 4 ' Ledger 17 x 11 in
Public Const DMPAPER_LEGAL = 5 ' Legal 8 1/2 x 14 in
Public Const DMPAPER_STATEMENT = 6 ' Statement 5 1/2 x 8 1/2 in
Public Const DMPAPER_EXECUTIVE = 7 ' Executive 7 1/4 x 10 1/2 in
Public Const DMPAPER_A3 = 8 ' A3 297 x 420 mm
Public Const DMPAPER_A4 = 9 ' A4 210 x 297 mm
Public Const DMPAPER_A4SMALL = 10 ' A4 Small 210 x 297 mm
Public Const DMPAPER_A5 = 11 ' A5 148 x 210 mm
Public Const DMPAPER_B4 = 12 ' B4 250 x 354
Public Const DMPAPER_B5 = 13 ' B5 182 x 257 mm
Public Const DMPAPER_FOLIO = 14 ' Folio 8 1/2 x 13 in
Public Const DMPAPER_QUARTO = 15 ' Quarto 215 x 275 mm
Public Const DMPAPER_10X14 = 16 ' 10x14 in
Public Const DMPAPER_11X17 = 17 ' 11x17 in
Public Const DMPAPER_NOTE = 18 ' Note 8 1/2 x 11 in
Public Const DMPAPER_ENV_9 = 19 ' Envelope #9 3 7/8 x 8 7/8
Public Const DMPAPER_ENV_10 = 20 ' Envelope #10 4 1/8 x 9 1/2
Public Const DMPAPER_ENV_11 = 21 ' Envelope #11 4 1/2 x 10 3/8
Public Const DMPAPER_ENV_12 = 22 ' Envelope #12 4 \276 x 11
Public Const DMPAPER_ENV_14 = 23 ' Envelope #14 5 x 11 1/2
Public Const DMPAPER_CSHEET = 24 ' C size sheet
Public Const DMPAPER_DSHEET = 25 ' D size sheet
Public Const DMPAPER_ESHEET = 26 ' E size sheet
Public Const DMPAPER_ENV_DL = 27 ' Envelope DL 110 x 220mm
Public Const DMPAPER_ENV_C5 = 28 ' Envelope C5 162 x 229 mm
Public Const DMPAPER_ENV_C3 = 29 ' Envelope C3 324 x 458 mm
Public Const DMPAPER_ENV_C4 = 30 ' Envelope C4 229 x 324 mm
Public Const DMPAPER_ENV_C6 = 31 ' Envelope C6 114 x 162 mm
Public Const DMPAPER_ENV_C65 = 32 ' Envelope C65 114 x 229 mm
Public Const DMPAPER_ENV_B4 = 33 ' Envelope B4 250 x 353 mm
Public Const DMPAPER_ENV_B5 = 34 ' Envelope B5 176 x 250 mm
Public Const DMPAPER_ENV_B6 = 35 ' Envelope B6 176 x 125 mm
Public Const DMPAPER_ENV_ITALY = 36 ' Envelope 110 x 230 mm
Public Const DMPAPER_ENV_MONARCH = 37 ' Envelope Monarch 3.875 x 7.5 in
Public Const DMPAPER_ENV_PERSONAL = 38 ' 6 3/4 Envelope 3 5/8 x 6 1/2 in
Public Const DMPAPER_FANFOLD_US = 39 ' US Std Fanfold 14 7/8 x 11 in
Public Const DMPAPER_FANFOLD_STD_GERMAN = 40 ' German Std Fanfold 8 1/2 x 12 in
Public Const DMPAPER_FANFOLD_LGL_GERMAN = 41 ' German Legal Fanfold 8 1/2 x 13 in

Public Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN

Public Const DMPAPER_USER = 256

' bin selections
Public Const DMBIN_UPPER = 1
Public Const DMBIN_FIRST = DMBIN_UPPER

Public Const DMBIN_ONLYONE = 1
Public Const DMBIN_LOWER = 2
Public Const DMBIN_MIDDLE = 3
Public Const DMBIN_MANUAL = 4
Public Const DMBIN_ENVELOPE = 5
Public Const DMBIN_ENVMANUAL = 6
Public Const DMBIN_AUTO = 7
Public Const DMBIN_TRACTOR = 8
Public Const DMBIN_SMALLFMT = 9
Public Const DMBIN_LARGEFMT = 10
Public Const DMBIN_LARGECAPACITY = 11
Public Const DMBIN_CASSETTE = 14
Public Const DMBIN_LAST = DMBIN_CASSETTE

Public Const DMBIN_USER = 256 ' device specific bins start here

' print qualities
Public Const DMRES_DRAFT = (-1)
Public Const DMRES_LOW = (-2)
Public Const DMRES_MEDIUM = (-3)
Public Const DMRES_HIGH = (-4)

' color enable/disable for color printers
Public Const DMCOLOR_MONOCHROME = 1
Public Const DMCOLOR_COLOR = 2

' duplex enable
Public Const DMDUP_SIMPLEX = 1
Public Const DMDUP_VERTICAL = 2
Public Const DMDUP_HORIZONTAL = 3

' TrueType options
Public Const DMTT_BITMAP = 1 ' print TT fonts as graphics
Public Const DMTT_DOWNLOAD = 2 ' download TT fonts as soft fonts
Public Const DMTT_SUBDEV = 3 ' substitute device fonts for TT fonts

' Collation selections
Public Const DMCOLLATE_FALSE = 0
Public Const DMCOLLATE_TRUE = 1

' DEVMODE dmDisplayFlags flags

Public Const DM_GRAYSCALE = &H1
Public Const DM_INTERLACED = &H2

Public Type PRINTER_INFO_2
  pServerName As String
  'The name of the network server which contols the printer, if any.
  pPrinterName As String
  'The name of the printer.
  pShareName As String
  'The name of the sharepoint of the printer on the network, if any.
  pPortName As String
  'A comma-separated list of the printer port(s) the printer is connected to, such as LPT1:.
  pDriverName As String
  'The name of the printer driver.
  pComment As String
  'A comment about or a brief description of the printer.
  pLocation As String
  'The physical location of the printer (usually applies to network printers).
  pDevMode As DEVMODE
  'Various default settings and attributes of the printer.
  pSepFile As String
  'The file that contains the separator page printed between jobs.
  pPrintProcessor As String
  'The name of the print processor the printer uses.
  pDatatype As String
  'The name of the data type used to record the print jobs.
  pParameters As String
  'Default parameters for the print processor.
  pSecurityDescriptor As Long
  'Security information about the printer.
  Attributes As Long
  'One or more of the following flags specifying various attributes of the printer:
  Priority As Long
  'The priority given to the printer by the print spooler.
  DefaultPriority As Long
  'The default priority for a print job.
  StartTime As Long
  'The earliest time the printer will print a job, specified in minutes after midnight UTC (GMT or Zulu time).
  UntilTime As Long
  'The latest time the printer will print a job, specified in minutes after midnight UTC (GMT or Zulu time).
  Status As Long
  'One or more flags specifying the printer's current status (Win NT only supports the PRINTER_STATUS_PAUSED and PRINTER_STATUS_PENDING_DELETION flags)
  cJobs As Long
  'Specifies the number of print jobs that have been queued for the printer.
  AveragePPM As Long
  'The average number of pages the printer can print per minute.
End Type

Public Type DEVNAMES
  wDriverOffset As Integer
  'The offset of the string in extra identifying the name of the device driver filename (without the extension).
  wDeviceOffset As Integer
  'The offset of the string in extra identifying the name of the device.
  wOutputOffset As Integer
  'The offset of the string in extra identifying the output port(s) which the device uses, separated by commas.
  wDefault As Integer
  'If non-zero, the information in the structure identifies the default device of its type. If zero, the information does not necessarily descibe the default device.
  extra As String * 100
  'Buffer which holds the three strings identified by wDriverOffset, wDeviceOffset, and wOutputOffset.
End Type

Public Const DC_PAPERS = 2
Public Const DC_FIELDS = 1
Public Const DC_PAPERSIZE = 3
Public Const DC_MINEXTENT = 4
Public Const DC_MAXEXTENT = 5
Public Const DC_BINS = 6
Public Const DC_DUPLEX = 7
Public Const DC_SIZE = 8
Public Const DC_EXTRA = 9
Public Const DC_VERSION = 10
Public Const DC_DRIVER = 11
Public Const DC_BINNAMES = 12
Public Const DC_ENUMRESOLUTIONS = 13
Public Const DC_FILEDEPENDENCIES = 14
Public Const DC_TRUETYPE = 15
Public Const DC_PAPERNAMES = 16
Public Const DC_ORIENTATION = 17
Public Const DC_COPIES = 18

Public TopMargin As Integer
Public BottomMargin As Integer
Public LeftMargin As Integer
Public RightMargin As Integer
Public HeaderText As String
Public FooterText As String
Public PageNum As Integer
Public HAlign As Integer
Public FAlign As Integer
Public Sub PrintText(TheText As String)

Dim i As Long
Dim j As Long
Dim CurrentWord As String

Screen.MousePointer = vbHourglass
StartNewPage False
i = 1
Do Until i > Len(TheText)
    CurrentWord = ""
    Do Until i > Len(TheText) Or Mid(TheText, i, 1) <= " "
        CurrentWord = CurrentWord & Mid(TheText, i, 1)
        i = i + 1
    Loop
    
    If (Printer.CurrentX + Printer.TextWidth(CurrentWord)) > (Printer.ScaleWidth - RightMargin) Then
        Printer.Print
        If Printer.CurrentY > (Printer.ScaleHeight - BottomMargin) Then
        Printer.Print FooterText;
        StartNewPage
        Else
        Printer.CurrentX = LeftMargin
        End If
    End If
    Printer.Print CurrentWord;
    Do Until i > Len(TheText) Or Mid(TheText, i, 1) > " "
        Select Case Mid(TheText, i, 1)
        Case " "
        Printer.Print " ";
        Case Chr(10)
        Printer.Print
            If Printer.CurrentY > (Printer.ScaleHeight - BottomMargin) Then
            Printer.Print FooterText;
            StartNewPage
            Else
            Printer.CurrentX = LeftMargin
            End If
        Case Chr(9)
        j = (Printer.CurrentX - LeftMargin) / Printer.TextWidth("0")
        j = j + (10 - (j Mod 10))
        Printer.CurrentX = LeftMargin + (j * Printer.TextWidth("0"))
        Case Else
        End Select
        i = i + 1
    Loop
Loop
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(FooterText)) / FAlign
Printer.CurrentY = (Printer.ScaleHeight - BottomMargin - Printer.TextHeight(FooterText))
Printer.Print FooterText;
Printer.EndDoc
Screen.MousePointer = vbDefault
End Sub

Private Sub StartNewPage(Optional Eject As Boolean = True)

Dim Spot As Long

If Eject Then
Printer.NewPage
End If
   
Printer.CurrentY = (TopMargin - Printer.TextHeight(HeaderText)) / HAlign
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(HeaderText)) / HAlign
Printer.Print HeaderText;
    
Printer.CurrentX = LeftMargin
Printer.CurrentY = Printer.ScaleHeight - (BottomMargin / 2)

Printer.CurrentX = LeftMargin
Printer.CurrentY = TopMargin


PageNum = PageNum + 1

Spot = InStr(1, HeaderText, "Page")
If Spot > 0 Then
Replace FooterText, Mid(HeaderText, Spot + 6, 1), CStr(PageNum)
End If

Spot = InStr(1, FooterText, "Page")
If Spot > 0 Then
Replace FooterText, Mid(FooterText, Spot + 6, 1), CStr(PageNum)
End If

End Sub



