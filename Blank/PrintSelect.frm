VERSION 5.00
Begin VB.Form PrintSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Page Setup"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6435
   ClipControls    =   0   'False
   Icon            =   "PrintSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5160
      TabIndex        =   2
      Top             =   2370
      Width           =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4020
      TabIndex        =   1
      Top             =   2370
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printer"
      Height          =   2070
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   6120
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   990
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   330
         Width           =   3420
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Properties"
         Height          =   345
         Left            =   4560
         TabIndex        =   3
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   285
         Left            =   1020
         TabIndex        =   13
         Top             =   1635
         Width           =   4500
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   285
         Left            =   1020
         TabIndex        =   12
         Top             =   1350
         Width           =   4500
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   285
         Left            =   1020
         TabIndex        =   11
         Top             =   1065
         Width           =   4500
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   285
         Left            =   1020
         TabIndex        =   10
         Top             =   780
         Width           =   4500
      End
      Begin VB.Label Label2 
         Caption         =   "Status:"
         Height          =   285
         Left            =   195
         TabIndex        =   9
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Type:"
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   1065
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Where:"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "Comment:"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   1635
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "&Name"
         Height          =   255
         Left            =   165
         TabIndex        =   5
         Top             =   390
         Width           =   690
      End
   End
End
Attribute VB_Name = "PrintSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrinterApi Lib "winspool.drv" Alias _
    "GetPrinterA" (ByVal hPrinter As Long, _
    ByVal Level As Long, _
    Buffer As Long, _
    ByVal pbSize As Long, _
    pbSizeNeeded As Long) As Long

Private Type PRINTER_DEFAULTS
  pDatatype As String
  pDevMode As DEVMODE
  DesiredAccess As Long
End Type

Private Declare Function OpenPrinter Lib "winspool.drv" _
    Alias "OpenPrinterA" (ByVal pPrinterName As String, _
    phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long

Private Enum Printer_Status
   PRINTER_STATUS_READY = &H0
   PRINTER_STATUS_PAUSED = &H1
   PRINTER_STATUS_ERROR = &H2
   PRINTER_STATUS_PENDING_DELETION = &H4
   PRINTER_STATUS_PAPER_JAM = &H8
   PRINTER_STATUS_PAPER_OUT = &H10
   PRINTER_STATUS_MANUAL_FEED = &H20
   PRINTER_STATUS_PAPER_PROBLEM = &H40
   PRINTER_STATUS_OFFLINE = &H80
   PRINTER_STATUS_IO_ACTIVE = &H100
   PRINTER_STATUS_BUSY = &H200
   PRINTER_STATUS_PRINTING = &H400
   PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
   PRINTER_STATUS_NOT_AVAILABLE = &H1000
   PRINTER_STATUS_WAITING = &H2000
   PRINTER_STATUS_PROCESSING = &H4000
   PRINTER_STATUS_INITIALIZING = &H8000
   PRINTER_STATUS_WARMING_UP = &H10000
   PRINTER_STATUS_TONER_LOW = &H20000
   PRINTER_STATUS_NO_TONER = &H40000
   PRINTER_STATUS_PAGE_PUNT = &H80000
   PRINTER_STATUS_USER_INTERVENTION = &H100000
   PRINTER_STATUS_OUT_OF_MEMORY = &H200000
   PRINTER_STATUS_DOOR_OPEN = &H400000
   PRINTER_STATUS_SERVER_UNKNOWN = &H800000
   PRINTER_STATUS_POWER_SAVE = &H1000000
End Enum

 Private Type PRINTER_INFO_2
   pServerName As String
   pPrinterName As String
   pShareName As String
   pPortName As String
   pDriverName As String
   pComment As String
   pLocation As String
   pDevMode As Long
   pSepFile As String
   pPrintProcessor As String
   pDatatype As String
   pParameters As String
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   JobsCount As Long
   AveragePPM As Long
End Type

Dim Original As String
Private Sub ShowProperties()
Dim hPrinter As Long
Dim pDev As PRINTER_DEFAULTS
OpenPrinter Printer.DeviceName, hPrinter, pDev
PrinterProperties Me.hwnd, hPrinter
ClosePrinter hPrinter
End Sub
Private Function GetPrinterInfo() As Long

Dim lRet As Long
Dim SizeNeeded As Long
Dim Buffer() As Long
Dim pDef As PRINTER_DEFAULTS
Dim hPrinter As Long
Dim PrintInfo As PRINTER_INFO_2

lRet = OpenPrinter(Printer.DeviceName, hPrinter, pDef)
ReDim Preserve Buffer(0 To 1) As Long
lRet = GetPrinterApi(hPrinter, 2, Buffer(0), UBound(Buffer), SizeNeeded)
ReDim Preserve Buffer(0 To (SizeNeeded / 4) + 3) As Long
lRet = GetPrinterApi(hPrinter, 2, Buffer(0), UBound(Buffer) * 4, SizeNeeded)

ClosePrinter hPrinter
With PrintInfo
   .pPrinterName = StringFromPointer(Buffer(1), 1024)
   .pPortName = StringFromPointer(Buffer(3), 1024)
   .pDriverName = StringFromPointer(Buffer(4), 1024)
   .pComment = StringFromPointer(Buffer(5), 1024)
   .pLocation = StringFromPointer(Buffer(6), 1024)
   .pDevMode = Buffer(7)
   .Status = Buffer(18)
End With
Label6 = Status(PrintInfo.Status)
Label7 = PrintInfo.pDriverName
Label8 = PrintInfo.pPortName
Label9 = PrintInfo.pComment
End Function



  




Private Function Status(StatusNum As Long) As String
Select Case StatusNum
Case &H0
Status = "Ready"
Case &H1
Status = "Paused"
Case &H2
Status = "Error"
Case &H4
Status = "Pending deletion"
Case &H8
Status = "Paper jam"
Case &H10
Status = "Out of paper"
Case &H20
Status = "Manual feed"
Case &H40
Status = "Paper problem"
Case &H80
Status = "Offline"
Case &H100
Status = "I/O active"
Case &H200
Status = "Busy"
Case &H400
Status = "Printing"
Case &H800
Status = "Output bin full"
Case &H1000
Status = "Not available"
Case &H2000
Status = "Waiting"
Case &H4000
Status = "Processing"
Case &H8000
Status = "Initializing"
Case &H10000
Status = "Warming up"
Case &H20000
Status = "Toner low"
Case &H40000
Status = "No toner"
Case &H80000
Status = "Page punt"
Case &H100000
Status = "User intervention"
Case &H200000
Status = "Out of memory"
Case &H400000
Status = "Door open"
Case &H800000
Status = "Server unknown"
Case &H1000000
Status = "Power save"
End Select

End Function

Public Function StringFromPointer(lpString As Long, lMaxLength As Long) As String

Dim sRet As String

If lpString = 0 Then
StringFromPointer = ""
Exit Function
End If

If IsBadStringPtrByLong(lpString, lMaxLength) Then
StringFromPointer = ""
Exit Function
End If

sRet = Space(lMaxLength)
CopyMemory ByVal sRet, ByVal lpString, ByVal Len(sRet)
If Err.LastDllError = 0 Then
    If InStr(sRet, Chr(0)) > 0 Then
    sRet = Left(sRet, InStr(sRet, Chr(0)) - 1)
    End If
End If

StringFromPointer = sRet
End Function

Private Sub Combo1_Click()

Dim Default As String
Dim p As Printer
For Each p In Printers
    If p.DeviceName = Combo1 Then
    Set Printer = p
    Exit For
    End If
Next
GetPrinterInfo
End Sub


Private Sub Command1_Click()
PrintSet.LoadBins
PrintSet.LoadPapers
Unload Me
End Sub

Private Sub Command2_Click()
Dim p As Printer
For Each p In Printers
    If p.DeviceName = Original Then
    Set Printer = p
    Exit For
    End If
Next
Unload Me
End Sub


Private Sub Command3_Click()
ShowProperties
End Sub

Private Sub Form_Load()

Dim p As Printer
Dim pDefault As String
Original = Printer.DeviceName
pDefault = Original
    For Each p In Printers
    Combo1.AddItem p.DeviceName
    Next p
Combo1.Text = pDefault
GetPrinterInfo
End Sub


