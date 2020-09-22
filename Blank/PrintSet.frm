VERSION 5.00
Begin VB.Form PrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Page Setup"
   ClientHeight    =   4755
   ClientLeft      =   1125
   ClientTop       =   2040
   ClientWidth     =   8010
   Icon            =   "PrintSet.frx":0000
   LinkTopic       =   "PrintSet"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleMode       =   0  'User
   ScaleTop        =   400
   ScaleWidth      =   8010
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame4 
      Caption         =   "Preview"
      Height          =   3855
      Index           =   1
      Left            =   5400
      TabIndex        =   27
      Top             =   195
      Visible         =   0   'False
      Width           =   2430
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1485
         Index           =   1
         Left            =   240
         Picture         =   "PrintSet.frx":030A
         ScaleHeight     =   1485
         ScaleWidth      =   1950
         TabIndex        =   28
         Top             =   1170
         Width           =   1950
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   1485
         Index           =   1
         Left            =   375
         ScaleHeight     =   1485
         ScaleWidth      =   1935
         TabIndex        =   29
         Top             =   1275
         Width           =   1935
      End
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1260
      TabIndex        =   24
      Text            =   "Page &p"
      Top             =   3735
      Width           =   3915
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   1260
      TabIndex        =   23
      Text            =   "&f"
      Top             =   3270
      Width           =   3915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Printer..."
      Height          =   345
      Left            =   6705
      TabIndex        =   6
      Top             =   4245
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5490
      TabIndex        =   5
      Top             =   4245
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   345
      Left            =   4275
      TabIndex        =   4
      Top             =   4245
      Width           =   1125
   End
   Begin VB.Frame Frame4 
      Caption         =   "Preview"
      Height          =   3855
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   195
      Width           =   2430
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1935
         Index           =   0
         Left            =   465
         Picture         =   "PrintSet.frx":9AE4
         ScaleHeight     =   1935
         ScaleWidth      =   1500
         TabIndex        =   25
         Top             =   945
         Width           =   1500
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   1935
         Index           =   0
         Left            =   585
         ScaleHeight     =   1935
         ScaleWidth      =   1500
         TabIndex        =   26
         Top             =   1065
         Width           =   1500
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Margins (inches)"
      Height          =   1365
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   3420
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   2610
         TabIndex        =   20
         Text            =   "1"""
         Top             =   810
         Width           =   645
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   2610
         TabIndex        =   19
         Text            =   "0.75"""
         Top             =   315
         Width           =   630
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Text            =   "1"""
         Top             =   825
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   900
         TabIndex        =   17
         Text            =   "0.75"""
         Top             =   315
         Width           =   630
      End
      Begin VB.Label Label6 
         Caption         =   "&Bottom:"
         Height          =   210
         Left            =   1890
         TabIndex        =   16
         Top             =   855
         Width           =   705
      End
      Begin VB.Label Label5 
         Caption         =   "&Right:"
         Height          =   240
         Left            =   1890
         TabIndex        =   15
         Top             =   390
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "&Top:"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   855
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "&Left:"
         Height          =   240
         Left            =   180
         TabIndex        =   13
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Orientation"
      Height          =   1365
      Left            =   180
      TabIndex        =   1
      Top             =   1680
      Width           =   1440
      Begin VB.OptionButton Option2 
         Caption         =   "L&andscape"
         Height          =   345
         Left            =   180
         TabIndex        =   10
         Top             =   810
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "P&ortrait"
         Height          =   330
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paper"
      Height          =   1365
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   5040
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         Top             =   825
         Width           =   3600
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1260
         Sorted          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Heeeelp!!!"
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label Label2 
         Caption         =   "&Source:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Si&ze:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   390
         Width           =   675
      End
   End
   Begin VB.Label Label8 
      Caption         =   "&Footer:"
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   3750
      Width           =   630
   End
   Begin VB.Label Label7 
      Caption         =   "&Header:"
      Height          =   240
      Left            =   180
      TabIndex        =   21
      Top             =   3285
      Width           =   720
   End
End
Attribute VB_Name = "PrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DeviceCapabilities Lib "winspool.drv" _
   Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, _
   ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, _
   ByVal dev As Long) As Long

Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12
Private Const DC_PAPERS = 2
Private Const DC_PAPERNAMES = 16
Dim NumPrinters As Long
Dim PrintInfo() As PRINTER_INFO_2
Dim CurrentPrinter As String


Dim Original As String

Public Sub LoadPapers()

Dim i As Integer
Dim Names As String
Dim CapRet As Long
Dim Papers As Long
Dim PaperName As String
Dim First As String

Combo1.Clear
Papers = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERNAMES, ByVal vbNullString, 0)
Names = StringBuffer(Papers * 64)
CapRet = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERNAMES, ByVal Names, 0)
    
    For i = 1 To Papers
    PaperName = Mid(Names, 64 * (i - 1) + 1, 64)
    PaperName = Left(PaperName, InStr(1, PaperName, Chr(0)) - 1)
    Combo1.AddItem PaperName
    If i = 1 Then First = PaperName
    Next

Combo1 = First
End Sub

Public Sub LoadBins()

Dim i As Integer
Dim Names As String
Dim CapRet As Long
Dim Bins As Long
Dim BinName As String
Dim First As String

Combo2.Clear
Bins = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINS, ByVal vbNullString, 0)
Names = StringBuffer(Bins * 24)
CapRet = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_BINNAMES, ByVal Names, 0)
    
    For i = 1 To Bins
    BinName = Mid(Names, 24 * (i - 1) + 1, 24)
    BinName = Left(BinName, InStr(1, BinName, Chr(0)) - 1)
    Combo2.AddItem BinName
    If i = 1 Then First = BinName
    Next
Combo2 = First
End Sub


Private Sub ShowProperties(hPrinter As Long)

OpenPrinter Printer.DeviceName, hPrinter, ByVal 0&
PrinterProperties Me.hwnd, hPrinter
ClosePrinter hPrinter

End Sub


Private Sub Command1_Click()

Dim Box As TextBox
Dim Spot As Integer
Dim AlignDone As Boolean

If Option1 Then
Printer.Orientation = 1
Else
Printer.Orientation = 2
End If

LeftMargin = Val(Text1) * 720
RightMargin = Val(Text3) * 720
TopMargin = Val(Text2) * 1440
BottomMargin = Val(Text4) * 1440

Do
DoEvents
Spot = InStr(Spot + 1, Text5, "&")
If Spot > 0 And Len(Text5) > Spot Then
    Select Case LCase(Mid(Text5, Spot + 1, 1))
        Case "f"
        HeaderText = TitleFromPath(FileName)
        Case "d"
        HeaderText = Format(Date, "Short Date")
        Case "t"
        HeaderText = Format(Now, "h:m:s")
        End Select
        
        If Not AlignDone Then
        Select Case LCase(Mid(Text5, Spot + 1, 1))
            Case "l"
            HAlign = 4
            AlignDone = True
            Case "c"
            HAlign = 2
            AlignDone = True
            Case "r"
            HAlign = 1
            AlignDone = True
        End Select
    End If
End If
Loop While Spot > 0
Spot = 0
AlignDone = False

Do
DoEvents
Spot = InStr(Spot + 1, Text6, "&")
If Spot > 0 And Len(Text6) > Spot Then
    Select Case LCase(Mid(Text6, Spot + 1, 1))
        Case "f"
        FooterText = TitleFromPath(FileName)
        Case "d"
        FooterText = Format(Date, "Short Date")
        Case "t"
        FooterText = Format(Now, "h:m:s")
    End Select
    If Not AlignDone Then
        Select Case LCase(Mid(Text6, Spot + 1, 1))
            Case "l"
            FAlign = 4
            AlignDone = True
            Case "c"
            FAlign = 2
            AlignDone = True
            Case "r"
            FAlign = 1
            AlignDone = True
        End Select
    End If
End If
Loop While Spot > 0
Unload Me
End Sub

Private Sub Command2_Click()

Dim P As Printer

For Each P In Printers
    If P.DeviceName = Original Then
    Set P = Printer
    Exit For
    End If
Next

Unload Me
End Sub

Private Sub Command3_Click()

PrintSelect.Show vbModal, Me

End Sub

Private Sub Form_Load()
MousePointer = vbHourglass
Original = Printer.DeviceName
LoadBins
LoadPapers
MousePointer = vbDefault
End Sub

Private Sub Option1_Click()
Frame4(0).Visible = True
Frame4(1).Visible = False
End Sub

Private Sub Option2_Click()
Frame4(1).Visible = True
Frame4(0).Visible = False
End Sub





