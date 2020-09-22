VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Notepad"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3885
      TabIndex        =   8
      Top             =   3435
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   35
      Left            =   1170
      TabIndex        =   5
      Top             =   2460
      Width           =   3735
   End
   Begin VB.Label Label7 
      Caption         =   "Resources:"
      Height          =   270
      Left            =   1170
      TabIndex        =   7
      Top             =   3060
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Physical memory available to Windows:"
      Height          =   225
      Left            =   1170
      TabIndex        =   6
      Top             =   2715
      Width           =   3825
   End
   Begin VB.Label Label5 
      Caption         =   "1"
      Height          =   315
      Left            =   1170
      TabIndex        =   4
      Top             =   1785
      Width           =   2040
   End
   Begin VB.Label Label4 
      Caption         =   "This product is licensed to:"
      Height          =   270
      Left            =   1170
      TabIndex        =   3
      Top             =   1425
      Width           =   2820
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright (C) 1991-1998 Microsoft Corp."
      Height          =   270
      Left            =   1170
      TabIndex        =   2
      Top             =   750
      Width           =   3465
   End
   Begin VB.Label Label2 
      Caption         =   "Windows 98"
      Height          =   240
      Left            =   1170
      TabIndex        =   1
      Top             =   495
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Microsoft (R) Notepad"
      Height          =   270
      Left            =   1170
      TabIndex        =   0
      Top             =   255
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
Private Declare Function GetFreeResources Lib "RSRC32.dll" Alias "_MyGetFreeSystemResources32@4" (ByVal lWhat As Long) As Long
Private Const REG_SZ = 1
Private Const LOCALMACHINE = &H80000002

Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long
    PlatformID As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Const GFSR_SYSTEMRESOURCES = 0
Private Const GFSR_GDIRESOURCES = 1
Private Const GFSR_USERRESOURCES = 2
Private Function OwnerName() As String

Dim BufferKey As Long
Dim NameBuffer As String

NameBuffer = StringBuffer(128)
If OSName = "Windows NT" Or OSName = "Windows 2000" Or OSName = "Windows XP" Then
RegOpenKey LOCALMACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", BufferKey
Else
RegOpenKey LOCALMACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", BufferKey
End If

RegQueryValueEx BufferKey, "RegisteredOwner", 0, REG_SZ, NameBuffer, Len(NameBuffer)
OwnerName = NameBuffer
End Function
Public Function OSName() As String

Dim OSV As OSVERSIONINFO

OSV.OSVSize = Len(OSV)
GetVersionEx OSV

If OSV.PlatformID = 1 Then
    Select Case OSV.dwVerMinor
    Case 0
    OSName = "Windows 95"
    Case 10
    OSName = "Windows 98"
    Case 90
    OSName = "Windows ME"
    End Select
ElseIf OSV.PlatformID = 2 Then
    Select Case OSV.dwVerMajor
    Case 3, 4
    OSName = "Windows NT"
    Case 5
        Select Case OSV.dwVerMinor
        Case 0
        OSName = "Windows 2000"
        Case 1
        OSName = "Windows XP"
        End Select
    End Select
End If
If Len(OSName) = 0 Then
OSName = "Unknown"
End If
End Function



Private Function AddCommas(Number As Long) As String

Dim Length As Integer

Length = Len(CStr(Number))
If Length = 5 Then
AddCommas = Left(CStr(Number), 2) & "," & Right(CStr(Number), 3) & " KB"
ElseIf Length = 6 Then
AddCommas = Left(CStr(Number), 3) & "," & Right(CStr(Number), 3) & " KB"
ElseIf Length > 6 Then
AddCommas = Left(CStr(Number), 1) & "," & Mid(CStr(Number), 2, 3) & "," & Right(CStr(Number), 3) & " KB"
Else
AddCommas = "Error"
End If
End Function

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Form_Load()

Dim MS As MEMORYSTATUS
Dim Display As String
Dim System As Integer
Dim User As Integer
Dim GDI As Integer
Dim Average As Long
On Local Error Resume Next
Icon = Editor.Icon
Image1.Picture = Editor.Icon
MS.dwLength = Len(MS)
GlobalMemoryStatus MS
Display = AddCommas(MS.dwTotalPhys / 1024)
System = GetFreeResources(GFSR_SYSTEMRESOURCES)
GDI = GetFreeResources(GFSR_GDIRESOURCES)
User = GetFreeResources(GFSR_USERRESOURCES)
Average = (System + GDI + User) \ 3
Label2 = OSName
Label5 = OwnerName
Label6 = "Physical memory available to Windows:   " & Display
Label7 = "Resources:                                               " & Average & "% Free"
End Sub

