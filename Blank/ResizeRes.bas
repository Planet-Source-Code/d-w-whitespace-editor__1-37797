Attribute VB_Name = "ResizeRes"
Option Explicit

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
 lParam As Any) As Long
Public Const LP_HT_CAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Function Resize() As Single
Select Case Screen.Width
Case 9600
Resize = 1
Case 12000
Resize = 1.25
Case 15360
Resize = 1.6
Case 19200
Resize = 2
Case Else
Resize = 1
End Select
End Function


