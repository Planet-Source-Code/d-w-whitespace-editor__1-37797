Attribute VB_Name = "Window"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Function ExplorerDirectory() As String

Dim TxtLen  As Long
Dim wHwnd As Long
Dim WindowCaption As String

wHwnd = FindWindow("CabinetWClass", vbNullString)

If wHwnd = 0 Then
wHwnd = FindWindow("ExploreWClass", vbNullString)
End If

TxtLen = GetWindowTextLength(wHwnd) + 1
WindowCaption = StringBuffer(TxtLen - 1)
GetWindowText wHwnd, WindowCaption, TxtLen

If Left(WindowCaption, 12) = "Exploring - " Then
WindowCaption = Mid(WindowCaption, 13)
End If

If Len(WindowCaption) > 0 And WindowCaption <> "C:\" Then
    If Dir(WindowCaption, vbDirectory) <> ".." And Dir(WindowCaption, vbDirectory) <> "." Then
    ExplorerDirectory = WindowCaption
    End If
End If
End Function



