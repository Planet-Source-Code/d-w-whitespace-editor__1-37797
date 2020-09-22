Attribute VB_Name = "Blank"
Option Explicit

Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal olestr As Long, ByVal BLen As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes As Long)

Public Marker As String * 8
Public Function BlankToString(BlankString As String) As String

Dim i As Long
Dim n As Long
Dim AscNum As Long
Dim Length As Long
Dim Spot As Long
Dim BlankLetter As String
Dim Buffer As String
BlankString = Mid(BlankString, InStr(1, BlankString, Marker) + 8)
Length = Len(BlankString)
Buffer = StringBuffer(Length / 8)

For n = 1 To Length Step 8
Spot = Spot + 1
BlankLetter = Mid(BlankString, n, 8)

If Len(BlankLetter) <> 8 Then
GoTo Out:
End If

AscNum = 0
    For i = 8 To 1 Step -1
        If Mid(BlankLetter, i, 1) = Chr(&HA0) Then
        AscNum = AscNum + (2 ^ (8 - i))
        End If
    Next
Mid(Buffer, Spot, 1) = Chr(AscNum)
Next
Out:
BlankToString = StringBuffer(Length / 8)
BlankToString = Mid(Buffer, 1, Length / 8)
End Function

Public Function StringToBlank(TheString As String) As String

Dim n As Long
Dim i As Long
Dim AscNum As Long
Dim Length As Long
Dim Spot As Long
Dim Buffer As String

Length = Len(TheString) + 1
Buffer = StringBuffer(Length * 8)
Mid(Buffer, 1, 8) = Marker
Spot = 8
For n = 1 To Length - 1
AscNum = Asc(Mid(TheString, n, 1))
    For i = 7 To 0 Step -1
    Spot = Spot + 1
        If AscNum And (2 ^ i) Then
        Mid(Buffer, Spot, 1) = Chr(&HA0)
        Else
        Mid(Buffer, Spot, 1) = Chr(&H20)
        End If
    Next
Next
StringToBlank = StringBuffer(Length * 8)
StringToBlank = Mid(Buffer, 1, Length * 8)
End Function



Public Function StringBuffer(ByVal Size As Long) As String
Dim Allocate As Long
Dim Buffer As String
Allocate = SysAllocStringByteLen(0, Size * 2)
RtlMoveMemory ByVal VarPtr(Buffer), Allocate, 4
StringBuffer = Buffer
End Function
