Attribute VB_Name = "SearchDir"
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long


Private Const MAX_PATH As Long = 260
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FileTime
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FileTime
  ftLastAccessTime As FileTime
  ftLastWriteTime As FileTime
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Public ThePath As String

Public Function GetPath(ByVal Name As String, Optional Path As Variant) As String

If IsMissing(Path) Then
Path = SpecialFolder(DOCUMENTS)
End If

GetFullPath Name, Path
GetPath = ThePath
ThePath = ""

End Function

Public Function IsFile(FileString As String) As Boolean
Dim FileNumber As Integer 'The Dir function may be
On Error Resume Next 'in use so use this.
FileNumber = FreeFile()
Open FileString For Input As #FileNumber
If Err Then
IsFile = False
Exit Function
End If
IsFile = True
Close #FileNumber
End Function

Public Function GetFullPath(ByVal Name As String, ByVal SearchPath As String) As Long

Dim FileResult As Long
Dim DirName As String
Dim WFD As WIN32_FIND_DATA
Dim Result As Long
Dim Attempt As Long
Dim FullPath As String
Dim PathSearch As String
Dim FileString As String

FileString = Dir(SlashPath(SearchPath), 39)
    
    If LCase(FileString) = LCase(Name) Then
    ThePath = SlashPath(SearchPath) & FileString
    Exit Function
    End If
    
    Do While Len(FileString) > 0
    DoEvents
    FileString = Dir
    
        If LCase(FileString) = LCase(Name) Then
        ThePath = SlashPath(SearchPath) & FileString
        Screen.MousePointer = vbDefault
        FindClose FileResult
        Exit Function
        End If
    Loop


FullPath = SlashPath(SearchPath) & "*.*" & Chr(0)
FileResult = FindFirstFile(FullPath, WFD)
Do
DoEvents
If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
DirName = TrimNull(WFD.cFileName)
    If Not (DirName = "." Or DirName = "..") Then
        If Not (WFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
            
            If DirName <> "." Then
            PathSearch = SlashPath(SlashPath(SearchPath) & DirName)
            Else
            PathSearch = SlashPath(SearchPath)
            End If
        
        Attempt = GetFullPath(Name, PathSearch)
        End If
    End If
End If
WFD.cFileName = ""
Result = FindNextFile(FileResult, WFD)
Loop Until Result = 0

FindClose FileResult
End Function
Public Function SlashPath(ByVal sPath As String) As String
If Len(sPath) = 0 Then Exit Function 'adds backslash
If Right(sPath, 1) = "\" Then   'if it is needed
SlashPath = sPath
Exit Function
Else
SlashPath = sPath & "\"
End If
End Function

Public Function TrimNull(ByVal Full As String) As String
Dim Nul As Integer
Nul = InStr(Full, Chr(0))
If Nul > 0 Then
TrimNull = Left(Full, Nul - 1)
Else
TrimNull = Full
End If
End Function
 
