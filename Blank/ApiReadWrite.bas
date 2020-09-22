Attribute VB_Name = "ApiReadWrite"
Option Explicit

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Enum FileConstant
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_ALL_FILES = &H7
End Enum
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const OPEN_EXISTING = 3
Private Const INVALID_HANDLE_VALUE = -1

Private Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long


Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
   (ByVal hFile As Long, lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long
    
Private Declare Function CreateFile Lib _
   "kernel32" Alias "CreateFileA" _
   (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As FileConstant, _
   ByVal hTemplateFile As Long) As Long
   
Private Declare Function FlushFileBuffers Lib "kernel32" _
   (ByVal hFile As Long) As Long

Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_MOVE = &H1
Private Const FO_RENAME = &H4
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_FILESONLY = &H80
Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_SILENT = &H4
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_WANTMAPPINGHANDLE = &H20

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
   (lpFileOp As SHFILEOPSTRUCT) As Long
Private FC As FileConstant
Public Function FileDelete(PathName As String) As Long

Dim FileOp As SHFILEOPSTRUCT

With FileOp
    .hwnd = 0
    .wFunc = FO_DELETE
    .pFrom = PathName & Chr(0)
    .fFlags = FOF_NOCONFIRMATION + FOF_SILENT
End With
FileDelete = SHFileOperation(FileOp)
End Function

Public Function FileIsOpen(PathName As String) As Boolean
Dim FileHandle As Long

If Dir(PathName, 39) = "" Then
FileIsOpen = False
Exit Function
End If

FileHandle = CreateFile(PathName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM, 0)

If FileHandle = -1 Then
FileIsOpen = True
Else
FileIsOpen = False
End If

CloseHandle FileHandle

End Function

Public Function FileWrite(PathName As String, Content As String) As Long
    
Dim FileHandle As Long
Dim ByteArray() As Byte
Dim Length As Long
Dim BytesWritten As Long
Dim Success As Long
Dim i As Long

FileHandle = CreateFile(PathName, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, FC, 0)
If FileHandle <> INVALID_HANDLE_VALUE Then
Length = Len(Content)
    If Length > 0 Then
    ReDim ByteArray(1 To Length) As Byte
        For i = 1 To Length
        ByteArray(i) = Asc(Mid(Content, i, 1))
        Next
    
    Success = WriteFile(FileHandle, ByteArray(1), UBound(ByteArray), BytesWritten, 0)
        If Success <> 0 Then
        Success = FlushFileBuffers(FileHandle)
        End If
    End If
Success = CloseHandle(FileHandle)
FC = FILE_ATTRIBUTE_NORMAL
FileWrite = BytesWritten
End If
End Function

Public Function Contents(PathName As String) As String
    
Dim FileHandle As Long
Dim ByteArray() As Byte
Dim FileLength As Long
Dim Result As Long
Dim FileWasOpen As String
Dim Attributes As Long

If Dir(PathName, 39) = "" Then
Contents = ""
Exit Function
End If

Attributes = GetAttr(PathName)
If Attributes = 0 Then
FC = FILE_ATTRIBUTE_NORMAL
Else
FC = Attributes
End If

FileLength = FileLen(PathName)
FileHandle = CreateFile(PathName, GENERIC_READ, 0, 0, OPEN_EXISTING, FC, 0)

If FileLength > 0 Then
ReDim ByteArray(1 To FileLength) As Byte
ReadFile FileHandle, ByteArray(1), UBound(ByteArray), Result, ByVal 0&
    If Result = UBound(ByteArray) Then
    Contents = StrConv(ByteArray, vbUnicode)
    Else
    Contents = ""
    End If
Else
Contents = ""
End If

Result = CloseHandle(FileHandle)
Out:
End Function
