Attribute VB_Name = "SubMain"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC As Long = (-4)

Public PrevProc1 As Long
Public PrevProc2 As Long

Public Const WM_USER = &H400
Public Const EM_CANPASTE = WM_USER + 50
Public Const EM_UNDO = &HC7
Public Const EM_CANUNDO = &HC6
Public Const WM_COPY& = &H301
Public Const WM_CUT& = &H300
Public Const WM_PASTE& = &H302

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206


Public White As Boolean
Public Wrapped As Boolean
Public FileTitle As String
Public FileName As String
Public DataModified As Boolean

Public Function RelayMessage(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim Retval As Long

If uMsg = WM_RBUTTONDOWN Then
With Editor
    If Not Wrapped Then
    .mnuUndoR.Enabled = SendMessage(.Page.hwnd, EM_CANUNDO, vbNull, vbNull)
    .mnuDelR.Enabled = .Page.SelLength > 0
    .mnuAllR.Enabled = Len(.Page) > 0
    Else
    .mnuUndoR.Enabled = SendMessage(.Wrap.hwnd, EM_CANUNDO, vbNull, vbNull)
    .mnuDelR.Enabled = .Wrap.SelLength > 0
    .mnuAllR.Enabled = Len(.Wrap) > 0
    End If
    .mnuCutR.Enabled = .mnuDelR.Enabled
    .mnuCopyR.Enabled = .mnuDelR.Enabled
    .mnuPasteR.Enabled = Len(Clipboard.GetText(vbCFText)) > 0
    .PopupMenu Editor.mnuPopup
End With
Else
    If Wrapped Then
    RelayMessage = CallWindowProc(PrevProc2, hwnd, uMsg, wParam, lParam)
    Else
    RelayMessage = CallWindowProc(PrevProc1, hwnd, uMsg, wParam, lParam)
    End If
End If
End Function

Private Function WindowPosition() As Integer

Dim Pos As Integer

Pos = GetSetting("Editor", "Settings", "Position", 0)
WindowPosition = Pos
Pos = Pos + 1

If Pos > 4 Then Pos = 0

SaveSetting "Editor", "Settings", "Position", Pos
End Function


Public Sub Main()

Dim Pos As Integer
Dim Cmd As String
Dim Path As String
Dim Action As VbMsgBoxResult

Load Editor
PrevProc1 = SetWindowLong(Editor.Page.hwnd, GWL_WNDPROC, AddressOf RelayMessage)
PrevProc2 = SetWindowLong(Editor.Wrap.hwnd, GWL_WNDPROC, AddressOf RelayMessage)
Marker = String(8, Chr(&HA0))

PageNum = 1
HAlign = 2
FAlign = 2

With Editor
.mnuFull.Checked = GetSetting("Editor", "Settings", "Fullscreen", False)
.mnuWhite.Checked = GetSetting("Editor", "Settings", "OpenWhite", False)
.mnuWrap.Checked = GetSetting("Editor", "Settings", "Wrap", False)
.mnuDir.Checked = GetSetting("Editor", "Settings", "Previous", False)
.Wrap.Visible = .mnuWrap.Checked
.Page.Visible = Not .mnuWrap.Checked
Wrapped = .mnuWrap.Checked

    If .mnuFull.Checked Then
    .WindowState = vbMaximized
    Else
    Pos = WindowPosition
    .Move Pos * 350 + 40, Pos * 350, .ScaleWidth * Resize, .Height * Resize
    End If

    If .mnuWhite.Checked Then
    .mnuSpace.Checked = GetSetting("Editor", "Settings", "WhiteSpace", False)
    End If

White = .mnuSpace.Checked

Cmd = Command
Restart:
Select Case Cmd
    Case ""
    FileTitle = "Untitled"
        If White Then
        .Caption = "Untitled - White Space"
        Else
        .Caption = "Untitled - Notepad"
        End If
    Case Else
    Cmd = Replace(Cmd, """", "")
        If Dir(Cmd, 39) = "" Then
          If LCase(Right(Cmd, 4)) <> ".txt" Then
          Cmd = Cmd & ".txt"
          End If
          Path = GetPath(Cmd, SpecialFolder(DESKTOP))
            If Len(Path) > 0 Then
            Cmd = Path
            Else
            Path = GetPath(Cmd, SpecialFolder(DOCUMENTS))
              If Len(Path) > 0 Then
              Cmd = Path
              Else
              Path = GetPath(Cmd, ExplorerDirectory)
                  If Len(Path) > 0 Then
                  Cmd = Path
                  Else
                  Action = MsgBox("Cannot find the " & Cmd & " file." & vbCrLf & vbCrLf _
                  & "Do you want to create a new file?", vbYesNo + vbExclamation, "NOTEPAD")
                      If Action = vbYes Then
                      Path = SpecialFolder(DESKTOP) & "\" & Cmd
                        Open Path For Output As #1
                        Close
                      Cmd = Path
                      GoTo Restart:
                      Else
                      Cmd = ""
                      GoTo Restart:
                      End If
                  End If
              End If
           End If
        Else
        Cmd = LongName(Cmd)
        End If
        If FileIsOpen(Cmd) Then
        MsgBox Cmd & vbCrLf _
            & "This file is already in use." & vbCrLf _
            & "Select a new name or close the file in use by another application.", _
             vbOKOnly + vbExclamation, "Open"
        Cmd = ""
        GoTo Restart:
        End If
            If White Then
            LoadWhite Cmd
                If Len(FileName) > 0 Then
                .Caption = TitleFromPath(Cmd) & " - White Space"
                Else
                .Caption = "Untitled - White Space"
                End If
            Else
            LoadData Cmd
                If Len(FileName) > 0 Then
                .Caption = TitleFromPath(Cmd) & " - Notepad"
                Else
                .Caption = "Untitled - Notepad"
                End If
            End If
    End Select

    If .mnuDir.Checked Then
    .FileDialog.InitDir = GetSetting("Editor", "Settings", "SaveDir", SpecialFolder(DOCUMENTS))
    Else
    .FileDialog.InitDir = SpecialFolder(DOCUMENTS)
    End If
End With

SetFont
Editor.Show
DataModified = False
End Sub
Public Sub LoadData(PathName As String)

Dim i As Long
Dim Temp As String * 1
Dim AllSpace As String
Dim Display As String
Dim VSpace As String
Dim WSpace As String
Dim Spot As Long
Dim Log As Boolean

AllSpace = StringBuffer(FileLen(PathName))
AllSpace = Contents(PathName)
Spot = InStr(1, AllSpace, Marker)
    If Spot > 0 Then
    WSpace = Mid(AllSpace, Spot)
    VSpace = Mid(AllSpace, 1, Len(AllSpace) - Len(WSpace))
    Else
    VSpace = AllSpace
    End If

Display = StringBuffer(Len(VSpace))
    For i = 1 To Len(VSpace)
    Temp = Mid(VSpace, i, 1)
    
        If Asc(Temp) = 0 Then
        Temp = " "
        End If
    
    Mid(Display, i, 1) = Temp
    Next
    
    If Left(Display, 4) = ".LOG" Then
    Display = Display & vbCrLf & Format(Now, "h:mm AMPM m/d/yy")
    Log = True
    End If
    
    If Wrapped Then
    Editor.Wrap = Display
    Else
    Editor.Page = Display
    End If

FileName = PathName
Editor.Caption = Editor.FileDialog.FileTitle & " - Notepad"

If Log Then
DataModified = True
Else
DataModified = False
End If

PageNum = 1
HAlign = 2
FAlign = 2
HeaderText = HeaderText & TitleFromPath(FileName)
FooterText = FooterText & "Page " & CStr(PageNum)
LeftMargin = 720
RightMargin = 720
TopMargin = 1440
BottomMargin = 1440
End Sub






Public Function TitleFromPath(FullPath As String) As String

TitleFromPath = Right(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))

End Function
Public Sub LoadWhite(PathName As String)

Dim i As Long
Dim Temp As String * 1
Dim AllSpace As String
Dim Display As String
Dim VSpace As String
Dim WSpace As String
Dim TheSpot As Long

AllSpace = StringBuffer(FileLen(PathName))
AllSpace = Contents(PathName)
TheSpot = InStr(1, AllSpace, Marker)
    
If TheSpot > 0 Then
WSpace = Mid(AllSpace, TheSpot)
VSpace = Mid(AllSpace, 1, TheSpot - 1)
    If Wrapped Then
    Editor.Wrap = BlankToString(WSpace)
    Else
    Editor.Page = BlankToString(WSpace)
    End If
Else
Editor.Page = ""
Editor.Wrap = ""
End If

FileName = PathName
Editor.Caption = Editor.FileDialog.FileTitle & " - White Space"
DataModified = False

PageNum = 1
HAlign = 2
FAlign = 2
HeaderText = HeaderText & TitleFromPath(FileName)
FooterText = FooterText & "Page " & PageNum
LeftMargin = 720
RightMargin = 720
TopMargin = 1440
BottomMargin = 1440
End Sub


Private Sub SetFont()

With Editor

.Page.Font = GetSetting("Editor", "Font", "Name", "Fixedsys")
.Page.Font.Bold = GetSetting("Editor", "Font", "Bold", False)
.Page.Font.Italic = GetSetting("Editor", "Font", "Italic", False)
.Page.Font.Size = GetSetting("Editor", "Font", "Size", 9)
.Page.Font.Underline = GetSetting("Editor", "Font", "Under", False)
.Wrap.Font = GetSetting("Editor", "Font", "Name", "Fixedsys")
.Wrap.Font.Bold = GetSetting("Editor", "Font", "Bold", False)
.Wrap.SelItalic = GetSetting("Editor", "Font", "Italic", False)
.Wrap.Font.Size = GetSetting("Editor", "Font", "Size", 9)
.Wrap.Font.Underline = GetSetting("Editor", "Font", "Under", False)
End With

End Sub

