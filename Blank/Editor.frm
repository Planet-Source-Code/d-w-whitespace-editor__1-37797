VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Editor 
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   7245
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4425
   ScaleWidth      =   7245
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5805
      Top             =   645
   End
   Begin RichTextLib.RichTextBox Hidden 
      Height          =   690
      Left            =   5355
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1217
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Editor.frx":1CFA
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5790
      Top             =   1125
   End
   Begin RichTextLib.RichTextBox Wrap 
      Height          =   2160
      Left            =   2745
      TabIndex        =   1
      Top             =   225
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   3810
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Editor.frx":1DC3
   End
   Begin RichTextLib.RichTextBox Page 
      Height          =   2145
      Left            =   345
      TabIndex        =   0
      Top             =   225
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3784
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   1.24000e5
      TextRTF         =   $"Editor.frx":1E8C
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPage 
         Caption         =   "Page Se&tup..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDel 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuDate 
         Caption         =   "Time/&Date"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWrap 
         Caption         =   "&Word Wrap"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Set &Font..."
      End
      Begin VB.Menu mnuFull 
         Caption         =   "&Fullscreen"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "W&hite Space"
      End
      Begin VB.Menu mnuWhite 
         Caption         =   "&Save White Setting"
      End
      Begin VB.Menu mnuDir 
         Caption         =   "Save &Directory"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTopic 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Notepad"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuUndoR 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuSepR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCutR 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopyR 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPasteR 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelR 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSepR2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAllR 
         Caption         =   "Select &All"
      End
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_USER = &H400
Private Const EM_CANPASTE = WM_USER + 50
Private Const EM_UNDO = &HC7
Private Const EM_CANUNDO = &HC6
Private Const WM_COPY& = &H301
Private Const WM_CUT& = &H300
Private Const WM_PASTE& = &H302
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETLINE = &HC4
Private Const EM_SCROLLCARET As Long = &HB7
Private Const WM_SETREDRAW  As Long = &HB
Private Const EM_GETSEL  As Long = &HB0
Private Const EM_SETSEL  As Long = &HB1

Dim CurPos As Long
Dim NewPos As Long
Dim FilterIndex As Integer
Private Sub SaveWhite(PathName As String, Optional OldPath As String)

Dim AllSpace As String
Dim VSpace As String
Dim WSpace As String
Dim Spot As Long
Dim BytesWritten As Long

If FileIsOpen(PathName) Then
MsgBox PathName & vbCrLf _
& "This file is already in use." & vbCrLf _
& "Select a new name or close the file in use by another application.", _
 vbOKOnly + vbExclamation, "Open"
mnuFileSaveAs_Click
Exit Sub
End If

If Len(OldPath) = 0 Then
AllSpace = Contents(PathName)
    If Len(AllSpace) > 0 Then FileDelete PathName
Else
AllSpace = Contents(OldPath)
End If

Spot = InStr(1, AllSpace, Marker)
If Spot > 0 Then
VSpace = StringBuffer(Spot - 1)
VSpace = Mid(AllSpace, 1, Spot - 1)
Else
VSpace = StringBuffer(Len(AllSpace))
VSpace = AllSpace
End If

WSpace = StringBuffer(Len(StringToBlank(IIf(Wrapped, Wrap.Text, Page.Text))))
WSpace = StringToBlank(IIf(Wrapped, Wrap.Text, Page.Text))

AllSpace = StringBuffer(Len(VSpace) + Len(WSpace))
Mid(AllSpace, 1, Len(VSpace)) = VSpace
Mid(AllSpace, Len(VSpace) + 1) = WSpace

BytesWritten = FileWrite(PathName, AllSpace)
FileName = PathName
DataModified = False


End Sub

Private Sub PrintPrint()

If Wrapped Then
Printer.Font.Bold = Editor.Wrap.Font.Bold
Printer.Font.Name = Editor.Wrap.Font.Name
Printer.Font.Size = Editor.Wrap.Font.Size * Resize
Printer.Font.Italic = Editor.Wrap.Font.Italic
Printer.Font.Strikethrough = Editor.Wrap.Font.Strikethrough
Printer.Font.Underline = Editor.Wrap.Font.Underline
PrintText Editor.Wrap.Text
Else
Printer.Font.Bold = Editor.Page.Font.Bold
Printer.Font.Name = Editor.Page.Font.Name
Printer.Font.Size = Editor.Page.Font.Size * Resize
Printer.Font.Italic = Editor.Page.Font.Italic
Printer.Font.Strikethrough = Editor.Page.Font.Strikethrough
Printer.Font.Underline = Editor.Page.Font.Underline
PrintText Editor.Page.Text

End If




End Sub


Private Function DataSafe() As Boolean

Dim Msg As String

If Len(Page.Text) = 0 And Len(Wrap.Text) = 0 And Len(FileName) = 0 Then
DataSafe = True
Exit Function
End If

If Not DataModified Then
DataSafe = True
Exit Function
End If

If Len(FileName) = 0 Then
Msg = "Untitled"
Else
Msg = FileName
End If

Select Case MsgBox("The text in the " & Msg & _
" file has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", _
    vbYesNoCancel + vbExclamation, IIf(White, "White Space", "Notepad"))
Case vbYes
DataSafe = True
mnuFileSave_Click
Case vbNo
DataSafe = True
Case vbCancel
DataSafe = False
End Select

End Function


Private Function FolderFromPath(PathName As String) As String

Dim Spot As Integer

Spot = InStrRev(PathName, "\", 1, vbTextCompare)
FolderFromPath = Left(PathName, (Len(PathName) - Spot - 2))

End Function

Private Sub SaveVisible(PathName As String, Optional OldPath As String)

Dim Temp As String
Dim AllSpace As String
Dim VSpace As String
Dim WSpace As String
Dim Length As Long
Dim Spot As Long
Dim BytesWritten As Long

If FileIsOpen(PathName) Then
MsgBox FileDialog.FileName & vbCrLf _
& "This file is already in use." & vbCrLf _
& "Select a new name or close the file in use by another application.", _
 vbOKOnly + vbExclamation, "Open"
mnuFileSaveAs_Click
Exit Sub
End If

If Len(OldPath) = 0 Then
AllSpace = Contents(PathName)
    If Len(AllSpace) > 0 Then FileDelete PathName
Else
AllSpace = Contents(OldPath)
End If

Spot = InStr(1, AllSpace, Marker)
If Spot > 0 Then
WSpace = StringBuffer(Len(AllSpace) - Spot + 1)
WSpace = Mid(AllSpace, Spot)
Else
WSpace = ""
End If

VSpace = StringBuffer(Len(IIf(Wrapped, Wrap.Text, Page.Text)))
VSpace = IIf(Wrapped, Wrap.Text, Page.Text)

AllSpace = StringBuffer(Len(VSpace) + Len(WSpace))

If Len(AllSpace) > 0 Then
Mid(AllSpace, 1, Len(VSpace)) = VSpace
Else
AllSpace = VSpace
End If

If Len(WSpace) > 0 Then
Mid(AllSpace, Len(VSpace) + 1) = WSpace
End If

BytesWritten = FileWrite(PathName, AllSpace)
FileName = PathName
DataModified = False
End Sub




Private Sub SaveFont()

SaveSetting "Editor", "Font", "Name", Page.Font
SaveSetting "Editor", "Font", "Bold", Page.Font.Bold
SaveSetting "Editor", "Font", "Italic", Page.Font.Italic
SaveSetting "Editor", "Font", "Size", Page.Font.Size
SaveSetting "Editor", "Font", "Under", Page.Font.Underline

End Sub

Private Sub SetDialogPath()
Dim FilePath As String
FilePath = FileDialog.FileName
Do While Right(FilePath, 1) <> "\"
FilePath = Left(FilePath, Len(FilePath) - 1)
Loop
FileDialog.InitDir = FilePath
SaveSetting "Editor", "Settings", "SaveDir", FileDialog.InitDir
End Sub

  




Private Sub Form_Activate()
Page.RightMargin = 124000 * Resize
End Sub

Private Sub Form_Load()
Dim Swap As String
If Wrapped Then
Wrap_Click
Else
Swap = Page.Text
Page.Text = ""
Page.Text = Swap
End If
End Sub


Private Sub Form_Paint()
Static Loaded As Boolean
If Not Loaded Then
Form_Load
Loaded = True
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

SetWindowLong Page.hwnd, GWL_WNDPROC, PrevProc1
SetWindowLong Wrap.hwnd, GWL_WNDPROC, PrevProc2

End Sub

Private Sub mnuAbout_Click()
About.Show vbModal
End Sub

Private Sub mnuAll_Click()
If Page.Visible Then
Page.SelStart = 0
Page.SelLength = Len(Page.Text)
Page.SetFocus
Else
Wrap.SelStart = 0
Wrap.SelLength = Len(Wrap.Text)
Wrap.SetFocus
End If
End Sub

Private Sub mnuAllR_Click()
mnuAll_Click
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
If Page.Visible Then
SendMessage Page.hwnd, WM_COPY, vbNull, vbNull
Else
SendMessage Wrap.hwnd, WM_COPY, vbNull, vbNull
End If
End Sub

Private Sub mnuCopyR_Click()
mnuCopy_Click
End Sub


Private Sub mnuCut_Click()
Clipboard.Clear
If Not Wrapped Then
SendMessage Page.hwnd, WM_CUT, vbNull, vbNull
Page.SelText = vbNullString
Else
SendMessage Wrap.hwnd, WM_CUT, vbNull, vbNull
Wrap.SelText = vbNullString
End If
End Sub

Private Sub mnuCutR_Click()
mnuCut_Click
End Sub

Private Sub mnuDate_Click()
Dim Stamp As String
Stamp = Format(Now, "h:mm AMPM m/d/yy")
If Page.Visible Then
Page.SelText = Stamp
Else
Wrap.SelText = Stamp
End If
End Sub

Private Sub mnuDel_Click()
If Not Wrapped Then
Page.SelText = vbNullString
Else
Wrap.SelText = vbNullString
End If
End Sub


Private Sub mnuDelR_Click()
mnuDel_Click
End Sub


Private Sub mnuDir_Click()
mnuDir.Checked = Not mnuDir.Checked
SaveSetting "Editor", "Settings", "Previous", mnuDir.Checked
End Sub

Private Sub mnuEdit_Click()
If Not Wrapped Then
mnuUndo.Enabled = SendMessage(Page.hwnd, EM_CANUNDO, vbNull, vbNull)
mnuDel.Enabled = Page.SelLength > 0
mnuAll.Enabled = Len(Page) > 0
Else
mnuUndo.Enabled = SendMessage(Wrap.hwnd, EM_CANUNDO, vbNull, vbNull)
mnuDel.Enabled = Wrap.SelLength > 0
mnuAll.Enabled = Len(Wrap) > 0
End If
mnuCut.Enabled = mnuDel.Enabled
mnuCopy.Enabled = mnuDel.Enabled
mnuPaste.Enabled = Len(Clipboard.GetText(vbCFText)) > 0
End Sub

Private Sub mnuFile_Click()
Dim p As String
On Error GoTo Out:
p = Printer.DeviceName
mnuPrint.Enabled = True
Exit Sub
Out:
mnuPrint.Enabled = False
End Sub

Private Sub mnuFind_Click()
If Page.Visible Then
ShowFind Me, Page
Else
ShowFind Me, Wrap
End If
End Sub

Private Sub mnuFont_Click()

FileDialog.FontName = Page.Font
FileDialog.FontBold = Page.Font.Bold
FileDialog.FontItalic = Page.Font.Italic
FileDialog.FontSize = Page.Font.Size
FileDialog.FontUnderline = Page.Font.Underline
FileDialog.flags = cdlCFScreenFonts
FileDialog.ShowFont
Page.Font = FileDialog.FontName
Page.Font.Bold = FileDialog.FontBold
Page.Font.Italic = FileDialog.FontItalic
Page.Font.Size = FileDialog.FontSize
Page.Font.Underline = FileDialog.FontUnderline
Wrap.Font = FileDialog.FontName
Wrap.Font.Bold = FileDialog.FontBold
Wrap.Font.Italic = FileDialog.FontItalic
Wrap.Font.Size = FileDialog.FontSize
Wrap.Font.Underline = FileDialog.FontUnderline
SaveFont
End Sub

Private Sub mnuFull_Click()
mnuFull.Checked = Not mnuFull.Checked
If mnuFull.Checked Then
WindowState = vbMaximized
Else
WindowState = vbNormal
End If
SaveSetting "Editor", "Settings", "Fullscreen", mnuFull.Checked
End Sub

Private Sub mnuNext_Click()
If FindStarted Then
    If Up Then
    FindPrevWord
    Else
    FindNextWord
    End If
Else
    If Page.Visible Then
    ShowFind Me, Page
    Else
    ShowFind Me, Wrap
    End If
End If
End Sub

Private Sub mnuPage_Click()
Dim p As String
On Error GoTo Out:
p = Printer.DeviceName
PrintSet.Show vbModal
Exit Sub
Out:
MsgBox "Before you can print, you need to install a printer." _
& vbCrLf & "To do this, click Start, point to Settings, click Printers, " _
& "and then double-click Add Printer.", vbOKOnly + vbExclamation, Editor.Caption
End Sub

Private Sub mnuPaste_Click()
mnuPasteR_Click
End Sub

Private Sub mnuPasteR_Click()

Dim Length As Long
Dim SelLength As Long
Hidden.Text = Clipboard.GetText(vbCFText)
Clipboard.Clear
Clipboard.SetText Hidden.Text, vbCFText
Hidden.Text = ""
If Wrapped Then
Length = Len(Wrap.Text)
    If Wrap.SelLength > 0 Then
    SelLength = Wrap.SelLength
    Wrap.SelText = Clipboard.GetText(vbCFText)
        If Length - SelLength + Len(Clipboard.GetText(vbCFText)) <> Len(Wrap.Text) Then
        GoTo ShowErrMessage:
        Exit Sub
        End If
    Else
    SendMessage Wrap.hwnd, WM_PASTE, vbNull, vbNull
        If Length + Len(Clipboard.GetText(vbCFText)) <> Len(Wrap.Text) Then
        GoTo ShowErrMessage:
        Exit Sub
        End If
    End If
Else
Length = Len(Page.Text)
    If Page.SelLength > 0 Then
    SelLength = Page.SelLength
    Page.SelText = Clipboard.GetText(vbCFText)
        If Length - SelLength + Len(Clipboard.GetText(vbCFText)) <> Len(Page.Text) Then
        GoTo ShowErrMessage:
        Exit Sub
        End If
    Else
    SendMessage Page.hwnd, WM_PASTE, vbNull, vbNull
        If Length + Len(Clipboard.GetText(vbCFText)) <> Len(Page.Text) Then
        GoTo ShowErrMessage:
        Exit Sub
        End If
    End If
End If

Exit Sub
ShowErrMessage:
MsgBox "Not enough memory available to complete this operation. " _
& "Quit one or more applications to increase available " _
& "memory, and then try again.", vbOKOnly + vbExclamation, "NotePad"
End Sub

Private Sub mnuPrint_Click()
PrintPrint
End Sub

Private Sub mnuSpace_Click()

If Not DataSafe Then Exit Sub

mnuSpace.Checked = Not mnuSpace.Checked

SaveSetting "Editor", "Settings", "WhiteSpace", mnuSpace.Checked
White = mnuSpace.Checked
If Len(FileName) > 0 Then
    If White Then
    LoadWhite FileName
    Caption = TitleFromPath(FileName) & " - White Space"
    Else
    LoadData FileName
    Caption = TitleFromPath(FileName) & " - Notepad"
    End If
Else
    Page = ""
    If White Then
    Caption = "Untitled - White Space"
    Else
    Caption = "Untitled - Notepad"
    End If
End If
End Sub

Private Sub mnuTopic_Click()
'works on Win98 and up
Dim Path As String
Path = SpecialFolder
Shell Path & "\Command\Start.exe " & Path _
 & "\Help\Notepad.chm", vbMinimizedFocus
End Sub

Private Sub mnuUndo_Click()
If Not Wrapped Then
SendMessage Me.Page.hwnd, EM_UNDO, vbNull, vbNull
Page.SetFocus
Else
SendMessage Me.Wrap.hwnd, EM_UNDO, vbNull, vbNull
Wrap.SetFocus
End If
End Sub

Private Sub mnuUndoR_Click()
mnuUndo_Click
End Sub

Private Sub mnuWhite_Click()
mnuWhite.Checked = Not mnuWhite.Checked
SaveSetting "Editor", "Settings", "OpenWhite", mnuWhite.Checked
End Sub

Private Sub mnuWrap_Click()

Dim Swap As String

mnuWrap.Checked = Not mnuWrap.Checked
SaveSetting "Editor", "Settings", "Wrap", mnuWrap.Checked

If mnuWrap.Checked Then
Swap = Page
Page = ""
Page.Visible = False
Wrap.Visible = True
Wrap = Swap
Wrapped = True
Else
Swap = Wrap
Wrap = ""
Wrap.Visible = False
Page.Visible = True
Page = Swap
Wrapped = False
End If

End Sub


Private Sub Page_Change()
CreateCaret Page.hwnd, 0, 2 * Resize, 16 * Resize
ShowCaret Page.hwnd
DataModified = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = Not DataSafe
End Sub
Private Sub Form_Resize()

Page.Move 0, 0, ScaleWidth, ScaleHeight
Wrap.Move 0, 0, ScaleWidth, ScaleHeight

End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub



Private Sub mnuFileNew_Click()

If Not DataSafe Then Exit Sub
Page = ""
FileTitle = "Untitled"
FileName = ""

If White Then
Caption = "Untitled - White Space"
Else
Caption = "Untitled - Notepad"
End If

DataModified = False
End Sub

 
Private Sub mnuFileOpen_Click()
If Not DataSafe Then Exit Sub
If Len(FileDialog.InitDir) = 0 Then
FileDialog.InitDir = SpecialFolder(DOCUMENTS)
End If
FileDialog.flags = cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    If White Then
    FileDialog.Filter = "All Files (*.*)|*.*"
    FileDialog.FileName = ""
    Else
    FileDialog.Filter = "Text Documents |*.txt|All Files (*.*)|*.*"
        If FilterIndex > 0 Then
        FileDialog.FilterIndex = FilterIndex
        End If
        If FilterIndex < 2 Then
        FileDialog.FileName = "*.txt"
        Else
        FileDialog.FileName = ""
        End If
    End If
TryAgain:
On Error Resume Next
FileDialog.ShowOpen

If Err.Number <> 0 Then
On Error GoTo 0
Exit Sub
End If
    If FileIsOpen(FileDialog.FileName) Then
    MsgBox FileDialog.FileName & vbCrLf _
    & "This file is already in use." & vbCrLf _
    & "Select a new name or close the file in use by another application.", _
     vbOKOnly + vbExclamation, "Open"
    
    On Error Resume Next
    FileDialog.ShowOpen
        If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Sub
        End If
    End If

SetDialogPath
FileTitle = FileDialog.FileTitle
FilterIndex = FileDialog.FilterIndex
    If White Then
    LoadWhite FileDialog.FileName
    Else
    LoadData FileDialog.FileName
    End If

End Sub

 
Private Sub mnuFileSave_Click()

If Len(FileName) = 0 Then
mnuFileSaveAs_Click
Exit Sub
End If

If Dir(FileName, vbNormal + vbArchive + vbHidden + vbReadOnly) = "" Then
mnuFileSaveAs_Click
Exit Sub
End If

If (GetAttr(FileName) And vbReadOnly) > 0 Then
mnuFileSaveAs_Click
Exit Sub
End If

If FileIsOpen(FileName) Then
mnuFileSaveAs_Click
Exit Sub
End If

If White Then
SaveWhite FileName
Else
SaveVisible FileName
End If

End Sub

 
Private Sub mnuFileSaveAs_Click()

Dim PathName As String
Dim OldPath As String

OldPath = FileName

If Len(FileDialog.InitDir) = 0 Then
FileDialog.InitDir = SpecialFolder(DOCUMENTS)
End If

FileDialog.FileName = TitleFromPath(FileName)
FileDialog.flags = cdlOFNPathMustExist + _
 cdlOFNLongNames + cdlOFNNoReadOnlyReturn + _
 cdlOFNOverwritePrompt + cdlOFNHideReadOnly
FileDialog.Filter = "Text Documents |*.txt|All Files (*.*)|*.*"
FileDialog.FilterIndex = FilterIndex

If FileDialog.FilterIndex < 2 Then
FileDialog.DefaultExt = ".txt"
Else
FileDialog.DefaultExt = vbNull
End If

On Error Resume Next
FileDialog.ShowSave

If Err.Number = cdlCancel Then
On Error GoTo 0
Exit Sub
ElseIf Err.Number <> 0 Then
On Error GoTo 0
Exit Sub
End If



SetDialogPath
FilterIndex = FileDialog.FilterIndex
FileTitle = FileDialog.FileTitle
PathName = FileDialog.FileName

If Dir(PathName, vbNormal + vbArchive + vbHidden + vbReadOnly) <> "" Then
    If (GetAttr(PathName) And vbReadOnly) > 0 Then
    MsgBox FileName & vbCrLf & _
    "This file exists with Read Only attributes" & _
    "Please use a different filename.", vbOKOnly, "Save As"
    mnuFileSaveAs_Click
    Exit Sub
    End If
End If

If Len(OldPath) = 0 Then
    If White Then
    SaveWhite PathName
    Caption = FileTitle & " - White Space"
    Else
    SaveVisible PathName
    Caption = FileTitle & " - Notepad"
    End If
Else
    If White Then
    SaveWhite PathName, OldPath
    Caption = FileTitle & " - White Space"
    Else
    SaveVisible PathName, OldPath
    Caption = FileTitle & " - Notepad"
    End If
End If
FileDialog.FileName = FileTitle
End Sub
 
Private Sub Page_Click()
CreateCaret Page.hwnd, 0, 2 * Resize, 16 * Resize
ShowCaret Page.hwnd
End Sub

Private Sub Page_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then
If Up Then
FindPrevWord
Else
FindNextWord
End If
End If
End Sub

Private Sub Page_KeyPress(KeyAscii As Integer)
CreateCaret Page.hwnd, 0, 2 * Resize, 16 * Resize
ShowCaret Page.hwnd
End Sub


Private Sub Timer1_Timer()
Timer1.Enabled = False
If Wrapped Then
Wrap_Click
Else
Page_Click
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Page_Change
DataModified = False
End Sub

Private Sub Wrap_Change()
CreateCaret Wrap.hwnd, 0, 2 * Resize, 16 * Resize
ShowCaret Wrap.hwnd
DataModified = True
End Sub


Private Sub Wrap_Click()
CreateCaret Wrap.hwnd, 0, 2 * Resize, 16 * Resize
ShowCaret Wrap.hwnd
End Sub


Private Sub Wrap_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then
If Up Then
FindPrevWord
Else
FindNextWord
End If
End If
End Sub


Private Sub Wrap_KeyPress(KeyAscii As Integer)
CreateCaret Wrap.hwnd, 0, 2 * Resize, 16 * Resize
ShowCaret Wrap.hwnd
End Sub


