Attribute VB_Name = "modMain"
Option Explicit

Public gstrImageName        As String
Public gstrFileName         As String
Public gblnCancelProcces    As Boolean
Public gblnIsRunning        As Boolean
Public gblnTextHasChanged   As Boolean
Public gstrActiveKey        As String
'printer settings
Public glngLeftMarginPrint  As Long
Public glngRightMarginPrint As Long
Public glngTopMarginPrint   As Long
Public glngBottMarginPrint  As Long
Public glngPrintSize        As Long
Public gintPrintSelection   As Integer
Public gblnChangePageSetup  As Boolean
Public gblnCancelPrint      As Boolean
Public gblnPrinterPresent   As Boolean
Public gblnCancelKey        As Boolean
Public gblnReadOnly         As Boolean
Public retVal               As String
Public vbCrLfLf             As String

'Run or open file
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Browsing folders
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type
'common dialog flags
Public Const CMDLG_NOCHECK = &H4
Public Const CMDLG_NOOVERWRITE = &H2
Public Const CMDLG_PATHMUSTEXIST = &H800
Public Const CMDLG_FILEMUSTEXIST = &H1000
'icon extraction
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    id((123)) As Byte
End Type
Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long                      '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1

Sub Main()
Dim dummy
Load frmMain
Load frmAbout
Load frmRead
Load frmWrite
Load frmPrintProgress
Load frmPrintSetup
Call GetWindowPos
vbCrLfLf = vbCrLf & vbCrLf
On Error Resume Next

' set helpfile
App.HelpFile = App.path + "\" & "STEGANOTE.hlp"
frmMain.ComDlg.HelpFile = App.path + "\" & "STEGANOTE.hlp"

'check if printer is present
dummy = Printer.DeviceName
If dummy <> "" And Err = 0 Then
    gblnPrinterPresent = True
    Else
    gblnPrinterPresent = False
    End If

'page setup
glngTopMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintTop", "5"))
glngBottMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintBottom", "5"))
glngLeftMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintLeft", "5"))
glngRightMarginPrint = Val(GetSetting(App.EXEName, "Config", "PrintRight", "5"))
glngPrintSize = Val(GetSetting(App.EXEName, "Config", "PrintSize", "11"))

Call CheckMenus
gstrImageName = ""
gstrFileName = ""

frmMain.Show

Call CheckMenus
End Sub

Sub ChangeWindowSize()
Dim Th
With frmMain
Th = .Toolbar1.Height
If .WindowState <> 1 And .ScaleHeight > 4000 Then
    If .ScaleHeight - Th - 300 > 0 Then  '''
        .picPreview.Top = .Height - .picPreview.Height - Th - 400
        .picInfo.Top = .picPreview.Top
        .picStatus.Top = .picPreview.Top - .picStatus.Height - 125
        .picFileContainer.Top = .picStatus.Top - .picFileContainer.Height '- 10
        .txtMain.Height = .ScaleHeight - .picInfo.Height - .picFileContainer.Height - .picStatus.Height - Th - 225
        .txtMain.Top = Th + 50
    End If
End If
If .WindowState <> 1 And .ScaleWidth > 3000 Then
    .txtMain.Width = .ScaleWidth
    .picFileContainer.Width = .ScaleWidth
    .picInfo.Width = .ScaleWidth - .picPreview.Width - 50
    .lblRND.Left = .picStatus.Width - .lblRND.Width - 200
    .picStatus.Width = .ScaleWidth
    .lblImageFile.Width = .picInfo.Width - 200
    .lblFile.Width = .picFileContainer.Width - 1000
If gstrImageName <> "" Then .lblImageFile.Caption = TrimPath(gstrImageName, .lblImageFile.Width)
If gstrFileName <> "" Then .lblFile.Caption = TrimPath(gstrFileName, .lblFile.Width)
End If
End With
End Sub

Public Function Browse(ByVal aTitle As String) As String
Dim bInfo As BROWSEINFO
Dim rtn&, pidl&, path$, Pos%
Dim BrowsePath As String
Dim t
bInfo.hOwner = frmMain.hwnd
bInfo.lpszTitle = aTitle
'the type of folder(s) to return
bInfo.ulFlags = &H1
'show the dialog box
pidl& = SHBrowseForFolder(bInfo)
'set the maximum characters
path = Space(512)
t = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'gets the selected path
Pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
'set the extracted path to SpecIn
Browse = Left(path$, Pos - 1)
'make sure that "\" is at the end of the path
If Right$(Browse, 1) = "\" Then
    Browse = Browse
    Else
    Browse = Browse + "\"
End If
If Browse = "\" Then Browse = ""
End Function

Sub SaveWindowPos()
'save last pos and state window
With frmMain
SaveSetting App.EXEName, "Config", "WindowState", .WindowState
If .WindowState <> vbNormal Then Exit Sub
SaveSetting App.EXEName, "Config", "Height", .Height
SaveSetting App.EXEName, "Config", "Width", .Width
SaveSetting App.EXEName, "Config", "Left", .Left
SaveSetting App.EXEName, "Config", "Top", .Top
End With
End Sub

Sub GetWindowPos()
Dim x
'get last pos and state window
With frmMain
x = CSng(GetSetting(App.EXEName, "Config", "Windowstate", vbNormal))
If x = vbMinimized Then x = vbNormal
.WindowState = x
If .WindowState <> vbNormal Then Exit Sub
.Height = CSng(GetSetting(App.EXEName, "Config", "Height", .Height))
.Width = CSng(GetSetting(App.EXEName, "Config", "Width", .Width))
.Left = CSng(GetSetting(App.EXEName, "Config", "Left", ((Screen.Width - .Width) / 2)))
.Top = CSng(GetSetting(App.EXEName, "Config", "Top", ((Screen.Height - .Height) / 2)))
End With
Call ChangeWindowSize
End Sub

Public Sub CheckMenus()
With frmMain
If gstrImageName = "" Then
    'no file
    .mnuRead.Enabled = False
    .mnuWrite.Enabled = False
    .mnuErase.Enabled = False
    If .Toolbar1.Buttons("read").Enabled <> False Then .Toolbar1.Buttons("read").Enabled = False
    If .Toolbar1.Buttons("write").Enabled <> False Then .Toolbar1.Buttons("write").Enabled = False
    Else
    If gblnReadOnly = True Then
        'read only
        .mnuRead.Enabled = True
        .mnuWrite.Enabled = False
        .mnuErase.Enabled = False
        If .Toolbar1.Buttons("read").Enabled <> True Then .Toolbar1.Buttons("read").Enabled = True
        If .Toolbar1.Buttons("write").Enabled <> False Then .Toolbar1.Buttons("write").Enabled = False
        Else
        If frmMain.txtMain.Text = "" And gstrFileName = "" Then
            .mnuWrite.Enabled = False
            If .Toolbar1.Buttons("write").Enabled <> False Then .Toolbar1.Buttons("write").Enabled = False
            Else
            .mnuWrite.Enabled = True
            If .Toolbar1.Buttons("write").Enabled <> True Then .Toolbar1.Buttons("write").Enabled = True
        End If
        If Right(gstrImageName, 4) <> ".bmp" Then
            .mnuRead.Enabled = False
            .mnuErase.Enabled = False
            If .Toolbar1.Buttons("read").Enabled <> False Then .Toolbar1.Buttons("read").Enabled = False
            Else
            .mnuRead.Enabled = True
            .mnuErase.Enabled = True
            If .Toolbar1.Buttons("read").Enabled <> True Then .Toolbar1.Buttons("read").Enabled = True
        End If
    End If
End If
End With


'clipboard
On Error Resume Next
If Clipboard.GetText <> "" Then
    If frmMain.Toolbar1.Buttons("paste").Enabled <> True Then frmMain.Toolbar1.Buttons("paste").Enabled = True
    frmMain.mnuPaste.Enabled = True
    Else
    frmMain.mnuPaste.Enabled = False
    If frmMain.Toolbar1.Buttons("paste").Enabled <> False Then frmMain.Toolbar1.Buttons("paste").Enabled = False
    End If

With frmMain
'printer
If gblnPrinterPresent = True Then
    .mnuPageSetup.Enabled = True
    If Len(.txtMain.Text) > 0 Then
        .mnuPrint.Enabled = True
        .Toolbar1.Buttons("print").Enabled = True
        Else
        .mnuPrint.Enabled = False
        .Toolbar1.Buttons("print").Enabled = False
        End If
    Else
    .mnuPageSetup.Enabled = False
    .mnuPrint.Enabled = False
    .Toolbar1.Buttons("print").Enabled = False
    End If
'text edit
If frmMain.txtMain.Text = "" Then
    .mnuCopy.Enabled = False
    .mnuCopyAll.Enabled = False
    .mnuSelectAll.Enabled = False
    .mnuCut.Enabled = False
    .mnuDelete.Enabled = False
    If .Toolbar1.Buttons("copy").Enabled <> False Then .Toolbar1.Buttons("copy").Enabled = False
    If .Toolbar1.Buttons("cut").Enabled <> False Then .Toolbar1.Buttons("cut").Enabled = False
    Else
    .mnuSelectAll.Enabled = True
    .mnuCopyAll.Enabled = True
    If frmMain.txtMain.SelLength <> 0 Then
        .mnuCopy.Enabled = True
        .mnuCut.Enabled = True
        .mnuDelete.Enabled = True
        If .Toolbar1.Buttons("copy").Enabled <> True Then .Toolbar1.Buttons("copy").Enabled = True
        If .Toolbar1.Buttons("cut").Enabled <> True Then .Toolbar1.Buttons("cut").Enabled = True
        Else
        .mnuCopy.Enabled = False
        .mnuCut.Enabled = False
        .mnuDelete.Enabled = False
        If .Toolbar1.Buttons("copy").Enabled <> False Then .Toolbar1.Buttons("copy").Enabled = False
        If .Toolbar1.Buttons("cut").Enabled <> False Then .Toolbar1.Buttons("cut").Enabled = False
        End If
    End If
'undo
If gblnTextHasChanged = True Then
    If .Toolbar1.Buttons("undo").Enabled <> True Then .Toolbar1.Buttons("undo").Enabled = True
    Else
    If .Toolbar1.Buttons("undo").Enabled <> False Then .Toolbar1.Buttons("undo").Enabled = False
    End If

End With
End Sub

Public Function KeyQuality(ByVal aKey As String) As Integer
' returns an integer value (0 to 100) rating the key quality
Dim QC As Integer
Dim LN As Integer
Dim k As Integer
Dim Uc As Boolean
Dim Lc As Boolean
Dim Wid As Integer
Dim ValidKey As Boolean
LN = Len(aKey)
QC = LN * 4
'check key lenght (at least 5 chars!)
If Len(aKey) < 5 Then KeyQuality = 0: Exit Function
' check for repetitions (abcabc, aaaaa, 121212, etc.)
For Wid = 1 To Int(Len(aKey) / 2)
    ValidKey = False
    For k = Wid + 1 To Len(aKey) Step Wid
        If Mid(aKey, 1, Wid) <> Mid(aKey, k, Wid) Then ValidKey = True: Exit For
    Next
If ValidKey = False Then Exit For
Next
If ValidKey = False Then KeyQuality = 0: Exit Function
'check ucases and lcases
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 64 And Asc(Mid(aKey, k, 1)) < 91 Then Uc = True
    If Asc(Mid(aKey, k, 1)) > 96 And Asc(Mid(aKey, k, 1)) < 123 Then Lc = True
Next
If Uc = True And Lc = True Then QC = QC * 1.5
'check numbers
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 47 And Asc(Mid(aKey, k, 1)) < 58 Then
        If Uc = True Or Lc = True Then QC = QC * 1.5
        Exit For
        End If
Next
'check signs
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) < 48 Or Asc(Mid(aKey, k, 1)) > 122 Or (Asc(Mid(aKey, k, 1)) > 57 And Asc(Mid(aKey, k, 1)) < 65) Then QC = QC * 1.5: Exit For
Next
If QC > 100 Then QC = 100
KeyQuality = Int(QC)
End Function


Public Function GetFileExt(strFile As String) As String
'returns extension of filename
Dim k   As Integer
Dim Pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "." Then Pos = k
Next
If Pos = Len(strFile) Then Pos = 0
If Pos = 0 Then
    GetFileExt = ""
    Else
    GetFileExt = LCase(Mid(strFile, Pos + 1))
    End If
End Function

Public Function GetFilePath(strFile As String) As String
'returns only the path without filename
Dim k As Integer
Dim Pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "\" Then Pos = k
Next
If Pos < 2 Then
    GetFilePath = ""
    Else
    GetFilePath = Left(strFile, Pos)
    End If
End Function

Public Function CutFileExt(strFile As String) As String
'returns full path and filename without extension
Dim k As Integer
Dim Pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "." Then Pos = k
Next
If Pos = 0 Then
    CutFileExt = strFile
    Else
    CutFileExt = Left(strFile, Pos - 1)
    End If
End Function

Public Function CutFilePath(strFile As String) As String
'returns only the filename without full path
Dim k As Integer
Dim Pos As Integer
For k = 1 To Len(strFile)
    If Mid(strFile, k, 1) = "\" Then Pos = k
Next
If Pos = 0 Then
    CutFilePath = strFile
    Else
    CutFilePath = Mid(strFile, Pos + 1)
    End If
End Function

Public Sub ProgressShow(pic As PictureBox, ByVal sngPercent As Single)
Dim strPercent  As String
Dim intX        As Integer
Dim intY        As Integer
Dim intWidth    As Integer
Dim intHeight   As Integer
Dim intPercent
'Format percentage and get attributes of text
intPercent = Int(100 * sngPercent) ' + 0.5)
'get the forecolor of the picbox
'pic.ForeColor = pic.ForeColor ' UpdateForeCol
'allways white background
'Draw filled box
pic.DrawMode = 13
If sngPercent = 0 Then
    pic.Cls
    Else
    pic.Line (-10, -10)-(pic.Width * sngPercent, pic.Height), pic.ForeColor, BF
    End If
pic.Refresh
End Sub

Public Sub ShowHelpFile(strCont As String)
On Error Resume Next
' &h3 = help
' &hb =cdlHelpContents
' &h1 = cdlHelpContext
If strCont = "" Then
    frmMain.ComDlg.HelpCommand = &H3
    Else
    frmMain.ComDlg.HelpCommand = &H1
    frmMain.ComDlg.HelpContext = strCont
    End If
frmMain.ComDlg.ShowHelp
If Err > 0 Then
    MsgBox "Helpfile not found, vbCritical "
    Err.Clear
    End If
End Sub

Public Function TrimPath(ByVal Text As String, Size As Integer)
'trim file path according to label length ai "C:\...\test.bmp"
Dim TW
Dim Part1 As String
Dim Part2 As String
Dim Pos As Integer
Size = Size - 200
TW = frmMain.TextWidth(Text)
If TW < Size Then
    TrimPath = Text
    Exit Function
    End If
Part1 = Left(Text, 3) & "...\"
Part2 = Mid(Text, 4)
Text = Part1 & Part2
Do
TW = frmMain.TextWidth(Text)
If TW >= (Size) Then
    Pos = InStr(1, Part2, "\")
    If Pos <> 0 And Pos < Len(Part2) Then
        Part2 = Mid(Part2, Pos + 1)
        Else
        Part2 = Mid(Part2, 2)
        End If
    Text = Part1 & Part2
    End If
Loop While TW > (Size)
TrimPath = Text
End Function

Public Sub ControlsBlock()
'lock controls when read/write starts
With frmMain
.txtMain.Enabled = False
.picPreview.Enabled = False
.picInfo.Enabled = False
.picFileContainer.Enabled = False
.mnuImage.Enabled = False
.mnuEdit.Enabled = False
.mnuHelpa.Enabled = False
.Toolbar1.Enabled = False
End With
gblnIsRunning = True
End Sub

Public Sub ControlsFree()
'free controls when read/write finished
With frmMain
.txtMain.Enabled = True
.picPreview.Enabled = True
.picInfo.Enabled = True
.picFileContainer.Enabled = True
.mnuImage.Enabled = True
.mnuEdit.Enabled = True
.mnuHelpa.Enabled = True
.Toolbar1.Enabled = True
End With
gblnIsRunning = False
End Sub

Public Function FileExist(FileName As String) As Boolean
'checks weither a file exists
    On Error GoTo FileDoesNotExist
    Call FileLen(FileName)
    FileExist = True
    Exit Function
FileDoesNotExist:
    FileExist = False
End Function

Public Sub GetFile()
'get file information and show file icon
Dim fName As String
Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO
Dim cls_id As CLSID
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown
On Error Resume Next
'set name
If gstrFileName <> "" Then
    frmMain.lblFile.Caption = TrimPath(gstrFileName, frmMain.lblFile.Width)
    frmMain.lblFileInfo.Caption = Format(FileLen(gstrFileName), "###,##0") & " Bytes"
    frmMain.imgIcon.MousePointer = 99
    Else
    frmMain.lblFile.Caption = "Drag and drop file to this field..."
    frmMain.lblFileInfo.Caption = ""
    frmMain.imgIcon.Picture = Nothing
    frmMain.imgIcon.MousePointer = 0
    Exit Sub
    End If
'set file icon
fName = Trim(gstrFileName)
If fName = "" Then frmMain.imgIcon.Picture = Nothing: Exit Sub
SHGetFileInfo fName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_LARGEICON
hIcon = sh_info.hIcon
With new_icon
    .cbSize = Len(new_icon)
    .picType = vbPicTypeIcon
    .hIcon = hIcon
End With
With cls_id
    .id(8) = &HC0
    .id(15) = &H46
End With
hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
If hRes = 0 Then Set icon_pic = lpUnk
frmMain.imgIcon = icon_pic
Call CheckMenus
End Sub

Public Sub StartFile(ByVal FileName As String)
'run, open or start any file
Dim RunCmd As String
Dim fExt As String
Dim x As Long
Dim RetS
On Error Resume Next
fExt = UCase(GetFileExt(FileName))
Select Case fExt
Case "WAV", "MP2", "MP3", "MID", "AVI"
    RunCmd = "play"
Case Else
    RunCmd = "open"
End Select
If FileExist(FileName) = False Or FileName = "" Then Exit Sub
'open file
RetS = ShellExecute(frmMain.hwnd, RunCmd, FileName, "", App.path, 1)
If RetS = 31 Then
    'if open fails, open with...
    RetS = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & FileName)
    If RetS = 31 Then
        MsgBox "Can 't open or run this file.", vbInformation
        End If
    End If
End Sub

