VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "StegaNote"
   ClientHeight    =   6180
   ClientLeft      =   1935
   ClientTop       =   1470
   ClientWidth     =   8130
   FillColor       =   &H8000000A&
   HelpContextID   =   1
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   8130
   Begin VB.PictureBox picStatus 
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   7635
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Width           =   7695
      Begin VB.Label lblRND 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6480
         TabIndex        =   15
         Top             =   20
         Width           =   1095
      End
      Begin VB.Label lblBytes 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Bytes Text"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   20
         Width           =   1695
      End
   End
   Begin VB.PictureBox picFileContainer 
      Height          =   855
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   795
      ScaleWidth      =   7635
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3480
      Width           =   7695
      Begin VB.Label lblFileInfo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   960
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag and drop file to this field..."
         Height          =   195
         Left            =   960
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         Top             =   120
         Width           =   5295
      End
      Begin VB.Image imgIcon 
         Appearance      =   0  'Flat
         Height          =   585
         Left            =   240
         MouseIcon       =   "frmMain.frx":030A
         OLEDropMode     =   1  'Manual
         ToolTipText     =   " Click Icon to Open or Start this file "
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   6360
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picStegano 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7080
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picInfo 
      Height          =   1200
      Left            =   1560
      ScaleHeight     =   1140
      ScaleWidth      =   5835
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4800
      Width           =   5895
      Begin VB.PictureBox picProgBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1785
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblImageInfo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label lblImageFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Drag and drop carrier image into the black field..."
         Height          =   200
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H00000000&
      Height          =   1200
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   1140
      ScaleWidth      =   1440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1500
      Begin VB.Image imgPreview 
         Height          =   1065
         Left            =   120
         OLEDropMode     =   1  'Manual
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0614
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B58
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":109C
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15E0
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B24
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2068
            Key             =   "read"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25AC
            Key             =   "write"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AF0
            Key             =   "undo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Description     =   "clear image and text "
            Object.ToolTipText     =   " New Image And Text "
            ImageKey        =   "clear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "read"
            Description     =   "read data"
            Object.ToolTipText     =   " Read Data From Image"
            ImageKey        =   "read"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "write"
            Description     =   "write data"
            Object.ToolTipText     =   " Write Data To Image "
            ImageKey        =   "write"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Description     =   "Print"
            Object.ToolTipText     =   " Print "
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Description     =   "Cut"
            Object.ToolTipText     =   " Cut "
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Description     =   "Copy"
            Object.ToolTipText     =   " Copy "
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Description     =   "Paste"
            Object.ToolTipText     =   " Paste "
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            Description     =   "Undo"
            Object.ToolTipText     =   " Undo "
            ImageKey        =   "undo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3015
      HideSelection   =   0   'False
      Left            =   0
      MaxLength       =   32000
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   3  'Both
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   3375
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      WindowList      =   -1  'True
      Begin VB.Menu mnuSelectImage 
         Caption         =   "&Select..."
      End
      Begin VB.Menu mnuRead 
         Caption         =   "&Read..."
      End
      Begin VB.Menu mnuWrite 
         Caption         =   "&Write..."
      End
      Begin VB.Menu ln4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuErase 
         Caption         =   "&Erase..."
      End
      Begin VB.Menu lijn1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu ln3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu lijn4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu lijn9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "All to Clip&board"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clea&r"
         Begin VB.Menu mnuClearAll 
            Caption         =   "&All"
         End
         Begin VB.Menu mnuClearText 
            Caption         =   "&Text"
         End
         Begin VB.Menu mnuClearFile 
            Caption         =   "&File"
         End
         Begin VB.Menu mnuClearImage 
            Caption         =   "&Carrier Image"
         End
      End
      Begin VB.Menu ln6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertFile 
         Caption         =   "&Insert File.."
      End
   End
   Begin VB.Menu mnuHelpa 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu lijn12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'------------------------ MENUS --------------------------------

Private Sub mnuClearAll_Click()
If gblnTextHasChanged = True Then
    retVal = MsgBox("The changes to the text are not saved." & vbCrLfLf & "Do you want to clear image and text anyway?", vbQuestion + vbYesNo)
    If retVal = vbNo Then Exit Sub
    End If
gstrImageName = ""
gstrFileName = ""
gstrActiveKey = ""
gblnTextHasChanged = False
With Me
    .txtMain.Text = ""
    .imgPreview.Picture = Nothing
    .lblImageFile.Caption = "Drag and drop carrier image into the black field..."
    .lblImageInfo.Caption = ""
    .imgIcon.Picture = Nothing
End With
Call GetFile
Call CheckMenus
End Sub

Private Sub mnuClearFile_Click()
gstrFileName = ""
Call GetFile
End Sub

Private Sub mnuClearImage_Click()
gstrImageName = ""
Call GetImage
End Sub

Private Sub mnuClearText_Click()
If gblnTextHasChanged = True Then
    retVal = MsgBox("The changes to the text are not saved." & vbCrLfLf & "Do you want to clear the text anyway?", vbQuestion + vbYesNo)
    If retVal = vbNo Then Exit Sub
    End If
Me.txtMain.Text = ""
End Sub

Private Sub mnuErase_Click()
If gstrImageName = "" Then Exit Sub
retVal = MsgBox("Are you sure you want to erase all data in """ & CutFilePath(gstrImageName) & """ ?", vbYesNo + vbExclamation)
If retVal = vbNo Then Exit Sub
Me.picProgBar.Visible = True
Call ControlsBlock
Call EraseImage
Me.picProgBar.Visible = False
Call ControlsFree
End Sub

Private Sub mnuInsertFile_Click()
Dim fName As String
On Error Resume Next
With frmMain.ComDlg
.DialogTitle = "Insert File..."
.Flags = CMDLG_FILEMUSTEXIST Or CMDLG_NOCHECK
.FileName = ""
.InitDir = gstrCurDir
.Filter = "All Files (*.*)|*.*"
.ShowOpen
If Err = 32755 Then Exit Sub
gstrFileName = .FileName
End With
Call GetFile
End Sub

Private Sub mnuRead_Click()
If gstrImageName = "" Then Exit Sub
frmRead.Show (vbModal)
If gstrActiveKey = "" Or gblnCancelKey = True Then Exit Sub
Me.picProgBar.Visible = True
Me.lblProgress.Caption = "Reading data..."
Call ControlsBlock
Call LoadDataPicture
Me.picProgBar.Visible = False
Me.lblProgress.Caption = ""
Call ControlsFree
End Sub

Private Sub mnuSelectImage_Click()
Dim fName As String
On Error Resume Next
With frmMain.ComDlg
.DialogTitle = "Select Carrier Image..."
.Flags = CMDLG_FILEMUSTEXIST Or CMDLG_NOCHECK
.FileName = ""
.InitDir = gstrCurDir
.Filter = "Image Files (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
.ShowOpen
If Err = 32755 Then Exit Sub
gstrImageName = .FileName
End With
Call GetImage
End Sub

Private Sub mnuWrite_Click()
If gstrImageName = "" Then Exit Sub
frmWrite.Show (vbModal)
If gstrActiveKey = "" Or gblnCancelKey = True Then Exit Sub
Me.picProgBar.Visible = True
Me.lblProgress.Caption = "Writing data..."
Call ControlsBlock
Call SaveDataPicture
Me.lblProgress.Caption = ""
Me.picProgBar.Visible = False
Call ControlsFree
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show (vbModal)
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuHelp_Click()
Call ShowHelpFile("")
End Sub

Private Sub mnuPageSetup_Click()
frmPrintSetup.Show (vbModal)
End Sub

Private Sub mnuPrint_Click()
If gblnPrinterPresent = False Then Exit Sub
frmPrintProgress.Show (vbModal)
End Sub

Private Sub mnuCut_Click()
SendKeys "^{x}"
frmMain.Refresh
Exit Sub
Clipboard.SetText Me.txtMain.SelText
Me.txtMain.SelText = ""
frmMain.Refresh
Call CheckMenus
End Sub

Private Sub mnuCopy_Click()
'SendKeys "^{c}"
Clipboard.SetText Me.txtMain.SelText
frmMain.Refresh
Call CheckMenus
End Sub

Private Sub mnuPaste_Click()
SendKeys "^{v}"
frmMain.Refresh
Call CheckMenus
End Sub

Private Sub mnuDelete_Click()
SendKeys "{del}"
frmMain.Refresh
Call CheckMenus
End Sub

Private Sub mnuUndo_Click()
SendKeys "^{z}"
frmMain.Refresh
Call CheckMenus
End Sub

Private Sub mnuSelectAll_Click()
frmMain.txtMain.SelStart = 0
frmMain.txtMain.SelLength = Len(frmMain.txtMain.Text)
Call CheckMenus
End Sub

Private Sub mnuCopyAll_Click()
Clipboard.Clear
Clipboard.SetText frmMain.txtMain
Call CheckMenus
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'test buttons toolbar
Select Case Button.Key
Case "new"
    mnuClearAll_Click
Case "read"
    mnuRead_Click
Case "write"
    mnuWrite_Click
Case "print"
    mnuPrint_Click
Case "cut"
    mnuCut_Click
Case "copy"
    mnuCopy_Click
Case "paste"
    mnuPaste_Click
Case "undo"
    SendKeys "^{z}"
    frmMain.Refresh
End Select
Call CheckMenus
End Sub

'------------------------------------------------------------------

Private Sub imgIcon_Click()
'run, open or start file
If gstrFileName <> "" Then StartFile (gstrFileName)
End Sub

Private Sub picInfo_GotFocus()
If Me.txtMain.Enabled = True And Me.txtMain.Visible = True Then Me.txtMain.SetFocus
End Sub

Private Sub picFileContainer_GotFocus()
If Me.txtMain.Enabled = True And Me.txtMain.Visible = True Then Me.txtMain.SetFocus
End Sub

Private Sub picPreview_GotFocus()
If Me.txtMain.Enabled = True And Me.txtMain.Visible = True Then Me.txtMain.SetFocus
End Sub

Private Sub picStatus_GotFocus()
If Me.txtMain.Enabled = True And Me.txtMain.Visible = True Then Me.txtMain.SetFocus
End Sub

Private Sub txtMain_Change()
gblnTextHasChanged = True
Me.lblBytes.Caption = Format(Len(Me.txtMain), "###,##0") & " Bytes Text"
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
Call CheckMenus
End Sub

Private Sub txtMain_KeyUp(KeyCode As Integer, Shift As Integer)
Call CheckMenus
End Sub

Private Sub txtMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Call CheckMenus
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
Dim a As String
'controle input textbox
Call CheckMenus
If KeyAscii = 1 Then ' ctrl+A = select all
    frmMain.txtMain.SelStart = 0
    frmMain.txtMain.SelLength = Len(frmMain.txtMain)
    KeyAscii = 0
End If
If Len(frmMain.txtMain.Text) > 31999 Then
    Select Case KeyAscii
    Case 3, 8, 24
        'allow deleting and stuff
    Case Else
        KeyAscii = 0
        'Maximum lenghte text reached
        MsgBox "Maximum lenght of text is reached.", vbExclamation
        Exit Sub
    End Select
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'cancel write or read process
If KeyCode = 27 Then gblnCancelProcces = True
End Sub

Private Sub Form_Resize()
Call ChangeWindowSize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'if running, cancel exit
If gblnIsRunning = True Then Cancel = True: Exit Sub
'warn if text is changed
If gblnTextHasChanged = True Then
    retVal = MsgBox("The changes to the text are not saved." & vbCrLfLf & "Do you want to close the program anyway?", vbQuestion + vbYesNo)
    If retVal = vbNo Then Cancel = True: Exit Sub
    End If
Call SaveWindowPos
Unload frmAbout
Unload frmRead
Unload frmWrite
Unload frmPrintProgress
Unload frmPrintSetup
Clipboard.Clear
End
End Sub

'------------------------ Get rnd seed values ----------------------------

Private Sub txtMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub imgPreview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub picFileContainer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

'------------------------ Drag and Drop----------------------------

Private Sub imgIcon_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpFile As String
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
gstrFileName = Data.Files(1)
Call GetFile
End Sub

Private Sub imgIcon_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim tmp As String
If Data.GetFormat(vbCFFiles) Then
    Effect = vbDropEffectCopy
    Else
    Effect = vbDropEffectNone
    End If
End Sub

Private Sub picFileContainer_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpFile As String
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
gstrFileName = Data.Files(1)
Call GetFile
End Sub

Private Sub picFileContainer_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim tmp As String
If Data.GetFormat(vbCFFiles) Then
    Effect = vbDropEffectCopy
    Else
    Effect = vbDropEffectNone
    End If
End Sub

Private Sub lblFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpFile As String
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
gstrFileName = Data.Files(1)
Call GetFile
End Sub

Private Sub lblFile_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim tmp As String
If Data.GetFormat(vbCFFiles) Then
    Effect = vbDropEffectCopy
    Else
    Effect = vbDropEffectNone
    End If
End Sub

Private Sub lblFileInfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpFile As String
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
gstrFileName = Data.Files(1)
Call GetFile
End Sub

Private Sub lblFileInfo_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim tmp As String
If Data.GetFormat(vbCFFiles) Then
    Effect = vbDropEffectCopy
    Else
    Effect = vbDropEffectNone
    End If
End Sub

Private Sub imgPreview_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpFile As String
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
gstrImageName = Data.Files(1)
Call GetImage
End Sub

Private Sub imgPreview_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim tmp As String
If Data.GetFormat(vbCFFiles) Then tmp = Right(Data.Files(1), 4)
If tmp = ".bmp" Or tmp = ".jpg" Or tmp = ".gif" Then
    Effect = vbDropEffectCopy
    Else
    Effect = vbDropEffectNone
    End If
End Sub

Private Sub picPreview_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tmpFile As String
If Not Data.GetFormat(vbCFFiles) Then Exit Sub
gstrImageName = Data.Files(1)
Call GetImage
End Sub

Private Sub picPreview_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Dim tmp As String
If Data.GetFormat(vbCFFiles) Then tmp = Right(Data.Files(1), 4)
If tmp = ".bmp" Or tmp = ".jpg" Or tmp = ".gif" Then
    Effect = vbDropEffectCopy
    Else
    Effect = vbDropEffectNone
    End If
End Sub

'------------------- Set Carrier Image Preview --------------------

Private Sub GetImage()
Dim picFactor
Dim x
Dim y
On Error GoTo errHandle

Screen.MousePointer = 11

With frmMain
'set image
.picStegano.AutoSize = True
.picStegano.Picture = LoadPicture(gstrImageName)
.picStegano.Refresh
.picStegano.AutoSize = False

Screen.MousePointer = 0

'warn on bmp files, larger than 1 MB
DataPicX = .picStegano.ScaleWidth
DataPicY = .picStegano.ScaleHeight
BmpSize = (DataPicX * DataPicY * 3) / 1024
If BmpSize > 1024 And GetFileExt(gstrImageName) <> "bmp" Then MsgBox "StegaNote will save this image as bitmap (bmp)." & vbCrLfLf & _
"The new file size will be " & Format(BmpSize, "###,##0") & " Kb.", vbInformation

'set the image preview picture
.imgPreview.Visible = False
.imgPreview.Stretch = False
.imgPreview.Picture = .picStegano.Picture
picFactor = .imgPreview.Width / .imgPreview.Height
x = .picPreview.Width
y = .picPreview.Height
If Int(x / picFactor) <= y Then
    .imgPreview.Width = x
    .imgPreview.Height = Int(x / picFactor)
    .imgPreview.Top = (.picPreview.Height - .imgPreview.Height - 75) / 2
    .imgPreview.Left = 0
    Else
    .imgPreview.Height = y
    .imgPreview.Width = Int(y * picFactor)
    .imgPreview.Left = (.picPreview.Width - .imgPreview.Width - 75) / 2
    .imgPreview.Top = 0
    End If
.imgPreview.Stretch = True
.imgPreview.Visible = True
If gstrImageName <> "" Then
    .lblImageFile.Caption = TrimPath(gstrImageName, Me.lblImageFile.Width)
    .lblImageInfo.Caption = Format(FileLen(gstrImageName) / 1024, "###,##0") & " Kb"
    If GetAttr(gstrImageName) And 1 Then
        .lblImageInfo.Caption = .lblImageInfo.Caption & " (Read-Only)"
        gblnReadOnly = True
        Else
        gblnReadOnly = False
        If Right(gstrImageName, 4) <> ".bmp" Then
            .lblImageInfo.Caption = Format(BmpSize, "###,##0") & " Kb when converted to bitmap"
        End If
    End If
Call CheckMenus
End If
End With

Exit Sub

errHandle:
Screen.MousePointer = 0
'Failed loading image
MsgBox "Failed loading" & CutFilePath(gstrImageName) & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub


