VERSION 5.00
Begin VB.Form frmPrintSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Print Setup"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5790
   ControlBox      =   0   'False
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrinter 
      Caption         =   "&Printer..."
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtSize 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "11"
         Top             =   2325
         Width           =   375
      End
      Begin VB.PictureBox picBlad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   3240
         ScaleHeight     =   2175
         ScaleWidth      =   1575
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   1575
         Begin VB.Image imgtekst 
            Height          =   1905
            Left            =   120
            Picture         =   "frmPrintSetup.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1320
         End
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "0"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "0"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtMarge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "0"
         Top             =   720
         Width           =   375
      End
      Begin VB.PictureBox PicShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   3360
         ScaleHeight     =   2175
         ScaleWidth      =   1575
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Font size"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2375
         Width           =   735
      End
      Begin VB.Label lblMarges 
         Alignment       =   2  'Center
         Caption         =   "Marges"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblExample 
         Alignment       =   2  'Center
         Caption         =   "Example"
         Height          =   255
         Left            =   3240
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblRight 
         Alignment       =   1  'Right Justify
         Caption         =   "Right"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "Left"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblBottom 
         Alignment       =   1  'Right Justify
         Caption         =   "Bottom"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblTop 
         Alignment       =   1  'Right Justify
         Caption         =   "Top"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   11
         Top             =   1845
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   10
         Top             =   1485
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   765
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrintSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
glngTopMarginPrint = Val(Me.txtMarge(0).Text)
glngBottMarginPrint = Val(Me.txtMarge(1).Text)
glngLeftMarginPrint = Val(Me.txtMarge(2).Text)
glngRightMarginPrint = Val(Me.txtMarge(3).Text)
glngPrintSize = Val(Me.txtSize.Text)
If gblnChangePageSetup = True Then
    SaveSetting App.EXEName, "Config", "PrintLeft", glngLeftMarginPrint
    SaveSetting App.EXEName, "Config", "PrintRight", glngRightMarginPrint
    SaveSetting App.EXEName, "Config", "PrintTop", glngTopMarginPrint
    SaveSetting App.EXEName, "Config", "PrintBottom", glngBottMarginPrint
    SaveSetting App.EXEName, "Config", "PrintSize", glngPrintSize
    End If
Me.Hide
End Sub

Private Sub cmdPrinter_Click()
On Error Resume Next
frmMain.ComDlg.Flags = &H4 Or &H100000
frmMain.ComDlg.ShowPrinter
Call SetExample
End Sub

Private Sub Form_Activate()
Me.txtMarge(0).Text = Trim(Str(glngTopMarginPrint))
Me.txtMarge(1).Text = Trim(Str(glngBottMarginPrint))
Me.txtMarge(2).Text = Trim(Str(glngLeftMarginPrint))
Me.txtMarge(3).Text = Trim(Str(glngRightMarginPrint))
Me.txtSize.Text = Trim(Str(glngPrintSize))
Me.cmdOK.SetFocus
gblnChangePageSetup = False
Call SetExample
End Sub

Private Sub txtMarge_Change(Index As Integer)
gblnChangePageSetup = True
Call DrawMargesExample
End Sub

Private Sub txtMarge_GotFocus(Index As Integer)
Me.txtMarge(Index).SelStart = 0
Me.txtMarge(Index).SelLength = Len(Me.txtMarge(Index))
End Sub

Private Sub txtMarge_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii > 29 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txtSize_GotFocus()
Me.txtSize.SelStart = 0
Me.txtSize.SelLength = Len(Me.txtSize.Text)
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
If KeyAscii > 29 And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txtSize_Change()
gblnChangePageSetup = True
End Sub

Private Sub SetExample()
With Me
If Printer.Orientation = vbPRORPortrait Then
    'portrait
    .picBlad.Top = 480
    .picBlad.Left = 3240
    .picBlad.Width = 1575
    .picBlad.Height = 2175
    .PicShadow.Width = 1575
    .PicShadow.Height = 2175
    .PicShadow.Top = 600
    .PicShadow.Left = 3360
    Else
    'landscape
    .picBlad.Top = 720
    .picBlad.Left = 2880
    .picBlad.Width = 2175
    .picBlad.Height = 1575
    .PicShadow.Width = 2175
    .PicShadow.Height = 1575
    .PicShadow.Top = 840
    .PicShadow.Left = 3000
    End If
End With
Call DrawMargesExample
End Sub

Private Sub DrawMargesExample()
tm = Val(Me.txtMarge(0).Text)
bm = Val(Me.txtMarge(1).Text)
lm = Val(Me.txtMarge(2).Text)
rm = Val(Me.txtMarge(3).Text)
If tm + bm > 95 Then
    tm = 5
    bm = 5
    Me.txtMarge(0).Text = "5"
    Me.txtMarge(1).Text = "5"
    End If
If lm + rm > 95 Then
    lm = 5
    rm = 5
    Me.txtMarge(2).Text = "5"
    Me.txtMarge(3).Text = "5"
    End If
SheetWidht = Me.picBlad.Width
SheetHeight = Me.picBlad.Height
Me.imgtekst.Width = Int(SheetWidht / 100 * (100 - lm - rm))
Me.imgtekst.Height = Int(SheetHeight / 100 * (100 - tm - bm))
Me.imgtekst.Top = Int(SheetHeight / 100 * tm)
Me.imgtekst.Left = Int(SheetWidht / 100 * lm)
End Sub
