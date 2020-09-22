VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrintProgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Printing"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4590
   ControlBox      =   0   'False
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblDevice 
      Caption         =   "[device name]"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblPrintComment 
      Caption         =   "Sending to..."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmPrintProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
gblnCancelPrint = True
End Sub

Private Sub Form_Activate()
gblnCancelPrint = False
Me.lblDevice.Caption = Printer.DeviceName
Me.ProgressBar1.Value = 0
Me.Refresh
'set printer font
Printer.FontName = frmMain.txtMain.FontName
Printer.FontSize = glngPrintSize
Printer.FontBold = frmMain.txtMain.FontBold
Printer.FontItalic = frmMain.txtMain.FontItalic
If gintPrintSelection = 1 Or gintPrintSelection = 33 Then
        'print selection
        Call PrintString(frmMain.txtMain.SelText, glngLeftMarginPrint, glngRightMarginPrint, glngTopMarginPrint, glngBottMarginPrint)
        Else
        'print all text
        Call PrintString(frmMain.txtMain.Text, glngLeftMarginPrint, glngRightMarginPrint, glngTopMarginPrint, glngBottMarginPrint)
        End If
Me.Hide
End Sub

