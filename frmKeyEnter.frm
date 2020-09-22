VERSION 5.00
Begin VB.Form frmRead 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Read data from image..."
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHide 
      Caption         =   "&Hide Typing"
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Read"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1170
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4050
   End
   Begin VB.Label lblCode 
      Caption         =   "Passphraze"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4245
   End
End
Attribute VB_Name = "frmRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Me.chkHide.Value = 1
Me.txtCode.PasswordChar = "*"
gblnCancelKey = False
If gstrActiveKey = "" Then
    Me.txtCode.Text = ""
    Me.txtCode.Text = ""
    Me.txtCode.SetFocus
    Else
    Me.txtCode.Text = gstrActiveKey
    Me.txtCode.Text = gstrActiveKey
    If Me.cmdOK.Enabled = True Then
        Me.cmdOK.SetFocus
        Else
        Me.txtCode.SetFocus
    End If
End If
End Sub

Private Sub chkHide_Click()
If Me.chkHide.Value = 1 Then
    Me.txtCode.PasswordChar = "*"
    Else
    Me.txtCode.PasswordChar = ""
    End If
End Sub

Private Sub cmdOK_Click()
If Me.txtCode.Text = "" Then Exit Sub
gstrActiveKey = Me.txtCode.Text
Me.Hide
End Sub

Private Sub cmdCancel_Click()
gblnCancelKey = True
Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub txtCode_Change()
If Len(Me.txtCode.Text) > 0 Then
    Me.cmdOK.Enabled = True
    Else
    Me.cmdOK.Enabled = False
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.txtCode <> "" And Me.cmdOK.Enabled = True Then cmdOK_Click
    End If
End Sub

Private Sub txtCode_GotFocus()
Me.txtCode.SelStart = 0
Me.txtCode.SelLength = Len(Me.txtCode.Text)
End Sub

