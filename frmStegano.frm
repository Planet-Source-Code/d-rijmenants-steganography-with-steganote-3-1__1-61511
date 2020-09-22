VERSION 5.00
Begin VB.Form frmWrite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Save data to Image..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ControlBox      =   0   'False
   HelpContextID   =   1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picProgBar 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2280
      ScaleHeight     =   135
      ScaleWidth      =   1995
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.CheckBox chkHide 
      Caption         =   "&Hide typing"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1620
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Write"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblQuality 
      Alignment       =   1  'Right Justify
      Caption         =   "Key Quality"
      Height          =   225
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Confirmation"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Passphraze"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkHide_Click()
If Me.chkHide.Value = 1 Then
    Me.txtKey.PasswordChar = "*"
    Me.txtConfirm.PasswordChar = "*"
    Else
    Me.txtKey.PasswordChar = ""
    Me.txtConfirm.PasswordChar = ""
    End If
End Sub

Private Sub cmdSave_Click()
If KeyQuality(Me.txtKey.Text) < 20 Then
    MsgBox "Passphrase too short or contains repetitions.", vbCritical
    Exit Sub
    End If
If Me.txtConfirm.Text <> Me.txtKey.Text Then
    MsgBox "Passphrase and confirmation do not match.", vbCritical
    Exit Sub
    End If
If gstrImageName = "Unsaved" Then
    MsgBox "Please select image to save the text.", vbCritical
    Exit Sub
    End If
gstrActiveKey = Me.txtKey.Text
Me.Hide
End Sub

Private Sub Form_Activate()
gblnCancelKey = False
If gstrActiveKey <> "" Then
    Me.txtKey.Text = gstrActiveKey
    Me.txtConfirm.Text = gstrActiveKey
    If Me.cmdSave.Enabled = True Then
        Me.cmdSave.SetFocus
        Else
        Me.txtKey.SetFocus
        End If
    Else
    Me.txtKey.Text = ""
    Me.txtConfirm.Text = ""
    Me.txtKey.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
gblnCancelKey = True
Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub imgPreview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub txtCode_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call RandomFeed(x, y)
End Sub

Private Sub txtKey_Change()
ProgressShow Me.picProgBar, KeyQuality(Me.txtKey.Text) / 100
Me.txtConfirm.Text = ""
If Len(Me.txtKey.Text) < 5 Then
    Me.cmdSave.Enabled = False
    Else
    Me.cmdSave.Enabled = True
    End If
End Sub

Private Sub txtKey_GotFocus()
Me.txtKey.SelStart = 0
Me.txtKey.SelLength = Len(Me.txtKey.Text)
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0: Me.txtConfirm.SetFocus
End Sub

Private Sub txtConfirm_GotFocus()
Me.txtConfirm.SelStart = 0
Me.txtConfirm.SelLength = Len(Me.txtConfirm.Text)
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.cmdSave.Enabled = True Then Me.cmdSave.SetFocus
    End If
End Sub


