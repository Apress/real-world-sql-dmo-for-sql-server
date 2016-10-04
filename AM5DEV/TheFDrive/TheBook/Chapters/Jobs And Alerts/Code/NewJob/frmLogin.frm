VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2820
   ClientLeft      =   6105
   ClientTop       =   4830
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1666.149
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkTrusted 
      Alignment       =   1  'Right Justify
      Caption         =   "Trusted Connection"
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   1620
      Width           =   2025
   End
   Begin VB.TextBox txtServer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   2
      Top             =   795
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   510
      TabIndex        =   5
      Top             =   2070
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2130
      TabIndex        =   6
      Top             =   2070
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1185
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   450
      Width           =   1155
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   1200
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkTrusted_Click()
If chkTrusted.Value = vbChecked Then
    txtUserName.Enabled = False
    txtPassword.Enabled = False
Else
    txtUserName.Enabled = True
    txtPassword.Enabled = True
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOK_Click()

On Error GoTo err_handler


If chkTrusted.Value = vbChecked Then
    oServer.LoginSecure = True
    oServer.Connect txtServer.Text
Else
    If txtServer.Text <> "" And txtUserName <> "" Then
        oServer.Login = txtUserName.Text
        oServer.Password = txtPassword.Text
    Else
        Exit Sub
    End If
End If


frmJobBuilder.Visible = True
Unload Me

Exit Sub

err_handler:

MsgBox "An Error Occured"
Exit Sub




End Sub
