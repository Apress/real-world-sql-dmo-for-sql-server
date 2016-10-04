VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2835
   ClientLeft      =   7050
   ClientTop       =   6435
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1675.012
   ScaleMode       =   0  'User
   ScaleWidth      =   4239.34
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chktrusted 
      Caption         =   "Use trusted Connection."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   3855
   End
   Begin VB.ComboBox cboServer 
      Height          =   315
      Left            =   1290
      TabIndex        =   7
      Top             =   270
      Width           =   2985
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   1245
      Width           =   2955
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   810
      TabIndex        =   4
      Top             =   2190
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2415
      TabIndex        =   5
      Top             =   2190
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1635
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   300
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   1260
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   1650
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oRserver As SQLDMO.RegisteredServer
Dim oApp As SQLDMO.Application
Dim oGroup As SQLDMO.ServerGroup

Private Sub ShowServers()

    Set oApp = New SQLDMO.Application
    Set oGroup = New SQLDMO.ServerGroup

    For Each oGroup In oApp.ServerGroups

        For Each oRserver In oGroup.RegisteredServers

            cboServer.AddItem oRserver.Name

        Next oRserver

    Next oGroup

End Sub

Private Sub chktrusted_Click()

    If chktrusted.Value = vbChecked Then

        txtPassword.Locked = True
        txtUserName.Locked = True
        txtPassword.Enabled = False
        txtUserName.Enabled = False

    Else

        txtPassword.Locked = False
        txtUserName.Locked = False
        txtPassword.Enabled = True
        txtUserName.Enabled = True

    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If cboServer.Text <> "" Then

        If chktrusted.Value = vbChecked Then

            Logon cboServer.Text, 1

        Else

            Logon cboServer.Text, 0

        End If

    End If

End Sub

Private Sub Form_Load()

    ShowServers

End Sub

Private Sub Logon(servername, Integrated As Integer)

    On Error GoTo err_handler

    Set oServer = New SQLDMO.SQLServer

    'check whether we require trusted connection or not

    If Integrated = 1 Then

        oServer.LoginSecure = True
        oServer.Connect servername

    Else

        oServer.Connect servername, txtUserName.Text, txtPassword.Text

    End If

    b_IsConnected = True
    glb_Server = oServer.Name
    Me.Visible = False

    Exit Sub

err_handler:
    MsgBox Err.Description, vbCritical, "Incorrect Login"
    Exit Sub

End Sub

