VERSION 5.00
Begin VB.Form frmNewLogin 
   Caption         =   "New Login"
   ClientHeight    =   3945
   ClientLeft      =   4215
   ClientTop       =   3720
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5625
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   4230
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   525
      Left            =   2700
      TabIndex        =   6
      Top             =   3390
      Width           =   1335
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   270
      TabIndex        =   3
      Top             =   960
      Width           =   4845
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "?"
      TabIndex        =   2
      Top             =   2550
      Width           =   4785
   End
   Begin VB.OptionButton optNT 
      Caption         =   "NT"
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   1860
      Width           =   2955
   End
   Begin VB.OptionButton optSS 
      Caption         =   "SQL Server"
      Height          =   345
      Left            =   240
      TabIndex        =   0
      Top             =   1500
      Width           =   2925
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   270
      TabIndex        =   5
      Top             =   750
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   2340
      Width           =   3285
   End
End
Attribute VB_Name = "frmNewLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

If optNT.Value = True And txtUserName.Text <> "" Then
    AddnewUser txtUserName, 1
    frmAddGroups.lstLogins.AddItem txtUserName
    ClearDown
    
    
ElseIf optSS.Value = True And txtUserName.Text <> "" Then
    AddnewUser txtUserName, 0, IIf(txtPassword.Text <> "", txtPassword.Text, "password")
    frmAddGroups.lstLogins.AddItem txtUserName
    ClearDown
    
End If


End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub AddnewUser(UserName As String, NTSQL As Integer, Optional Password As String)
Dim exists As Integer
Dim oNewLogin As SQLDMO.Login
Set oNewLogin = New SQLDMO.Login
exists = 0

For Each oLogin In oServer.Logins
    If UCase(oLogin.Name) = UCase(UserName) Then
        exists = 1
        MsgBox "Login already exists", vbInformation, "Login already Exists"
        Exit Sub
    End If
Next oLogin
    
 If exists = 0 Then
    If NTSQL = 1 Then
        oNewLogin.Name = UserName
        oNewLogin.Type = SQLDMOLogin_NTUser
    Else
        oNewLogin.Name = UserName
        oNewLogin.Type = SQLDMOLogin_Standard
        oNewLogin.SetPassword "", Password
    End If
End If


    
oServer.Logins.Add oNewLogin

    
    
End Sub

Private Sub ClearDown()
txtPassword.Text = ""
txtUserName.Text = ""
End Sub
