VERSION 5.00
Begin VB.Form frmpermissions 
   Caption         =   "Permissions"
   ClientHeight    =   7065
   ClientLeft      =   3375
   ClientTop       =   2625
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   7365
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   6810
      TabIndex        =   7
      Top             =   3840
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmpermissions.frx":0000
      Top             =   240
      Width           =   6975
   End
   Begin VB.ComboBox cboDatabases 
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
      Left            =   660
      TabIndex        =   2
      Top             =   1920
      Width           =   6105
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3960
      Left            =   660
      TabIndex        =   1
      Top             =   2910
      Width           =   5985
   End
   Begin VB.ComboBox cboServer 
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
      Left            =   660
      TabIndex        =   0
      Top             =   1080
      Width           =   6075
   End
   Begin VB.Label Label3 
      Caption         =   "Users"
      Height          =   285
      Left            =   660
      TabIndex        =   5
      Top             =   2670
      Width           =   2925
   End
   Begin VB.Label Label2 
      Caption         =   "Databases"
      Height          =   225
      Left            =   660
      TabIndex        =   4
      Top             =   1170
      Width           =   2475
   End
   Begin VB.Label Label1 
      Caption         =   "Servers"
      Height          =   315
      Left            =   660
      TabIndex        =   3
      Top             =   810
      Width           =   3915
   End
End
Attribute VB_Name = "frmpermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PopulateServers()
cboServer.Clear

For Each oGroup In SQLDMO.ServerGroups
    For Each oRServer In oGroup.RegisteredServers
        cboServer.AddItem oRServer.Name
    Next oRServer
Next oGroup

End Sub

Private Sub PopulateDatabases(ServerName As String)
cboDatabases.Clear

If ServerName <> "" Then
    Set oServer = New SQLDMO.SQLServer
    oServer.LoginSecure = True
    oServer.Connect ServerName
        For Each oDatabase In oServer.Databases
            cboDatabases.AddItem oDatabase.Name
        Next oDatabase
   
End If

End Sub


Private Sub PopulateUsers(ServerName As String, DatabaseName As String)
lstUsers.Clear

If ServerName <> "" And DatabaseName <> "" Then

        For Each oUser In oServer.Databases(DatabaseName).Users
            lstUsers.AddItem oUser.Name
        Next oUser
End If


End Sub

Private Sub cboDatabases_Click()
PopulateUsers cboServer.Text, cboDatabases.Text
End Sub

Private Sub cboServer_Click()
PopulateDatabases (cboServer.Text)
End Sub

Private Sub Command1_Click()
Unload Me
Application.Quit
End Sub

Private Sub Form_Load()
PopulateServers
End Sub

Private Sub lstUsers_Click()
frmView.Visible = True
End Sub
