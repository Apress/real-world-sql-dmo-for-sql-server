VERSION 5.00
Begin VB.Form frmAddGroups 
   Caption         =   "Add Users And Groups"
   ClientHeight    =   9870
   ClientLeft      =   1245
   ClientTop       =   2250
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   12300
   Begin VB.CommandButton cmdSRoleManage 
      Caption         =   "..."
      Height          =   2145
      Left            =   630
      TabIndex        =   14
      Top             =   2160
      Width           =   525
   End
   Begin VB.CommandButton cmdMaintenance 
      Caption         =   "..."
      Height          =   2775
      Left            =   11820
      TabIndex        =   13
      Top             =   2790
      Width           =   405
   End
   Begin VB.CommandButton cmdNewLogin 
      Caption         =   "New Login"
      Height          =   525
      Left            =   1170
      TabIndex        =   12
      Top             =   8730
      Width           =   2535
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
      ForeColor       =   &H00C00000&
      Height          =   2460
      Left            =   7500
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   6060
      Width           =   4425
   End
   Begin VB.ListBox lstLogins 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3060
      ItemData        =   "frmAddGroups.frx":0000
      Left            =   1170
      List            =   "frmAddGroups.frx":0002
      MouseIcon       =   "frmAddGroups.frx":0004
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   5550
      Width           =   5325
   End
   Begin VB.ListBox lstDBRoles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2760
      Left            =   7470
      TabIndex        =   3
      Top             =   2790
      Width           =   4305
   End
   Begin VB.ListBox lstSRoles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2160
      Left            =   1170
      TabIndex        =   2
      Top             =   2130
      Width           =   5115
   End
   Begin VB.ListBox lstDatabases 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1560
      Left            =   7470
      TabIndex        =   1
      Top             =   780
      Width           =   4155
   End
   Begin VB.ComboBox cboServers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   1170
      TabIndex        =   0
      Top             =   750
      Width           =   4965
   End
   Begin VB.Label Label6 
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7470
      TabIndex        =   11
      Top             =   5760
      Width           =   3165
   End
   Begin VB.Label Label5 
      Caption         =   "Logins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1170
      TabIndex        =   10
      Top             =   5160
      Width           =   2835
   End
   Begin VB.Label Label4 
      Caption         =   "Database Roles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7470
      TabIndex        =   9
      Top             =   2460
      Width           =   2955
   End
   Begin VB.Label Label3 
      Caption         =   "Server Roles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   8
      Top             =   1830
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "Databases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7470
      TabIndex        =   7
      Top             =   480
      Width           =   2865
   End
   Begin VB.Label Label1 
      Caption         =   "Servers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   6
      Top             =   450
      Width           =   2325
   End
End
Attribute VB_Name = "frmAddGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function FindIfUserOwnsAnything(UserName As String, DatabaseName As String) As Integer


On Error GoTo err_handler
'0 means owns nothing
'1 means owns something
'2 means error

Dim oList As SQLDMO.SQLObjectList
Dim obj As Object
Dim oUsr As SQLDMO.User

Set oUsr = New SQLDMO.User

FindIfUserOwnsAnything = 0



        'let's check the database owner first

        
            If UCase$(oServer.Databases(DatabaseName).Owner) = UCase$(UserName) Then
                FindIfUserOwnsAnything = 1
                Exit Function
            End If

        'Now the user objects
      
         For Each oUsr In oServer.Databases(DatabaseName).Users
            If UCase$(oUsr.Login) = UCase$(UserName) Then
                Set oList = oServer.Databases(DatabaseName).Users(oUsr.Name).ListOwnedObjects(SQLDMOObj_AllDatabaseUserObjects, SQLDMOObjSort_Type)
                    If oList.Count > 0 Then
                        FindIfUserOwnsAnything = 1
                        Exit Function
                    End If
            End If
        Next oUsr
        
Exit Function

err_handler:

FindIfUserOwnsAnything = 2
Exit Function
  
 
End Function

Private Sub LoadServers()

cboServers.Clear

For Each oGroup In SQLDMO.ServerGroups
    For Each oRServer In oGroup.RegisteredServers
        cboServers.AddItem oRServer.Name
    Next oRServer
Next oGroup
End Sub

Public Sub LoadDatabases()

lstDatabases.Clear
For Each oDatabase In oServer.Databases
    lstDatabases.AddItem oDatabase.Name
Next oDatabase
End Sub

Private Sub loadDatabaseUsers(DatabaseName As String)

lstUsers.Clear

For Each oUser In oServer.Databases(DatabaseName).Users
    lstUsers.AddItem oUser.Name & " (" & oUser.Login & ")"
Next oUser

End Sub

Private Sub loadDatabaseRoles(DatabaseName As String)

lstDBRoles.Clear

For Each oDBRole In oServer.Databases(DatabaseName).DatabaseRoles
    lstDBRoles.AddItem oDBRole.Name
Next oDBRole

End Sub


Private Sub loadLogins()

lstLogins.Clear

For Each oLogin In oServer.Logins
    lstLogins.AddItem oLogin.Name
Next oLogin

End Sub

Private Function AddDatabaseUser(LoginName As String, DatabaseName As String) As Integer
On Error GoTo err_handler
Dim oNewUser As SQLDMO.User
Set oNewUser = New SQLDMO.User

Dim exists As Integer

exists = 0

For Each oUser In oServer.Databases(DatabaseName).Users
    If UCase(oUser.Login) = UCase(LoginName) Then
        exists = 1
    End If
Next oUser

If exists = 0 Then
    oNewUser.Login = LoginName
    oNewUser.Name = IIf(InStr(1, LoginName, "\") > 0, Right(LoginName, Len(LoginName) - InStr(1, LoginName, "\")), LoginName)
    oServer.Databases(DatabaseName).Users.Add oNewUser
    MsgBox "New User " & oNewUser.Name & "(" & (oNewUser.Login) & ") added", vbInformation, "New User Added"
    AddDatabaseUser = 1
    Exit Function
Else
    MsgBox "User already Exists", vbInformation, "User already Exists"
    AddDatabaseUser = 0
    Exit Function
End If

Exit Function

err_handler:
MsgBox "Error Adding User: " & Err.Description, vbCritical, "Error"
AddDatabaseUser = 0
Exit Function


End Function


Private Sub loadSRoles()

lstsroles.Clear

For Each oSRole In oServer.ServerRoles
    lstsroles.AddItem oSRole.Name
Next oSRole

End Sub

Private Sub cboServers_Click()

Set oServer = New SQLDMO.SQLServer

oServer.LoginSecure = True
oServer.Connect cboServers.Text

LoadDatabases
loadLogins
loadSRoles

End Sub



Private Sub cmdMaintenance_Click()
If lstDatabases.SelCount > 0 Then
    frmDBRoles.Visible = True
End If
End Sub

Private Sub cmdNewLogin_Click()
frmNewLogin.Visible = True
End Sub

Private Sub cmdSRoleManage_Click()
frmManageLogins.Visible = True
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) \ 2
Me.Top = (Screen.Height - Me.Height) \ 2

LoadServers
End Sub

Private Sub lstdatabases_Click()
loadDatabaseRoles lstDatabases.Text
loadDatabaseUsers lstDatabases.Text
End Sub

Private Function ShowIfWeCanDeleteLogin(LoginName As String) As Integer

Dim CanWe As Integer
Dim oDatabase As SQLDMO.Database
Dim oList As SQLDMO.SQLObjectList
Dim obj As Object
Dim oUsr As SQLDMO.User

CanWe = 1



        'let's check the database owner first

        For Each oDatabase In oServer.Databases
            If UCase$(oDatabase.Owner) = UCase$(LoginName) Then
              CanWe = 0
            End If

        'Now the user objects
      
         For Each oUsr In oDatabase.Users
            If UCase$(oUsr.Login) = UCase$(LoginName) Then
                Set oList = oServer.Databases(oDatabase.Name).Users(oUsr.Name).ListOwnedObjects(SQLDMOObj_AllDatabaseUserObjects, SQLDMOObjSort_Type)
                    If oList.Count > 0 Then
                        CanWe = 0
                    End If
            End If
        Next oUsr
    Next oDatabase
    
 ShowIfWeCanDeleteLogin = CanWe




End Function

Private Sub lstLogins_DblClick()

If MsgBox("Delete Login " & lstLogins.Text & "?", vbQuestion + vbYesNo, "Remove Login") = vbNo Then
    Exit Sub
Else
    If ShowIfWeCanDeleteLogin(lstLogins.Text) = 1 Then
        oServer.Logins.Remove lstLogins.Text
        lstLogins.RemoveItem lstLogins.ListIndex
    Else
        MsgBox "Login Owns objects within 1 or more databases", vbInformation, "Removal Aborted"
    End If
End If

End Sub

Private Sub lstLogins_OLESetData(Data As DataObject, DataFormat As Integer)
Data = lstLogins.Text
End Sub

Private Sub lstUsers_DblClick()


Dim NameToCheck As String

NameToCheck = Mid(lstUsers.Text, 1, InStr(1, lstUsers.Text, "(") - 2)

If FindIfUserOwnsAnything(oServer.Databases(lstDatabases.Text).Users(NameToCheck).Login, lstDatabases.Text) = 0 Then
    'oUsr.Name = Mid(lstUsers.Text, 1, InStr(1, lstUsers.Text, "(") - 2)
    lstUsers.RemoveItem lstUsers.ListIndex
    oServer.Databases(lstDatabases.Text).Users(NameToCheck).Remove
    MsgBox "User Removed", vbInformation, "User Removed"
ElseIf FindIfUserOwnsAnything(oServer.Databases(lstDatabases.Text).Users(NameToCheck).Login, lstDatabases.Text) = 1 Then
    MsgBox "This User Owns Objects", vbInformation, "Cannot Remove Login"
ElseIf FindIfUserOwnsAnything(oServer.Databases(lstDatabases.Text).Users(NameToCheck).Login, lstDatabases.Text) = 2 Then
    MsgBox "An Error Occured"
End If


End Sub

Private Sub lstUsers_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim UserToAdd As String
UserToAdd = IIf(InStr(1, Data.GetData(1), "\") > 0, Right(Data.GetData(1), Len(Data.GetData(1)) - InStr(1, Data.GetData(1), "\")), Data.GetData(1)) + " (" + Data.GetData(1) + ")"



    If AddDatabaseUser(Data.GetData(1), lstDatabases.Text) = 1 Then
        lstUsers.AddItem UserToAdd
    End If


End Sub
