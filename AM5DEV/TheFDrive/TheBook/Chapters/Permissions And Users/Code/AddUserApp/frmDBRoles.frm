VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDBRoles 
   Caption         =   "Roles"
   ClientHeight    =   5955
   ClientLeft      =   3300
   ClientTop       =   2595
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7125
   Begin TabDlg.SSTab SSTab1 
      Height          =   6075
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   10716
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Add To Role"
      TabPicture(0)   =   "frmDBRoles.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdExit"
      Tab(0).Control(1)=   "cmdValidate"
      Tab(0).Control(2)=   "lstRoles"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "New Role"
      TabPicture(1)   =   "frmDBRoles.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtRoleName"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "optApp"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optNormal"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtPassword"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdAdd"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdClose"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   615
         Left            =   -69630
         TabIndex        =   11
         Top             =   5100
         Width           =   1515
      End
      Begin VB.CommandButton cmdValidate 
         Caption         =   "Add"
         Height          =   615
         Left            =   -71430
         TabIndex        =   10
         Top             =   5100
         Width           =   1515
      End
      Begin VB.ListBox lstRoles 
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
         Height          =   3690
         Left            =   -74670
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   930
         Width           =   6525
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   645
         Left            =   5340
         TabIndex        =   8
         Top             =   5010
         Width           =   1485
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   645
         Left            =   3360
         TabIndex        =   7
         Top             =   5010
         Width           =   1485
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         Left            =   630
         TabIndex        =   5
         Top             =   3600
         Width           =   5025
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "Standard"
         Height          =   435
         Left            =   540
         TabIndex        =   4
         Top             =   2340
         Width           =   3765
      End
      Begin VB.OptionButton optApp 
         Caption         =   "Application Role"
         Height          =   375
         Left            =   540
         TabIndex        =   3
         Top             =   2850
         Width           =   3465
      End
      Begin VB.TextBox txtRoleName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   510
         TabIndex        =   1
         Top             =   1380
         Width           =   5835
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   6
         Top             =   3330
         Width           =   3405
      End
      Begin VB.Label Label1 
         Caption         =   "Role Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   510
         TabIndex        =   2
         Top             =   1080
         Width           =   4485
      End
   End
End
Attribute VB_Name = "frmDBRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
Dim i As Integer
If txtRoleName.Text <> "" Then
    If optNormal.Value = True Then
        AddDatabaseRole txtRoleName.Text, 1, frmAddGroups.lstDatabases.Text
        ClearUp
    Else
        AddDatabaseRole txtRoleName.Text, 0, frmAddGroups.lstDatabases.Text, IIf(txtPassword.Text <> "", txtPassword.Text, "Password")
        ClearUp
    End If
    
    showAllDbroles frmAddGroups.lstDatabases.Text

For i = 0 To lstRoles.ListCount - 1
    If ShowRoleUserMemberOf(Left(frmAddGroups.lstUsers.Text, InStr(1, frmAddGroups.lstUsers.Text, "(") - 2), frmAddGroups.lstDatabases.Text, lstRoles.List(i)) = True Then
        lstRoles.Selected(i) = True
    End If
Next i
    
End If



End Sub

Private Sub ClearUp()
txtPassword.Text = ""
txtRoleName.Text = ""
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub showAllDbroles(DatabaseName As String)
lstRoles.Clear
For Each oDBRole In oServer.Databases(DatabaseName).DatabaseRoles
    If oDBRole.AppRole = False Then
        lstRoles.AddItem oDBRole.Name
    End If
    
Next oDBRole
End Sub

Private Sub cmdValidate_Click()


Dim j As Integer
Dim i As Integer

For j = 0 To lstRoles.ListCount - 1

Select Case ReturnValuesofMembership(Left(frmAddGroups.lstUsers.Text, InStr(1, frmAddGroups.lstUsers.Text, "(") - 2), frmAddGroups.lstDatabases.Text, lstRoles.List(j), lstRoles.Selected(j))
    
       Case NEEDS_ADDING
            oServer.Databases(frmAddGroups.lstDatabases.Text).DatabaseRoles(lstRoles.List(j)).AddMember Left(frmAddGroups.lstUsers.Text, InStr(1, frmAddGroups.lstUsers.Text, "(") - 2)

    Case NEEDS_DELETING
            oServer.Databases(frmAddGroups.lstDatabases.Text).DatabaseRoles(lstRoles.List(j)).DropMember Left(frmAddGroups.lstUsers.Text, InStr(1, frmAddGroups.lstUsers.Text, "(") - 2)

End Select

Next j

End Sub

Private Sub Form_Load()
lstRoles.Clear
Dim i As Integer
showAllDbroles frmAddGroups.lstDatabases.Text
optNormal.Value = True

For i = 0 To lstRoles.ListCount - 1
    If ShowRoleUserMemberOf(Left(frmAddGroups.lstUsers.Text, InStr(1, frmAddGroups.lstUsers.Text, "(") - 2), frmAddGroups.lstDatabases.Text, lstRoles.List(i)) = True Then
        lstRoles.Selected(i) = True
    End If
Next i
    
    

End Sub

Private Sub AddDatabaseRole(rolename As String, RoleType As Integer, DatabaseName As String, Optional Password As String)
'1 = standard
'0 = app

Dim exists As Integer
Dim onewdbrole As SQLDMO.DatabaseRole
Set onewdbrole = New SQLDMO.DatabaseRole
exists = 0

For Each oDBRole In oServer.Databases(DatabaseName).DatabaseRoles
    If UCase(oDBRole.Name) = UCase(rolename) Then
        exists = 1
        MsgBox "Role already Exists", vbInformation, "Exists"
        Exit Sub
    End If
Next oDBRole

If exists = 0 Then
    If RoleType = 1 Then
        onewdbrole.Name = rolename
        onewdbrole.AppRole = False
        oServer.Databases(DatabaseName).DatabaseRoles.Add onewdbrole
    ElseIf RoleType = 0 Then
        onewdbrole.Name = rolename
        onewdbrole.AppRole = True
        onewdbrole.Password = Password
        oServer.Databases(DatabaseName).DatabaseRoles.Add onewdbrole
    End If
End If


End Sub

Private Function ShowRoleUserMemberOf(UserName As String, DatabaseName As String, DBRole As String) As Boolean
Dim i As Integer
Dim j As Integer
ShowRoleUserMemberOf = False



Dim oQryresults As SQLDMO.QueryResults

Set oQryresults = oServer.Databases(DatabaseName).DatabaseRoles(DBRole).EnumDatabaseRoleMember

For i = 1 To oQryresults.Rows
        If oQryresults.GetColumnString(i, 1) = UserName Then
           ShowRoleUserMemberOf = True
        End If
  
Next i

End Function


Public Function ReturnValuesofMembership(strUserName As String, strDatabaseName As String, strRoleName As String, booChecked As Boolean) As Integer

Dim i As Integer
Dim booInList As Boolean


Dim oQList As SQLDMO.QueryResults

Set oQList = oServer.Databases(strDatabaseName).DatabaseRoles(strRoleName).EnumDatabaseRoleMember


    booInList = False   'assume we will not find the name
    
    For i = 1 To oQList.Rows
        If UCase(Trim(oQList.GetColumnString(i, 1))) = UCase(Trim(strUserName)) Then
            booInList = True
            Exit For
        End If
    Next i
       
    If booInList = True And booChecked = True Then
            'in the list and checkbox ticked = still a member
            ReturnValuesofMembership = IS_MEMBER
    End If
            'in the list and checkbox not ticked = needs removing
    If booInList = True And booChecked = False Then
        ReturnValuesofMembership = NEEDS_DELETING
    End If
    
    If booInList = False And booChecked = True Then
    'not in the list and checkbox ticked = needs adding
            ReturnValuesofMembership = NEEDS_ADDING
    End If
    
End Function


