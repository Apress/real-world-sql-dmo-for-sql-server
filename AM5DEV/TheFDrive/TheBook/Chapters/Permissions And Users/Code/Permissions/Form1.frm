VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "View"
   ClientHeight    =   6180
   ClientLeft      =   2895
   ClientTop       =   5880
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11610
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7950
      TabIndex        =   7
      Top             =   5430
      Width           =   1515
   End
   Begin VB.CheckBox chkInsert 
      Caption         =   "INSERT"
      Height          =   345
      Left            =   7110
      TabIndex        =   6
      Top             =   2451
      Width           =   3915
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "DELETE"
      Height          =   345
      Left            =   7110
      TabIndex        =   5
      Top             =   1874
      Width           =   3915
   End
   Begin VB.CheckBox chkUpdate 
      Caption         =   "UPDATE"
      Height          =   345
      Left            =   7110
      TabIndex        =   4
      Top             =   1297
      Width           =   3915
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "SELECT"
      Height          =   345
      Left            =   7110
      TabIndex        =   3
      Top             =   720
      Width           =   3915
   End
   Begin VB.ListBox lstTables 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4260
      Left            =   150
      TabIndex        =   2
      Top             =   690
      Width           =   6375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9630
      TabIndex        =   0
      Top             =   5430
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Tables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   330
      Width           =   3135
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ValidatePermissions(DatabaseName As String, tableName As String, UserName As String)
DoEvents



Dim i As Integer


Set oList = oServer.Databases(DatabaseName).Tables(tableName).ListUserPermissions(UserName)

For Each objPermission In oList

    Select Case objPermission.PrivilegeTypeName

    Case "Select"
    chkSelect.Value = vbChecked
    
    Case "Insert"
    chkInsert.Value = vbChecked
    
    Case "Update"
    chkUpdate.Value = vbChecked
    
    Case "Delete"
    chkDelete.Value = vbChecked
    
End Select
    
Next objPermission

End Sub

Sub AlterPermissionOrNot(DatabaseName As String, tableName As String, UserName As String)

Dim Sselect As Integer
Dim Iinsert As Integer
Dim Uupdate As Integer
Dim Ddelete As Integer


Sselect = 0
Ddelete = 0
Iinsert = 0
Uupdate = 0



For Each objPermission In oList

If objPermission.PrivilegeTypeName = "Select" Then
    Sselect = 1
End If

If objPermission.PrivilegeTypeName = "Update" Then
    Uupdate = 1
End If

If objPermission.PrivilegeTypeName = "Delete" Then
    Ddelete = 1
End If

If objPermission.PrivilegeTypeName = "Insert" Then
    Iinsert = 1
End If

Next objPermission


If Sselect = 1 And chkSelect.Value = vbUnchecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Revoke SQLDMOPriv_Select, UserName
ElseIf Sselect = 0 And chkSelect.Value = vbChecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Grant SQLDMOPriv_Select, UserName
End If

If Uupdate = 1 And chkUpdate.Value = vbUnchecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Revoke SQLDMOPriv_Update, UserName
ElseIf Uupdate = 0 And chkUpdate.Value = vbChecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Grant SQLDMOPriv_Update, UserName
End If

If Iinsert = 1 And chkInsert.Value = vbUnchecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Revoke SQLDMOPriv_Insert, UserName
ElseIf Iinsert = 0 And chkInsert.Value = vbChecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Grant SQLDMOPriv_Insert, UserName
End If

If Ddelete = 1 And chkDelete.Value = vbUnchecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Revoke SQLDMOPriv_Delete, UserName
ElseIf Ddelete = 0 And chkDelete.Value = vbChecked Then
    oServer.Databases(DatabaseName).Tables(tableName).Grant SQLDMOPriv_Delete, UserName
End If













End Sub


Private Sub ClearBoxes()
chkInsert.Value = vbUnchecked
chkSelect.Value = vbUnchecked
chkUpdate.Value = vbUnchecked
chkDelete.Value = vbUnchecked

End Sub

Private Sub cmdClose_Click()
Me.Visible = False
End Sub
Private Sub PopulateTables()

For Each oTable In oServer.Databases(frmpermissions.cboDatabases.Text).Tables
    lstTables.AddItem oTable.Name
Next oTable

End Sub


Private Sub cmdUpdate_Click()
AlterPermissionOrNot frmpermissions.cboDatabases.Text, lstTables.Text, frmpermissions.lstUsers.Text
End Sub

Private Sub Form_Load()

PopulateTables
End Sub


Private Sub lstTables_Click()
ClearBoxes
ValidatePermissions frmpermissions.cboDatabases.Text, lstTables.Text, frmpermissions.lstUsers.Text

End Sub
