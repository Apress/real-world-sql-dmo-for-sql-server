VERSION 5.00
Begin VB.Form frmManageLogins 
   Caption         =   "Manage Logins"
   ClientHeight    =   5610
   ClientLeft      =   4335
   ClientTop       =   2595
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   4770
      TabIndex        =   2
      Top             =   4980
      Width           =   1305
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   2940
      TabIndex        =   1
      Top             =   4980
      Width           =   1305
   End
   Begin VB.ListBox lstsroles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   480
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   570
      Width           =   5625
   End
End
Attribute VB_Name = "frmManageLogins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub populateRoles()

lstsroles.Clear

For Each oSRole In oServer.ServerRoles
    lstsroles.AddItem oSRole.Name
Next oSRole
End Sub

Private Function ShowSRoleLoginMemberOf(LoginName As String, SRole As String) As Boolean
Dim i As Integer

ShowSRoleLoginMemberOf = False


Dim oQryresults As SQLDMO.QueryResults

Set oQryresults = oServer.ServerRoles(SRole).EnumServerRoleMember

For i = 1 To oQryresults.Rows
        If oQryresults.GetColumnString(i, 1) = LoginName Then
           ShowSRoleLoginMemberOf = True
        End If
  
Next i

End Function


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSubmit_Click()
Dim j As Integer
Dim i As Integer

For j = 0 To lstsroles.ListCount - 1

Select Case ReturnValuesofSMembership(frmAddGroups.lstLogins.Text, lstsroles.List(j), lstsroles.Selected(j))
    
       Case NEEDS_ADDING
            oServer.ServerRoles(lstsroles.List(j)).AddMember frmAddGroups.lstLogins.Text
            
    Case NEEDS_DELETING
               oServer.ServerRoles(lstsroles.List(j)).DropMember frmAddGroups.lstLogins.Text
           
End Select

Next j

End Sub

Private Sub Form_Load()
lstsroles.Clear
Dim i As Integer
populateRoles

For i = 0 To lstsroles.ListCount - 1
    If ShowSRoleLoginMemberOf(frmAddGroups.lstLogins.Text, lstsroles.List(i)) = True Then
        lstsroles.Selected(i) = True
    End If
Next i
End Sub

Public Function ReturnValuesofSMembership(strLoginName As String, strSRoleName As String, booChecked As Boolean) As Integer

Dim i As Integer
Dim booInList As Boolean


Dim oQList As SQLDMO.QueryResults

Set oQList = oServer.ServerRoles(strSRoleName).EnumServerRoleMember


    booInList = False   'assume we will not find the name
    
    For i = 1 To oQList.Rows
        If UCase(Trim(oQList.GetColumnString(i, 1))) = UCase(Trim(strLoginName)) Then
            booInList = True
            Exit For
        End If
    Next i
       
    If booInList = True And booChecked = True Then
            'in the list and checkbox ticked = still a member
            ReturnValuesofSMembership = IS_MEMBER
    End If
            'in the list and checkbox not ticked = needs removing
    If booInList = True And booChecked = False Then
        ReturnValuesofSMembership = NEEDS_DELETING
    End If
    
    If booInList = False And booChecked = True Then
    'not in the list and checkbox ticked = needs adding
            ReturnValuesofSMembership = NEEDS_ADDING
    End If
    
End Function

