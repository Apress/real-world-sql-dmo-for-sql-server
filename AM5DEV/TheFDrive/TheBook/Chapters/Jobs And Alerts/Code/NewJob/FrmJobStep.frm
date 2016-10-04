VERSION 5.00
Begin VB.Form FrmJobStep 
   Caption         =   "Job Step"
   ClientHeight    =   6300
   ClientLeft      =   3225
   ClientTop       =   3225
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   9195
   Begin VB.ComboBox cboOnfailureAction 
      Height          =   315
      Left            =   5850
      TabIndex        =   10
      Top             =   4380
      Width           =   1905
   End
   Begin VB.ComboBox cboOnSuccessAction 
      Height          =   315
      Left            =   5850
      TabIndex        =   9
      Top             =   3810
      Width           =   1905
   End
   Begin VB.TextBox txtStepText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1425
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1920
      Width           =   5535
   End
   Begin VB.ComboBox cboDatabase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox txtStepName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   2190
      TabIndex        =   2
      Top             =   570
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   585
      Left            =   7290
      TabIndex        =   1
      Top             =   5550
      Width           =   1515
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   585
      Left            =   5610
      TabIndex        =   0
      Top             =   5550
      Width           =   1515
   End
   Begin VB.Label Label6 
      Caption         =   "Failure"
      Height          =   195
      Left            =   5190
      TabIndex        =   12
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label Label5 
      Caption         =   "Success"
      Height          =   225
      Left            =   5160
      TabIndex        =   11
      Top             =   3870
      Width           =   645
   End
   Begin VB.Label Label4 
      Caption         =   "Step Text"
      Height          =   345
      Left            =   1050
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Database"
      Height          =   255
      Left            =   990
      TabIndex        =   6
      Top             =   1500
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Transact SQL Step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1740
      TabIndex        =   4
      Top             =   960
      Width           =   5205
   End
   Begin VB.Label Label1 
      Caption         =   "Step Name"
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   630
      Width           =   1245
   End
End
Attribute VB_Name = "FrmJobStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub GetJobStepCredentials()

Dim OjobSuccessAction As SQLDMO.SQLDMO_JOBSTEPACTION_TYPE
Dim OjobFailureAction As SQLDMO.SQLDMO_JOBSTEPACTION_TYPE


Select Case cboOnSuccessAction.Text

Case "Quit With Success"
OjobSuccessAction = 1
Case "Quit With Failure"
OjobSuccessAction = 2
Case "Goto Next Step"
OjobSuccessAction = 3
End Select


Select Case cboOnfailureAction.Text

Case "Quit With Success"
OjobFailureAction = 1
Case "Quit With Failure"
OjobFailureAction = 2
Case "Goto Next Step"
OjobFailureAction = 3
End Select



Set OJobStep = New SQLDMO.JobStep




OJobStep.stepid = stepid + 1
OJobStep.Name = txtStepName.Text
OJobStep.DatabaseName = cboDatabase.Text
OJobStep.SubSystem = "TSQL"
OJobStep.Command = txtStepText.Text
OJobStep.OnSuccessAction = OjobSuccessAction
OJobStep.OnFailAction = OjobFailureAction


oServer.JobServer.Jobs(strJobName).JobSteps.Add OJobStep
AddToStepList stepid, OJobStep.Name, OJobStep.SubSystem, cboOnSuccessAction.Text, cboOnfailureAction.Text

stepid = stepid + 1

Unload Me


End Sub

Private Sub AddToStepList(intstepid As Integer, strStepname As String, strStepType As String, strOnsuccess As String, strOnfailure As String)
frmJobBuilder.msf_JobStep.AddItem intstepid & vbTab & strStepname & vbTab & strStepType & vbTab & strOnsuccess & vbTab & strOnfailure
End Sub

Private Sub cmdAdd_Click()
If txtStepName.Text <> "" Then
    GetJobStepCredentials
End If

End Sub

Private Sub cmdCancel_Click()
Me.Visible = False
End Sub

 Private Sub LoadDatabases()
 For Each odatabase In oServer.Databases
    cboDatabase.AddItem odatabase.Name
Next odatabase

cboDatabase.ListIndex = 1
 End Sub

Sub loadActions()
cboOnfailureAction.AddItem "Quit With Success"
cboOnfailureAction.AddItem "Quit With Failure"
cboOnfailureAction.AddItem "Goto Next Step"

cboOnSuccessAction.AddItem "Quit With Success"
cboOnSuccessAction.AddItem "Quit With Failure"
cboOnSuccessAction.AddItem "Goto Next Step"

cboOnfailureAction.ListIndex = 1
cboOnSuccessAction.ListIndex = 0

End Sub

Private Sub Form_Load()
LoadDatabases
loadActions

End Sub
