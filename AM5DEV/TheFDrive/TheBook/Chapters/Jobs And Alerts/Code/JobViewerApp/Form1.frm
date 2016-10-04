VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmJobsMain 
   Caption         =   "Main Job Viewer Form"
   ClientHeight    =   10005
   ClientLeft      =   735
   ClientTop       =   570
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   14085
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   150
      TabIndex        =   18
      Top             =   9210
      Width           =   13785
   End
   Begin MSFlexGridLib.MSFlexGrid fg_JobHistory 
      Height          =   1725
      Left            =   120
      TabIndex        =   15
      Top             =   7380
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   3043
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      FillStyle       =   1
      AllowUserResizing=   3
   End
   Begin VB.CommandButton cmdJobHistory 
      Caption         =   "Job History"
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   9630
      Width           =   13785
   End
   Begin VB.Frame frJobDetails 
      Caption         =   "Job Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   8040
      TabIndex        =   4
      Top             =   210
      Width           =   5865
      Begin VB.TextBox txtSchedule 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   5430
         Width           =   5385
      End
      Begin VB.OptionButton optEnabled 
         Alignment       =   1  'Right Justify
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   810
         Width           =   2025
      End
      Begin VB.Label Label2 
         Caption         =   "Schedule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   17
         Top             =   5220
         Width           =   2955
      End
      Begin VB.Label lblCurrentRunStatus 
         Caption         =   "Current Run Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   4800
         Width           =   5565
      End
      Begin VB.Label lblNextRunDate 
         Caption         =   "Next Run Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   13
         Top             =   4174
         Width           =   5505
      End
      Begin VB.Label lblCategory 
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   3642
         Width           =   5475
      End
      Begin VB.Label lblLastRunDate 
         Caption         =   "Last Run Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1282
         Width           =   5415
      End
      Begin VB.Label lblOwner 
         Caption         =   "Owner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3170
         Width           =   5415
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2698
         Width           =   5415
      End
      Begin VB.Label lblDateCreated 
         Caption         =   "Date Created:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2226
         Width           =   5415
      End
      Begin VB.Label lblLastRunTime 
         Caption         =   "Last Run Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1754
         Width           =   5415
      End
   End
   Begin MSComctlLib.ImageList il1 
      Left            =   3900
      Top             =   5730
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":031A
            Key             =   "server"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":076C
            Key             =   "sjob"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BBE
            Key             =   "fjob"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1010
            Key             =   "sstep"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1462
            Key             =   "fstep"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtStepCommand 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   5910
      Width           =   7785
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Tree"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4980
      Width           =   7665
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   4725
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   8334
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "il1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Job History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Top             =   7140
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Step Command"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   5520
      Width           =   4725
   End
End
Attribute VB_Name = "frmJobsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadRegisteredServers()

For Each oGroup In SQLDMO.ServerGroups
    For Each oRserver In oGroup.RegisteredServers
        tv1.Nodes.Add "ROOT", tvwChild, "SERVER_" & oRserver.Name, oRserver.Name, il1.ListImages(2).Key
            tv1.Nodes.Add "SERVER_" & oRserver.Name, tvwChild, "DUMMY_" & oRserver.Name, "DUMMY"
    Next oRserver
Next oGroup
End Sub


Private Sub LoadJobList(strServerclicked As String, strServerKey As String, intCountOfChildren As Integer, strFirstChildName As String, intFirstChildIndex As Integer)

On Error GoTo err_handler


'Check if the server we are logged onto is
'the server we last used. If not log onto it
'if yes then carry on

If strLastServerLoggedOnTo = "GOBBLEDYGOOK" Then
    oServer.LoginSecure = True
    oServer.Connect strServerclicked
    strLastServerLoggedOnTo = strServerclicked
ElseIf strServerclicked <> strLastServerLoggedOnTo Then
    oServer.DisConnect
    oServer.LoginSecure = True
    oServer.Connect strServerclicked
    strLastServerLoggedOnTo = strServerclicked
End If


'check if the node has children and the first node's text is not DUMMY
'if yes then prepopulated so get out

If intCountOfChildren > 0 And strFirstChildName <> "DUMMY" Then
    Exit Sub
End If

'If the first node is DUMMY then let's attempt a population

If strFirstChildName = "DUMMY" Then
    tv1.Nodes.Remove intFirstChildIndex
        For Each oJob In oServer.JobServer.Jobs
            If oJob.LastRunOutcome = SQLDMOJobOutcome_Succeeded Or oJob.LastRunOutcome = SQLDMOJobOutcome_Unknown Then
                tv1.Nodes.Add "SERVER_" & strServerclicked, tvwChild, "JOB_" & strServerclicked & "_" & oJob.Name, oJob.Name, il1.ListImages(3).Key
            ElseIf oJob.LastRunOutcome = SQLDMOJobOutcome_Failed Then
                tv1.Nodes.Add "SERVER_" & strServerclicked, tvwChild, "JOB_" & strServerclicked & "_" & oJob.Name, oJob.Name, il1.ListImages(4).Key
            End If
            
            tv1.Nodes.Add "JOB_" & strServerclicked & "_" & oJob.Name, tvwChild, "DUMMY_" & strServerclicked & "_" & oJob.Name, "DUMMY"
        Next oJob
End If

Exit Sub

err_handler:
tv1.Nodes.Remove intFirstChildIndex
tv1.Nodes.Add "SERVER_" & strServerclicked, tvwChild, "ERROR_" & strServerclicked, "ERROR LOGGING INTO SERVER"
            
Exit Sub
End Sub



Private Sub ShowJobHistory(ServerName As String, JobName As String, booFailedOrAll As Boolean)

Dim oJobHistory As SQLDMO.QueryResults
Dim i As Integer
Dim j As Integer


fg_JobHistory.Rows = 1

oServer.JobServer.JobHistoryFilter.JobName = JobName

If booFailedOrAll = True Then
    oServer.JobServer.JobHistoryFilter.OutcomeTypes = SQLDMOJobOutcome_Failed
End If


Set oJobHistory = oServer.JobServer.EnumJobHistory(oServer.JobServer.JobHistoryFilter)


    For i = 1 To oJobHistory.Rows
        fg_JobHistory.AddItem oJobHistory.GetColumnString(i, 3) & vbTab & oJobHistory.GetColumnString(i, 4) & vbTab & oJobHistory.GetColumnString(i, 5) & vbTab & oJobHistory.GetColumnString(i, 6) & vbTab & oJobHistory.GetColumnString(i, 7) & vbTab & oJobHistory.GetColumnString(i, 8) & vbTab & oJobHistory.GetColumnString(i, 9) & vbTab & oJobHistory.GetColumnString(i, 10)
    Next i


    




End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdJobHistory_Click()

If Left(tv1.SelectedItem.Key, 4) = "JOB_" Then
fg_JobHistory.Rows = 1
ShowJobHistory tv1.SelectedItem.Parent.Text, tv1.SelectedItem.Text, False
End If

End Sub

Private Sub cmdRefresh_Click()

If strLastServerLoggedOnTo <> "GOBBLEDYGOOK" Then
    oServer.DisConnect
    strLastServerLoggedOnTo = "GOBBLEDYGOOK"
End If
tv1.Nodes.Clear
tv1.Nodes.Add , , "ROOT", "JOB SERVERS", il1.ListImages(1).Key
LoadRegisteredServers
ClearStepBoxes

End Sub

Private Sub cmdSchedule_Click()

If Left(tv1.SelectedItem.Key, 4) = "JOB_" Then
    ShowJobSchedule tv1.SelectedItem.Parent.Text, tv1.SelectedItem.Text
End If

End Sub

Private Sub Form_Load()
tv1.Nodes.Add , , "ROOT", "JOB SERVERS", il1.ListImages(1).Key
LoadRegisteredServers
End Sub

Private Sub tv1_Expand(ByVal Node As MSComctlLib.Node)

If Left(Node.Key, 6) = "SERVER" Then
    LoadJobList Node.Text, Node.Key, Node.Children, Node.Child.Text, Node.Child.Index
    ClearStepBoxes

End If

If Left(Node.Key, 3) = "JOB" Then
    LoadJobStepList Node.Text, Node.Parent.Text, Node.Key, Node.Children, Node.Child.Text, Node.Child.Index
End If

End Sub

Private Sub ClearStepBoxes()
lblDateCreated = "Date Created:"
lblDescription = "Description"
lblLastRunDate = "Last Run Date"
lblLastRunTime = "Last Run Time"
lblOwner = "Owner"
lblCategory = "Category"
lblNextRunDate = "Next Run Date"
lblCurrentRunStatus = "Current Run Status"
txtStepCommand.Text = ""

End Sub

Private Sub LoadJobStepList(JobName As String, strServer As String, strJobKey As String, intCountOfChildren As Integer, strFirstChildName As String, intFirstChildIndex As Integer)

On Error GoTo err_handler

'do not need to check if the strLastServerLoggedOnTo is set to GOOBLEDYGOOK
'as if we can see the steps we've already logged in somewhere at least


If strServer <> strLastServerLoggedOnTo Then
    oServer.DisConnect
    oServer.LoginSecure = True
    oServer.Connect strServer
    strLastServerLoggedOnTo = strServer
End If

'if count of children > 0 and first child is not dummy then we've populated
'so get out

If intCountOfChildren > 0 And strFirstChildName <> "DUMMY" Then
    Exit Sub
End If

'If the first child is dummy then

If strFirstChildName = "DUMMY" Then
    tv1.Nodes.Remove intFirstChildIndex
        For Each oJobStep In oServer.JobServer.Jobs(JobName).JobSteps
            If oJobStep.LastRunOutcome = SQLDMOJobOutcome_Succeeded Or oJobStep.LastRunOutcome = SQLDMOJobOutcome_Unknown Then
                tv1.Nodes.Add strJobKey, tvwChild, "JOBSTEP_" & JobName & "_" & oJobStep.Name, oJobStep.Name, il1.ListImages(5).Key
            ElseIf oJobStep.LastRunOutcome = SQLDMOJobOutcome_Failed Then
                tv1.Nodes.Add strJobKey, tvwChild, "JOBSTEP_" & JobName & "_" & oJobStep.Name, oJobStep.Name, il1.ListImages(6).Key
            End If
        Next oJobStep
End If


Exit Sub



err_handler:

Exit Sub
End Sub

Public Sub ShowJobCommand(strServer As String, strJob As String, strJobStep As String)

On Error GoTo err_handler

Dim strStepCommand As String
txtStepCommand = ""

'do not need to check if the strLastServerLoggedOnTo is set to GOOBLEDYGOOK
'as if we can see the steps we've already logged in somewhere at least
If strServer <> strLastServerLoggedOnTo Then
    oServer.DisConnect
    oServer.LoginSecure = True
    oServer.Connect strServer
    strLastServerLoggedOnTo = strServer
End If

strStepCommand = oServer.JobServer.Jobs(strJob).JobSteps(strJobStep).Command

txtStepCommand.Text = strStepCommand


Exit Sub

err_handler:
MsgBox "error"
Exit Sub

End Sub

Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)

If Left(Node.Key, 7) = "JOBSTEP" Then
    ShowJobCommand Node.Parent.Parent.Text, Node.Parent.Text, Node.Text
End If

If Left(Node.Key, 4) = "JOB_" Then
    ShowJobDetails Node.Parent.Text, Node.Text
    txtStepCommand.Text = ""
End If

End Sub


Private Sub ShowJobDetails(strServer As String, JobName As String)

Dim strCurrentRunStatus As String

If strServer <> strLastServerLoggedOnTo Then
    oServer.DisConnect
    oServer.LoginSecure = True
    oServer.Connect strServer
    strLastServerLoggedOnTo = strServer
End If

If oServer.JobServer.Jobs(JobName).Enabled = True Then
    optEnabled.Value = True
Else
    optEnabled.Value = False
End If

lblDateCreated = "Date Created:      " & oServer.JobServer.Jobs(JobName).DateCreated
lblDescription = "Description:      " & oServer.JobServer.Jobs(JobName).Description
lblLastRunDate = "Last Run Date:      " & Mid(oServer.JobServer.Jobs(JobName).LastRunDate, 7, 2) & "/" & Mid(oServer.JobServer.Jobs(JobName).LastRunDate, 5, 2) & "/" & Left(oServer.JobServer.Jobs(JobName).LastRunDate, 4)
lblLastRunTime = "Last Run Time:      " & Mid(oServer.JobServer.Jobs(JobName).LastRunTime, 1, 2) & ":" & Mid(oServer.JobServer.Jobs(JobName).LastRunTime, 3, 2) & ":" & Mid(oServer.JobServer.Jobs(JobName).LastRunTime, 5, 2)
lblOwner = "Owner:      " & oServer.JobServer.Jobs(JobName).Owner
lblCategory = "Category:      " & oServer.JobServer.Jobs(JobName).Category
lblNextRunDate = "Next Run Date:      " & Mid(oServer.JobServer.Jobs(JobName).NextRunDate, 7, 2) & "/" & Mid(oServer.JobServer.Jobs(JobName).NextRunDate, 5, 2) & "/" & Left(oServer.JobServer.Jobs(JobName).NextRunDate, 4)




Select Case oServer.JobServer.Jobs(JobName).CurrentRunStatus

Case 0
strCurrentRunStatus = "Unknown - Probably never run"
Case 1
strCurrentRunStatus = "Executing"
Case 2
strCurrentRunStatus = "Waiting for worker thread"
Case 3
strCurrentRunStatus = "Between retries"
Case 4
strCurrentRunStatus = "Idle"
Case 5
strCurrentRunStatus = "Suspended"
Case 6
strCurrentRunStatus = "Waiting for Step to Finish"
Case 7
strCurrentRunStatus = "Performaing Completion Actions"
End Select

lblCurrentRunStatus = "Current Run Status      " & strCurrentRunStatus

txtSchedule.Text = ShowJobSchedule(strServer, JobName)


End Sub

Function IsBitSet(Src As Integer, Bit As Integer) As Boolean
    IsBitSet = (Src And (2 ^ Bit)) <> 0
End Function


Public Function ShowJobSchedule(strServerclicked As String, Job As String) As String

Dim strFrequencyInterval As String
Dim strFreqTypeConstant As String
Dim strFreqInterval As String
Dim strSubType As String
Dim strCompleteScheduleString As String
Dim freqRelativeInterval As String


If oServer.JobServer.Jobs(Job).HasSchedule = False Then
    ShowJobSchedule = "No Schedule Available"
    Exit Function
End If


Select Case oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDay
Case 0
strSubType = "Unknown"
Case 1
strSubType = "Once"
Case 4
strSubType = "Minutes"
Case 8
strSubType = "Hourly"
End Select

Select Case oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyType
Case 0
strFreqTypeConstant = "Unknown"
Case 1
strFreqTypeConstant = "One Time"
strCompleteScheduleString = "This job will executute Once only on: " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartDate & " at: " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay
Case 4
strFreqTypeConstant = "Daily"

If strSubType = "Once" Then
    strCompleteScheduleString = "This job will execute every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval & " days at " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay
ElseIf strSubType = "Minutes" Then
        strCompleteScheduleString = "This job will execute every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval & " days and every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " minutes between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
ElseIf strSubType = "Hourly" Then
         strCompleteScheduleString = "This job will execute every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval & " days and every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " hours between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
End If
Case 8
strFreqTypeConstant = "Weekly"

If (oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval And 1) Then
strFrequencyInterval = strFrequencyInterval + "Sunday,"
End If

If (oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval And 2) Then
strFrequencyInterval = strFrequencyInterval + "Monday,"
End If

If (oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval And 4) Then
strFrequencyInterval = strFrequencyInterval + "Tuesday,"
End If

If (oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval And 8) Then
strFrequencyInterval = strFrequencyInterval + "Wednesday,"
End If

If (oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval And 16) Then
strFrequencyInterval = strFrequencyInterval + "Thursday,"
End If

If (oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval And 32) Then
strFrequencyInterval = strFrequencyInterval + "Friday,"
End If

If (oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval And 64) Then
strFrequencyInterval = strFrequencyInterval + "Saturday,"
End If

If strSubType = "Once" Then
    strCompleteScheduleString = "This job will execute on " & Left(strFrequencyInterval, Len(strFrequencyInterval) - 1) & " once only at " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay
ElseIf strSubType = "Minutes" Then
        strCompleteScheduleString = "This job will execute on " & Left(strFrequencyInterval, Len(strFrequencyInterval) - 1) & " and every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " minutes between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
ElseIf strSubType = "Hourly" Then
         strCompleteScheduleString = "This job will execute on " & Left(strFrequencyInterval, Len(strFrequencyInterval) - 1) & " and every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " hours between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
End If

Case 16
strFreqTypeConstant = "Monthly"

If strSubType = "Once" Then
    strCompleteScheduleString = "This job will execute on day " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval & " of the month every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months once only at " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay
ElseIf strSubType = "Minutes" Then
        strCompleteScheduleString = "This job will execute on day " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval & " of the month every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " minutes between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
ElseIf strSubType = "Hourly" Then
        strCompleteScheduleString = "This job will execute on day " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval & " of the month every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " hours between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
End If
Case 32
strFreqTypeConstant = "Monthly Relative"

If oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 1 Then
    strFrequencyInterval = "Sunday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 2 Then
    strFrequencyInterval = "Monday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 3 Then
    strFrequencyInterval = "Tuesday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 4 Then
    strFrequencyInterval = "Wednesday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 5 Then
    strFrequencyInterval = "Thursday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 6 Then
    strFrequencyInterval = "Friday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 7 Then
    strFrequencyInterval = "Saturday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 8 Then
    strFrequencyInterval = "Day"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 9 Then
    strFrequencyInterval = "Weekday"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyInterval = 10 Then
     strFrequencyInterval = "Weekendday"
End If

If oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRelativeInterval = 1 Then
    freqRelativeInterval = "First"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRelativeInterval = 2 Then
    freqRelativeInterval = "Second"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRelativeInterval = 4 Then
    freqRelativeInterval = "Third"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRelativeInterval = 8 Then
    freqRelativeInterval = "Fourth"
ElseIf oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRelativeInterval = 16 Then
    freqRelativeInterval = "Last"
End If



If strSubType = "Once" Then
    If strFrequencyInterval = "Day" Then
        strCompleteScheduleString = "This job executes on the " & freqRelativeInterval & " day of every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months once only at " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay
    Else
        strCompleteScheduleString = "This job executes on the " & freqRelativeInterval & " " & strFrequencyInterval & " of every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months once only at " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay
    End If
    
ElseIf strSubType = "Minutes" Then
    If strFrequencyInterval = "Day" Then
        strCompleteScheduleString = "This job executes on the " & freqRelativeInterval & " day of every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " minutes between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
    Else
        strCompleteScheduleString = "This job executes on the " & freqRelativeInterval & " " & strFrequencyInterval & " of every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " minutes between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
    End If
    
    
ElseIf strSubType = "Hourly" Then
    If strFrequencyInterval = "Day" Then
        strCompleteScheduleString = "This job executes on the " & freqRelativeInterval & " day of every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " hours between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
    Else
        strCompleteScheduleString = "This job executes on the " & freqRelativeInterval & " " & strFrequencyInterval & " of every " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencyRecurrenceFactor & " months every  " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.FrequencySubDayInterval & " hours between " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveStartTimeOfDay & " and " & oServer.JobServer.Jobs(Job).JobSchedules(1).Schedule.ActiveEndTimeOfDay
    End If
End If


Case 64
strFreqTypeConstant = "Autostart"
    strCompleteScheduleString = "This job will execute when the SQL Server Agent Starts Up"
Case 128
strFreqTypeConstant = "On Idle"
    strCompleteScheduleString = "This job will execute when the CPU becomes idle"
End Select


ShowJobSchedule = strCompleteScheduleString

End Function

