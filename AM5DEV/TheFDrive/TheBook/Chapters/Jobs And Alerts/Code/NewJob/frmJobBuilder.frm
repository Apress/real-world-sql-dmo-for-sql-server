VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJobBuilder 
   Caption         =   "JobBuilder"
   ClientHeight    =   6780
   ClientLeft      =   3495
   ClientTop       =   4455
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   13080
   Begin TabDlg.SSTab STAB1 
      Height          =   6825
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12039
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmJobBuilder.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdExit"
      Tab(0).Control(1)=   "cmdAddJob"
      Tab(0).Control(2)=   "txtDescription"
      Tab(0).Control(3)=   "cboOwner"
      Tab(0).Control(4)=   "cboCategory"
      Tab(0).Control(5)=   "txtJobName"
      Tab(0).Control(6)=   "chkEnabled"
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(8)=   "lblDateCreated"
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(11)=   "Label1"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Job"
      TabPicture(1)   =   "frmJobBuilder.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "msf_JobStep"
      Tab(1).Control(1)=   "cmdAddStep"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Schedule"
      TabPicture(2)   =   "frmJobBuilder.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblEvery"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblInterval"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Line2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblOnceADay"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblOccursAt"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblStart"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblEnd"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblAt"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lblOn"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtScheduleName"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtWeekDayInterval"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtOnceADay"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtRecurringEvery"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "lstMinuteHour"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtStartsAt"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtEndsAt"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "chkSunday"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "chkMonday"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "chkTuesday"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "chkWednesday"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "chkThursday"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "chkFriday"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "chkSaturday"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "cmdAddSchedule"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "optOnceADay"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Frame1"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtAt"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "dtOn"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "optRecurring"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).ControlCount=   30
      TabCaption(3)   =   "Notifications"
      TabPicture(3)   =   "frmJobBuilder.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdNotification"
      Tab(3).Control(1)=   "cboNetSendLevel"
      Tab(3).Control(2)=   "cboPagelevel"
      Tab(3).Control(3)=   "cboMailLevel"
      Tab(3).Control(4)=   "cboNetSendOperator"
      Tab(3).Control(5)=   "cboPageOperator"
      Tab(3).Control(6)=   "cboEmailOperator"
      Tab(3).Control(7)=   "chkNetSend"
      Tab(3).Control(8)=   "chkPage"
      Tab(3).Control(9)=   "chkEmail"
      Tab(3).Control(10)=   "Label6"
      Tab(3).Control(11)=   "Label5"
      Tab(3).ControlCount=   12
      Begin VB.OptionButton optRecurring 
         Height          =   225
         Left            =   3240
         TabIndex        =   60
         Top             =   4980
         Width           =   345
      End
      Begin MSComCtl2.DTPicker dtOn 
         Height          =   435
         Left            =   6510
         TabIndex        =   59
         Top             =   1890
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   22740993
         CurrentDate     =   37291
      End
      Begin VB.TextBox txtAt 
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
         Left            =   9150
         TabIndex        =   56
         Top             =   1890
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         Caption         =   "Intervals"
         Height          =   3555
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Width           =   2715
         Begin VB.OptionButton optAutoRun 
            Caption         =   "Run when SQL Server starts"
            Height          =   225
            Left            =   180
            TabIndex        =   55
            Top             =   330
            Width           =   2355
         End
         Begin VB.OptionButton optCPUIdle 
            Caption         =   "When CPU Becomes Idle"
            Height          =   195
            Left            =   180
            TabIndex        =   54
            Top             =   990
            Width           =   2145
         End
         Begin VB.OptionButton optOneTimeOnly 
            Caption         =   "One Time Only"
            Height          =   225
            Left            =   180
            TabIndex        =   53
            Top             =   1680
            Width           =   1425
         End
         Begin VB.OptionButton optDaily 
            Caption         =   "Daily"
            Height          =   225
            Left            =   180
            TabIndex        =   52
            Top             =   2340
            Width           =   1065
         End
         Begin VB.OptionButton optWeekly 
            Caption         =   "Weekly"
            Height          =   285
            Left            =   180
            TabIndex        =   51
            Top             =   3000
            Width           =   1095
         End
      End
      Begin VB.OptionButton optOnceADay 
         Height          =   255
         Left            =   3240
         TabIndex        =   49
         Top             =   4140
         Width           =   345
      End
      Begin VB.CommandButton cmdAddSchedule 
         Caption         =   "Add Schedule"
         Height          =   615
         Left            =   10440
         TabIndex        =   48
         Top             =   5970
         Width           =   1635
      End
      Begin VB.CheckBox chkSaturday 
         Caption         =   "Saturday"
         Height          =   255
         Left            =   10140
         TabIndex        =   47
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CheckBox chkFriday 
         Caption         =   "Friday"
         Height          =   255
         Left            =   9080
         TabIndex        =   46
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CheckBox chkThursday 
         Caption         =   "Thursday"
         Height          =   255
         Left            =   8020
         TabIndex        =   45
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CheckBox chkWednesday 
         Caption         =   "Wednesday"
         Height          =   255
         Left            =   6750
         TabIndex        =   44
         Top             =   2970
         Width           =   1215
      End
      Begin VB.CheckBox chkTuesday 
         Caption         =   "Tuesday"
         Height          =   255
         Left            =   5690
         TabIndex        =   43
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CheckBox chkMonday 
         Caption         =   "Monday"
         Height          =   255
         Left            =   4630
         TabIndex        =   42
         Top             =   2970
         Width           =   1005
      End
      Begin VB.CheckBox chkSunday 
         Caption         =   "Sunday"
         Height          =   255
         Left            =   3570
         TabIndex        =   41
         Top             =   2970
         Width           =   1005
      End
      Begin VB.TextBox txtEndsAt 
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
         Left            =   9450
         TabIndex        =   38
         Top             =   5190
         Width           =   1185
      End
      Begin VB.TextBox txtStartsAt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9450
         TabIndex        =   37
         Top             =   4650
         Width           =   1185
      End
      Begin VB.ListBox lstMinuteHour 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmJobBuilder.frx":0070
         Left            =   6630
         List            =   "frmJobBuilder.frx":007A
         TabIndex        =   36
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox txtRecurringEvery 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5730
         TabIndex        =   35
         Top             =   4860
         Width           =   675
      End
      Begin VB.TextBox txtOnceADay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5820
         TabIndex        =   33
         Top             =   4050
         Width           =   945
      End
      Begin VB.TextBox txtWeekDayInterval 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5700
         TabIndex        =   29
         Top             =   2100
         Width           =   645
      End
      Begin VB.TextBox txtScheduleName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4350
         TabIndex        =   27
         Top             =   1200
         Width           =   6015
      End
      Begin VB.CommandButton cmdNotification 
         Caption         =   "Set Notification"
         Height          =   585
         Left            =   -64230
         TabIndex        =   26
         Top             =   5940
         Width           =   1665
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   465
         Left            =   -64080
         TabIndex        =   25
         Top             =   6030
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid msf_JobStep 
         Height          =   3795
         Left            =   -73020
         TabIndex        =   24
         Top             =   1200
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   6694
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         AllowUserResizing=   3
      End
      Begin VB.CommandButton cmdAddJob 
         Caption         =   "Add Job"
         Height          =   465
         Left            =   -64080
         TabIndex        =   23
         Top             =   5430
         Width           =   1755
      End
      Begin VB.ComboBox cboNetSendLevel 
         Height          =   315
         Left            =   -66870
         TabIndex        =   22
         Top             =   2970
         Width           =   2925
      End
      Begin VB.ComboBox cboPagelevel 
         Height          =   315
         Left            =   -66870
         TabIndex        =   21
         Top             =   2460
         Width           =   2925
      End
      Begin VB.ComboBox cboMailLevel 
         Height          =   315
         Left            =   -66870
         TabIndex        =   20
         Top             =   1890
         Width           =   2955
      End
      Begin VB.ComboBox cboNetSendOperator 
         Height          =   315
         Left            =   -72030
         TabIndex        =   17
         Top             =   3000
         Width           =   3105
      End
      Begin VB.ComboBox cboPageOperator 
         Height          =   315
         Left            =   -72030
         TabIndex        =   16
         Top             =   2490
         Width           =   3135
      End
      Begin VB.ComboBox cboEmailOperator 
         Height          =   315
         Left            =   -72030
         TabIndex        =   15
         Top             =   1860
         Width           =   3105
      End
      Begin VB.CheckBox chkNetSend 
         Caption         =   "Net Send"
         Height          =   345
         Left            =   -73950
         TabIndex        =   14
         Top             =   3030
         Width           =   1095
      End
      Begin VB.CheckBox chkPage 
         Caption         =   "Page"
         Height          =   255
         Left            =   -73920
         TabIndex        =   13
         Top             =   2460
         Width           =   765
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   -73920
         TabIndex        =   12
         Top             =   1860
         Width           =   735
      End
      Begin VB.CommandButton cmdAddStep 
         Caption         =   "Add Step"
         Height          =   525
         Left            =   -64620
         TabIndex        =   11
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2025
         Left            =   -72240
         TabIndex        =   9
         Top             =   4020
         Width           =   5505
      End
      Begin VB.ComboBox cboOwner 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         TabIndex        =   5
         Top             =   3240
         Width           =   4695
      End
      Begin VB.ComboBox cboCategory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72270
         TabIndex        =   4
         Top             =   2610
         Width           =   4725
      End
      Begin VB.TextBox txtJobName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72240
         TabIndex        =   2
         Top             =   1020
         Width           =   5565
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   -65760
         TabIndex        =   1
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label lblOn 
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   58
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label lblAt 
         Caption         =   "At"
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
         Left            =   8640
         TabIndex        =   57
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblEnd 
         Caption         =   "Ends"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8850
         TabIndex        =   40
         Top             =   5280
         Width           =   435
      End
      Begin VB.Label lblStart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8850
         TabIndex        =   39
         Top             =   4740
         Width           =   435
      End
      Begin VB.Label lblOccursAt 
         Caption         =   "Occurs Every"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3750
         TabIndex        =   34
         Top             =   4950
         Width           =   1725
      End
      Begin VB.Label lblOnceADay 
         Caption         =   "Once a Day at"
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
         Left            =   3840
         TabIndex        =   32
         Top             =   4050
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   3840
         X2              =   9990
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Label lblInterval 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6540
         TabIndex        =   31
         Top             =   2100
         Width           =   1725
      End
      Begin VB.Label lblEvery 
         Caption         =   "Every"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4770
         TabIndex        =   30
         Top             =   2250
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Name:"
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
         Left            =   3660
         TabIndex        =   28
         Top             =   1260
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "On Action"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -66510
         TabIndex        =   19
         Top             =   1140
         Width           =   3225
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Operator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -71970
         TabIndex        =   18
         Top             =   1140
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Description"
         Height          =   285
         Left            =   -73320
         TabIndex        =   10
         Top             =   4050
         Width           =   1065
      End
      Begin VB.Label lblDateCreated 
         Caption         =   "Date Created"
         Height          =   195
         Left            =   -72270
         TabIndex        =   8
         Top             =   1830
         Width           =   4965
      End
      Begin VB.Label Label3 
         Caption         =   "Owner"
         Height          =   255
         Left            =   -73320
         TabIndex        =   7
         Top             =   3270
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Category"
         Height          =   225
         Left            =   -73290
         TabIndex        =   6
         Top             =   2610
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Job Name:"
         Height          =   285
         Left            =   -73320
         TabIndex        =   3
         Top             =   1050
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmJobBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEmail_Click()
ShallWeMail
End Sub

Private Sub chkNetSend_Click()
ShallWeNetSend
End Sub

Private Sub chkPage_Click()
ShallWePage
End Sub

Private Function JobExists(JobName) As Boolean

JobExists = False

For Each ojob In oServer.JobServer.Jobs
    If ojob.Name = JobName Then
        JobExists = True
    End If
Next ojob

End Function

Private Sub AnyNotifications()

Dim oOutcomeLevel As SQLDMO.SQLDMO_COMPLETION_TYPE


'Email
If chkEmail.Value = vbChecked And cboMailLevel.Text <> "" And cboEmailOperator.Text <> "" Then

Select Case cboMailLevel.Text

    Case "On Success"
    oOutcomeLevel = 1
    Case "On Failure"
    oOutcomeLevel = 2
    Case "On Completion"
    oOutcomeLevel = 3
End Select

    oServer.JobServer.Jobs(strJobName).BeginAlter
        oServer.JobServer.Jobs(strJobName).EmailLevel = oOutcomeLevel
        oServer.JobServer.Jobs(strJobName).OperatorToEmail = "ALLAN"
    oServer.JobServer.Jobs(strJobName).DoAlter
    
End If


'Page
If chkPage.Value = vbChecked And cboPagelevel.Text <> "" And cboPageOperator.Text <> "" Then
    

    Select Case cboPagelevel.Text

    Case "On Success"
    oOutcomeLevel = 1
    Case "On Failure"
    oOutcomeLevel = 2
    Case "On Completion"
    oOutcomeLevel = 3
    End Select

    oServer.JobServer.Jobs(strJobName).BeginAlter
        oServer.JobServer.Jobs(strJobName).OperatorToPage = cboPageOperator.Text
        oServer.JobServer.Jobs(strJobName).PageLevel = oOutcomeLevel
    oServer.JobServer.Jobs(strJobName).DoAlter
    
End If



'Net Send
If chkNetSend.Value = vbChecked And cboNetSendLevel.Text <> "" And cboNetSendOperator.Text <> "" Then
   
    
    Select Case cboNetSendLevel.Text
    
    Case "On Success"
    oOutcomeLevel = 1
    Case "On Failure"
    oOutcomeLevel = 2
    Case "On Completion"
    oOutcomeLevel = 3
    End Select
    
    oServer.JobServer.Jobs(strJobName).BeginAlter
        oServer.JobServer.Jobs(strJobName).OperatorToNetSend = cboNetSendOperator.Text
        oServer.JobServer.Jobs(strJobName).NetSendLevel = oOutcomeLevel
    oServer.JobServer.Jobs(strJobName).DoAlter
 
End If
End Sub


Private Sub cmdAddJob_Click()

If txtJobName.Text <> "" And JobExists(txtJobName.Text) = False Then
    Set ojob = New SQLDMO.Job
    ojob.Name = txtJobName.Text
    ojob.Category = cboCategory.Text
    ojob.Owner = cboOwner.Text
    ojob.Description = txtDescription.Text
    
    'ojob.EmailLevel = SQLDMOComp_Success
    'ojob.OperatorToEmail = "ALLAN"
    
    If chkEnabled.Value = vbChecked Then
        ojob.Enabled = True
    Else
        ojob.Enabled = False
    End If
    
    
    oServer.JobServer.Jobs.Add ojob
    strJobName = ojob.Name
End If
    
    
End Sub

Private Sub cmdAddSchedule_Click()

If optAutoRun.Value = True Then
    SQLServerStartSchedule txtScheduleName.Text
ElseIf optCPUIdle.Value = True Then
    CPUIdleSchedule txtScheduleName.Text
ElseIf optOneTimeOnly.Value = True Then
    OneTimeOnlySchedule txtScheduleName.Text, dtOn.Year & IIf(Len(dtOn.Month) <> 2, "0" & dtOn.Month, dtOn.Month) & IIf(Len(dtOn.Day) <> 2, "0" & dtOn.Day, dtOn.Day), txtAt.Text
ElseIf optDaily.Value = True Then
    If optOnceADay.Value = True Then
        DailySchedule txtScheduleName.Text, txtWeekDayInterval.Text, True, txtOnceADay.Text
    Else
        If lstMinuteHour.Text = "Minute" Then
            DailySchedule txtScheduleName.Text, txtWeekDayInterval.Text, False, , txtRecurringEvery.Text, 1, txtStartsAt.Text, txtEndsAt.Text
        Else
            DailySchedule txtScheduleName.Text, txtWeekDayInterval.Text, False, , txtRecurringEvery.Text, 2, txtStartsAt.Text, txtEndsAt.Text
        End If
    End If

ElseIf optWeekly.Value = True Then

    If optOnceADay.Value = True Then
        WeeklySchedule txtScheduleName.Text, txtWeekDayInterval.Text, True, txtOnceADay.Text, , , , , DaysOfWeekToRun
    Else
        If lstMinuteHour.Text <> "Hour" Then
            WeeklySchedule txtScheduleName.Text, txtWeekDayInterval.Text, False, , txtRecurringEvery.Text, 1, txtStartsAt.Text, txtEndsAt.Text, DaysOfWeekToRun
        Else
            WeeklySchedule txtScheduleName.Text, txtWeekDayInterval.Text, False, , txtRecurringEvery.Text, 2, txtStartsAt.Text, txtEndsAt.Text, DaysOfWeekToRun
        End If
    End If


End If



End Sub

Private Sub cmdAddStep_Click()
FrmJobStep.Visible = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdNotification_Click()
AnyNotifications
End Sub

Private Sub Form_Load()
lblDateCreated.Caption = "Date Created:  " & Now()
LoadCategories
LoadOwners
LoadOperators
loadActions
ShallWeMail
ShallWeNetSend
ShallWePage
msf_JobStep.Rows = 0
HideAllScheduleControls
optAutoRun.Value = True

cboCategory.ListIndex = 1




End Sub

Private Sub ShallWeMail()
If chkEmail.Value = vbChecked Then
    cboEmailOperator.Enabled = True
    cboMailLevel.Enabled = True
Else
    cboEmailOperator.Enabled = False
    cboMailLevel.Enabled = False
End If
End Sub

Private Sub ShallWeNetSend()
If chkNetSend.Value = vbChecked Then
    cboNetSendOperator.Enabled = True
    cboNetSendLevel.Enabled = True
Else
    cboNetSendOperator.Enabled = False
    cboNetSendLevel.Enabled = False
End If
End Sub

Private Sub ShallWePage()
If chkPage.Value = vbChecked Then
    cboPageOperator.Enabled = True
    cboPagelevel.Enabled = True
Else
    cboPageOperator.Enabled = False
    cboPagelevel.Enabled = False
End If
End Sub

Private Sub loadActions()

cboMailLevel.AddItem "On Success"
cboMailLevel.AddItem "On Failure"
cboMailLevel.AddItem "On Completion"

cboPagelevel.AddItem "On Success"
cboPagelevel.AddItem "On Failure"
cboPagelevel.AddItem "On Completion"

cboNetSendLevel.AddItem "On Success"
cboNetSendLevel.AddItem "On Failure"
cboNetSendLevel.AddItem "On Completion"




End Sub


Private Sub LoadOperators()
Dim ooperator As New SQLDMO.Operator
For Each ooperator In oServer.JobServer.Operators
    cboPageOperator.AddItem ooperator.Name
    cboNetSendOperator.AddItem ooperator.Name
    cboEmailOperator.AddItem ooperator.Name
Next ooperator


End Sub

Private Sub LoadOwners()
Dim ologin As New SQLDMO.Login

For Each ologin In oServer.Logins
    cboOwner.AddItem ologin.Name
Next ologin

End Sub
Private Sub LoadCategories()

Dim ocategory As New SQLDMO.Category

For Each ocategory In oServer.JobServer.JobCategories
    cboCategory.AddItem ocategory.Name
Next ocategory
End Sub


Private Sub HideAllScheduleControls()
lblOn.Visible = False
dtOn.Visible = False
lblAt.Visible = False
txtAt.Visible = False
lblEvery.Visible = False
txtWeekDayInterval.Visible = False
lblInterval.Visible = False
chkSunday.Visible = False
chkMonday.Visible = False
chkTuesday.Visible = False
chkWednesday.Visible = False
chkThursday.Visible = False
chkFriday.Visible = False
chkSaturday.Visible = False
optOnceADay.Visible = False
optRecurring.Visible = False
lblOnceADay.Visible = False
txtOnceADay.Visible = False
lblOccursAt.Visible = False
txtRecurringEvery.Visible = False
lstMinuteHour.Visible = False
lblStart.Visible = False
lblEnd.Visible = False
txtStartsAt.Visible = False
txtEndsAt.Visible = False



End Sub


Private Sub HideAllScheduleControlsExceptWeekly()
lblOn.Visible = False
dtOn.Visible = False
lblAt.Visible = False
txtAt.Visible = False
lblEvery.Visible = True
txtWeekDayInterval.Visible = True
lblInterval.Visible = True
lblInterval.Caption = "Weeks"
chkSunday.Visible = True
chkMonday.Visible = True
chkTuesday.Visible = True
chkWednesday.Visible = True
chkThursday.Visible = True
chkFriday.Visible = True
chkSaturday.Visible = True
optOnceADay.Visible = True
optRecurring.Visible = True
lblOnceADay.Visible = True
txtOnceADay.Visible = True
lblOccursAt.Visible = True
txtRecurringEvery.Visible = True
lstMinuteHour.Visible = True
lblStart.Visible = True
lblEnd.Visible = True
txtStartsAt.Visible = True
txtEndsAt.Visible = True



End Sub




Private Sub HideAllScheduleControlsExceptDaily()
lblOn.Visible = False
dtOn.Visible = False
lblAt.Visible = False
txtAt.Visible = False
lblEvery.Visible = True
txtWeekDayInterval.Visible = True
lblInterval.Visible = True
lblInterval.Caption = "Days"
chkSunday.Visible = False
chkMonday.Visible = False
chkTuesday.Visible = False
chkWednesday.Visible = False
chkThursday.Visible = False
chkFriday.Visible = False
chkSaturday.Visible = False
optOnceADay.Visible = True
optRecurring.Visible = True
lblOnceADay.Visible = True
txtOnceADay.Visible = True
lblOccursAt.Visible = True
txtRecurringEvery.Visible = True
lstMinuteHour.Visible = True
lblStart.Visible = True
lblEnd.Visible = True
txtStartsAt.Visible = True
txtEndsAt.Visible = True



End Sub



Private Sub HideAllScheduleControlsExceptOnceOnly()
lblOn.Visible = True
dtOn.Visible = True
lblAt.Visible = True
txtAt.Visible = True
lblEvery.Visible = False
txtWeekDayInterval.Visible = False
lblInterval.Visible = False
chkSunday.Visible = False
chkMonday.Visible = False
chkTuesday.Visible = False
chkWednesday.Visible = False
chkThursday.Visible = False
chkFriday.Visible = False
chkSaturday.Visible = False
optOnceADay.Visible = False
optRecurring.Visible = False
lblOnceADay.Visible = False
txtOnceADay.Visible = False
lblOccursAt.Visible = False
txtRecurringEvery.Visible = False
lstMinuteHour.Visible = False
lblStart.Visible = False
lblEnd.Visible = False
txtStartsAt.Visible = False
txtEndsAt.Visible = False



End Sub


Private Sub optAutoRun_Click()
HideAllScheduleControls
End Sub

Private Sub optCPUIdle_Click()
HideAllScheduleControls
End Sub

Private Sub optDaily_Click()
HideAllScheduleControlsExceptDaily
End Sub

Private Sub optOneTimeOnly_Click()
HideAllScheduleControlsExceptOnceOnly
End Sub

Private Sub optWeekly_Click()
HideAllScheduleControlsExceptWeekly
End Sub


Private Sub SQLServerStartSchedule(strschedulename As String)


oJobSchedule.Schedule.FrequencyType = SQLDMOFreq_Autostart
oJobSchedule.Name = strschedulename


With oServer.JobServer.Jobs(strJobName)
    .BeginAlter
    .JobSchedules.Add oJobSchedule
End With

    
End Sub

Private Sub CPUIdleSchedule(strschedulename As String)
oJobSchedule.Schedule.FrequencyType = SQLDMOFreq_OnIdle
oJobSchedule.Name = strschedulename


With oServer.JobServer.Jobs(strJobName)
    .BeginAlter
    .JobSchedules.Add oJobSchedule
End With
End Sub

Private Sub OneTimeOnlySchedule(strschedulename As String, strWhen As String, strAtTime As String)
oJobSchedule.Schedule.FrequencyType = SQLDMOFreq_OneTime
oJobSchedule.Schedule.ActiveStartDate = strWhen
oJobSchedule.Schedule.ActiveStartTimeOfDay = strAtTime
oJobSchedule.Name = strschedulename

With oServer.JobServer.Jobs(strJobName)
    .BeginAlter
    .JobSchedules.Add oJobSchedule
End With
End Sub


Private Function DaysOfWeekToRun() As Integer

DaysOfWeekToRun = 0

If chkSunday.Value = vbChecked Then
    DaysOfWeekToRun = DaysOfWeekToRun + 1
End If

If chkMonday.Value = vbChecked Then
    DaysOfWeekToRun = DaysOfWeekToRun + 2
End If

If chkTuesday.Value = vbChecked Then
    DaysOfWeekToRun = DaysOfWeekToRun + 4
End If

If chkWednesday.Value = vbChecked Then
    DaysOfWeekToRun = DaysOfWeekToRun + 8
End If

If chkThursday.Value = vbChecked Then
    DaysOfWeekToRun = DaysOfWeekToRun + 16
End If

If chkFriday.Value = vbChecked Then
    DaysOfWeekToRun = DaysOfWeekToRun + 32
End If

If chkSaturday.Value = vbChecked Then
    DaysOfWeekToRun = DaysOfWeekToRun + 64
End If




End Function

Private Sub DailySchedule(strschedulename As String, intDailyInterval As Integer, booOnce As Boolean, Optional strOnceOnlyTime As String, Optional intRecurrenceInterval As Integer, Optional intMinHour As Integer, Optional strStartTime As String, Optional strEndTime As String)
oJobSchedule.Schedule.FrequencyType = SQLDMOFreq_Daily
oJobSchedule.Schedule.FrequencyInterval = intDailyInterval
oJobSchedule.Name = strschedulename


If booOnce = True Then
    oJobSchedule.Schedule.ActiveStartTimeOfDay = strOnceOnlyTime
Else
    
    oJobSchedule.Schedule.FrequencySubDayInterval = intRecurrenceInterval
    
    If intMinHour = 1 Then 'Minute
        oJobSchedule.Schedule.FrequencySubDay = 4
    Else 'hours
        oJobSchedule.Schedule.FrequencySubDay = 8
    End If
    
    oJobSchedule.Schedule.ActiveStartTimeOfDay = strStartTime
    oJobSchedule.Schedule.ActiveEndTimeOfDay = strEndTime

End If




    oServer.JobServer.Jobs(strJobName).BeginAlter
    oServer.JobServer.Jobs(strJobName).JobSchedules.Add oJobSchedule
    oServer.JobServer.Jobs(strJobName).DoAlter


End Sub


Private Sub WeeklySchedule(strschedulename As String, intWeeklyInterval As Integer, booOnce As Boolean, Optional strOnceOnlyTime As String, Optional intRecurrenceInterval As Integer, Optional intMinHour As Integer, Optional strStartTime As String, Optional strEndTime As String, Optional intOnDays As Integer)
oJobSchedule.Schedule.FrequencyType = SQLDMOFreq_Weekly
oJobSchedule.Schedule.FrequencyInterval = intOnDays
oJobSchedule.Name = strschedulename
oJobSchedule.Schedule.FrequencyRecurrenceFactor = intWeeklyInterval

If booOnce = True Then
    oJobSchedule.Schedule.ActiveStartTimeOfDay = strOnceOnlyTime
Else
    
    
    
    If intMinHour = 1 Then 'Minute
        oJobSchedule.Schedule.FrequencySubDay = 4
    Else 'hours
        oJobSchedule.Schedule.FrequencySubDay = 8
    End If
    
    oJobSchedule.Schedule.ActiveStartTimeOfDay = strStartTime
    oJobSchedule.Schedule.ActiveEndTimeOfDay = strEndTime
    oJobSchedule.Schedule.FrequencySubDayInterval = intRecurrenceInterval
    

End If




    oServer.JobServer.Jobs(strJobName).BeginAlter
    oServer.JobServer.Jobs(strJobName).JobSchedules.Add oJobSchedule
    oServer.JobServer.Jobs(strJobName).DoAlter


End Sub
