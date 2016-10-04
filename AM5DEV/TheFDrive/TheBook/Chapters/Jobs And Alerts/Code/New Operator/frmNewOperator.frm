VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNewOperator 
   Caption         =   "New Operator"
   ClientHeight    =   7935
   ClientLeft      =   5895
   ClientTop       =   3405
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   6480
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmNewOperator.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCheckName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtmail"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPager"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtNetSend"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkMonday"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkTuesday"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkWednesday"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkFriday"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkSaturday"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chkSunday"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboWeekdayStart"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cboWeekDayEnd"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboWeekendStart"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cboWeekendEnd"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdAdd"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkThursday"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      Begin VB.CheckBox chkThursday 
         Caption         =   "Thursday"
         Height          =   255
         Left            =   510
         TabIndex        =   26
         Tag             =   "16"
         Top             =   5190
         Width           =   2085
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   405
         Left            =   4740
         TabIndex        =   25
         Top             =   7320
         Width           =   1065
      End
      Begin VB.ComboBox cboWeekendEnd 
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
         Left            =   3720
         TabIndex        =   22
         Top             =   6630
         Width           =   1515
      End
      Begin VB.ComboBox cboWeekendStart 
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
         Left            =   3720
         TabIndex        =   21
         Top             =   6300
         Width           =   1515
      End
      Begin VB.ComboBox cboWeekDayEnd 
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
         Left            =   3780
         TabIndex        =   18
         Top             =   5100
         Width           =   1455
      End
      Begin VB.ComboBox cboWeekdayStart 
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
         Left            =   3780
         TabIndex        =   17
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CheckBox chkSunday 
         Caption         =   "Sunday"
         Height          =   165
         Left            =   480
         TabIndex        =   16
         Tag             =   "1"
         Top             =   6660
         Width           =   2355
      End
      Begin VB.CheckBox chkSaturday 
         Caption         =   "Saturday"
         Height          =   195
         Left            =   480
         TabIndex        =   15
         Tag             =   "64"
         Top             =   6360
         Width           =   2355
      End
      Begin VB.CheckBox chkFriday 
         Caption         =   "Friday"
         Height          =   195
         Left            =   510
         TabIndex        =   14
         Tag             =   "32"
         Top             =   5520
         Width           =   2355
      End
      Begin VB.CheckBox chkWednesday 
         Caption         =   "Wednesday"
         Height          =   195
         Left            =   510
         TabIndex        =   13
         Tag             =   "8"
         Top             =   4920
         Width           =   2355
      End
      Begin VB.CheckBox chkTuesday 
         Caption         =   "Tuesday"
         Height          =   195
         Left            =   510
         TabIndex        =   12
         Tag             =   "4"
         Top             =   4650
         Width           =   2355
      End
      Begin VB.CheckBox chkMonday 
         Caption         =   "Monday"
         Height          =   195
         Left            =   510
         TabIndex        =   11
         Tag             =   "2"
         Top             =   4350
         Width           =   2355
      End
      Begin VB.TextBox txtNetSend 
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
         Left            =   1350
         TabIndex        =   9
         Top             =   3240
         Width           =   3525
      End
      Begin VB.TextBox txtPager 
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
         Left            =   1350
         TabIndex        =   7
         Top             =   2490
         Width           =   3555
      End
      Begin VB.TextBox txtmail 
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
         Left            =   1350
         TabIndex        =   4
         Top             =   1740
         Width           =   3525
      End
      Begin VB.CommandButton cmdCheckName 
         Caption         =   "Check"
         Height          =   375
         Left            =   4980
         TabIndex        =   3
         Top             =   930
         Width           =   1065
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   930
         Width           =   3525
      End
      Begin VB.Label Label9 
         Caption         =   "End"
         Height          =   255
         Left            =   5370
         TabIndex        =   24
         Top             =   6690
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "Start"
         Height          =   195
         Left            =   5340
         TabIndex        =   23
         Top             =   6360
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "End"
         Height          =   255
         Left            =   5370
         TabIndex        =   20
         Top             =   5130
         Width           =   405
      End
      Begin VB.Label Label6 
         Caption         =   "Start"
         Height          =   225
         Left            =   5340
         TabIndex        =   19
         Top             =   4830
         Width           =   945
      End
      Begin VB.Label Label5 
         Caption         =   "Pager Days"
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
         Left            =   360
         TabIndex        =   10
         Top             =   3960
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Net Send"
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
         Left            =   270
         TabIndex        =   8
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Pager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   6
         Top             =   2490
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "E-Mail"
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
         Left            =   330
         TabIndex        =   5
         Top             =   1770
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   300
         TabIndex        =   2
         Top             =   1020
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmNewOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PagerDays As String

Private Sub FillComboBoxes()

Dim hr As Integer
Dim min As Integer
Dim sec As Integer
Dim i As Integer
Dim strTime As String
Dim ctl As Control


For i = 0 To 24


    For Each ctl In frmNewOperator.Controls
        If TypeOf ctl Is ComboBox Then
            If i <> 24 Then
                ctl.AddItem CStr(i) & ":" & "00" & ":" & "00"
                ctl.AddItem CStr(i) & ":" & "30" & ":" & "00"
            End If
        End If
    Next ctl


Next i
End Sub

Private Sub DisablePagerDays()
Dim ctl As Control

For Each ctl In Me.Controls
    If TypeOf ctl Is CheckBox Or TypeOf ctl Is ComboBox Then
        ctl.Enabled = False
    End If
Next ctl

End Sub

Private Sub EnablePagerDays()
Dim ctl As Control

For Each ctl In Me.Controls
    If TypeOf ctl Is CheckBox Or TypeOf ctl Is ComboBox Then
        ctl.Enabled = True
    End If
Next ctl

End Sub

Private Sub ClearUpAfterwards()

Dim ctl As Control


For Each ctl In Controls
    If TypeOf ctl Is CheckBox Then
        ctl.Value = vbUnchecked
    End If
    
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    End If

Next ctl



End Sub

Private Sub cmdAdd_Click()

On Error GoTo err_handler



If txtName.Text = "" Or CheckOperators(txtName.Text) = True Then
    Exit Sub
Else

Set oOperator = New SQLDMO.Operator

With oOperator
    .Name = txtName

    If txtmail.Text <> "" Then
        
        .EmailAddress = txtmail.Text
    
    End If
    
    If txtPager.Text <> "" Then
        .PagerAddress = txtPager.Text
        .PagerDays = CheckPagerDays
        
            If HaveWeGotPagerDaysInWeek = True Then
                .WeekdayPagerStartTime = cboWeekdayStart.Text
                .WeekdayPagerEndTime = cboWeekDayEnd.Text
            End If
            
            If HaveWeGotPagerDaysInWeekend = True Then
                .SaturdayPagerStartTime = cboWeekendStart.Text
                .SaturdayPagerEndTime = cboWeekendEnd.Text
                .SundayPagerStartTime = cboWeekendStart.Text
                .SundayPagerEndTime = cboWeekendEnd.Text
            End If
    End If

End With

End If





oServer.JobServer.Operators.Add oOperator
ClearUpAfterwards

Exit Sub

err_handler:

MsgBox Err.Description
Exit Sub

End Sub

Private Sub cmdCheckName_Click()
CheckOperators (txtName.Text)

End Sub


Private Function CheckOperators(strOperatorName As String) As Boolean

CheckOperators = False

For Each oOperator In oServer.JobServer.Operators
    If oOperator.Name = strOperatorName Then
        CheckOperators = True
        MsgBox "An Operator With That Name Already Exists", vbInformation, "Already Exists"
    End If
Next oOperator


End Function

Private Sub Form_Load()
FillComboBoxes
DisablePagerDays

End Sub

Private Function CheckPagerDays() As Integer


Dim ctl As Control
    
CheckPagerDays = 0
    
For Each ctl In frmNewOperator.Controls
    If TypeOf ctl Is CheckBox Then
        If ctl.Value = vbChecked Then
        MsgBox ctl.Tag
            CheckPagerDays = CheckPagerDays + ctl.Tag
        MsgBox CheckPagerDays
        End If
    End If
Next ctl


    
End Function

Private Function HaveWeGotPagerDaysInWeek() As Boolean

HaveWeGotPagerDaysInWeek = False

If chkMonday.Value = vbChecked Or chkTuesday.Value = vbChecked Or chkWednesday.Value = vbChecked Or chkThursday.Value = vbChecked Or chkFriday.Value = vbChecked Then
    HaveWeGotPagerDaysInWeek = True
End If

End Function

Private Function HaveWeGotPagerDaysInWeekend() As Boolean

HaveWeGotPagerDaysInWeekend = False

If chkSaturday.Value = vbChecked Or chkSunday.Value Then
    HaveWeGotPagerDaysInWeekend = True
End If

End Function


Private Sub txtPager_Change()
If Len(txtPager.Text) > 0 Then
    EnablePagerDays
Else
    DisablePagerDays
End If

End Sub


