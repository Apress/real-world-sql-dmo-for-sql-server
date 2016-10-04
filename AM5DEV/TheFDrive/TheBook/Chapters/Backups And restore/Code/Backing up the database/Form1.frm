VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Backup"
   ClientHeight    =   9495
   ClientLeft      =   5190
   ClientTop       =   2430
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   7770
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "Form1.frx":0000
      Top             =   330
      Width           =   645
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove"
      Height          =   405
      Left            =   3450
      TabIndex        =   18
      Top             =   4530
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   405
      Left            =   3450
      TabIndex        =   17
      Top             =   3990
      Width           =   1695
   End
   Begin VB.ComboBox cboRetain 
      Height          =   315
      ItemData        =   "Form1.frx":000A
      Left            =   420
      List            =   "Form1.frx":002C
      TabIndex        =   15
      Top             =   7200
      Width           =   525
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   3570
      TabIndex        =   14
      Top             =   7740
      Width           =   3465
   End
   Begin VB.OptionButton OptLog 
      Caption         =   "LOG"
      Height          =   285
      Left            =   3630
      TabIndex        =   13
      Top             =   6570
      Width           =   675
   End
   Begin VB.OptionButton OptDiff 
      Caption         =   "DIFFERENTIAL"
      Height          =   345
      Left            =   3630
      TabIndex        =   12
      Top             =   6000
      Width           =   1455
   End
   Begin VB.OptionButton optFull 
      Caption         =   "FULL"
      Height          =   285
      Left            =   3630
      TabIndex        =   11
      Top             =   5520
      Width           =   825
   End
   Begin VB.ComboBox cboDatabases 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   2220
      Width           =   4755
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   3060
      Width           =   4755
   End
   Begin VB.Frame frINIT 
      Caption         =   "Backup Types"
      Height          =   1065
      Left            =   360
      TabIndex        =   2
      Top             =   5580
      Width           =   2535
      Begin VB.OptionButton OptOverwrite 
         Caption         =   "Overwrite"
         Height          =   225
         Left            =   270
         TabIndex        =   10
         Top             =   690
         Width           =   1485
      End
      Begin VB.OptionButton optAppend 
         Caption         =   "Append"
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   300
         Width           =   1545
      End
   End
   Begin VB.ListBox lstdevices 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   360
      TabIndex        =   1
      Top             =   3990
      Width           =   2895
   End
   Begin VB.ComboBox cboServers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   1050
      Width           =   4755
   End
   Begin VB.Label Label5 
      Caption         =   "Retain Days"
      Height          =   195
      Left            =   390
      TabIndex        =   16
      Top             =   6990
      Width           =   1425
   End
   Begin VB.Label Label4 
      Caption         =   "Name for the backup"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2820
      Width           =   1905
   End
   Begin VB.Label Label3 
      Caption         =   "Devices"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "Databases"
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   1980
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Servers"
      Height          =   315
      Left            =   360
      TabIndex        =   5
      Top             =   810
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oServer As SQLDMO.SQLServer
Public oDevice As SQLDMO.BackupDevice
Public oDatabase As SQLDMO.Database
Public oBackup As SQLDMO.Backup
Public oRegisteredServer As SQLDMO.RegisteredServer
Public oApp As SQLDMO.Application
Public oGroup As SQLDMO.ServerGroup
Public oRestore As SQLDMO.Restore


Private Sub ShowServers()

Set oApp = New SQLDMO.Application
Set oRegisteredServer = New SQLDMO.RegisteredServer

cboServers.Clear

For Each oGroup In oApp.ServerGroups

For Each oRegisteredServer In oApp.ServerGroups(oGroup.Name).RegisteredServers
    cboServers.AddItem oRegisteredServer.Name
Next oRegisteredServer

Next oGroup

Set oGroup = Nothing
Set oRegisteredServer = Nothing
Set oApp = Nothing

End Sub

Private Sub ShowDatabases(servername As String)

Set oServer = New SQLDMO.SQLServer

oServer.LoginSecure = True
oServer.Connect servername

cboDatabases.Clear

For Each oDatabase In oServer.Databases
    cboDatabases.AddItem oDatabase.Name
Next oDatabase


oServer.DisConnect
Set oServer = Nothing
Set oDatabase = Nothing


End Sub

Private Sub ShowBackupDevices(servername As String)

Set oServer = New SQLDMO.SQLServer

oServer.LoginSecure = True
oServer.Connect servername

lstdevices.Clear

For Each oDevice In oServer.BackupDevices
    lstdevices.AddItem oDevice.Name
Next oDevice

oServer.DisConnect
Set oServer = Nothing
Set oDevice = Nothing

End Sub

Private Sub BackupTheDatabase(servername As String, DatabaseName As String, Location As String, DeviceYN As Integer, BackupType As SQLDMO_BACKUP_TYPE, RetainDays As Integer, InitFirst As Boolean, BackupName As String)
On Error GoTo Err_handler

Set oServer = New SQLDMO.SQLServer
Set oBackup = New SQLDMO.Backup
Set oRestore = New SQLDMO.Restore

oServer.LoginSecure = True
oServer.Connect servername

oBackup.Action = BackupType
oBackup.BackupSetName = BackupName
oBackup.Database = DatabaseName
oBackup.RetainDays = RetainDays
oBackup.Initialize = InitFirst




If DeviceYN = 1 Then
    oBackup.Devices = Location
    oRestore.Devices = Location
Else
    oBackup.Files = Location
    oRestore.Files = Location
End If



oBackup.SQLBackup oServer
oRestore.SQLVerify oServer

MsgBox "The database Backup Succeeded", vbOKOnly, "Backup Completed"

oServer.DisConnect
Set oServer = Nothing
Exit Sub


Err_handler:

If Err.Number = -2147218262 Then
    MsgBox "The backup you have just performed would appear to be corrupted", vbCritical, "Backup problem"
Else
    MsgBox "The database Backup Failed" & vbCrLf & Err.Description, vbOKOnly, "Backup Completed"
End If

oServer.DisConnect
Set oServer = Nothing
Exit Sub

End Sub

Private Sub cboServers_Click()
ShowDatabases cboServers.Text
ShowBackupDevices cboServers.Text
End Sub

Private Sub cmdAdd_Click()
Dim strNewLocation As String

strNewLocation = InputBox("Enter A New Location", "New Backup Location")

If strNewLocation <> "" Then
    lstdevices.AddItem strNewLocation
End If


End Sub

Private Sub cmdCreate_Click()

Dim servername As String
Dim BackupType As SQLDMO_BACKUP_TYPE
Dim Init As Boolean
Dim DeviceCounter As Integer
Dim BackupName As String




Init = False

If optFull.Value = True Then
    BackupType = SQLDMOBackup_Database
ElseIf OptDiff.Value = True Then
    BackupType = SQLDMOBackup_Differential
ElseIf OptLog.Value = True Then
    BackupType = SQLDMOBackup_Log
End If

If OptOverwrite.Value = True Then
    Init = True
End If
    

If (cboServers.Text <> "" Or cboDatabases.Text <> "" Or lstdevices.ListIndex <> -1) Then

        Set oServer = New SQLDMO.SQLServer
Set oBackup = New SQLDMO.Backup


oServer.LoginSecure = True
oServer.Connect cboServers.Text


        For Each oDevice In oServer.BackupDevices
            If oDevice.Name = lstdevices.Text Then
                DeviceCounter = 1
            End If
        Next oDevice
        
        If txtName = "" Then
            BackupName = cboDatabases.Text & "_" & Format(Now(), "yyyymmdd") & Format(Now(), "hhmmss")
        Else
            BackupName = txtName.Text
        End If

        
BackupTheDatabase cboServers.Text, cboDatabases.Text, lstdevices.Text, DeviceCounter, BackupType, cboRetain.Text, Init, BackupName



End If


End Sub

Private Sub cmdremove_Click()
If lstdevices.ListIndex <> -1 Then
    lstdevices.RemoveItem lstdevices.ListIndex
End If
End Sub

Private Sub Form_Load()
cboRetain.ListIndex = 0
ShowServers
optAppend.Value = True
optFull.Value = True
End Sub
