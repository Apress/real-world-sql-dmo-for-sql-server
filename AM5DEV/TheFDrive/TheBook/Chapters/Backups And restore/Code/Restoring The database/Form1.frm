VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Restore"
   ClientHeight    =   9420
   ClientLeft      =   5610
   ClientTop       =   2715
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   10740
   Begin TabDlg.SSTab SSTab1 
      Height          =   9465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   16695
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "msf_Available"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboServers"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboDatabases"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lstdevices"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdChooseDeviceFile"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdRDDL(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdRDDL(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkForce"
      Tab(1).Control(1)=   "msf_FileList"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "Label6"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "DDL"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtDDL"
      Tab(2).Control(1)=   "Label7"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdRDDL 
         Caption         =   "Restore"
         Height          =   855
         Index           =   1
         Left            =   8490
         TabIndex        =   26
         Top             =   8430
         Width           =   1695
      End
      Begin VB.CommandButton cmdRDDL 
         Caption         =   "Generate"
         Height          =   855
         Index           =   0
         Left            =   6690
         TabIndex        =   25
         Top             =   8430
         Width           =   1695
      End
      Begin VB.TextBox txtDDL 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   6945
         Left            =   -74520
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   1530
         Width           =   9765
      End
      Begin VB.Frame Frame2 
         Caption         =   "Type Of Restore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2280
         TabIndex        =   20
         Top             =   4560
         Width           =   4845
         Begin VB.OptionButton optLog 
            Caption         =   "Log"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   22
            Top             =   330
            Width           =   2625
         End
         Begin VB.OptionButton optDatabase 
            Caption         =   "Database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   330
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkForce 
         Caption         =   "Force restore Over Existing Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73710
         TabIndex        =   19
         Top             =   1140
         Width           =   3915
      End
      Begin VB.CommandButton cmdChooseDeviceFile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   9630
         TabIndex        =   18
         Top             =   2670
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid msf_FileList 
         Height          =   2325
         Left            =   -73830
         TabIndex        =   16
         Top             =   2190
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   4101
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   3
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   270
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "Form1.frx":0054
         Top             =   720
         Width           =   705
      End
      Begin VB.Frame Frame1 
         Caption         =   "Recovery State options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Left            =   -73770
         TabIndex        =   9
         Top             =   5700
         Width           =   8085
         Begin VB.OptionButton optrecover 
            Caption         =   "Recover the Database to operational State"
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   420
            Width           =   6105
         End
         Begin VB.OptionButton optNonop 
            Caption         =   "Do not recover the database  but able to apply more log backups"
            Height          =   375
            Left            =   210
            TabIndex        =   12
            Top             =   1080
            Width           =   6045
         End
         Begin VB.OptionButton optStandby 
            Caption         =   "Do not recover the database but leave it as Read Only and  able to apply more log backups"
            Height          =   375
            Left            =   210
            TabIndex        =   11
            Top             =   1680
            Width           =   6855
         End
         Begin VB.TextBox txtStandby 
            Height          =   285
            Left            =   1830
            TabIndex        =   10
            Top             =   2250
            Width           =   4635
         End
         Begin VB.Label Label5 
            Caption         =   "Standby File"
            Height          =   255
            Left            =   750
            TabIndex        =   14
            Top             =   2250
            Width           =   945
         End
      End
      Begin VB.ListBox lstdevices 
         Height          =   1425
         Left            =   2280
         TabIndex        =   6
         Top             =   2700
         Width           =   7275
      End
      Begin VB.ComboBox cboDatabases 
         Height          =   315
         Left            =   4350
         TabIndex        =   2
         Top             =   1620
         Width           =   5925
      End
      Begin VB.ComboBox cboServers 
         Height          =   315
         Left            =   4380
         TabIndex        =   1
         Top             =   1110
         Width           =   5895
      End
      Begin MSFlexGridLib.MSFlexGrid msf_Available 
         Height          =   1665
         Left            =   2250
         TabIndex        =   5
         Top             =   5670
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2937
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowUserResizing=   3
      End
      Begin VB.Label Label7 
         Caption         =   "Generated DDL"
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
         Left            =   -74520
         TabIndex        =   24
         Top             =   1170
         Width           =   6285
      End
      Begin VB.Label Label6 
         Caption         =   "Restore Database Files As:"
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
         Left            =   -73740
         TabIndex        =   17
         Top             =   1890
         Width           =   5355
      End
      Begin VB.Label Label2 
         Caption         =   "Backup Location Or Device"
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
         Left            =   2190
         TabIndex        =   8
         Top             =   2460
         Width           =   2955
      End
      Begin VB.Label Label3 
         Caption         =   "Available Backups"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2190
         TabIndex        =   7
         Top             =   5460
         Width           =   3555
      End
      Begin VB.Label Label1 
         Caption         =   "Restore As Database:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2190
         TabIndex        =   4
         Top             =   1680
         Width           =   2205
      End
      Begin VB.Label Label4 
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub GenerateDDLAndRestore(servername As String, DatabaseName As String, Force As Integer, location As String, deviceYN As Integer, DBRestore As Integer, filenumber As Integer, LeftInState As Integer, MoveFiles As Integer, movestring As String, standbyfilesloc As String)

On Error GoTo err_handler

Set oRestore = New SQLDMO.Restore


'check to see if using Device or file
If deviceYN = 1 Then
    oRestore.Devices = location
Else
    oRestore.Files = location
End If

'Should we force this backup over the top of the existing Database

If Force = 1 Then
    oRestore.ReplaceDatabase = True
End If

'What type of backup are we doing

If DBRestore = 1 Then
    oRestore.Action = SQLDMORestore_Database
Else
    oRestore.Action = SQLDMORestore_Log
End If

'Are we moving the files at all

If MoveFiles = 1 Then
    oRestore.RelocateFiles = movestring
End If

'What database do we want to restore
oRestore.Database = DatabaseName

Select Case LeftInState

'Recover the database
Case 1
oRestore.LastRestore = True
'Leave non-operational
Case 2
oRestore.LastRestore = False
'leave in standby
Case 3
oRestore.LastRestore = False
oRestore.StandbyFiles = standbyfilesloc

End Select

oRestore.filenumber = filenumber

'Because we want to reuse this procedure for DDL only or DDl and
'actual restore we include this clause here.
If servername <> "" Then
    Set oServer = New SQLDMO.SQLServer
    oServer.LoginSecure = True
    oServer.Connect servername
    
    oRestore.SQLRestore oServer
End If


txtDDL.Text = oRestore.GenerateSQL

Exit Sub

err_handler:
MsgBox "The restore has failed because... " & vbCrLf & Err.Description
Exit Sub






End Sub


Private Sub cboServers_Click()

'Show the databases on the server selected
ShowDatabases cboServers.Text
End Sub

Private Sub cmdChooseDeviceFile_Click()
Form2.Visible = True
End Sub




Private Sub cmdRDDL_Click(Index As Integer)



If cboServers.Text <> "" And cboDatabases.Text <> "" And lstdevices.Text <> "" And msf_Available.Rows > 1 Then

Set oServer = New SQLDMO.SQLServer
Set oRestore = New SQLDMO.Restore

oServer.LoginSecure = True
oServer.Connect cboServers.Text



Dim iLeftInState As Integer
Dim iDBRestore As Integer
Dim iForce As Integer
Dim ideviceYN As Integer
Dim iMoveFiles As Integer
Dim iBackupType As Integer
Dim strMoveFiles As String
Dim qry_Comparison As SQLDMO.QueryResults
Dim i As Integer

iForce = 0
ideviceYN = 0
iBackupType = 1
strMoveFiles = ""
iMoveFiles = 0


Select Case chkForce.Value

Case vbChecked
iForce = 1
End Select

For Each oDevice In oServer.BackupDevices
    If oDevice.Name = lstdevices.Text Then
        ideviceYN = 1
    End If
Next oDevice


Select Case optDatabase.Value

Case False
iBackupType = 0
End Select


If optrecover.Value = True Then
    iLeftInState = 1
ElseIf optNonop.Value = True Then
    iLeftInState = 2
ElseIf optStandby = True Then
    iLeftInState = 3
End If

If iLeftInState = 3 And txtStandby = "" Then
    MsgBox "You have chosen to place the database in standby." & vbCrLf & "You need to specify an undo file", vbInformation, "Missing Undo File"
    Exit Sub
End If

'Now the tricky part.  We need to compare the values for the placement of the
'files in the grid with those that we know are in the backupset row for row.
'If we find any that are different then this will need to be our
'move string for the restore

If ideviceYN = 1 Then
    oRestore.Devices = lstdevices.Text
Else
    oRestore.Files = lstdevices.Text
End If

oRestore.filenumber = msf_Available.TextMatrix(msf_Available.Row, 1)

Set qry_Comparison = oRestore.ReadFileList(oServer)

For i = 1 To qry_Comparison.Rows
    If msf_FileList.TextMatrix(i, 1) <> qry_Comparison.GetColumnString(i, 2) Then
        strMoveFiles = strMoveFiles & "[" & msf_FileList.TextMatrix(i, 0) & "],[" & msf_FileList.TextMatrix(i, 1) & "],"
    End If
Next i

If Len(strMoveFiles) > 0 Then
    strMoveFiles = Mid(strMoveFiles, 1, Len(strMoveFiles) - 1)
    iMoveFiles = 1
End If


Select Case Index

Case 0 'Generate DDL only
GenerateDDLAndRestore "", cboDatabases.Text, iForce, lstdevices.Text, ideviceYN, iBackupType, oRestore.filenumber, iLeftInState, iMoveFiles, strMoveFiles, txtStandby.Text

Case 1
GenerateDDLAndRestore "", cboDatabases.Text, iForce, lstdevices.Text, ideviceYN, iBackupType, oRestore.filenumber, iLeftInState, iMoveFiles, strMoveFiles, txtStandby.Text
GenerateDDLAndRestore cboServers.Text, cboDatabases.Text, iForce, lstdevices.Text, ideviceYN, iBackupType, oRestore.filenumber, iLeftInState, iMoveFiles, strMoveFiles, txtStandby.Text


End Select

End If


End Sub

Private Sub Form_Load()

'Set the headings for our Flexgrid which is
'going to show any backups on the media

msf_Available.TextMatrix(0, 0) = "Backup Type"
msf_Available.TextMatrix(0, 1) = "Position"
msf_Available.TextMatrix(0, 2) = "Database Name"
msf_Available.TextMatrix(0, 3) = "Backup Finish Date"
msf_FileList.TextMatrix(0, 0) = "Logical File Name"
msf_FileList.TextMatrix(0, 1) = "Physical File Name"

'Set the widths of the cells to a decent first width
'these can be adjusted.  This is on both the gris that
'will be used to display the results of the RESTORE HEADERONLY
'and the RESTORE FILELISTONLY

msf_Available.ColWidth(0) = 1215
msf_Available.ColWidth(1) = 735
msf_Available.ColWidth(2) = 3255
msf_Available.ColWidth(3) = 2685
msf_FileList.ColWidth(0) = 3952
msf_FileList.ColWidth(1) = 3953

'Make sure the first tab we show is the general tab

SSTab1.Tab = 0

msf_FileList.Rows = 1
msf_Available.Rows = 1

'Populate our servers combobox with the names of all our servers
ShowServers

optDatabase.Value = True
optrecover.Value = True


End Sub




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

Private Sub ShowContentsOfBackup(servername As String, devicename As String)

Dim oHeader As SQLDMO.QueryResults
Dim DeviceCounter As Integer
Dim i As Integer
Dim Backup_type As String
DeviceCounter = 0
msf_Available.Rows = 1

Set oServer = New SQLDMO.SQLServer
Set oRestore = New SQLDMO.Restore


oServer.LoginSecure = True
oServer.Connect servername

'See if the name we passed is a device or an adhoc file

For Each oDevice In oServer.BackupDevices
    If oDevice.Name = devicename Then
        DeviceCounter = 1
    End If
Next oDevice



If DeviceCounter = 0 Then
    oRestore.Files = devicename
Else
    oRestore.Devices = devicename
End If


'here is where we read the contents of the backup header into a Queryresults Object
'ready for processing.  A Queryresults object is very much like a table so we
' just go through it processing rows and columns.  this is the RESTORE HEADERONLY part

Set oHeader = oRestore.ReadBackupHeader(oServer)

For i = 1 To oHeader.Rows

'Here we convery the unintelligible integer value for backup type into
'something we can read and comprehend without opening the reference books

Select Case oHeader.GetColumnString(i, 3)
    Case 1
    Backup_type = "FULL"
    Case 2
    Backup_type = "LOG"
    Case 4
    Backup_type = "FILE"
    Case 5
    Backup_type = "DIFF DB"
    Case 6
    Backup_type = "DIFF FILE"
End Select

'Add our required column and row values to the flexgrid showing us the contents
'of the backup device we specified
    msf_Available.AddItem Backup_type & vbTab & oHeader.GetColumnString(i, 6) & vbTab & oHeader.GetColumnString(i, 10) & vbTab & oHeader.GetColumnString(i, 19)
Next i

oServer.DisConnect
Set oServer = Nothing
Set oRestore = Nothing


End Sub

Private Sub lstdevices_Click()

'I don't want to see the devices if they are not linked to a server
If cboServers.Text <> "" Then
    ShowContentsOfBackup cboServers.Text, lstdevices.Text
End If
End Sub


Private Sub ShowMeFileDetails(servername As String, devicename As String, filenumber As Integer)


Dim oFileResults As SQLDMO.QueryResults

Dim DeviceCounter As Integer
Dim i As Integer

DeviceCounter = 0

msf_FileList.Rows = 1

Set oServer = New SQLDMO.SQLServer
Set oRestore = New SQLDMO.Restore


oServer.LoginSecure = True
oServer.Connect servername

'See if the name we passed is a device or an adhoc file

For Each oDevice In oServer.BackupDevices
    If oDevice.Name = devicename Then
        DeviceCounter = 1
    End If
Next oDevice


If DeviceCounter = 0 Then
    oRestore.Files = devicename
Else
    oRestore.Devices = devicename
End If

oRestore.filenumber = filenumber


'Up until now the code has been exactly the same as when
'we wanted to populate the other grid with details of our backup headers

'this is the part that we do the equivalent of RESTORE FILELISTONLY

Set oFileResults = oRestore.ReadFileList(oServer)

For i = 1 To oFileResults.Rows


'Add our required column and row values to the flexgrid showing us the file details
'of the backup device we specified
    msf_FileList.AddItem oFileResults.GetColumnString(i, 1) & vbTab & oFileResults.GetColumnString(i, 2)
Next i

End Sub

Private Sub msf_Available_Click()

If msf_Available.Row <> 0 Then
    ShowMeFileDetails cboServers.Text, lstdevices.Text, msf_Available.TextMatrix(msf_Available.Row, 1)
End If
End Sub

Private Sub msf_FileList_Click()
Dim strNewFile As String
'Here we want to set any new values in the

If msf_FileList.Col = 1 Then
    strNewFile = InputBox("Enter New Location Of File", "New File Location", msf_FileList.Text)
        If strNewFile = "" Or msf_FileList.Col <> 1 Then
            msf_FileList.Text = msf_FileList.Text
        Else
            msf_FileList.Text = strNewFile
        End If
End If

End Sub


