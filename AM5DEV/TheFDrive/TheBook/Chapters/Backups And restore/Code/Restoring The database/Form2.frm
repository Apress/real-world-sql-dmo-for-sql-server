VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Choose Location"
   ClientHeight    =   2370
   ClientLeft      =   6090
   ClientTop       =   5055
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   ScaleHeight     =   2370
   ScaleWidth      =   6270
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   3060
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   285
      Left            =   4980
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optFile 
      BackColor       =   &H00C0C0C0&
      Caption         =   "File"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1110
      Width           =   1845
   End
   Begin VB.OptionButton optDevice 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Device"
      Height          =   225
      Left            =   270
      TabIndex        =   3
      Top             =   420
      Width           =   1695
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "..."
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   1380
      Width           =   555
   End
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Top             =   1380
      Width           =   4335
   End
   Begin VB.ComboBox cboDevices 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   720
      Width           =   4305
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ShowBackupDevices(servername As String)

Set oServer = New SQLDMO.SQLServer

oServer.LoginSecure = True
oServer.Connect servername

cboDevices.Clear

For Each oDevice In oServer.BackupDevices
    cboDevices.AddItem oDevice.Name
Next oDevice

oServer.DisConnect
Set oServer = Nothing
Set oDevice = Nothing

End Sub


Private Sub cmdClose_Click()
If optDevice.Value = True And cboDevices.Text <> "" Then
    Form1.lstdevices.AddItem cboDevices.Text
ElseIf optFile.Value = True And txtFilename.Text <> "" Then
    Form1.lstdevices.AddItem txtFilename.Text
End If
    
Me.Visible = False
txtFilename.Text = ""
End Sub


Private Sub cmdFind_Click()

If txtFilename.Enabled = False Then
    Exit Sub
Else

cdl1.DialogTitle = "Find Your Backup"
cdl1.InitDir = "c:\"
cdl1.ShowOpen

If (cdl1.FileName <> "" And cdl1.CancelError = False) Then
    txtFilename.Text = cdl1.FileName
End If

End If
    
End Sub

Private Sub Form_Load()
optDevice.Value = True
cboDevices.Enabled = True
txtFilename.Enabled = False
txtFilename.BackColor = &H8000000F
cboDevices.BackColor = &H80000005
ShowBackupDevices Form1.cboServers.Text
End Sub

Private Sub Form_Paint()
'optDevice.Value = True
'cboDevices.Enabled = True
'txtFilename.Enabled = False
'txtFilename.BackColor = &H8000000F
'cboDevices.BackColor = &H80000005
End Sub

Private Sub optDevice_Click()
cboDevices.Enabled = True
txtFilename.Enabled = False
txtFilename.BackColor = &H8000000F
cboDevices.BackColor = &H80000005
End Sub

Private Sub optFile_Click()

cboDevices.Enabled = False
txtFilename.Enabled = True
cboDevices.BackColor = &H8000000F
txtFilename.BackColor = &H80000005
End Sub
