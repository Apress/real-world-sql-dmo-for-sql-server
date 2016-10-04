VERSION 5.00
Begin VB.Form frmCreateDevice 
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   6090
   ClientTop       =   5835
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   4215
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create device"
      Height          =   585
      Left            =   870
      TabIndex        =   0
      Top             =   180
      Width           =   2535
   End
End
Attribute VB_Name = "frmCreateDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreate_Click()
Dim oServer As SQLDMO.SQLServer
Dim oDevice As SQLDMO.BackupDevice

Set oServer = New SQLDMO.SQLServer

oServer.LoginSecure = True
oServer.Connect "AM2"

Set oDevice = New SQLDMO.BackupDevice

oDevice.Name = "My_Backup_Device"
oDevice.Type = SQLDMODevice_DiskDump
oDevice.PhysicalLocation = "f:\DBBackups\MyBackups.bak"

oServer.BackupDevices.Add oDevice


oServer.DisConnect

Set oServer = Nothing


End Sub
