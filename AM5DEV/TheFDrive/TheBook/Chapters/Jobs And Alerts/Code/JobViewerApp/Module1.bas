Attribute VB_Name = "Module1"
Option Explicit

Public strLastServerLoggedOnTo As String
Public oServer As New SQLDMO.SQLServer
Public oRserver As SQLDMO.RegisteredServer
Public oJob As SQLDMO.job
Public oJobStep As SQLDMO.JobStep
Public oGroup As SQLDMO.ServerGroup
Public oJobSchedule As SQLDMO.JobSchedule



Sub main()

strLastServerLoggedOnTo = "GOBBLEDYGOOK"
frmJobsMain.Show
End Sub




