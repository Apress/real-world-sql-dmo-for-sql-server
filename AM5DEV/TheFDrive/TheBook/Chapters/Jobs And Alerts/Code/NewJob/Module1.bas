Attribute VB_Name = "Module1"
Option Explicit

Public oServer As New SQLDMO.SQLServer
Public odatabase As SQLDMO.Database
Public ojob As SQLDMO.Job
Public stepid As Integer
Public OJobStep As SQLDMO.JobStep
Public strJobName As String

Public oJobSchedule As New SQLDMO.JobSchedule




