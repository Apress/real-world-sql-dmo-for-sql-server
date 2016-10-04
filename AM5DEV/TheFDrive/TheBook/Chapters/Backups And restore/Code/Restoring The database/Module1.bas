Attribute VB_Name = "Module1"
Option Explicit


Public oServer As SQLDMO.SQLServer
Public oDevice As SQLDMO.BackupDevice
Public oDatabase As SQLDMO.Database
Public oBackup As SQLDMO.Backup
Public oRegisteredServer As SQLDMO.RegisteredServer
Public oApp As SQLDMO.Application
Public oRestore As SQLDMO.Restore
Public oGroup As SQLDMO.ServerGroup
