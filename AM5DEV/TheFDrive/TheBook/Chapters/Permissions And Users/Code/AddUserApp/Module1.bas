Attribute VB_Name = "Module1"
Option Explicit

Public oGroup As SQLDMO.ServerGroup
Public oRServer As SQLDMO.RegisteredServer
Public oServer As SQLDMO.SQLServer
Public oDatabase As SQLDMO.Database
Public oLogin As SQLDMO.Login
Public oUser As SQLDMO.User
Public oSRole As SQLDMO.ServerRole
Public oDBRole As SQLDMO.DatabaseRole
Public oNameList As SQLDMO.NameList
Public Const DO_NOTHING = 0
Public Const IS_MEMBER = 1
Public Const NEEDS_ADDING = 2
Public Const NEEDS_DELETING = 3
