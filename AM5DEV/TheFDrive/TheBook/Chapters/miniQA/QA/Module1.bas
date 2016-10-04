Attribute VB_Name = "Module1"
Option Explicit

Public oServer As SQLDMO.SQLServer
Public glb_Server As String
Public b_IsConnected As Boolean


Sub main()
App.EXEName = "Lite QA"
App.Title = "Lite QA"
App.ProductName = "Lite QA"
b_IsConnected = False
End Sub
