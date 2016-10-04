Attribute VB_Name = "Module1"
Option Explicit
Sub main()
ListNewDatabases 15
End Sub


Private Sub ListNewDatabases(MinsAgo As Integer)
Dim oServer As SQLDMO.SQLServer
Dim oBackup As SQLDMO.Backup
Dim oDatabase As SQLDMO.Database
Dim oList As SQLDMO.SQLObjectList
Dim obj As Object
Dim oUsr As SQLDMO.User
Dim oJob As SQLDMO.Job
Dim OutputString As String
Dim oObj As Object

On Error GoTo err_handler

Set oServer = New SQLDMO.SQLServer
Set oBackup = New SQLDMO.Backup


    oServer.LoginSecure = True
    oServer.Connect "AM2"

        For Each oDatabase In oServer.Databases
            If oDatabase.Name <> "master" And oDatabase.Name <> "tempdb" And oDatabase.Name <> "model" And oDatabase.Name <> "msdb" Then
                'Check for created databases
                If DateDiff("n", Left(oDatabase.CreateDate, 19), Format(Now(), "dd-mm-yyyy hh:mm:ss")) <= MinsAgo Then
                    OutputString = OutputString & "Database " & oDatabase.Name & " on " & oServer.Name & " was created on " & Left(oDatabase.CreateDate, 19) & " by " & oDatabase.Owner & vbCrLf & "The database was backed Up to \\MyNetworkServer\FullBackupShare\" & oServer.Name & "\" & oDatabase.Name & "_" & Format(Now(), "yyyymmyy") & ".bak" & vbCrLf
                    oBackup.Database = oDatabase.Name
                    oBackup.Files = "f:\DBBackups\" & oServer.Name & "\" & oDatabase.Name & "_" & Format(Now(), "yyyymmyy") & ".bak"
                    oBackup.SQLBackup oServer
                End If
                
            End If
          Next oDatabase
          
          oServer.DisConnect
    


If OutputString <> "" Then

Open "c:\NewDBs.txt" For Output As #1

Print #1, OutputString

End If



Exit Sub
Close #1

err_handler:
Resume Next

End Sub
