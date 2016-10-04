VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIntro 
   BackColor       =   &H0000FFFF&
   ClientHeight    =   10215
   ClientLeft      =   930
   ClientTop       =   2595
   ClientWidth     =   16755
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   16755
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Grid"
      Height          =   525
      Left            =   5610
      TabIndex        =   35
      Top             =   7710
      Width           =   1425
   End
   Begin VB.TextBox txtIncrement 
      Height          =   315
      Left            =   6510
      TabIndex        =   33
      Top             =   4530
      Width           =   705
   End
   Begin VB.TextBox txtSeed 
      Height          =   315
      Left            =   5700
      TabIndex        =   31
      Top             =   4530
      Width           =   705
   End
   Begin VB.CheckBox chkIdent 
      BackColor       =   &H0000C0C0&
      Caption         =   "Identity Col"
      Height          =   285
      Left            =   4320
      TabIndex        =   30
      Top             =   4530
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddRelToDB 
      Caption         =   "Add to Database"
      Height          =   405
      Left            =   11910
      TabIndex        =   27
      Top             =   7140
      Width           =   3135
   End
   Begin VB.CommandButton cmdAddrelgrid 
      Caption         =   "Add to grid"
      Height          =   615
      Left            =   12420
      TabIndex        =   26
      Top             =   4050
      Width           =   2535
   End
   Begin VB.ListBox lstcol2 
      Height          =   1425
      Left            =   13710
      TabIndex        =   25
      Top             =   2400
      Width           =   2895
   End
   Begin VB.ListBox lstcol1 
      Height          =   1425
      Left            =   10740
      TabIndex        =   24
      Top             =   2400
      Width           =   2805
   End
   Begin MSFlexGridLib.MSFlexGrid relgrid 
      Height          =   2115
      Left            =   10860
      TabIndex        =   23
      Top             =   4860
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3731
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.ComboBox cbotables2 
      Height          =   315
      Left            =   13650
      TabIndex        =   22
      Top             =   1470
      Width           =   2865
   End
   Begin VB.TextBox txtdefault 
      Height          =   315
      Left            =   4320
      TabIndex        =   20
      Top             =   4110
      Width           =   2535
   End
   Begin VB.OptionButton optKey 
      BackColor       =   &H0000C0C0&
      Caption         =   "Primary Key"
      Height          =   285
      Left            =   4350
      TabIndex        =   19
      Top             =   3420
      Width           =   1275
   End
   Begin VB.ComboBox cbotables 
      Height          =   315
      Left            =   10680
      TabIndex        =   17
      Top             =   1440
      Width           =   2865
   End
   Begin VB.TextBox txtcolname 
      Height          =   285
      Left            =   150
      TabIndex        =   15
      Top             =   2640
      Width           =   3075
   End
   Begin VB.CommandButton cmdAddCol 
      Caption         =   "Add Column"
      Height          =   675
      Left            =   3810
      TabIndex        =   14
      Top             =   5010
      Width           =   1515
   End
   Begin MSFlexGridLib.MSFlexGrid TableBuilderGrid 
      Height          =   1845
      Left            =   120
      TabIndex        =   12
      Top             =   5730
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   3254
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.ComboBox cboServers 
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   1080
      Width           =   3285
   End
   Begin VB.ComboBox cboDatabases 
      Height          =   315
      Left            =   3690
      TabIndex        =   8
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   525
      Left            =   7500
      TabIndex        =   7
      Top             =   7680
      Width           =   1395
   End
   Begin VB.CheckBox chkNull 
      BackColor       =   &H0000C0C0&
      Caption         =   "Nullable"
      Height          =   375
      Left            =   4350
      TabIndex        =   6
      Top             =   2910
      Width           =   1185
   End
   Begin VB.TextBox txtLength 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   2520
      Width           =   1125
   End
   Begin VB.ListBox lstDatatypes 
      Height          =   1620
      Left            =   150
      TabIndex        =   2
      Top             =   3300
      Width           =   4065
   End
   Begin VB.TextBox txtTableName 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   1980
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "Increment"
      Height          =   285
      Left            =   6510
      TabIndex        =   34
      Top             =   4860
      Width           =   795
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Seed"
      Height          =   255
      Left            =   5700
      TabIndex        =   32
      Top             =   4860
      Width           =   645
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "referencing table"
      Height          =   285
      Left            =   13710
      TabIndex        =   29
      Top             =   2100
      Width           =   2355
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "referenced table"
      Height          =   285
      Left            =   10710
      TabIndex        =   28
      Top             =   2100
      Width           =   2385
   End
   Begin VB.Label De 
      BackColor       =   &H0000C0C0&
      Caption         =   "Default"
      Height          =   285
      Left            =   4350
      TabIndex        =   21
      Top             =   3840
      Width           =   2235
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Tables"
      Height          =   255
      Left            =   12720
      TabIndex        =   18
      Top             =   690
      Width           =   1995
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "Column Name"
      Height          =   255
      Left            =   150
      TabIndex        =   16
      Top             =   2460
      Width           =   2265
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "Table Creation"
      Height          =   315
      Left            =   90
      TabIndex        =   13
      Top             =   5490
      Width           =   3675
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Servers"
      Height          =   285
      Left            =   150
      TabIndex        =   11
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Databases"
      Height          =   255
      Left            =   3690
      TabIndex        =   9
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label l 
      BackColor       =   &H0000C0C0&
      Caption         =   "Length"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Datatypes"
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   3090
      Width           =   3315
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Table Name"
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Top             =   1710
      Width           =   2565
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim database As SQLDMO.database
Dim server As SQLDMO.SQLServer
Dim srvRole As SQLDMO.ServerRole
Dim dbroles As SQLDMO.DatabaseRole
Dim srvRoleMembers As SQLDMO.QueryResults

Private Sub cboDatabases_Click()
ListTables cboServers.Text, cboDatabases.Text
ListTables2 cboServers.Text, cboDatabases.Text
End Sub

Private Sub cboServers_Click()
ListDataBases cboServers.Text
ListDataTypes cboServers.Text
End Sub

Private Sub AddColToGrid(KeyCol As String, colname As String, datatype As String, collength As String, nullability As String, default As String)

Dim i As Integer
Dim iCountofids As Integer
Dim idYes As String
iCountofids = 0


'Do we have any id cols already
    
    For i = 0 To TableBuilderGrid.Rows - 1
        If TableBuilderGrid.TextMatrix(i, 6) = "YES" Then
            iCountofids = 1
        End If
    Next i
    
    If iCountofids = 1 And chkIdent.Value = vbChecked Then
        MsgBox "You already have an identity Field", vbInformation, "Identity Error"
        Exit Sub
    End If
    
    
    If chkIdent.Value = vbChecked And IsNumeric(txtSeed.Text) And IsNumeric(txtIncrement.Text) Then
        idYes = "YES"
        TableBuilderGrid.AddItem KeyCol & vbTab & colname & vbTab & datatype & vbTab & collength & vbTab & nullability & vbTab & default & vbTab & idYes & vbTab & txtSeed.Text & vbTab & txtIncrement.Text
    ElseIf chkIdent.Value = vbChecked And (IsNumeric(txtSeed.Text) = False Or IsNumeric(txtIncrement.Text) = False) Then
        MsgBox "Not a valid seed or increment"
        Exit Sub
    Else
        TableBuilderGrid.AddItem KeyCol & vbTab & colname & vbTab & datatype & vbTab & collength & vbTab & nullability & vbTab & default
    End If
    


End Sub

Private Sub cbotables_Click()
Listcols cboServers.Text, cboDatabases.Text, cbotables.Text
End Sub
Private Sub cbotables2_Click()
Listcols2 cboServers.Text, cboDatabases.Text, cbotables2.Text
End Sub

Private Sub cmdAddCol_Click()
Dim nullability As String
Dim KeyCol As String
Select Case chkNull.Value

Case 1
nullability = "NULL"
Case 0
nullability = "NOT NULL"
End Select

If optKey.Value = True Then
    KeyCol = "PK"
End If



AddColToGrid KeyCol, txtcolname.Text, lstDatatypes.Text, txtLength.Text, nullability, txtdefault.Text

txtcolname.Text = ""
txtLength.Text = ""
chkNull.Value = 0
optKey.Value = False
txtdefault.Text = ""
End Sub

Private Sub cmdAddrelgrid_Click()
AddFieldsToGrid lstcol1.Text, lstcol2.Text
End Sub

Private Sub AddRelationshipToTable(servername As String, databasename As String, reffingTable As String, reffedTable As String)
Dim dmo_server As SQLDMO.SQLServer
Dim dmo_user As SQLDMO.User
Dim oKey As SQLDMO.Key
Dim i As Integer
 
Set dmo_server = New SQLDMO.SQLServer
 
dmo_server.LoginSecure = True
dmo_server.Connect servername

If relgrid.Rows > 0 Then
    Set oKey = New SQLDMO.Key
        oKey.Type = SQLDMOKey_Foreign
        
        'we only need to set the referenced table name
        'as we will be adding the key to the referencing table
        
        oKey.ReferencedTable = reffedTable
        

        For i = 0 To relgrid.Rows - 1
            oKey.KeyColumns.Add relgrid.TextMatrix(i, 2)
            oKey.ReferencedColumns.Add relgrid.TextMatrix(i, 0)
        Next i

        dmo_server.Databases(databasename).Tables(reffingTable).Keys.Add oKey
End If


End Sub


Private Sub cmdAddRelToDB_Click()
AddRelationshipToTable cboServers.Text, cboDatabases.Text, cbotables2.Text, cbotables.Text
relgrid.Clear
relgrid.Rows = 0
End Sub

Private Sub cmdClear_Click()
TableBuilderGrid.Clear
TableBuilderGrid.Rows = 0
End Sub

Private Sub cmdCreate_Click()
AddTableToDatabase cboServers.Text, cboDatabases.Text, txtTableName.Text
ListTables cboServers.Text, cboDatabases.Text
ListTables2 cboServers.Text, cboDatabases.Text
End Sub







Private Sub AddTableToDatabase(servername As String, databasename As String, tablename As String)
On Error GoTo err_handler
Dim i As Integer 'for rows
Dim j As Integer 'for columns
Dim oSrv As SQLDMO.SQLServer
Dim oTable As SQLDMO.Table
Dim ocol As SQLDMO.Column
Dim colname As String
Dim datatype As String
Dim collength As Integer
Dim nullability As Boolean
Dim objPK As SQLDMO.Key
Dim objDefault As SQLDMO.default

Set objPK = New SQLDMO.Key

Set oSrv = New SQLDMO.SQLServer

oSrv.LoginSecure = True
oSrv.Connect servername

Set oTable = New SQLDMO.Table
oTable.Name = tablename


'loop through the rows of the grid
For i = 0 To TableBuilderGrid.Rows - 1
'initialise a new column every time we go to a new row
    Set ocol = New SQLDMO.Column
'is it a textual column that will require a length ?
        If TableBuilderGrid.TextMatrix(i, 2) = "char" Or TableBuilderGrid.TextMatrix(i, 2) = "varchar" Or TableBuilderGrid.TextMatrix(i, 2) = "nvarchar" Or TableBuilderGrid.TextMatrix(i, 2) = "nchar" Then
            ocol.Name = TableBuilderGrid.TextMatrix(i, 1)
            ocol.datatype = TableBuilderGrid.TextMatrix(i, 2)
            ocol.Length = CStr(TableBuilderGrid.TextMatrix(i, 3))
                If TableBuilderGrid.TextMatrix(i, 4) = "NOT NULL" Then
                    ocol.AllowNulls = False
                Else
                    ocol.AllowNulls = True
                End If
        Else
            ocol.Name = TableBuilderGrid.TextMatrix(i, 1)
            ocol.datatype = TableBuilderGrid.TextMatrix(i, 2)
                If TableBuilderGrid.TextMatrix(i, 4) = "NOT NULL" Then
                    ocol.AllowNulls = False
                Else
                    ocol.AllowNulls = True
                End If
        End If
        
        If TableBuilderGrid.TextMatrix(i, 6) = "YES" Then
            ocol.Identity = True
            ocol.IdentitySeed = TableBuilderGrid.TextMatrix(i, 7)
            ocol.IdentityIncrement = TableBuilderGrid.TextMatrix(i, 8)
        End If
        
    oTable.Columns.Add ocol
        
Next i

oSrv.Databases(databasename).Tables.Add oTable

'Now for Primary Keys and defaults

'Primary Keys
Set objPK = New SQLDMO.Key
For i = 0 To TableBuilderGrid.Rows - 1
    If TableBuilderGrid.TextMatrix(i, 0) = "PK" Then
        objPK.KeyColumns.Add TableBuilderGrid.TextMatrix(i, 1)
    End If
Next i

If objPK.KeyColumns.Count > 0 Then
    
objPK.Name = "My_PK_" & CStr(Minute(Now())) & "_" & CStr(Second(Now()))


objPK.FillFactor = 85
objPK.Type = SQLDMOKey_Primary

oSrv.Databases(databasename).Tables(tablename).Keys.Add objPK

End If


'Defaults
For i = 0 To TableBuilderGrid.Rows - 1
 Set objDefault = New SQLDMO.default
    objDefault.Name = "DEFAULT_" & TableBuilderGrid.TextMatrix(i, 1)
   If TableBuilderGrid.TextMatrix(i, 5) <> "" Then
   oSrv.Databases(databasename).Tables(tablename).BeginAlter
   oSrv.Databases(databasename).Tables(tablename).Columns(TableBuilderGrid.TextMatrix(i, 1)).DRIDefault.Text = "'" & TableBuilderGrid.TextMatrix(i, 5) & "'"
   oSrv.Databases(databasename).Tables(tablename).DoAlter
   End If
   Next i
   
TableBuilderGrid.Rows = 0

Exit Sub

err_handler:

MsgBox Err.Description, vbInformation, "Mistake made"
Exit Sub

End Sub
Private Sub ListDataBases(servername As String)
Dim dmoSrv As SQLDMO.SQLServer
Dim dmoDB As SQLDMO.database

Set dmoSrv = New SQLDMO.SQLServer
cboDatabases.Clear
dmoSrv.LoginSecure = True
dmoSrv.Connect servername

For Each dmoDB In dmoSrv.Databases
    cboDatabases.AddItem dmoDB.Name
Next dmoDB
End Sub



Private Sub ListTables(servername As String, databasename As String)
Dim dmoSrv As SQLDMO.SQLServer
Dim dmoDB As SQLDMO.database
Dim oTable As SQLDMO.Table

Set dmoSrv = New SQLDMO.SQLServer
cbotables.Clear
dmoSrv.LoginSecure = True
dmoSrv.Connect servername

For Each oTable In dmoSrv.Databases(databasename).Tables
    If oTable.SystemObject = False Then
        cbotables.AddItem oTable.Name
    End If
Next oTable
End Sub
Private Sub ListTables2(servername As String, databasename As String)
Dim dmoSrv As SQLDMO.SQLServer
Dim dmoDB As SQLDMO.database
Dim oTable As SQLDMO.Table

Set dmoSrv = New SQLDMO.SQLServer
cbotables2.Clear
dmoSrv.LoginSecure = True
dmoSrv.Connect servername

For Each oTable In dmoSrv.Databases(databasename).Tables
    If oTable.SystemObject = False Then
        cbotables2.AddItem oTable.Name
    End If
Next oTable
End Sub

Private Sub Listcols2(servername As String, databasename As String, tablename As String)
Dim dmoSrv As SQLDMO.SQLServer
Dim ocol As SQLDMO.Column

Set dmoSrv = New SQLDMO.SQLServer
lstcol2.Clear
dmoSrv.LoginSecure = True
dmoSrv.Connect servername

For Each ocol In dmoSrv.Databases(databasename).Tables(tablename).Columns
    lstcol2.AddItem ocol.Name
Next ocol
End Sub
Private Sub Listcols(servername As String, databasename As String, tablename As String)
Dim dmoSrv As SQLDMO.SQLServer
Dim ocol As SQLDMO.Column

Set dmoSrv = New SQLDMO.SQLServer
lstcol1.Clear
dmoSrv.LoginSecure = True
dmoSrv.Connect servername

For Each ocol In dmoSrv.Databases(databasename).Tables(tablename).Columns
    lstcol1.AddItem ocol.Name
Next ocol
End Sub



Private Sub ListDataTypes(servername As String)
Dim dt As SQLDMO.SystemDatatype
Dim dts As SQLDMO.SystemDatatypes
Dim dmoSrv As SQLDMO.SQLServer
lstDatatypes.Clear

Set dmoSrv = New SQLDMO.SQLServer

dmoSrv.LoginSecure = True
dmoSrv.Connect servername

For Each dt In dmoSrv.Databases("master").SystemDatatypes
    lstDatatypes.AddItem dt.Name
Next dt
End Sub
Private Sub AddFieldsToGrid(leftCol As String, rightcol As String)

relgrid.AddItem leftCol & vbTab & "-->" & vbTab & rightcol

End Sub

Private Sub CreateRelationship(serverame As String, databasename As String, tablename As String)

Dim dmoSrv As SQLDMO.SQLServer

Set dmoSrv = New SQLDMO.SQLServer

dmoSrv.LoginSecure = True
dmoSrv.Connect servername




End Sub


Private Sub listServers()

Dim oServer As SQLDMO.SQLServer
Dim oApp As New SQLDMO.Application
Dim oServerGroup As SQLDMO.ServerGroup
Dim oRegisteredServer As SQLDMO.RegisteredServer

cboServers.Clear

 Set oServer = New SQLDMO.SQLServer
 
For Each oRegisteredServer In oApp.ServerGroups(1).RegisteredServers
    cboServers.AddItem oRegisteredServer.Name
Next oRegisteredServer
 End Sub


Private Sub Form_Load()
TableBuilderGrid.Rows = 0
relgrid.Rows = 0
relgrid.ColWidth(0) = 2500
relgrid.ColWidth(1) = 500
relgrid.ColWidth(2) = 2500
listServers
End Sub





