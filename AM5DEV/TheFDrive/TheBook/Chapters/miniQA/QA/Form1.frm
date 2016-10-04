VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmQA 
   Caption         =   "Query Analyser Lite"
   ClientHeight    =   9450
   ClientLeft      =   1680
   ClientTop       =   3210
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   14115
   Begin MSFlexGridLib.MSFlexGrid msf_grid 
      Height          =   3795
      Left            =   270
      TabIndex        =   12
      Top             =   5400
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   6694
      _Version        =   393216
   End
   Begin VB.TextBox txtDelimiter 
      Height          =   315
      Left            =   6570
      TabIndex        =   10
      Top             =   1380
      Width           =   315
   End
   Begin VB.OptionButton optGrid 
      Caption         =   "In grid"
      Height          =   285
      Left            =   4800
      TabIndex        =   9
      Top             =   1890
      Width           =   1935
   End
   Begin VB.OptionButton optText 
      Caption         =   "In Text"
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   1380
      Width           =   1485
   End
   Begin VB.CommandButton cmdLogon 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   270
      TabIndex        =   7
      Top             =   630
      Width           =   3825
   End
   Begin VB.CommandButton cmdRun 
      Height          =   525
      Left            =   9960
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   870
      Width           =   795
   End
   Begin VB.ComboBox cboDatabases 
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Top             =   960
      Width           =   4065
   End
   Begin VB.TextBox txtResults 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3705
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   5430
      Width           =   13635
   End
   Begin VB.CommandButton cmdParse 
      Height          =   555
      Left            =   9090
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      Width           =   795
   End
   Begin VB.TextBox txtQuery 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   2955
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   2340
      Width           =   13605
   End
   Begin VB.Label Label2 
      Caption         =   "Delimiter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6900
      TabIndex        =   11
      Top             =   1410
      Width           =   1065
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   4830
      TabIndex        =   5
      Top             =   720
      Width           =   2925
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mini QA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11130
      TabIndex        =   2
      Top             =   870
      Width           =   2715
   End
End
Attribute VB_Name = "frmQA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TWIPS_PER_CHAR = 110

Dim oDatabase As SQLDMO.Database

Private Sub ExecuteQuery(DatabaseName As String, Query As String, EXECType As Integer, i_IsGridOrText As Integer, Optional str_delimiter As String)

    Debug.Print "Executing a new Query " & Now() & " " & Query
    
    If i_IsGridOrText = 1 Then

        Debug.Print "To Text Output"

    Else

        Debug.Print "To Grid Output"

    End If

    Dim str_OutputText As String
    Dim oQryresults As SQLDMO.QueryResults
    Dim i As Integer
    Dim j As Integer
    Dim txt_Results As String

    txt_Results = ""
    txtResults.Text = ""

    If str_delimiter = "" Then

        str_delimiter = ","

    End If

    On Error GoTo err_handler

    Select Case EXECType

            'Only parse

        Case 1
        
            msf_grid.Visible = False
            txtResults.Visible = True
        
            oServer.Databases(DatabaseName).ExecuteImmediate Query, SQLDMOExec_ParseOnly
            txtResults.Text = "Parse of query Completed successfully"

            'This will actually execute the code
            'We need to see if the user has selected to
            'have the output to text or grid
            'we make one or the other visible.

        Case 2

            Set oQryresults = oServer.Databases(DatabaseName).ExecuteWithResultsAndMessages(Query, Len(Query), str_OutputText)

            'Regardless of selection if the rows returned are 0 then we use the text box

            If oQryresults.Rows = 0 Then

                txtResults.Text = "No Row(s) affected"

            Else
    
                'Here we go
    
                If i_IsGridOrText = 1 Then 'Text
                
                    msf_grid.Visible = False
                    txtResults.Visible = True
    
                    'First run through we want the column names.
    
                    For j = 1 To oQryresults.Columns

                        txt_Results = txt_Results & oQryresults.ColumnName(j) & str_delimiter

                    Next j

                    txt_Results = Left(txt_Results, Len(txt_Results) - 1) & vbCrLf
    
                    'Now we want the data
    
                    For i = 1 To oQryresults.Rows

                        For j = 1 To oQryresults.Columns

                            txt_Results = txt_Results & oQryresults.GetColumnString(i, j) & str_delimiter

                        Next j

                        txt_Results = Left(txt_Results, Len(txt_Results) - 1) & vbCrLf

                    Next i

                    txtResults.Text = txt_Results

                Else
                    
                    'We now want to do the grid work
                    
                    txtResults.Visible = False
                    
                    msf_grid.Visible = True
                    
                    msf_grid.Rows = 0
                    msf_grid.Cols = oQryresults.Columns
                    msf_grid.FixedCols = 0
                    
                    'First run through we want the column names.
    
                    For j = 1 To oQryresults.Columns

                        txt_Results = txt_Results & oQryresults.ColumnName(j) & vbTab
                        
                        'check column widths and resize
                        
                        If msf_grid.ColWidth(j - 1) < (Len(oQryresults.ColumnName(j)) * TWIPS_PER_CHAR) Then

                            msf_grid.ColWidth(j - 1) = Len(oQryresults.ColumnName(j)) * TWIPS_PER_CHAR

                        End If

                    Next j

                    msf_grid.AddItem txt_Results
    
                    txt_Results = ""
    
                    'Now we want the data
    
                    For i = 1 To oQryresults.Rows

                        For j = 1 To oQryresults.Columns

                            txt_Results = txt_Results & oQryresults.GetColumnString(i, j) & vbTab
                           
                            'check column widths and resize
                           
                            If msf_grid.ColWidth(j - 1) < (Len(oQryresults.GetColumnString(i, j)) * TWIPS_PER_CHAR) Then

                                msf_grid.ColWidth(j - 1) = Len(oQryresults.GetColumnString(i, j)) * TWIPS_PER_CHAR

                            End If

                        Next j

                        msf_grid.AddItem txt_Results
                        
                        txt_Results = ""

                    Next i

                End If
    
            End If

    End Select

    If str_OutputText <> "" Then

        txtResults.Text = txtResults.Text & vbCrLf & str_OutputText

    End If

    Exit Sub

err_handler:
    txtResults.Text = Err.Source & vbCrLf & Err.Description
    Exit Sub

End Sub

Private Sub cmdLogon_Click()

    frmLogin.Show 1, Me

    If b_IsConnected = True Then

        MsgBox "Connected to " & glb_Server
        ShowDatabases

    End If

End Sub

Private Sub ShowDatabases()

    cboDatabases.Clear

    If glb_Server <> "" Then

        For Each oDatabase In oServer.Databases

            cboDatabases.AddItem oDatabase.Name

        Next oDatabase

    End If

End Sub

Private Sub cmdParse_Click()

    If txtQuery.Text = "" Then

        txtResults.Text = "Command Completed Successfully - No Query Offered"

    ElseIf txtQuery.SelLength > 0 Then

        ExecuteQuery cboDatabases.Text, txtQuery.SelText, 1, 1

    Else

        ExecuteQuery cboDatabases.Text, txtQuery.Text, 1, 1

    End If

End Sub

Private Sub cmdRun_Click()

    'If the user does not enter a query return polite message

    If txtQuery.Text = "" Then
    
        txtResults.Visible = True

        txtResults.Text = "Command Completed Successfully - No Query Offered"
        
        'If they have actively selected text and the output is text then

    ElseIf txtQuery.SelLength > 0 And optText.Value = True Then

        ExecuteQuery cboDatabases.Text, txtQuery.SelText, 2, 1, txtDelimiter.Text
        
        'If they have actively selected text and the output is grid then
        
    ElseIf txtQuery.SelLength > 0 And optText.Value = False Then

        ExecuteQuery cboDatabases.Text, txtQuery.SelText, 2, 2, txtDelimiter.Text

        'If they want to execute the window and the output is text then
    
    ElseIf optText.Value = True Then

        ExecuteQuery cboDatabases.Text, txtQuery.Text, 2, 1, txtDelimiter.Text
    
        'If they want to execute the window and the output is grid then
    
    ElseIf optText.Value = False Then

        ExecuteQuery cboDatabases.Text, txtQuery.Text, 2, 2, txtDelimiter.Text

    End If

End Sub

Private Sub Form_Load()

    optText.Value = True
    txtResults.Visible = True
    msf_grid.Visible = False

End Sub



