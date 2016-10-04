VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   3600
   ClientTop       =   4485
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Attempting Login to..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   12465
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Show , frmLogin

End Sub

Public Sub InformUser(ByVal strMessage As String)

    lblInfo.Caption = strMessage
    'make sure the display refreshes
    DoEvents

End Sub

