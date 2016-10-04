VERSION 5.00
Begin VB.Form frmReplicate 
   Caption         =   "Replicate!"
   ClientHeight    =   4065
   ClientLeft      =   4740
   ClientTop       =   4350
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   3765
   Begin VB.CommandButton cmdSubscriber 
      Caption         =   "Install Subscriber"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1260
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdInstallMerge 
      Caption         =   "Install Merge"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "&Uninstall Repl"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Install Distributor"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   3495
   End
End
Attribute VB_Name = "frmReplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bConnected As Boolean
    
Dim objSrv1 As New SQLDMO.SQLServer
Dim objSrv2 As New SQLDMO.SQLServer

Dim objReplication As New SQLDMO.Replication
Dim objDistribDB As New SQLDMO.DistributionDatabase
Dim objPublisher As New SQLDMO.DistributionPublisher
Dim oPub As SQLDMO.Publisher
Dim mpsMergesubscription As SQLDMO.MergeSubscription
Dim subSubscriber As SQLDMO.RegisteredSubscriber
Dim bMergeExists As Boolean




Private Sub cmdGo_Click()

    'first let's check we have a publisher

    bConnected = False
    bMergeExists = False
    
        'log on to the server
    
        With objSrv1
            .LoginSecure = True
            .Connect "ALLAN" ' The publisher and distributor
        End With
    
    Debug.Print "Connected to Server Allan"


    bConnected = True
    
    'if the distributor is already available then
    'we can skip installing it.  And vice versa

    If objSrv1.Replication.Distributor.DistributorAvailable = False Then
    objDistribDB.Name = "distribution" ' set the distribution database name
        objSrv1.Replication.Distributor.DistributionDatabases.Add objDistribDB ' add the distribution database to the collection
        With objSrv1.Replication.Distributor
            .DistributionServer = objSrv1.TrueName ' set the distribution server to the true name of the local server
            .Install
        End With
    End If
   
   Debug.Print "Distributor Done"

    ' Now let's add a publisher
    
    If objSrv1.Replication.Distributor.IsDistributionPublisher = False Then

    objPublisher.Name = objSrv1.TrueName   ' set the publisher to be the local server
    objPublisher.DistributionDatabase = "distribution" 'Which database are we going to use for distribution
    objPublisher.DistributionWorkingDirectory = App.Path & "\ReplWorkingDir" 'Set the working directory
    objSrv1.Replication.Distributor.DistributionPublishers.Add objPublisher
    objPublisher.ThirdParty = True

    End If
    
    Debug.Print "Publisher Done"

    'once we've added the publisher and the distributor
    'we need to create a publication

        Dim objReplicationDB As SQLDMO.ReplicationDatabase
        Dim objMergeReplication As New SQLDMO.MergePublication
        Dim objMergeArticle1 As New SQLDMO.MergeArticle
            
        ' shows databases available for replication
        'For Each objReplicationDB In objSrv1.Replication.ReplicationDatabases
            'Debug.Print objReplicationDB.Name
        'Next
    
        If objSrv1.Replication.ReplicationDatabases("Northwind").EnableMergePublishing = False Then
            objSrv1.Replication.ReplicationDatabases("Northwind").EnableMergePublishing = True
        End If
        
        Debug.Print "Merge Publishing Enabled"
        
        ' Add a merge publication
        
        For Each objMergeReplication In objSrv1.Replication.ReplicationDatabases("Northwind").MergePublications
            If objMergeReplication.Name = "NewMergeReplication" Then
                bMergeExists = True
            End If
        Next
        
        Debug.Print "Check to see if publication exists Done"
        
        If bMergeExists = False Then
        objMergeReplication.Name = "NewMergeReplication"
        objMergeReplication.PublicationAttributes = SQLDMOPubAttrib_AllowPull + SQLDMOPubAttrib_AllowPush
        objSrv1.Replication.ReplicationDatabases("Northwind").MergePublications.Add objMergeReplication
    
        'add an article
        objMergeArticle1.Name = "Customers"
        objMergeArticle1.SourceObjectName = "Customers"
        objMergeArticle1.SourceObjectOwner = "dbo"
        objMergeArticle1.ColumnTracking = True ' set columntracking on to improve accuracy
        objMergeArticle1.Status = SQLDMOArtStat_Active
        objMergeReplication.MergeArticles.Add objMergeArticle1
       
       End If
       
       Debug.Print "Publication added and Article added to publication"
       
       
       Set subSubscriber = New SQLDMO.RegisteredSubscriber
       
       subSubscriber.Name = "ALLAN\Am3"
    
       
       objSrv1.Replication.Publisher.RegisteredSubscribers.Add subSubscriber
       
       Debug.Print "Subscriber Added"
       
       


       Set mpsMergesubscription = New SQLDMO.MergeSubscription

       With mpsMergesubscription

        .Subscriber = "ALLAN\Am3"
        .SubscriptionDB = "Northwind"
        '.MergeSchedule = 'Add a schedule object here
    

        End With


    objSrv1.Replication.ReplicationDatabases("Northwind").MergePublications("NewMergeReplication").MergeSubscriptions.Add mpsMergesubscription


    
    
    
    
End Sub
