VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Autoupdater"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":00D2
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTimeout 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picprogress 
      Appearance      =   0  'Flat
      BackColor       =   &H0041E7D2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      ScaleHeight     =   195
      ScaleWidth      =   6375
      TabIndex        =   4
      Top             =   5040
      Width           =   6375
   End
   Begin SHDocVwCtl.WebBrowser PicNews 
      Height          =   3855
      Left            =   90
      TabIndex        =   2
      Top             =   600
      Width           =   8535
      ExtentX         =   15055
      ExtentY         =   6800
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin InetCtlsObjects.Inet inetDownload 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   3
      Top             =   4710
      Width           =   3255
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   6840
      TabIndex        =   0
      Top             =   4680
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim Filename As String

    ' Load a GUI image, if it does not exist.. Exit out of the program.
    Form_LoadPicture (App.Path & "\Data Files\graphics\gui\updater\update.jpg")
    
    'Load the webpage
    PicNews.Navigate NewsURL
    
    Me.Show
    
    picprogress.Width = 0
    lblstatus.Caption = "Welcome to the " & GameName & " Launcher."
    
    'Everything's loaded.. Check for updates!
    CheckVersion
End Sub

Private Sub lblConnect_Click()
    If UpToDate <> 0 Then
        If FileExist(App.Path & "\" & ClientName) Then
            Shell App.Path & "\" & ClientName, vbNormalFocus
        Else
            MsgBox "Could not locate " & ClientName, vbCritical
        End If
        DestroyUpdater
    End If
End Sub

Private Sub lblExit_Click()
    DestroyUpdater
End Sub


Private Sub lblwebsite_Click()
    Shell "explorer.exe " & GameWebsite
End Sub

Private Sub tmrTimeout_Timer()
    MsgBox "The connection to the update server could not be made.", vbCritical, "Connection Error"
    DestroyUpdater
End Sub
