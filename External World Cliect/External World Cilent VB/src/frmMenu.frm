VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMenu.lblBlank(0).Visible = YesNo"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMenu.frx":324A
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   5280
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   36
      Top             =   4680
      Width           =   480
   End
   Begin VB.TextBox txtCName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      MaxLength       =   12
      TabIndex        =   30
      Top             =   4080
      Width           =   2775
   End
   Begin VB.OptionButton optFemale 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   5280
      Width           =   1095
   End
   Begin VB.OptionButton optMale 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   5295
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox picCharacter 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   13680
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   438
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   6570
      Begin VB.PictureBox picSprite_copy 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   4800
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   16
         Top             =   1680
         Width           =   480
      End
      Begin VB.ComboBox cmbClass_copy 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.OptionButton optMale_copy 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   7
         Top             =   2295
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optFemale_copy 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtCName_copy 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   0
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblSprite_copy 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ Change Sprite ]"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblBlank_copy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblBlank_copy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   13
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBlank_copy 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblCAccept_copy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkPass 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      Top             =   6000
      Width           =   195
   End
   Begin VB.TextBox txtLPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   6240
      MaxLength       =   20
      PasswordChar    =   "�"
      TabIndex        =   2
      Top             =   5280
      Width           =   3975
   End
   Begin VB.TextBox txtLUser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   6240
      MaxLength       =   12
      TabIndex        =   1
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox txtRUser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2880
      MaxLength       =   12
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtRPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   20
      PasswordChar    =   "�"
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtRPass2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   2880
      MaxLength       =   20
      PasswordChar    =   "�"
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   13800
      ScaleHeight     =   3195
      ScaleWidth      =   6570
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   6570
   End
   Begin VB.Label lblRegister1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2880
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblRegister 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Register"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2880
      TabIndex        =   37
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblCAccept 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   34
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblClass 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   33
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lblGender 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   32
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblSprite 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[ Change Sprite ]"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2640
      TabIndex        =   31
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label lblBlank 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "     Save Password?"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   12
      Left            =   6480
      TabIndex        =   26
      Top             =   6000
      Width           =   1965
   End
   Begin VB.Label lblLAccept 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8640
      TabIndex        =   25
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblBlank 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   300
      Index           =   3
      Left            =   14640
      TabIndex        =   24
      Top             =   7080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblBlank 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   300
      Index           =   0
      Left            =   14640
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblBlank 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   1560
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblBlank 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   1560
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label txtRAccept 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ACCEPT"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBlank 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Retype:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   1560
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   4
      Left            =   12240
      Picture         =   "frmMenu.frx":16FB8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   3
      Left            =   12240
      Picture         =   "frmMenu.frx":1AB3F
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   2
      Left            =   12240
      Picture         =   "frmMenu.frx":1E93E
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   1
      Left            =   12240
      Picture         =   "frmMenu.frx":228BD
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblNews 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   12360
      TabIndex        =   17
      Top             =   3840
      Width           =   2535
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Function MakeWindowedControlTransparent(ctlControl As Control) As Long
    Dim result As Long
    ctlControl.Visible = False
    result = SetWindowLong(ctlControl.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    ctlControl.Visible = True ' Use the visible property as a quick VB way of forcing a repaint with the new style
    MakeWindowedControlTransparent = result
End Function

Private Sub cmbClass_Click()
    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
End Sub

Private Sub cmbClass_KeyPress(KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn Then Exit Sub
    If cmbClass.text <> "cmbClass" And _
       Not cmbClass.ListIndex < 0 Then
        optMale.SetFocus
        optMale.value = True
    End If
End Sub

Private Sub Form_Load()
    Dim tmpTxt As String, tmpArray() As String, I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    
lblName.Visible = False
lblClass.Visible = False
lblGender.Visible = False
txtCName.Visible = False
lblSprite.Visible = False
cmbClass.Visible = False
optMale.Visible = False
optFemale.Visible = False
picSprite.Visible = True
lblCAccept.Visible = False
picSprite.Visible = False
txtLUser.Visible = True
txtLPass.Visible = True
lblBlank(12).Visible = True
chkPass.Visible = True
lblLAccept.Visible = True



    ' general menu stuff
    Me.Caption = Options.Game_Name
    MAX_SKILLS = 4
    ReDim Skill(1 To MAX_SKILLS)
    
    ' Set info texts
    PlayerInfoText(1) = "~Player Info~"
    PlayerInfoText(2) = "Level:        "
    PlayerInfoText(3) = "Strength:     "
    PlayerInfoText(4) = "Endurance:    "
    PlayerInfoText(5) = "Intelligence: "
    PlayerInfoText(6) = "Agility:      "
    PlayerInfoText(7) = "WillPower:    "
    
    'reload dx8 variabls
    frmMain.Width = 12090
    frmMain.Height = 9420
    Call LoadDX8Vars
    
    ' load news
    Open App.Path & "\data files\news.txt" For Input As #1
        Line Input #1, tmpTxt
    Close #1
    ' split breaks
    tmpArray() = Split(tmpTxt, "<br />")
    lblNews.Caption = vbNullString
    OpeningBook = True
    For I = 0 To UBound(tmpArray)
        lblNews.Caption = lblNews.Caption & tmpArray(I) & vbNewLine
    Next

    ' Load the username + pass
    txtLUser.text = Trim$(Options.Username)
    If Options.savePass = 1 Then
        txtLPass.text = Trim$(Options.Password)
        chkPass.value = Options.savePass
    End If
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not EnteringGame Then DestroyGame
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            'If Not picLogin.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                Show_Login Not txtLUser.Visible
                Show_Register False
                picCharacter.Visible = False
                
                
lblName.Visible = False
txtCName.Visible = False
lblBlank(8).Visible = False
txtRUser.Visible = False
lblBlank(9).Visible = False
txtRPass.Visible = False
lblClass.Visible = False
lblSprite.Visible = False
cmbClass.Visible = False
txtRPass2.Visible = False
picSprite.Visible = False
lblBlank(11).Visible = False
lblGender.Visible = False
optMale.Visible = False
optFemale.Visible = False
lblCAccept.Visible = False

                
lblName.Visible = False
txtCName.Visible = False
lblSprite.Visible = False
lblClass.Visible = False
cmbClass.Visible = False
picSprite.Visible = False
lblGender.Visible = False
optMale.Visible = False
optFemale.Visible = False
lblCAccept.Visible = False
                
                If txtLUser.Visible Then
                    txtLPass.SetFocus
                    txtLPass.SelStart = Len(txtLPass.text)
                End If
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        Case 2
            'If Not picRegister.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                Show_Login False
                Show_Register Not txtRUser.Visible
                picCharacter.Visible = False
                
lblName.Visible = False
txtCName.Visible = False
lblBlank(8).Visible = False
txtRUser.Visible = False
lblBlank(9).Visible = False
txtRPass.Visible = False
lblClass.Visible = False
lblSprite.Visible = False
cmbClass.Visible = False
txtRPass2.Visible = False
picSprite.Visible = False
lblBlank(11).Visible = False
lblGender.Visible = False
optMale.Visible = False
optFemale.Visible = False
lblCAccept.Visible = False
                
                If txtRUser.Visible Then
                    txtRUser.SetFocus
                End If
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        Case 3
            'If Not picCredits.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                'picCredits.Visible = Not picCredits.Visible
                Show_Login False
                Show_Register False
                picCharacter.Visible = False
                
lblName.Visible = False
txtCName.Visible = False
lblBlank(8).Visible = False
txtRUser.Visible = False
lblBlank(9).Visible = False
txtRPass.Visible = False
lblClass.Visible = False
lblSprite.Visible = False
cmbClass.Visible = False
txtRPass2.Visible = False
picSprite.Visible = False
lblBlank(11).Visible = False
lblGender.Visible = False
optMale.Visible = False
optFemale.Visible = False
lblCAccept.Visible = False
                
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        Case 4
            Call DestroyGame
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    changeButtonState_Menu Index, 2 ' clicked
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    If Not MenuButton(Index).state = 2 Then ' make sure we're not clicking
        changeButtonState_Menu Index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Menu = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Menu = Index
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
        
    ' reset all buttons
    resetButtons_Menu -1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_MouseUp", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub lblBlank_Click(Index As Integer)
    chkPass.value = Abs(Not CBool(chkPass.value))
End Sub

Private Sub lblLAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lblLAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblRegister_Click()
lblBlank(8).Visible = True
lblBlank(9).Visible = True
lblBlank(11).Visible = True
txtRUser.Visible = True
txtRPass.Visible = True
txtRPass2.Visible = True
txtRAccept.Visible = True
lblRegister1.Visible = True
lblRegister.Visible = False
End Sub

Private Sub lblRegister1_Click()
lblBlank(8).Visible = False
lblBlank(9).Visible = False
lblBlank(11).Visible = False
txtRUser.Visible = False
txtRPass.Visible = False
txtRPass2.Visible = False
txtRAccept.Visible = False
lblRegister1.Visible = False
lblRegister.Visible = True
End Sub

Private Sub lblSprite_Click()
Dim spritecount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If optMale.value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite >= spritecount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lblSprite_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optFemale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optFemale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optFemale_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lblCAccept_Click
    End If
    KeyAscii = 0
End Sub

Private Sub optMale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optMale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMale_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lblCAccept_Click
    End If
    KeyAscii = 0
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picCharacter_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCredits_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picCredits_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picLogin_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picMain_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picRegister_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub




Private Sub txtCName_KeyPress(KeyAscii As Integer)
    If Not Len(txtCName.text) > 0 Then Exit Sub
    If Not KeyAscii = vbKeyReturn Then Exit Sub
    
    cmbClass.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtLPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(txtLPass.text) > 0 Then
            If Len(txtLUser.text) > 0 Then
                Call lblLAccept_Click
            End If
        End If
        
        KeyAscii = 0
    End If
End Sub

Private Sub txtLUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(txtLUser.text) > 0 Then
            txtLPass.SetFocus
            txtLPass.SelStart = Len(txtLPass.text)
        End If
        KeyAscii = 0
    End If
End Sub

' Register
Private Sub txtRAccept_Click()
    Dim name As String
    Dim Password As String
    Dim PasswordAgain As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("Passwords don't match.")
            Exit Sub
        End If

        If Not isStringLegal(name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
        
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtRAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' New Char
Private Sub lblCAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MenuState(MENU_STATE_ADDCHAR)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lblCAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
