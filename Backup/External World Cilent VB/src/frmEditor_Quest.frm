VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest System"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6720
      TabIndex        =   93
      Top             =   7800
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quest Title"
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rewards"
         Height          =   180
         Index           =   2
         Left            =   3480
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tasks"
         Height          =   180
         Index           =   3
         Left            =   4920
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Requirements"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "General"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   7800
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save and Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   7800
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quest List"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstIndex 
         Height          =   7080
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraTasks 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tasks"
      Height          =   6495
      Left            =   3600
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   3375
         Begin VB.HScrollBar scrlEvent 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   91
            Top             =   4080
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.HScrollBar scrlNPC 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   49
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlItem 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   48
            Top             =   2280
            Width           =   3135
         End
         Begin VB.HScrollBar scrlAmount 
            Height          =   255
            Left            =   120
            Max             =   10
            TabIndex        =   47
            Top             =   5040
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   46
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox txtTaskSpeech 
            Height          =   270
            Left            =   120
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtTaskLog 
            Height          =   270
            Left            =   120
            MaxLength       =   200
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   43
            Top             =   3480
            Width           =   3135
         End
         Begin VB.CheckBox chkEnd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "End Quest Now?"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Left            =   120
            TabIndex        =   42
            Top             =   5400
            Width           =   1935
         End
         Begin VB.Label lblEvent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Event: None"
            Height          =   180
            Left            =   120
            TabIndex        =   92
            Top             =   3840
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label lblNPC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPC: None"
            Height          =   180
            Left            =   120
            TabIndex        =   56
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   55
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label lblAmount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   54
            Top             =   4800
            Width           =   795
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   3240
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label lblMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: None"
            Height          =   180
            Left            =   120
            TabIndex        =   53
            Top             =   2640
            Width           =   810
         End
         Begin VB.Label lblSpeech 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Task Speech:"
            Height          =   180
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblLog 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Task Log:"
            Height          =   180
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   750
         End
         Begin VB.Label lblResource 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resource: None"
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   3240
            Width           =   1200
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   3600
         TabIndex        =   31
         Top             =   600
         Width           =   2175
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Get from Event"
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   90
            Top             =   2520
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nothing"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Slay NPC"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Gather Items"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Talk to NPC"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Reach Map"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Give Item to NPC"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   1560
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Kill Player"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Train with Resource"
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   33
            Top             =   2040
            Width           =   1815
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Get from NPC"
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.HScrollBar scrlTotalTasks 
         Height          =   255
         Left            =   1680
         Max             =   10
         Min             =   1
         TabIndex        =   29
         Top             =   240
         Value           =   1
         Width           =   4095
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Task: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame fraRewards 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Rewards"
      Height          =   6495
      Left            =   3600
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cmbSkill 
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlSkillExp 
         Height          =   255
         LargeChange     =   50
         Left            =   3000
         TabIndex        =   94
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton cmdItemRewRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1680
         TabIndex        =   75
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ListBox lstItemRew 
         Height          =   2220
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   120
         List            =   "frmEditor_Quest.frx":0007
         TabIndex        =   59
         Top             =   1200
         Width           =   2775
      End
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         LargeChange     =   50
         Left            =   3000
         TabIndex        =   57
         Top             =   600
         Width           =   2775
      End
      Begin VB.HScrollBar scrlItemRew 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   27
         Top             =   600
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlItemRewValue 
         Height          =   135
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   26
         Top             =   960
         Value           =   1
         Width           =   2775
      End
      Begin VB.CommandButton cmdItemRew 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Skill:"
         Height          =   180
         Left            =   3000
         TabIndex        =   96
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Skill Exp Reward: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   95
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Experience Reward: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   58
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lblItemRew 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Item Reward: 0 (1)"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1425
      End
   End
   Begin VB.Frame fraRequirements 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requirements"
      Height          =   6495
      Left            =   3600
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.HScrollBar scrlReqSwitch 
         Height          =   255
         Left            =   120
         Max             =   70
         TabIndex        =   88
         Top             =   1680
         Width           =   2415
      End
      Begin VB.HScrollBar scrlReqClass 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   87
         Top             =   2400
         Value           =   1
         Width           =   2415
      End
      Begin VB.ListBox lstReqClass 
         Height          =   1140
         ItemData        =   "frmEditor_Quest.frx":0017
         Left            =   120
         List            =   "frmEditor_Quest.frx":0019
         TabIndex        =   86
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton cmdReqClassRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1320
         TabIndex        =   84
         Top             =   3960
         Width           =   1215
      End
      Begin VB.HScrollBar scrlReqItemValue 
         Height          =   135
         Left            =   2760
         Max             =   10
         Min             =   1
         TabIndex        =   81
         Top             =   840
         Value           =   1
         Width           =   3015
      End
      Begin VB.HScrollBar scrlReqItem 
         Height          =   255
         Left            =   2760
         Max             =   255
         TabIndex        =   80
         Top             =   480
         Value           =   1
         Width           =   3015
      End
      Begin VB.ListBox lstReqItem 
         Height          =   1860
         ItemData        =   "frmEditor_Quest.frx":001B
         Left            =   2760
         List            =   "frmEditor_Quest.frx":001D
         TabIndex        =   79
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmdReqItemRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   4560
         TabIndex        =   77
         Top             =   3000
         Width           =   1215
      End
      Begin VB.HScrollBar scrlReqLevel 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   21
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar scrlReqQuest 
         Height          =   255
         Left            =   120
         Max             =   70
         TabIndex        =   20
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdReqItem 
         Caption         =   "Update"
         Height          =   255
         Left            =   2760
         TabIndex        =   78
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmdReqClass 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblReqSwitch 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Switch: None"
         Height          =   180
         Left            =   120
         TabIndex        =   89
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label lblReqClass 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Class: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   83
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label lblReqItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Item Needed: 0 (1)"
         Height          =   180
         Left            =   2760
         TabIndex        =   82
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblReqLevel 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblReqQuest 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quest: None"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Frame fraGeneral 
      BackColor       =   &H00E0E0E0&
      Caption         =   "General"
      Height          =   6495
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdTakeItemRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   4560
         TabIndex        =   74
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdTakeItem 
         Caption         =   "Update"
         Height          =   255
         Left            =   3000
         TabIndex        =   73
         Top             =   6120
         Width           =   1575
      End
      Begin VB.ListBox lstTakeItem 
         Height          =   2040
         ItemData        =   "frmEditor_Quest.frx":001F
         Left            =   3000
         List            =   "frmEditor_Quest.frx":0021
         TabIndex        =   71
         Top             =   4080
         Width           =   2775
      End
      Begin VB.ListBox lstGiveItem 
         Height          =   2040
         ItemData        =   "frmEditor_Quest.frx":0023
         Left            =   120
         List            =   "frmEditor_Quest.frx":0025
         TabIndex        =   69
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtQuestLog 
         Height          =   270
         Left            =   1680
         MaxLength       =   200
         TabIndex        =   67
         Top             =   240
         Width           =   4095
      End
      Begin VB.CheckBox chkRepeat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Repetative Quest?"
         Height          =   255
         Left            =   3960
         TabIndex        =   64
         Top             =   600
         Width           =   1815
      End
      Begin VB.HScrollBar scrlTakeItem 
         Height          =   255
         Left            =   3000
         Max             =   255
         TabIndex        =   63
         Top             =   3480
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlTakeItemValue 
         Height          =   135
         Left            =   3000
         Max             =   10
         Min             =   1
         TabIndex        =   62
         Top             =   3840
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlGiveItemValue 
         Height          =   135
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   61
         Top             =   3840
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlGiveItem 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   60
         Top             =   3480
         Value           =   1
         Width           =   2775
      End
      Begin VB.TextBox txtSpeech 
         Height          =   270
         Index           =   1
         Left            =   120
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtSpeech 
         Height          =   270
         Index           =   2
         Left            =   120
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1800
         Width           =   5655
      End
      Begin VB.TextBox txtSpeech 
         Height          =   270
         Index           =   3
         Left            =   120
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2400
         Width           =   5655
      End
      Begin VB.CommandButton cmdGiveItemRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1680
         TabIndex        =   72
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdGiveItem 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Quest Log:"
         Height          =   180
         Left            =   120
         TabIndex        =   68
         Top             =   250
         Width           =   1485
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblTakeItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Take Item on the End: 0 (1)"
         Height          =   420
         Left            =   3000
         TabIndex        =   66
         Top             =   3000
         Width           =   2745
      End
      Begin VB.Label lblGiveItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Give Item on Start: 0 (1)"
         Height          =   420
         Left            =   120
         TabIndex        =   65
         Top             =   3000
         Width           =   2715
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lblQ1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Request Speech:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lblQ2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Meanwhile Speech:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblQ3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finished Speech:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////

Option Explicit
Private TempTask As Long

Private Sub cmbSkill_Click()
    If cmbSkill.ListIndex < 0 Then Exit Sub
    If EditorIndex < 1 Then Exit Sub
    Quest(EditorIndex).Skill = cmbSkill.ListIndex + 1
End Sub

Private Sub cmdSSave_Click()
    If Options.Debug Then On Error GoTo ErrHandler
    
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        QuestEditorOk False
    End If
    
    Exit Sub
ErrHandler:
    HandleError "cmdSSave", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Exit Sub
End Sub

Private Sub Form_Load()
    scrlTotalTasks.max = MAX_TASKS
    scrlNPC.max = MAX_NPCS
    scrlItem.max = MAX_ITEMS
    scrlMap.max = MAX_MAPS
    scrlResource.max = MAX_RESOURCES
    scrlAmount.max = MAX_INTEGER
    scrlReqLevel.max = MAX_LEVELS
    scrlReqQuest.max = MAX_QUESTS
    scrlReqItem.max = MAX_ITEMS
    scrlReqItemValue.max = MAX_INTEGER
    scrlGiveItem.max = MAX_ITEMS
    scrlGiveItemValue.max = MAX_INTEGER
    scrlTakeItem.max = MAX_ITEMS
    scrlTakeItemValue.max = MAX_INTEGER
    scrlExp.max = MAX_INTEGER 'Alatar v1.2
    scrlItemRew.max = MAX_ITEMS
    scrlItemRewValue.max = MAX_INTEGER
End Sub

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        QuestEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    QuestEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit
End Sub

Private Sub scrlSkillExp_Change()
    lblSkillExp.Caption = "Skill Exp Reward: " & scrlSkillExp.value
    Quest(EditorIndex).SkillExp = scrlSkillExp.value
End Sub

Private Sub lstIndex_Click()
    QuestEditorInit
End Sub

Private Sub scrlEvent_Change()
    If scrlEvent.value > 0 Then
        lblEvent.Caption = "Event: " & scrlEvent.value '& "-" & Map.Events(scrlEvent.Value).Name
    Else
        lblEvent.Caption = "Event: None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.value).Event = scrlEvent.value
End Sub

Private Sub scrlTotalTasks_Change()
    Dim I As Long
    
    lblSelected = "Selected Task: " & scrlTotalTasks.value
    
    LoadTask EditorIndex, scrlTotalTasks.value
End Sub

Private Sub optTask_Click(Index As Integer)
    Quest(EditorIndex).Task(scrlTotalTasks.value).order = Index
    LoadTask EditorIndex, scrlTotalTasks.value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long
    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtQuestLog_Change()
    Quest(EditorIndex).QuestLog = Trim$(txtQuestLog.text)
End Sub

Private Sub txtTaskLog_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.value).TaskLog = Trim$(txtTaskLog.text)
End Sub

Private Sub chkRepeat_Click()
    If chkRepeat.value = 1 Then
        Quest(EditorIndex).Repeat = 1
    Else
        Quest(EditorIndex).Repeat = 0
    End If
End Sub

Private Sub scrlReqLevel_Change()
    lblReqLevel.Caption = "Level: " & scrlReqLevel.value
    Quest(EditorIndex).RequiredLevel = scrlReqLevel.value
End Sub

Private Sub scrlReqQuest_Change()
    If Not scrlReqQuest.value = 0 Then
        If Not Trim$(Quest(scrlReqQuest.value).name) = "" Then
            lblReqQuest.Caption = "Quest: " & Trim$(Quest(scrlReqQuest.value).name)
        Else
            lblReqQuest.Caption = "Quest: None"
        End If
    Else
        lblReqQuest.Caption = "Quest: None"
    End If
    Quest(EditorIndex).RequiredQuest = scrlReqQuest.value
End Sub

'Alatar v1.2

Private Sub scrlReqItem_Change()
    If scrlReqItem.value > 0 Then
        lblReqItem.Caption = "Item Needed: " & Trim$(Item(scrlReqItem.value).name) & " (" & scrlReqItemValue.value & ")"
    Else
        scrlReqItemValue.value = 1
        lblReqItem.Caption = "Item Needed: None (" & scrlReqItemValue.value & ")"
    End If
End Sub

Private Sub scrlReqItemValue_Change()
    If scrlReqItem.value > 0 Then lblReqItem.Caption = "Item Needed: " & Trim$(Item(scrlReqItem.value).name) & " (" & scrlReqItemValue.value & ")"
End Sub

Private Sub cmdReqItem_Click()
    Dim Index As Long
    
    Index = lstReqItem.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlReqItem.value < 1 Or scrlReqItem.value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlReqItem.value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).RequiredItem(Index).Item = scrlReqItem.value
    Quest(EditorIndex).RequiredItem(Index).value = scrlReqItemValue.value
    UpdateQuestRequirementItems
End Sub

Private Sub cmdReqItemRemove_Click()
    Dim Index As Long
    
    Index = lstReqItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RequiredItem(Index).Item = 0
    Quest(EditorIndex).RequiredItem(Index).value = 1
    UpdateQuestRequirementItems
End Sub

Private Sub scrlReqClass_Change()
    If scrlReqClass.value < 1 Or scrlReqClass.value > Max_Classes Then
        lblReqClass.Caption = "Class: 0"
    Else
        lblReqClass.Caption = "Class: " & scrlReqClass.value & " (" & Trim$(Class(scrlReqClass.value).name) & ")"
    End If
End Sub

Private Sub cmdReqClass_Click()
    Dim Index As Long
    
    Index = lstReqClass.ListIndex + 1 'the selected class
    If Index = 0 Then Exit Sub
    If scrlReqClass.value < 1 Or scrlReqClass.value > Max_Classes Then Exit Sub
    If Trim$(Class(scrlReqClass.value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).RequiredClass(Index) = scrlReqClass.value
    UpdateQuestClass
End Sub

Private Sub cmdReqClassRemove_Click()
    Dim Index As Long
    
    Index = lstReqClass.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RequiredClass(Index) = 0
    UpdateQuestClass
End Sub

'/Alatar v1.2

Private Sub scrlExp_Change()
    lblExp = "Experience Reward: " & scrlExp.value
    Quest(EditorIndex).RewardExp = scrlExp.value
End Sub

Private Sub scrlItemRew_Change()
    If scrlItemRew.value > 0 Then
        lblItemRew.Caption = "Item Reward: " & Trim$(Item(scrlItemRew.value).name) & " (" & scrlItemRewValue.value & ")"
    Else
        lblItemRew.Caption = "Item Reward: None (" & scrlItemRewValue.value & ")"
    End If
End Sub

Private Sub scrlItemRewValue_Change()
    If scrlItemRew.value > 0 Then lblItemRew.Caption = "Item Reward: " & Trim$(Item(scrlItemRew.value).name) & " (" & scrlItemRewValue.value & ")"
End Sub

'Alatar v1.2
Private Sub cmdItemRew_Click()
    Dim Index As Long
    
    Index = lstItemRew.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlItemRew.value < 1 Or scrlItemRew.value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlItemRew.value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).RewardItem(Index).Item = scrlItemRew.value
    Quest(EditorIndex).RewardItem(Index).value = scrlItemRewValue.value
    UpdateQuestRewardItems
End Sub

Private Sub cmdItemRewRemove_Click()
    Dim Index As Long
    
    Index = lstItemRew.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RewardItem(Index).Item = 0
    Quest(EditorIndex).RewardItem(Index).value = 1
    UpdateQuestRewardItems
End Sub
'/Alatar v1.2

Private Sub txtSpeech_Change(Index As Integer)
    Quest(EditorIndex).Speech(Index) = Trim$(txtSpeech(Index).text)
End Sub

Private Sub txtTaskSpeech_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.value).Speech = Trim$(txtTaskSpeech.text)
End Sub

'Alatar v1.2
Private Sub scrlGiveItem_Change()
    If scrlGiveItem.value > 0 Then
        lblGiveItem = "Give Item on Start: " & Trim$(Item(scrlGiveItem.value).name) & " (" & scrlGiveItemValue.value & ")"
    Else
        scrlGiveItemValue.value = 1
        lblGiveItem = "Give Item on Start: None (" & scrlGiveItemValue.value & ")"
    End If
End Sub

Private Sub scrlGiveItemValue_Change()
    If scrlGiveItem.value > 0 Then lblGiveItem = "Give Item on Start: " & Trim$(Item(scrlGiveItem.value).name) & " (" & scrlGiveItemValue.value & ")"
End Sub

Private Sub cmdGiveItem_Click()
    Dim Index As Long
    
    Index = lstGiveItem.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlGiveItem.value < 1 Or scrlGiveItem.value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlGiveItem.value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).GiveItem(Index).Item = scrlGiveItem.value
    Quest(EditorIndex).GiveItem(Index).value = scrlGiveItemValue.value
    UpdateQuestGiveItems
End Sub

Private Sub cmdGiveItemRemove_Click()
    Dim Index As Long
    
    Index = lstGiveItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).GiveItem(Index).Item = 0
    Quest(EditorIndex).GiveItem(Index).value = 1
    UpdateQuestGiveItems
End Sub

Private Sub scrlTakeItem_Change()
    If scrlTakeItem.value > 0 Then
        lblTakeItem = "Take Item on the End: " & Trim$(Item(scrlTakeItem.value).name) & " (" & scrlTakeItemValue.value & ")"
    Else
        scrlTakeItemValue.value = 1
        lblTakeItem = "Take Item on the End: None (" & scrlTakeItemValue.value & ")"
    End If
End Sub

Private Sub scrlTakeItemValue_Change()
    If scrlTakeItem.value > 0 Then lblTakeItem = "Take Item on the End: " & Trim$(Item(scrlTakeItem.value).name) & " (" & scrlTakeItemValue.value & ")"
End Sub

Private Sub cmdTakeItem_Click()
    Dim Index As Long
    
    Index = lstTakeItem.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlTakeItem.value < 1 Or scrlTakeItem.value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlTakeItem.value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).TakeItem(Index).Item = scrlTakeItem.value
    Quest(EditorIndex).TakeItem(Index).value = scrlTakeItemValue.value
    UpdateQuestTakeItems
End Sub

Private Sub cmdTakeItemRemove_Click()
    Dim Index As Long
    
    Index = lstTakeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).TakeItem(Index).Item = 0
    Quest(EditorIndex).TakeItem(Index).value = 1
    UpdateQuestTakeItems
End Sub
'/Alatar v1.2

Private Sub scrlAmount_Change()
    lblAmount.Caption = "Amount: " & scrlAmount.value
    Quest(EditorIndex).Task(scrlTotalTasks.value).Amount = scrlAmount.value
End Sub

Private Sub scrlNPC_Change()
    If scrlNPC.value > 0 Then
        lblNPC.Caption = "NPC: " & scrlNPC.value & "-" & Trim$(NPC(scrlNPC.value).name)
    Else
        lblNPC.Caption = "NPC: None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.value).NPC = scrlNPC.value
End Sub

Private Sub scrlItem_Change()
    If scrlItem.value > 0 Then
        lblItem.Caption = "Item: " & Trim$(Item(scrlItem.value).name)
    Else
        lblItem.Caption = "Item: None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.value).Item = scrlItem.value
End Sub

Private Sub scrlMap_Change()
    If scrlMap.value > 0 Then
        lblMap.Caption = "Map: " & scrlMap.value
    Else
        lblMap.Caption = "Map: None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.value).Map = scrlMap.value
End Sub

Private Sub scrlResource_Change()
    If scrlResource.value > 0 Then
        lblResource.Caption = "Resource: " & scrlResource.value & "-" & Trim$(Resource(scrlResource.value).name)
    Else
        lblResource.Caption = "Resource: None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.value).Resource = scrlResource.value
End Sub

Private Sub chkEnd_Click()
    If chkEnd.value = 1 Then
        Quest(EditorIndex).Task(scrlTotalTasks.value).QuestEnd = True
    Else
        Quest(EditorIndex).Task(scrlTotalTasks.value).QuestEnd = False
    End If
End Sub

Private Sub optShowFrame_Click(Index As Integer)
    fraGeneral.Visible = False
    fraRequirements.Visible = False
    fraRewards.Visible = False
    fraTasks.Visible = False
    
    If optShowFrame(Index).value = True Then
        Select Case Index
            Case 0
                fraGeneral.Visible = True
            Case 1
                fraRequirements.Visible = True
            Case 2
                fraRewards.Visible = True
            Case 3
                fraTasks.Visible = True
        End Select
    End If
End Sub
