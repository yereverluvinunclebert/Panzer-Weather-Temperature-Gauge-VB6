VERSION 5.00
Object = "{BCE37951-37DF-4D69-A8A3-2CFABEE7B3CC}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form widgetPrefs 
   AutoRedraw      =   -1  'True
   Caption         =   "Panzer Weather Temperature Gauge Preferences"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8880
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10693.53
   ScaleMode       =   0  'User
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame fraDevelopmentButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   5490
      TabIndex        =   40
      Top             =   0
      Width           =   1065
      Begin VB.Label lblDevelopment 
         Caption         =   "Development"
         Height          =   240
         Left            =   45
         TabIndex        =   41
         Top             =   855
         Width           =   960
      End
      Begin VB.Image imgDevelopment 
         Height          =   600
         Left            =   150
         Picture         =   "frmPrefs.frx":10CA
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgDevelopmentClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1682
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Timer positionTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1170
      Top             =   9690
   End
   Begin VB.CheckBox chkEnableResizing 
      Caption         =   "Enable Corner Resize"
      Height          =   210
      Left            =   3240
      TabIndex        =   95
      Top             =   10125
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fraAboutButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   7695
      TabIndex        =   65
      Top             =   0
      Width           =   975
      Begin VB.Label lblAbout 
         Caption         =   "About"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   66
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgAbout 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1A08
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgAboutClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":1F90
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraConfigButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   1215
      TabIndex        =   42
      Top             =   -15
      Width           =   930
      Begin VB.Label lblConfig 
         Caption         =   "Config."
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   43
         Top             =   855
         Width           =   510
      End
      Begin VB.Image imgConfig 
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":247B
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgConfigClicked 
         Height          =   600
         Left            =   165
         Picture         =   "frmPrefs.frx":2A5A
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraPositionButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   4410
      TabIndex        =   38
      Top             =   0
      Width           =   930
      Begin VB.Label lblPosition 
         Caption         =   "Position"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   39
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgPosition 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":2F5F
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgPositionClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":3530
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save the changes you have made to the preferences"
      Top             =   10020
      Width           =   1320
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Open the help utility"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.Frame fraSoundsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   3315
      TabIndex        =   11
      Top             =   -15
      Width           =   930
      Begin VB.Label lblSounds 
         Caption         =   "Sounds"
         Height          =   240
         Left            =   210
         TabIndex        =   12
         Top             =   870
         Width           =   615
      End
      Begin VB.Image imgSounds 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":38CE
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgSoundsClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":3E8D
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Timer themeTimer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   645
      Top             =   9705
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Close the utility"
      Top             =   10020
      Width           =   1320
   End
   Begin VB.Frame fraWindowButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   6615
      TabIndex        =   4
      Top             =   0
      Width           =   975
      Begin VB.Label lblWindow 
         Caption         =   "Window"
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   855
         Width           =   615
      End
      Begin VB.Image imgWindow 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":435D
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Image imgWindowClicked 
         Height          =   600
         Left            =   160
         Picture         =   "frmPrefs.frx":4827
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
   End
   Begin VB.Frame fraFontsButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   930
      Begin VB.Label lblFonts 
         Caption         =   "Fonts"
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   855
         Width           =   510
      End
      Begin VB.Image imgFonts 
         Height          =   600
         Left            =   180
         Picture         =   "frmPrefs.frx":4BD3
         Stretch         =   -1  'True
         Top             =   195
         Width           =   600
      End
      Begin VB.Image imgFontsClicked 
         Height          =   600
         Left            =   180
         Picture         =   "frmPrefs.frx":5129
         Stretch         =   -1  'True
         Top             =   195
         Width           =   600
      End
   End
   Begin VB.Frame fraGeneralButton 
      Height          =   1140
      Left            =   240
      TabIndex        =   0
      Top             =   -15
      Width           =   930
      Begin VB.Image imgGeneral 
         Height          =   600
         Left            =   180
         Picture         =   "frmPrefs.frx":55C2
         Stretch         =   -1  'True
         Top             =   225
         Width           =   600
      End
      Begin VB.Label lblGeneral 
         Caption         =   "General"
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   855
         Width           =   705
      End
      Begin VB.Image imgGeneralClicked 
         Height          =   600
         Left            =   165
         Stretch         =   -1  'True
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position && Size"
      Height          =   8355
      Left            =   240
      TabIndex        =   44
      Top             =   1230
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraPositionInner 
         BorderStyle     =   0  'None
         Height          =   7755
         Left            =   150
         TabIndex        =   45
         Top             =   300
         Width           =   7680
         Begin VB.Frame fraGaugePosition 
            Caption         =   "Temperature Gauge "
            Height          =   3435
            Left            =   255
            TabIndex        =   157
            Top             =   4305
            Width           =   4410
            Begin VB.TextBox txtLandscapeHoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   166
               Top             =   1170
               Width           =   960
            End
            Begin VB.TextBox txtLandscapeVoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   165
               Top             =   1575
               Width           =   960
            End
            Begin VB.TextBox txtPortraitHoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   164
               Top             =   2460
               Width           =   960
            End
            Begin VB.TextBox txtPortraitVoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   163
               Top             =   2865
               Width           =   960
            End
            Begin VB.ComboBox cmbPortraitLocked 
               Height          =   315
               Left            =   1965
               Style           =   2  'Dropdown List
               TabIndex        =   160
               ToolTipText     =   "Choose the alarm sound."
               Top             =   2025
               Width           =   2160
            End
            Begin VB.ComboBox cmbLandscapeLocked 
               Height          =   315
               Left            =   1965
               Style           =   2  'Dropdown List
               TabIndex        =   159
               ToolTipText     =   "Choose the alarm sound."
               Top             =   750
               Width           =   2160
            End
            Begin VB.CheckBox chkPreventDragging 
               Caption         =   "This gauge Locked. *"
               Height          =   225
               Left            =   1980
               TabIndex        =   158
               ToolTipText     =   "Checking this box turns off the ability to drag the program with the mouse, locking it in position."
               Top             =   360
               Width           =   2250
            End
            Begin VB.Label lblPosition 
               Caption         =   "Locked in Portrait :"
               Height          =   375
               Index           =   11
               Left            =   405
               TabIndex        =   162
               Tag             =   "lblAlarmSound"
               Top             =   2070
               Width           =   2040
            End
            Begin VB.Label lblPosition 
               Caption         =   "Locked in Landscape :"
               Height          =   435
               Index           =   13
               Left            =   165
               TabIndex        =   161
               Tag             =   "lblAlarmSound"
               Top             =   795
               Width           =   2115
            End
         End
         Begin VB.ComboBox cmbGaugeType 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   2415
            Width           =   2430
         End
         Begin VB.ComboBox cmbWidgetPosition 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   153
            ToolTipText     =   "Choose the alarm sound."
            Top             =   1035
            Width           =   2430
         End
         Begin VB.Frame fraPositionBalloonBox 
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   120
            TabIndex        =   149
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cmbAspectHidden 
               Height          =   315
               Left            =   2145
               Style           =   2  'Dropdown List
               TabIndex        =   150
               ToolTipText     =   "Choose the alarm sound."
               Top             =   0
               Width           =   2430
            End
            Begin VB.Label lblPosition 
               Caption         =   "Aspect Ratio Hidden Mode :"
               Height          =   375
               Index           =   3
               Left            =   0
               TabIndex        =   152
               Tag             =   "lblAlarmSound"
               Top             =   45
               Width           =   2145
            End
            Begin VB.Label lblPosition 
               Caption         =   "Tablets only. Don't fiddle with this unless you really know what you are doing. Read the help before fiddling!"
               Height          =   765
               Index           =   6
               Left            =   2145
               TabIndex        =   151
               Tag             =   "lblAlarmSoundDesc"
               Top             =   420
               Width           =   5370
            End
         End
         Begin vb6projectCCRSlider.Slider sliGaugeSize 
            Height          =   390
            Left            =   2085
            TabIndex        =   167
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   2910
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   5
            Max             =   220
            Value           =   5
            TickFrequency   =   5
            LargeChange     =   5
            SelStart        =   5
         End
         Begin VB.Label lblPosition 
            Caption         =   "Select the Gauge first :"
            Height          =   375
            Index           =   2
            Left            =   525
            TabIndex        =   176
            Tag             =   "lblAlarmSound"
            Top             =   2460
            Width           =   1935
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "5"
            Height          =   315
            Index           =   0
            Left            =   2235
            TabIndex        =   175
            Top             =   3390
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "220 (%)"
            Height          =   315
            Index           =   5
            Left            =   5550
            TabIndex        =   174
            Top             =   3405
            Width           =   735
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "90"
            Height          =   315
            Index           =   2
            Left            =   3465
            TabIndex        =   173
            Top             =   3405
            Width           =   420
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Gauge Size :"
            Height          =   315
            Index           =   1
            Left            =   1125
            TabIndex        =   172
            Top             =   2955
            Width           =   975
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel. Immediate. *"
            Height          =   555
            Index           =   2
            Left            =   2235
            TabIndex        =   171
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   3720
            Width           =   3810
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "50"
            Height          =   315
            Index           =   1
            Left            =   2895
            TabIndex        =   170
            Top             =   3405
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "130"
            Height          =   315
            Index           =   3
            Left            =   4155
            TabIndex        =   169
            Top             =   3405
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "180"
            Height          =   315
            Index           =   4
            Left            =   4935
            TabIndex        =   168
            Top             =   3405
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "Widget Position by Percent :"
            Height          =   375
            Index           =   8
            Left            =   195
            TabIndex        =   155
            Tag             =   "lblAlarmSound"
            Top             =   1080
            Width           =   2355
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":5A2C
            Height          =   705
            Index           =   10
            Left            =   2250
            TabIndex        =   154
            Tag             =   "lblAlarmSoundDesc"
            Top             =   1485
            Width           =   5325
         End
         Begin VB.Label lblPosition 
            Caption         =   "*"
            Height          =   255
            Index           =   1
            Left            =   4545
            TabIndex        =   94
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "this text box is filled in the setPrefsLabels sub routine"
            Height          =   4065
            Index           =   12
            Left            =   5025
            TabIndex        =   60
            Tag             =   "lblAlarmSoundDesc"
            Top             =   4425
            Width           =   2520
         End
      End
   End
   Begin VB.Frame fraAbout 
      Caption         =   "About"
      Height          =   8580
      Left            =   255
      TabIndex        =   67
      Top             =   1185
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton btnGithubHome 
         Caption         =   "Github &Home"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   225
         Width           =   1470
      End
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   7950
         TabIndex        =   81
         Top             =   1995
         Width           =   420
      End
      Begin VB.TextBox txtAboutText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         Text            =   "frmPrefs.frx":5AD2
         Top             =   2205
         Width           =   8010
      End
      Begin VB.CommandButton btnAboutDebugInfo 
         Caption         =   "Debug &Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1290
         Width           =   1470
      End
      Begin VB.CommandButton btnFacebook 
         Caption         =   "&Facebook"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   930
         Width           =   1470
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   570
         Width           =   1470
      End
      Begin VB.CommandButton btnDonate 
         Caption         =   "&Donate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1635
         Width           =   1470
      End
      Begin VB.Label lblDotDot 
         BackStyle       =   0  'Transparent
         Caption         =   ".        ."
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2940
         TabIndex        =   85
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblRevisionNum 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3450
         TabIndex        =   84
         Top             =   510
         Width           =   525
      End
      Begin VB.Label lblMajorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2730
         TabIndex        =   83
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblMinorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3090
         TabIndex        =   82
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2715
         TabIndex        =   79
         Top             =   855
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Originator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1050
         TabIndex        =   78
         Top             =   855
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1065
         TabIndex        =   77
         Top             =   495
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   2715
         TabIndex        =   76
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Current Developer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   1050
         TabIndex        =   75
         Top             =   1215
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1050
         TabIndex        =   74
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Windows XP, Vista, 7, 8, 10  && 11 + ReactOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2715
         TabIndex        =   73
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label lblAbout 
         Caption         =   "(32bit WoW64)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3900
         TabIndex        =   72
         Top             =   510
         Width           =   2550
      End
   End
   Begin VB.Frame fraDevelopment 
      Caption         =   "Development"
      Height          =   6210
      Left            =   240
      TabIndex        =   46
      Top             =   1200
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraDevelopmentInner 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   870
         TabIndex        =   47
         Top             =   300
         Width           =   7455
         Begin VB.Frame fraDefaultEditor 
            BorderStyle     =   0  'None
            Height          =   2370
            Left            =   75
            TabIndex        =   98
            Top             =   3165
            Width           =   7290
            Begin VB.CommandButton btnDefaultEditor 
               Caption         =   "..."
               Height          =   300
               Left            =   5115
               Style           =   1  'Graphical
               TabIndex        =   100
               ToolTipText     =   "Click to select the .vbp file to edit the program - You need to have access to the source!"
               Top             =   210
               Width           =   315
            End
            Begin VB.TextBox txtDefaultEditor 
               Height          =   315
               Left            =   1440
               TabIndex        =   99
               Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
               Top             =   195
               Width           =   3660
            End
            Begin VB.Label lblGitHub 
               Caption         =   $"frmPrefs.frx":6A89
               ForeColor       =   &H8000000D&
               Height          =   915
               Left            =   1560
               TabIndex        =   104
               ToolTipText     =   "Double Click to visit github"
               Top             =   1440
               Width           =   4935
            End
            Begin VB.Label lblDebug 
               Caption         =   $"frmPrefs.frx":6B22
               Height          =   930
               Index           =   9
               Left            =   1545
               TabIndex        =   102
               Top             =   690
               Width           =   4785
            End
            Begin VB.Label lblDebug 
               Caption         =   "Default Editor :"
               Height          =   255
               Index           =   7
               Left            =   285
               TabIndex        =   101
               Tag             =   "lblSharedInputFile"
               Top             =   225
               Width           =   1350
            End
         End
         Begin VB.TextBox txtDblClickCommand 
            Height          =   315
            Left            =   1515
            TabIndex        =   57
            ToolTipText     =   "Enter a Windows command for the gauge to operate when double-clicked."
            Top             =   1095
            Width           =   3660
         End
         Begin VB.CommandButton btnOpenFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5175
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Click to select a particular file for the gauge to run or open when double-clicked."
            Top             =   2250
            Width           =   315
         End
         Begin VB.TextBox txtOpenFile 
            Height          =   315
            Left            =   1515
            TabIndex        =   53
            ToolTipText     =   "Enter a particular file for the gauge to run or open when double-clicked."
            Top             =   2235
            Width           =   3660
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6BC6
            Left            =   1530
            List            =   "frmPrefs.frx":6BC8
            Style           =   2  'Dropdown List
            TabIndex        =   50
            ToolTipText     =   "Choose to set debug mode."
            Top             =   -15
            Width           =   2160
         End
         Begin VB.Label lblDebug 
            Caption         =   "DblClick Command :"
            Height          =   510
            Index           =   1
            Left            =   -15
            TabIndex        =   59
            Tag             =   "lblPrefixString"
            Top             =   1155
            Width           =   1545
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Shift+double-clicking on the widget image will open this file. "
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   58
            Top             =   2730
            Width           =   3705
         End
         Begin VB.Label lblDebug 
            Caption         =   "Default command to run when the gauge receives a double-click eg %SystemRoot%/system32/ncpa.cpl"
            Height          =   570
            Index           =   5
            Left            =   1590
            TabIndex        =   56
            Tag             =   "lblSharedInputFileDesc"
            Top             =   1605
            Width           =   4410
         End
         Begin VB.Label lblDebug 
            Caption         =   "Open File :"
            Height          =   255
            Index           =   4
            Left            =   645
            TabIndex        =   55
            Tag             =   "lblSharedInputFile"
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label lblDebug 
            Caption         =   "Turning on the debugging will provide extra information in the debug window.  *"
            Height          =   495
            Index           =   2
            Left            =   1545
            TabIndex        =   52
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   4455
         End
         Begin VB.Label lblDebug 
            Caption         =   "Debug :"
            Height          =   375
            Index           =   0
            Left            =   855
            TabIndex        =   51
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   8550
      Left            =   75
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Frame fraGeneralInner 
         BorderStyle     =   0  'None
         Height          =   8040
         Left            =   450
         TabIndex        =   49
         Top             =   225
         Width           =   7215
         Begin VB.CheckBox chkGaugeFunctions 
            Caption         =   "Enable Gauges and METAR polling *"
            Height          =   225
            Left            =   2010
            TabIndex        =   177
            ToolTipText     =   "When checked this box enables the pointer. That's it!"
            Top             =   225
            Width           =   3405
         End
         Begin VB.TextBox txtAirportsURL 
            Height          =   315
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   141
            Text            =   "https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports.dat"
            Top             =   7185
            Width           =   4755
         End
         Begin VB.CommandButton btnLocation 
            Caption         =   "Select ICAO"
            Height          =   315
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   705
            Width           =   1215
         End
         Begin VB.ComboBox cmbMetricImperial 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6BCA
            Left            =   2010
            List            =   "frmPrefs.frx":6BCC
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   2880
            Width           =   1740
         End
         Begin VB.ComboBox cmbWindSpeedScale 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6BCE
            Left            =   2010
            List            =   "frmPrefs.frx":6BD0
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   2370
            Width           =   1740
         End
         Begin VB.ComboBox cmbPressureScale 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6BD2
            Left            =   2010
            List            =   "frmPrefs.frx":6BD4
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   1830
            Width           =   1740
         End
         Begin VB.TextBox txtIcao 
            Height          =   315
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   121
            Text            =   "EGSH"
            Top             =   705
            Width           =   1725
         End
         Begin VB.ComboBox cmbTemperatureScale 
            Height          =   315
            ItemData        =   "frmPrefs.frx":6BD6
            Left            =   2010
            List            =   "frmPrefs.frx":6BD8
            Style           =   2  'Dropdown List
            TabIndex        =   118
            Top             =   1260
            Width           =   1740
         End
         Begin vb6projectCCRSlider.Slider sliSamplingInterval 
            Height          =   390
            Left            =   1890
            TabIndex        =   112
            ToolTipText     =   "Setting the sampling interval affects the frequency of the pointer updates."
            Top             =   3765
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   60
            Max             =   4800
            Value           =   60
            TickFrequency   =   100
            SelStart        =   60
         End
         Begin vb6projectCCRSlider.Slider sliStormTestInterval 
            Height          =   390
            Left            =   1890
            TabIndex        =   135
            ToolTipText     =   "Setting the sampling interval affects the frequency of the pointer updates."
            Top             =   4995
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   1800
            Max             =   7200
            Value           =   1800
            TickFrequency   =   120
            SmallChange     =   5
            LargeChange     =   10
            SelStart        =   1800
         End
         Begin vb6projectCCRSlider.Slider sliErrorInterval 
            Height          =   390
            Left            =   1890
            TabIndex        =   144
            ToolTipText     =   "Setting the sampling interval affects the frequency of the pointer updates."
            Top             =   6240
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Max             =   8
            Value           =   1
            SmallChange     =   5
            LargeChange     =   10
            SelStart        =   8
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Gauge Functions :"
            Height          =   315
            Index           =   6
            Left            =   525
            TabIndex        =   178
            Top             =   210
            Width           =   1320
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "hours"
            Height          =   315
            Index           =   23
            Left            =   3615
            TabIndex        =   148
            Top             =   6720
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "This is the full URL giving the location of the airports.dat file "
            Height          =   285
            Index           =   20
            Left            =   2070
            TabIndex        =   143
            Top             =   7650
            Width           =   4635
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Error Message Interval :"
            Height          =   360
            Index           =   24
            Left            =   180
            TabIndex        =   147
            Top             =   6300
            Width           =   1770
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "8hrs"
            Height          =   300
            Index           =   22
            Left            =   5490
            TabIndex        =   146
            Top             =   6720
            Width           =   465
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "0 (disabled)"
            Height          =   300
            Index           =   21
            Left            =   1995
            TabIndex        =   145
            Top             =   6720
            Width           =   945
         End
         Begin VB.Label lblGeneral 
            Caption         =   $"frmPrefs.frx":6BDA
            Height          =   1005
            Index           =   12
            Left            =   3900
            TabIndex        =   129
            Top             =   2850
            Width           =   3270
         End
         Begin VB.Label lblGeneral 
            Caption         =   "ICAO Airports URL :"
            Height          =   255
            Index           =   15
            Left            =   255
            TabIndex        =   142
            Top             =   7230
            Width           =   1545
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "1800"
            Height          =   315
            Index           =   19
            Left            =   2070
            TabIndex        =   140
            Top             =   5475
            Width           =   540
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "7200"
            Height          =   315
            Index           =   18
            Left            =   5385
            TabIndex        =   139
            Top             =   5460
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "secs"
            Height          =   315
            Index           =   17
            Left            =   3615
            TabIndex        =   138
            Top             =   5460
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Storm Test Interval :"
            Height          =   315
            Index           =   16
            Left            =   300
            TabIndex        =   137
            Top             =   5055
            Width           =   1635
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Adjust to determine storm checking frequency."
            Height          =   540
            Index           =   15
            Left            =   2100
            TabIndex        =   136
            Top             =   5760
            Width           =   3810
         End
         Begin VB.Label lblGeneral 
            Caption         =   "The Wind Speed Scale"
            Height          =   480
            Index           =   14
            Left            =   3885
            TabIndex        =   131
            Top             =   2430
            Width           =   2610
         End
         Begin VB.Label lblGeneral 
            Caption         =   "The Air Pressure Scale"
            Height          =   480
            Index           =   13
            Left            =   3885
            TabIndex        =   130
            Top             =   1875
            Width           =   2610
         End
         Begin VB.Label lblGeneral 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric or Imperial :"
            Height          =   480
            Index           =   8
            Left            =   255
            TabIndex        =   128
            Top             =   2910
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Alignment       =   1  'Right Justify
            Caption         =   "Anemometer :"
            Height          =   480
            Index           =   7
            Left            =   255
            TabIndex        =   126
            Top             =   2400
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Alignment       =   1  'Right Justify
            Caption         =   "Barometer :"
            Height          =   345
            Index           =   4
            Left            =   255
            TabIndex        =   124
            Top             =   1860
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Caption         =   "ICAO Station ID :"
            Height          =   255
            Index           =   1
            Left            =   585
            TabIndex        =   122
            Top             =   765
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Temperature :"
            Height          =   480
            Index           =   5
            Left            =   810
            TabIndex        =   120
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Select Celsius / Fahrenheit / Kelvin"
            Height          =   480
            Index           =   10
            Left            =   3855
            TabIndex        =   119
            Top             =   1305
            Width           =   2610
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Adjust to determine gauge sampling frequency.*"
            Height          =   270
            Index           =   14
            Left            =   2100
            TabIndex        =   117
            Top             =   4530
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Sampling Interval :"
            Height          =   315
            Index           =   13
            Left            =   495
            TabIndex        =   116
            Top             =   3825
            Width           =   1410
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "secs"
            Height          =   315
            Index           =   12
            Left            =   3615
            TabIndex        =   115
            Top             =   4230
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "4800"
            Height          =   315
            Index           =   11
            Left            =   5385
            TabIndex        =   114
            Top             =   4230
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "60"
            Height          =   315
            Index           =   10
            Left            =   2070
            TabIndex        =   113
            Top             =   4230
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      Height          =   6405
      Left            =   240
      TabIndex        =   8
      Top             =   1185
      Width           =   7140
      Begin VB.Frame fraConfigInner 
         BorderStyle     =   0  'None
         Height          =   5745
         Left            =   435
         TabIndex        =   34
         Top             =   435
         Width           =   6450
         Begin VB.Frame fraPrefsTooltips 
            BorderStyle     =   0  'None
            Height          =   1125
            Index           =   0
            Left            =   1845
            TabIndex        =   183
            Top             =   2310
            Width           =   3150
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Prefs - Enable SquareTooltips *"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   186
               Top             =   450
               Width           =   2970
            End
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Prefs - Enable Balloon Tooltips *"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   185
               Top             =   120
               Width           =   2760
            End
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Disable Prefs Tooltips *"
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   184
               Top             =   750
               Width           =   2970
            End
         End
         Begin VB.Frame fraClockTooltips 
            BorderStyle     =   0  'None
            Height          =   1110
            Left            =   1755
            TabIndex        =   179
            Top             =   1140
            Width           =   3345
            Begin VB.OptionButton optClockTooltips 
               Caption         =   "Gauges - Enable Balloon Tooltips *"
               Height          =   315
               Index           =   0
               Left            =   225
               TabIndex        =   182
               Top             =   120
               Width           =   3060
            End
            Begin VB.OptionButton optClockTooltips 
               Caption         =   "Gauges - Enable Square Tooltips"
               Height          =   300
               Index           =   1
               Left            =   225
               TabIndex        =   181
               Top             =   450
               Width           =   2790
            End
            Begin VB.OptionButton optClockTooltips 
               Caption         =   "Disable Gauge Tooltips *"
               Height          =   300
               Index           =   2
               Left            =   225
               TabIndex        =   180
               Top             =   780
               Width           =   2790
            End
         End
         Begin VB.CheckBox chkGenStartup 
            Caption         =   "Run the Temperature Widget at Windows Startup "
            Height          =   465
            Left            =   1995
            TabIndex        =   133
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   4875
            Width           =   4020
         End
         Begin VB.CheckBox chkDpiAwareness 
            Caption         =   "DPI Awareness Enable *"
            Height          =   225
            Left            =   1995
            TabIndex        =   105
            ToolTipText     =   "Check the box to make the program DPI aware. RESTART required."
            Top             =   3885
            Width           =   3405
         End
         Begin VB.CheckBox chkShowTaskbar 
            Caption         =   "Show Gauges in Taskbar"
            Height          =   225
            Left            =   1995
            TabIndex        =   103
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   3495
            Width           =   3405
         End
         Begin VB.ComboBox cmbScrollWheelDirection 
            Height          =   315
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   61
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   120
            Width           =   2490
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Auto Start :"
            Height          =   375
            Index           =   11
            Left            =   960
            TabIndex        =   134
            Tag             =   "lblRefreshInterval"
            Top             =   4995
            Width           =   1740
         End
         Begin VB.Label lblConfiguration 
            Caption         =   $"frmPrefs.frx":6C78
            Height          =   930
            Index           =   0
            Left            =   1965
            TabIndex        =   106
            Top             =   4215
            Width           =   4335
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "The scroll-wheel resizing direction can be determined here. The direction chosen causes the gauge to grow. *"
            Height          =   690
            Index           =   6
            Left            =   2025
            TabIndex        =   86
            Top             =   540
            Width           =   3930
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Mouse Wheel Resize :"
            Height          =   345
            Index           =   3
            Left            =   255
            TabIndex        =   62
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   165
            Width           =   2055
         End
      End
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Fonts"
      Height          =   4320
      Left            =   255
      TabIndex        =   9
      Top             =   1230
      Width           =   7335
      Begin VB.Frame fraFontsInner 
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   690
         TabIndex        =   26
         Top             =   360
         Width           =   6105
         Begin VB.CommandButton btnResetMessages 
            Caption         =   "Reset"
            Height          =   300
            Left            =   1710
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   3405
            Width           =   885
         End
         Begin VB.TextBox txtPrefsFontCurrentSize 
            Height          =   315
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   96
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1065
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtPrefsFontSize 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "8"
            ToolTipText     =   "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
            Top             =   1065
            Width           =   510
         End
         Begin VB.CommandButton btnPrefsFont 
            Caption         =   "Font"
            Height          =   300
            Left            =   5025
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "The Font Selector."
            Top             =   75
            Width           =   585
         End
         Begin VB.TextBox txtPrefsFont 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "Times New Roman"
            ToolTipText     =   "Disabled for manual input. Choose a font via the font selector to be used only for this preferences window"
            Top             =   90
            Width           =   3285
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Hidden message boxes can be reactivated by pressing this reset button."
            Height          =   480
            Index           =   4
            Left            =   2700
            TabIndex        =   110
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   3345
            Width           =   3360
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Reset Pop ups :"
            Height          =   300
            Index           =   1
            Left            =   435
            TabIndex        =   108
            Tag             =   "lblPrefsFont"
            Top             =   3450
            Width           =   1470
         End
         Begin VB.Label lblCurrentFontsTab 
            Caption         =   "Resized Font"
            Height          =   315
            Left            =   4950
            TabIndex        =   97
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1110
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   $"frmPrefs.frx":6D2C
            Height          =   1710
            Index           =   0
            Left            =   1725
            TabIndex        =   64
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   1455
            Width           =   4455
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "The chosen font size *"
            Height          =   480
            Index           =   7
            Left            =   2370
            TabIndex        =   33
            ToolTipText     =   "Choose a font size that fits the text boxes"
            Top             =   1095
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Base Font Size :"
            Height          =   330
            Index           =   3
            Left            =   435
            TabIndex        =   32
            Tag             =   "lblPrefsFontSize"
            Top             =   1095
            Width           =   1230
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Prefs Window Font:"
            Height          =   300
            Index           =   2
            Left            =   15
            TabIndex        =   31
            Tag             =   "lblPrefsFont"
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in this preferences window, gauge tooltips and message boxes *"
            Height          =   480
            Index           =   6
            Left            =   1695
            TabIndex        =   30
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   480
            Width           =   4035
         End
      End
   End
   Begin VB.Frame fraSounds 
      Caption         =   "Sounds"
      Height          =   1965
      Left            =   240
      TabIndex        =   13
      Top             =   1230
      Visible         =   0   'False
      Width           =   7965
      Begin VB.Frame fraSoundsInner 
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   930
         TabIndex        =   25
         Top             =   285
         Width           =   5160
         Begin VB.CheckBox chkEnableSounds 
            Caption         =   "Enable Sounds for the Animations"
            Height          =   225
            Left            =   1485
            TabIndex        =   35
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   285
            Width           =   3405
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Audio :"
            Height          =   255
            Index           =   3
            Left            =   885
            TabIndex        =   63
            Tag             =   "lblSharedInputFile"
            Top             =   285
            Width           =   765
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "When checked, this box enables all the sounds used during any animation/interaction with the main program."
            Height          =   705
            Index           =   4
            Left            =   1515
            TabIndex        =   36
            Tag             =   "lblEnableSoundsDesc"
            Top             =   645
            Width           =   3705
         End
      End
   End
   Begin VB.Frame fraWindow 
      Caption         =   "Window"
      Height          =   6300
      Left            =   420
      TabIndex        =   10
      Top             =   1515
      Width           =   8280
      Begin VB.Frame fraWindowInner 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   1095
         TabIndex        =   14
         Top             =   345
         Width           =   5715
         Begin VB.Frame fraHiding 
            BorderStyle     =   0  'None
            Height          =   2010
            Left            =   480
            TabIndex        =   87
            Top             =   2325
            Width           =   5130
            Begin VB.ComboBox cmbHidingTime 
               Height          =   315
               Left            =   825
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   1680
               Width           =   3720
            End
            Begin VB.CheckBox chkWidgetHidden 
               Caption         =   "Hiding Widget *"
               Height          =   225
               Left            =   855
               TabIndex        =   88
               Top             =   225
               Width           =   2955
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   "Hiding :"
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   91
               Top             =   210
               Width           =   720
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   $"frmPrefs.frx":6E6A
               Height          =   975
               Index           =   1
               Left            =   855
               TabIndex        =   89
               Top             =   600
               Width           =   3900
            End
         End
         Begin VB.ComboBox cmbWindowLevel 
            Height          =   315
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   3720
         End
         Begin VB.CheckBox chkIgnoreMouse 
            Caption         =   "Ignore Mouse *"
            Height          =   225
            Left            =   1335
            TabIndex        =   15
            ToolTipText     =   "Checking this box causes the program to ignore all mouse events."
            Top             =   1500
            Width           =   2535
         End
         Begin vb6projectCCRSlider.Slider sliOpacity 
            Height          =   390
            Left            =   1200
            TabIndex        =   16
            ToolTipText     =   "Set the transparency of the Program."
            Top             =   4560
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   20
            Max             =   100
            Value           =   20
            SelStart        =   20
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "This setting controls the relative layering of this widget. You may use it to place it on top of other windows or underneath. "
            Height          =   660
            Index           =   3
            Left            =   1320
            TabIndex        =   93
            Top             =   570
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Window Level :"
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "20%"
            Height          =   315
            Index           =   7
            Left            =   1290
            TabIndex        =   23
            Top             =   5070
            Width           =   345
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "100%"
            Height          =   315
            Index           =   9
            Left            =   4650
            TabIndex        =   22
            Top             =   5070
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Opacity"
            Height          =   315
            Index           =   8
            Left            =   2775
            TabIndex        =   21
            Top             =   5070
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Opacity:"
            Height          =   315
            Index           =   6
            Left            =   555
            TabIndex        =   20
            Top             =   4620
            Width           =   780
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Set the program transparency level."
            Height          =   330
            Index           =   5
            Left            =   1335
            TabIndex        =   19
            Top             =   5385
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Checking this box causes the program to ignore all mouse events except right click menu interactions."
            Height          =   660
            Index           =   4
            Left            =   1320
            TabIndex        =   18
            Top             =   1890
            Width           =   3810
         End
      End
   End
   Begin VB.Label lblDragCorner 
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   8670
      TabIndex        =   111
      ToolTipText     =   "drag me"
      Top             =   10335
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSize 
      Caption         =   "Size in twips"
      Height          =   285
      Left            =   1875
      TabIndex        =   107
      Top             =   9780
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Label lblAsterix 
      Caption         =   "All controls marked with a * take effect immediately."
      Height          =   300
      Left            =   1920
      TabIndex        =   92
      Top             =   10155
      Width           =   3870
   End
   Begin VB.Menu prefsMnuPopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About Panzer Weather Widget"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with KoFi"
      End
      Begin VB.Menu mnuSupport 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButton 
         Caption         =   "Theme Colours"
         Begin VB.Menu mnuLight 
            Caption         =   "Light Theme Enable"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "High Contrast Theme Enable"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Auto Theme Selection"
         End
      End
      Begin VB.Menu mnuLicenceA 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuClosePreferences 
         Caption         =   "Close Preferences"
      End
   End
End
Attribute VB_Name = "widgetPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ArgumentWithIncompatibleObjectType, AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder

' gaugeForm_BubblingEvent ' leaving that here so I can copy/paste to find it

'---------------------------------------------------------------------------------------
' Module    : widgetPrefs
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/08/2022
' Purpose   : VB6 standard form to display the prefs
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
' Constants and APIs to create and subclass the dragCorner
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTBOTTOMRIGHT  As Long = 17
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Constants defined for setting a theme to the prefs
Private Const COLOR_BTNFACE As Long = 15

' APIs declared for setting a theme to the prefs
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Types for determining the timezone

'Private BiasAdjust As Boolean
'
'' results UDT
'Private Type TZ_LOOKUP_DATA
'   TimeZoneName As String
'   bias As Long
'   IsDST As Boolean
'End Type
'
'Private tzinfo() As TZ_LOOKUP_DATA
'
''holds the correct key for the OS version
'Private sTzKey As String

'windows constants and declares
'Private Const TIME_ZONE_ID_UNKNOWN As Long = 1
'Private Const TIME_ZONE_ID_STANDARD As Long = 1
'Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
'Private Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF
'Private Const VER_PLATFORM_WIN32_NT = 2
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1

'registry constants
'Private Const SKEY_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones"
'Private Const SKEY_9X = "SOFTWARE\Microsoft\Windows\CurrentVersion\Time Zones"
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const ERROR_SUCCESS = 0
''Private Const REG_SZ As Long = 1
''Private Const REG_BINARY = 3
''Private Const REG_DWORD As Long = 4
'Private Const STANDARD_RIGHTS_READ As Long = &H20000
'Private Const KEY_QUERY_VALUE As Long = &H1
'Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
'Private Const KEY_NOTIFY As Long = &H10
'Private Const SYNCHRONIZE As Long = &H100000
'Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or _
'                                   KEY_QUERY_VALUE Or _
'                                   KEY_ENUMERATE_SUB_KEYS Or _
'                                   KEY_NOTIFY) And _
'                                   (Not SYNCHRONIZE))

'Private Type SYSTEMTIME
'   wYear As Integer
'   wMonth As Integer
'   wDayOfWeek As Integer
'   wDay As Integer
'   wHour As Integer
'   wMinute As Integer
'   wSecond As Integer
'   wMilliseconds As Integer
'End Type
'
'Private Type FILETIME
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'Private Type REG_TIME_ZONE_INFORMATION
'   bias As Long
'   StandardBias As Long
'   DaylightBias As Long
'   StandardDate As SYSTEMTIME
'   DaylightDate As SYSTEMTIME
'End Type


'Private Type TIME_ZONE_INFORMATION
'    bias                    As Long
'    StandardName(0 To 63)   As Byte
'    StandardDate            As SYSTEMTIME
'    StandardBias            As Long
'    DaylightName(0 To 63)   As Byte
'    DaylightDate            As SYSTEMTIME
'    DaylightBias            As Long
'End Type

'Private Type OSVERSIONINFO
'   OSVSize As Long
'   dwVerMajor As Long
'   dwVerMinor As Long
'   dwBuildNumber As Long
'   PlatformID As Long
'   szCSDVersion As String * 128
'End Type

' APIs for determining the timezone

'Private Declare Function GetVersionEx Lib "kernel32" _
'   Alias "GetVersionExA" _
'  (lpVersionInformation As OSVERSIONINFO) As Long
'
'Private Declare Function GetTimeZoneInformation Lib "kernel32" _
'   (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
'   Alias "RegOpenKeyExA" _
'  (ByVal hKey As Long, _
'   ByVal lpsSubKey As String, _
'   ByVal ulOptions As Long, _
'   ByVal samDesired As Long, _
'   phkResult As Long) As Long
'
'Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
'   Alias "RegQueryValueExA" _
'  (ByVal hKey As Long, _
'   ByVal lpszValueName As String, _
'   ByVal lpdwReserved As Long, _
'   lpdwType As Long, _
'   lpData As Any, _
'   lpcbData As Long) As Long
'
'Private Declare Function RegQueryInfoKey Lib "advapi32.dll" _
'   Alias "RegQueryInfoKeyA" _
'  (ByVal hKey As Long, _
'   ByVal lpClass As String, _
'   lpcbClass As Long, _
'   ByVal lpReserved As Long, _
'   lpcsSubKeys As Long, _
'   lpcbMaxsSubKeyLen As Long, _
'   lpcbMaxClassLen As Long, _
'   lpcValues As Long, _
'   lpcbMaxValueNameLen As Long, _
'   lpcbMaxValueLen As Long, _
'   lpcbSecurityDescriptor As Long, _
'   lpftLastWriteTime As FILETIME) As Long
   
'Private Declare Function RegQueryValueExString Lib "advapi32.dll" _
'   Alias "RegQueryValueExA" _
'  (ByVal hKey As Long, _
'   ByVal lpValueName As String, _
'   ByVal lpReserved As Long, _
'   lpType As Long, _
'   ByVal lpData As String, _
'   lpcbData As Long) As Long

'Private Declare Function RegEnumKey Lib "advapi32.dll" _
'   Alias "RegEnumKeyA" _
'  (ByVal hKey As Long, _
'   ByVal dwIndex As Long, _
'   ByVal lpName As String, _
'   ByVal cbName As Long) As Long
'
'Private Declare Function RegCloseKey Lib "advapi32.dll" _
'  (ByVal hKey As Long) As Long

'Private Declare Function lstrlenW Lib "kernel32" _
'  (ByVal lpString As Long) As Long
'
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Private Types for determining prefs sizing
Private gblPrefsLoadedFlg As Boolean
Private pvtPrefsDynamicSizingFlg As Boolean
Private pvtLastFormHeight As Long
Private Const cPrefsFormHeight As Long = 11055
Private Const cPrefsFormWidth  As Long = 9090
'------------------------------------------------------ ENDS

Private pvtPrefsStartupFlg As Boolean
Private gblAllowSizeChangeFlg As Boolean

Private mIsLoaded As Boolean ' property

' module level balloon tooltip variables for subclassed comboBoxes ONLY.

Private pCmbTemperatureScaleBalloonTooltip As String
Private pCmbPressureScaleBalloonTooltip As String
Private pCmbWindSpeedScaleBalloonTooltip As String
Private pCmbMetricImperialBalloonTooltip As String

Private pCmbGaugeTypeBalloonTooltip As String
        
Private pCmbMultiMonitorResizeBalloonTooltip As String
Private pCmbScrollWheelDirectionBalloonTooltip As String
Private pCmbWindowLevelBalloonTooltip As String
Private pCmbHidingTimeBalloonTooltip As String
Private pCmbAspectHiddenBalloonTooltip As String
Private pCmbWidgetPositionBalloonTooltip As String
Private pcmbLandscapeLockedBalloonTooltip As String
Private pcmbPortraitLockedBalloonTooltip As String
Private pCmbDebugBalloonTooltip As String
Private pCmbAlarmDayBalloonTooltip As String
Private pCmbAlarmMonthBalloonTooltip As String
Private pCmbAlarmYearBalloonTooltip As String
Private pCmbAlarmHoursBalloonTooltip As String
Private pCmbAlarmMinutesBalloonTooltip As String


Private Sub btnAboutDebugInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnAboutDebugInfo.hWnd, "Here you can switch on Debug mode, not yet functional for this widget.", _
                  TTIconInfo, "Help on the Debug Info. Buttton", , , , True
End Sub

Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnClose.hWnd, "Close the Preference Utility", _
                  TTIconInfo, "Help on the Close Buttton", , , , True
End Sub

Private Sub btnDefaultEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnDefaultEditor.hWnd, "Field to hold the path to a Visual Basic Project (VBP) file you would like to execute on a right click menu, edit option, if you select the adjacent button a file explorer will appear allowing you to select the VBP file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the Default Editor Field", , , , True
End Sub

Private Sub btnDonate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnDonate.hWnd, "Here you can visit my KofI page and donate a Coffee if you like my creations.", _
                  TTIconInfo, "Help on the Donate Buttton", , , , True
End Sub

Private Sub btnFacebook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnFacebook.hWnd, "Here you can visit the Facebook page for the steampunk Widget community.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnHelp.hWnd, "Opens the help document, this will open as a compiled HTML file.", _
                  TTIconInfo, "Help on the Help Buttton", , , , True
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnSave.hWnd, "Save the changes you have made to the preferences", _
                  TTIconInfo, "Help on the Save Buttton", , , , True
End Sub

Private Sub btnUpdate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnUpdate.hWnd, "Here you can able to download a new version of the program from github, when you click the button it will open a browser window and take you to the github page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnLocation_Click()
    fSelector.SelectorForm.Show
End Sub

Private Sub btnLocation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnLocation.hWnd, "Press to select the current ICAO code used to identify the weather feed source data.*", _
                  TTIconInfo, "Select the current ICAO code", , , , True

End Sub

Private Sub btnOpenFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnOpenFile.hWnd, "Clicking on this button will cause a file explorer window to appear allowing you to select any file you would like to execute on a shift+DBlClick. Once selected the adjacent text field will be automatically filled with the chosen path and file.", _
                  TTIconInfo, "Help on the shift+DBlClick File Explorer Button", , , , True
End Sub

Private Sub btnGithubHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnGithubHome.hWnd, "Here you can visit the widget's home page on github, when you click the button it will open a browser window and take you to the github home page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub chkDpiAwareness_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkDpiAwareness.hWnd, "Check here to make the program DPI aware. NOT required on small to medium screens that are less than 1920 bytes wide. Try it and see which suits your system. RESTART required.", _
                  TTIconInfo, "Help on DPI Awareness Mode", , , , True
End Sub

Private Sub chkGenStartup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkGenStartup.hWnd, "Check this box to enable the automatic start of the program when Windows is started.", _
                  TTIconInfo, "Help on the Widget Automatic Start Toggle", , , , True
End Sub

Private Sub chkGaugeFunctions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkGaugeFunctions.hWnd, "Press to toggle the polling for METAR data.*", _
                  TTIconInfo, "Enable METAR polling", , , , True

End Sub


Private Sub chkPreventDragging_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkPreventDragging.hWnd, "" _
    & "Lock the gauges in a certain position in either landscape/portrait mode. This ensures that the widget always appears exactly where you want it to. Drag" _
    & "the gauge into position using the mouse and when the widget is locked in place (using the Widget lock button), this value is set automatically.", _
                  TTIconInfo, "Lock the gauges in position.", , , , True
End Sub

Private Sub chkShowTaskbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkShowTaskbar.hWnd, "Checking this box causes" _
        & " each gauge and form within the weather widget to appear in the taskbar. " _
        & " Disabling it allows for a much cleaner taskbar (recommended).", _
        TTIconInfo, "Help on taskbar visibility.", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbGaugeType_Click
' Author    : beededea
' Date      : 06/05/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbGaugeType_Click()
   On Error GoTo cmbGaugeType_Click_Error

    btnSave.Enabled = True ' enable the save button
    fraGaugePosition.Caption = cmbGaugeType.List(cmbGaugeType.ListIndex) & " Position"
    
    If cmbGaugeType.ListIndex = 0 Then ' fTemperature gauge
        If aspectRatio = "landscape" Then
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = gblTemperatureFormHighDpiXPos
                txtLandscapeVoffset.Text = gblTemperatureFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = gblTemperatureFormLowDpiXPos
                txtLandscapeVoffset.Text = gblTemperatureFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fTemperature.temperatureGaugeForm.Left
            txtPortraitVoffset.Text = fTemperature.temperatureGaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayTemperatureWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        widgetPrefs.sliGaugeSize.Value = Val(gblTemperatureGaugeSize)
    End If

    'Anemometer gauge
    
    If cmbGaugeType.ListIndex = 1 Then
        If aspectRatio = "landscape" Then
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = gblAnemometerFormHighDpiXPos
                txtLandscapeVoffset.Text = gblAnemometerFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblAnemometerFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblAnemometerFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = gblAnemometerFormLowDpiXPos
                txtLandscapeVoffset.Text = gblAnemometerFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblAnemometerFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblAnemometerFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fAnemometer.anemometerGaugeForm.Left
            txtPortraitVoffset.Text = fAnemometer.anemometerGaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblAnemometerFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblAnemometerFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblAnemometerFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblAnemometerFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayAnemoWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        
        widgetPrefs.sliGaugeSize.Value = Val(gblAnemometerGaugeSize)

    End If
    
    'humidity gauge
    
    If cmbGaugeType.ListIndex = 2 Then
        If aspectRatio = "landscape" Then
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = gblHumidityFormHighDpiXPos
                txtLandscapeVoffset.Text = gblHumidityFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblHumidityFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblHumidityFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = gblHumidityFormLowDpiXPos
                txtLandscapeVoffset.Text = gblHumidityFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblHumidityFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblHumidityFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fHumidity.humidityGaugeForm.Left
            txtPortraitVoffset.Text = fHumidity.humidityGaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblHumidityFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblHumidityFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblHumidityFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblHumidityFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayHumidWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        
        widgetPrefs.sliGaugeSize.Value = Val(gblHumidityGaugeSize)

    End If
    
    'barometer gauge
    
    If cmbGaugeType.ListIndex = 3 Then
        If aspectRatio = "landscape" Then
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = gblBarometerFormHighDpiXPos
                txtLandscapeVoffset.Text = gblBarometerFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblBarometerFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblBarometerFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = gblBarometerFormLowDpiXPos
                txtLandscapeVoffset.Text = gblBarometerFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblBarometerFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblBarometerFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fBarometer.barometerGaugeForm.Left
            txtPortraitVoffset.Text = fBarometer.barometerGaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblBarometerFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblBarometerFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblBarometerFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblBarometerFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayAnemoWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        
        widgetPrefs.sliGaugeSize.Value = Val(gblBarometerGaugeSize)

    End If
    
    
    'pictorial gauge
    
    If cmbGaugeType.ListIndex = 4 Then
        If aspectRatio = "landscape" Then
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = gblPictorialFormHighDpiXPos
                txtLandscapeVoffset.Text = gblPictorialFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblPictorialFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblPictorialFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = gblPictorialFormLowDpiXPos
                txtLandscapeVoffset.Text = gblPictorialFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblPictorialFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblPictorialFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fPictorial.pictorialGaugeForm.Left
            txtPortraitVoffset.Text = fPictorial.pictorialGaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblPictorialFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblPictorialFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblPictorialFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblPictorialFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayAnemoWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        
        widgetPrefs.sliGaugeSize.Value = Val(gblPictorialGaugeSize)

    End If
    
   On Error GoTo 0
   Exit Sub

cmbGaugeType_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbGaugeType_Click of Form widgetPrefs"
End Sub

' ----------------------------------------------------------------
' Procedure Name: cmbMetricImperial_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 20/03/2024
' ----------------------------------------------------------------
Private Sub cmbMetricImperial_Click()
    On Error GoTo cmbMetricImperial_Click_Error
    
    btnSave.Enabled = True ' enable the save button
    gblWindSpeedScale = LTrim$(Str$(cmbWindSpeedScale.ListIndex))
    sPutINISetting "Software\PzTemperatureGauge", "metricImperial", gblMetricImperial, gblSettingsFile
    
    On Error GoTo 0
    Exit Sub

cmbMetricImperial_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbMetricImperial_Click, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: cmbPressureScale_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 19/03/2024
' ----------------------------------------------------------------
Private Sub cmbPressureScale_Click()
    On Error GoTo cmbPressureScale_Click_Error
    
    btnSave.Enabled = True ' enable the save button

    
    overlayBaromWidget.thisFace = cmbPressureScale.ListIndex
    
    On Error GoTo 0
    Exit Sub

cmbPressureScale_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbPressureScale_Click, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: cmbTemperatureScale_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 17/01/2024
' ----------------------------------------------------------------
Private Sub cmbTemperatureScale_Click()
    
    On Error GoTo cmbTemperatureScale_Click_Error
    btnSave.Enabled = True ' enable the save button

    
    overlayTemperatureWidget.thisFace = cmbTemperatureScale.ListIndex
    
    On Error GoTo 0
    Exit Sub

cmbTemperatureScale_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbTemperatureScale_Click, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: cmbWindSpeedScale_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 19/03/2024
' ----------------------------------------------------------------
Private Sub cmbWindSpeedScale_Click()
    
    On Error GoTo cmbWindSpeedScale_Click_Error
    btnSave.Enabled = True ' enable the save button
'    gblWindSpeedScale = LTrim$(Str$(cmbWindSpeedScale.ListIndex))
'    sPutINISetting "Software\PzTemperatureGauge", "windSpeedScale", gblWindSpeedScale, gblSettingsFile
        
    overlayAnemoWidget.thisFace = cmbWindSpeedScale.ListIndex

    
    On Error GoTo 0
    Exit Sub

cmbWindSpeedScale_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWindSpeedScale_Click, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: Form_Initialize
' Purpose:
' Procedure Kind: Constructor (Initialize)
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 05/10/2023
' ----------------------------------------------------------------
Private Sub Form_Initialize()
    On Error GoTo Form_Initialize_Error
    
     ' initialise private variables
    Call initialisePrefsVars

    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form widgetPrefs"
    
    End Sub




' ---------------------------------------------------------------------------------------
' Procedure : initialisePrefsVars
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : initialise private variables
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrefsVars()

   On Error GoTo initialisePrefsVars_Error

    gblPrefsLoadedFlg = False
    pvtPrefsDynamicSizingFlg = False
    pvtLastFormHeight = 0
    pvtPrefsStartupFlg = False
'    pvtAllowSizeChangeFlg = False
    pCmbMultiMonitorResizeBalloonTooltip = vbNullString
    pCmbScrollWheelDirectionBalloonTooltip = vbNullString
    pCmbWindowLevelBalloonTooltip = vbNullString
    pCmbHidingTimeBalloonTooltip = vbNullString
    pCmbAspectHiddenBalloonTooltip = vbNullString
    pCmbWidgetPositionBalloonTooltip = vbNullString
    pcmbLandscapeLockedBalloonTooltip = vbNullString
    pcmbPortraitLockedBalloonTooltip = vbNullString
    pCmbDebugBalloonTooltip = vbNullString
    pCmbAlarmDayBalloonTooltip = vbNullString
    pCmbAlarmMonthBalloonTooltip = vbNullString
    pCmbAlarmYearBalloonTooltip = vbNullString
    pCmbAlarmHoursBalloonTooltip = vbNullString
    pCmbAlarmMinutesBalloonTooltip = vbNullString
    'pvtPrefsFormResizedByDrag = False
    mIsLoaded = False ' property

   On Error GoTo 0
   Exit Sub

initialisePrefsVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrefsVars of Form widgetPrefs"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 25/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    
    Dim prefsFormHeight As Long: prefsFormHeight = 0

    On Error GoTo Form_Load_Error
        
    pvtPrefsStartupFlg = True ' this is used to prevent some control initialisations from running code at startup
    pvtPrefsDynamicSizingFlg = False
    gblPrefsLoadedFlg = True ' this is a variable tested by an added form property to indicate whether the form is loaded or not
    gblWindowLevelWasChanged = False
    prefsFormHeight = prefsCurrentHeight
    
    With lblDragCorner
      .ForeColor = &H80000015
      .BackStyle = vbTransparent
      .AutoSize = True
      .Font.Size = 12
      .Font.Name = "Marlett"
      .Caption = "o"
      .Font.Bold = False
      .Visible = False
    End With
     
    btnSave.Enabled = False ' disable the save button
    Me.mnuAbout.Caption = "About Panzer Weather Gauge Cairo " & gblCodingEnvironment & " widget"
    
    If gblDpiAwareness = "1" Then
        pvtPrefsDynamicSizingFlg = True
        chkEnableResizing.Value = 1
        lblDragCorner.Visible = True
    End If
    
    ' subclass specific WidgetPrefs controls that need additional functionality that VB6 does not provide (scrollwheel/balloon tooltips)
    Call subClassControls
    
    ' reverts TwinBasic form themeing to that of the earlier classic look and feel
    #If TWINBASIC Then
       Call setVisualStyles
    #End If
      
    ' read the last saved position from the settings.ini
    Call readPrefsPosition
        
    ' determine the frame heights in dynamic sizing or normal mode
    Call setframeHeights
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips
    
    ' set the text in any labels that need a vbCrLf to space the text
    Call setPrefsLabels
    
    ' populate all the comboboxes in the prefs form
    Call populatePrefsComboBoxes
        
    ' adjust all the preferences and main program controls
    Call adjustPrefsControls
    
    ' adjust the theme used by the prefs alone
    Call adjustPrefsTheme
    
    ' size and position the frames and buttons
    Call positionPrefsFramesButtons
    
    ' make the last used tab appear on startup
    Call showLastTab
    
    ' load the about text and load into prefs
    Call loadPrefsAboutText
    
    ' load the preference icons from a previously populated CC imageList
    Call loadHigherResPrefsImages
    
    ' now cause a form_resize event and set the height of the whole form
    If gblDpiAwareness = "1" Then
        If prefsFormHeight < gblPhysicalScreenHeightTwips Then
            Me.Height = prefsFormHeight
        Else
            Me.Height = gblPhysicalScreenHeightTwips - 1000
        End If
    End If
    
    ' position the prefs on the current monitor
    Call positionPrefsMonitor
    
    ' set the Z order of the prefs form
    Call setPrefsFormZordering
    
    ' start the timer that records the prefs position every 10 seconds
    positionTimer.Enabled = True
    
    ' end the startup by un-setting the start flag
    pvtPrefsStartupFlg = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form widgetPrefs"

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : positionPrefsMonitor
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 06/10/2023
' Purpose   : position the prefs on the current monitor
'---------------------------------------------------------------------------------------
'
Public Sub positionPrefsMonitor()

    Dim formLeftTwips As Long: formLeftTwips = 0
    Dim formTopTwips As Long: formTopTwips = 0
    
    Dim monitorCount As Long: monitorCount = 0
    
    On Error GoTo positionPrefsMonitor_Error
    
    If gblDpiAwareness = "1" Then
        formLeftTwips = Val(gblPrefsFormHighDpiXPosTwips)
        formTopTwips = Val(gblPrefsFormHighDpiYPosTwips)
    Else
        formLeftTwips = Val(gblPrefsFormLowDpiXPosTwips)
        formTopTwips = Val(gblPrefsFormLowDpiYPosTwips)
    End If
    
    If formLeftTwips = 0 Then
        If ((fTemperature.temperatureGaugeForm.Left + fTemperature.temperatureGaugeForm.Width) * gblScreenTwipsPerPixelX) + 200 + Me.Width > gblVirtualScreenWidthPixels Then
            Me.Left = (fTemperature.temperatureGaugeForm.Left * gblScreenTwipsPerPixelX) - (Me.Width + 200)
        End If
    End If

    ' if a current location not stored then position to the middle of the screen
    If formLeftTwips <> 0 Then
        Me.Left = formLeftTwips
    Else
        Me.Left = gblVirtualScreenWidthPixels / 2 - Me.Width / 2
    End If
    
    If formTopTwips <> 0 Then
        Me.Top = formTopTwips
    Else
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
    
    monitorCount = fGetMonitorCount
    If monitorCount > 1 Then Call SetFormOnMonitor(Me.hWnd, formLeftTwips / fTwipsPerPixelX, formTopTwips / fTwipsPerPixelY)
    
    ' calculate the on-screen widget position
    If Me.Left < 0 Then
        Me.Left = 10
    End If
    If Me.Top < 0 Then
        Me.Top = 0
    End If
'    If Me.Left > gblVirtualScreenWidthPixels - 2500 Then
'        Me.Left = gblVirtualScreenWidthPixels - 2500
'    End If
'    If Me.Top > gblPhysicalScreenHeightTwips - 2500 Then
'        Me.Top = gblPhysicalScreenHeightTwips - 2500
'    End If
    
    If Me.Left > gblVirtualScreenWidthTwips - 2500 Then
        Me.Left = gblVirtualScreenWidthTwips - 2500
    End If
    If Me.Top > gblVirtualScreenHeightTwips - 2500 Then
        Me.Top = gblVirtualScreenHeightTwips - 2500
    End If
    
    On Error GoTo 0
    Exit Sub

positionPrefsMonitor_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsMonitor of Form widgetPrefs"
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 01/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo btnResetMessages_MouseMove_Error

    If gblPrefsTooltips = "0" Then CreateToolTip btnResetMessages.hWnd, "The various pop-up messages that this program generates can be manually hidden. This button restores them to their original visible state.", _
                  TTIconInfo, "Help on the message reset button", , , , True

    On Error GoTo 0
    Exit Sub

btnResetMessages_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_MouseMove of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkDpiAwareness_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 14/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkDpiAwareness_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo chkDpiAwareness_Click_Error

    btnSave.Enabled = True ' enable the save button
    If pvtPrefsStartupFlg = False Then ' don't run this on startup
                    
        answer = vbYes
        answerMsg = "You must close this widget and HARD restart it, in order to change the widget's DPI awareness (a simple soft reload just won't cut it), do you want me to close and restart this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "DpiAwareness Confirmation", True, "chkDpiAwarenessRestart")
        
        If chkDpiAwareness.Value = 0 Then
            gblDpiAwareness = "0"
        Else
            gblDpiAwareness = "1"
        End If

        sPutINISetting "Software\PzTemperatureGauge", "dpiAwareness", gblDpiAwareness, gblSettingsFile
        
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "OK, the widget is still DPI aware until you restart. Some forms may show abnormally."
            answer = msgBoxA(answerMsg, vbOKOnly, "DpiAwareness Notification", True, "chkDpiAwarenessAbnormal")
        
            Exit Sub
        Else

            sPutINISetting "Software\PzTemperatureGauge", "dpiAwareness", gblDpiAwareness, gblSettingsFile
            'Call reloadProgram ' this is insufficient, image controls still fail to resize and autoscale correctly
            Call hardRestart
        End If

    End If

   On Error GoTo 0
   Exit Sub

chkDpiAwareness_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkDpiAwareness_Click of Form widgetPrefs"
End Sub


'
''---------------------------------------------------------------------------------------
'' Procedure : chkPrefsTooltips_Click
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 07/09/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub chkPrefsTooltips_Click()
'
'   On Error GoTo chkPrefsTooltips_Click_Error
'
'    btnSave.Enabled = True ' enable the save button
'
'    If pvtPrefsStartupFlg = False Then
'        If chkEnablePrefsTooltips.Value = 1 Then
'            gblEnablePrefsTooltips = "1"
'        Else
'            gblEnablePrefsTooltips = "0"
'        End If
'
'        sPutINISetting "Software\PzTemperatureGauge", "enablePrefsTooltips", gblEnablePrefsTooltips, gblSettingsFile
'
'    End If
'
'    ' set the tooltips on the prefs screen
'    Call setPrefsTooltips
'
'   On Error GoTo 0
'   Exit Sub
'
'chkEnablePrefsTooltips_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnablePrefsTooltips_Click of Form widgetPrefs"
'
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkShowTaskbar_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 19/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkShowTaskbar_Click()

   On Error GoTo chkShowTaskbar_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkShowTaskbar.Value = 1 Then
        gblShowTaskbar = "1"
    Else
        gblShowTaskbar = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowTaskbar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowTaskbar_Click of Form widgetPrefs"
End Sub



'Private Sub cmbTickSwitchPref_Click()
'   btnSave.Enabled = True ' enable the save button
'
'    If cmbTickSwitchPref.ListIndex = 0 Then
'        overlayAnemoWidget.pointerAnimate = False
'        overlayTemperatureWidget.pointerAnimate = False
'        gblPointerAnimate = "0"
'    Else
'        overlayAnemoWidget.pointerAnimate = True
'        overlayTemperatureWidget.pointerAnimate = True
'        gblPointerAnimate = "1"
'    End If
'
'    sPutINISetting "Software\PzAnemometerGauge", "pointerAnimate", gblPointerAnimate, gblSettingsFile
'    widgetPrefs.cmbTickSwitchPref.ListIndex = Val(gblPointerAnimate)
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 01/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_Click()

    On Error GoTo btnResetMessages_Click_Error
        
    ' Clear all the message box "show again" entries in the registry
    Call clearAllMessageBoxRegistryEntries
    
    MsgBox "Message boxes fully reset, confirmation pop-ups will continue as normal."

    On Error GoTo 0
    Exit Sub

btnResetMessages_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnAboutDebugInfo_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/03/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnAboutDebugInfo_Click()

   On Error GoTo btnAboutDebugInfo_Click_Error
   'If debugflg = 1 Then Debug.Print "%btnAboutDebugInfo_Click"

    'mnuDebug_Click
    MsgBox "The debug mode is not yet enabled."

   On Error GoTo 0
   Exit Sub

btnAboutDebugInfo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAboutDebugInfo_Click of form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDonate_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDonate_Click()
   On Error GoTo btnDonate_Click_Error

    Call mnuCoffee_ClickEvent

   On Error GoTo 0
   Exit Sub

btnDonate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDonate_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnFacebook_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnFacebook_Click()
   On Error GoTo btnFacebook_Click_Error
   'If debugflg = 1 Then DebugPrint "%btnFacebook_Click"

    Call menuForm.mnuFacebook_Click
    

   On Error GoTo 0
   Exit Sub

btnFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFacebook_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnOpenFile_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnOpenFile_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnOpenFile_Click_Error


    
    Call addTargetFile(txtOpenFile.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtOpenFile.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        'answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Create file confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnOpenFile_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnOpenFile_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnUpdate_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 29/02/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnUpdate_Click()
   On Error GoTo btnUpdate_Click_Error
   'If debugflg = 1 Then DebugPrint "%btnUpdate_Click"

    'MsgBox "The update button is not yet enabled."
    menuForm.mnuLatest_Click

   On Error GoTo 0
   Exit Sub

btnUpdate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdate_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkGaugeFunctions_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkGaugeFunctions_Click()
    On Error GoTo chkGaugeFunctions_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    ' disable polling
    WeatherMeteo.Ticking = chkGaugeFunctions.Value

    On Error GoTo 0
    Exit Sub

chkGaugeFunctions_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGaugeFunctions_Click of Form widgetPrefs"
End Sub

Private Sub chkGenStartup_Click()
    btnSave.Enabled = True ' enable the save button
End Sub



'---------------------------------------------------------------------------------------
' Procedure : subClassControls
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : sub classing code to capture form movement and intercept messages to the comboboxes to provide missing balloon tooltips functionality
'---------------------------------------------------------------------------------------
'
Private Sub subClassControls()
    
   On Error GoTo subClassControls_Error

    If InIDE Then
        MsgBox "NOTE: Running in IDE so Sub classing is disabled" & vbCrLf & "Mousewheel will not scroll icon maps and balloon tooltips will not display on comboboxes" & vbCrLf & vbCrLf & _
            "In addition, the display screen will not show messages as it currently crashes when run within the IDE."
    Else
        ' sub classing code to intercept messages to the form itself in order to capture WM_EXITSIZEMOVE messages that occur AFTER the form has been resized
        
        Call SubclassForm(widgetPrefs.hWnd, ObjPtr(widgetPrefs))
        
        'now the comboboxes in order to capture the mouseOver and display the balloon tooltips
        
        
        Call SubclassComboBox(cmbGaugeType.hWnd, ObjPtr(cmbGaugeType))
        Call SubclassComboBox(cmbTemperatureScale.hWnd, ObjPtr(cmbTemperatureScale))
        Call SubclassComboBox(cmbPressureScale.hWnd, ObjPtr(cmbPressureScale))
        Call SubclassComboBox(cmbWindSpeedScale.hWnd, ObjPtr(cmbWindSpeedScale))
        Call SubclassComboBox(cmbMetricImperial.hWnd, ObjPtr(cmbMetricImperial))
        
        'Call SubclassComboBox(cmbMultiMonitorResize.hWnd, ObjPtr(cmbMultiMonitorResize))
        Call SubclassComboBox(cmbScrollWheelDirection.hWnd, ObjPtr(cmbScrollWheelDirection))
        Call SubclassComboBox(cmbWindowLevel.hWnd, ObjPtr(cmbWindowLevel))
        Call SubclassComboBox(cmbHidingTime.hWnd, ObjPtr(cmbHidingTime))
        
        Call SubclassComboBox(cmbLandscapeLocked.hWnd, ObjPtr(cmbLandscapeLocked))
        'Call SubclassComboBox(cmbPortraitLocked.hWnd, ObjPtr(cmbPortraitLocked))
        Call SubclassComboBox(cmbWidgetPosition.hWnd, ObjPtr(cmbWidgetPosition))
        Call SubclassComboBox(cmbAspectHidden.hWnd, ObjPtr(cmbAspectHidden))
        Call SubclassComboBox(cmbDebug.hWnd, ObjPtr(cmbDebug))
                
    End If

    On Error GoTo 0
    Exit Sub

subClassControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure subClassControls of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MouseMoveOnComboText
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Add a balloon tooltip dynamically to combo boxes using subclassing, called by combobox_proc
'             (VB6 will not allow Elroy's advanced tooltips to show on VB6 comboboxes, we must subclass the controls)
'             Note: Each control must also be added to the subClassControls routine
'---------------------------------------------------------------------------------------
'
Public Sub MouseMoveOnComboText(sComboName As String)
    Dim sTitle As String
    Dim sText As String

    On Error GoTo MouseMoveOnComboText_Error
    
    Select Case sComboName
        Case "cmbTemperatureScale"
            sTitle = "Help on the Drop Down Temperature Scale"
            sText = pCmbTemperatureScaleBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbTemperatureScale.hWnd, sText, , sTitle, , , , True
        Case "cmbPressureScale"
            sTitle = "Help on the Drop Down Pressure Scale"
            sText = pCmbPressureScaleBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbPressureScale.hWnd, sText, , sTitle, , , , True
        Case "cmbWindSpeedScale"
            sTitle = "Help on the Drop Down Wind Speed Scale"
            sText = pCmbWindSpeedScaleBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbWindSpeedScale.hWnd, sText, , sTitle, , , , True
        Case "cmbMetricImperial"
            sTitle = "Help on the Measurement Standard Drop Down"
            sText = pCmbMetricImperialBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbMetricImperial.hWnd, sText, , sTitle, , , , True
'
        Case "cmbGaugeType"
            sTitle = "Help on the Position and Size Gauge Selector"
            sText = pCmbGaugeTypeBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbGaugeType.hWnd, sText, , sTitle, , , , True
    
'        Case "cmbMultiMonitorResize"
'            sTitle = "Help on the Drop Down Icon Filter"
'            sText = pCmbMultiMonitorResizeBalloonTooltip
'            If gblPrefsTooltips = "0" Then CreateToolTip cmbMultiMonitorResize.hWnd, sText, , sTitle, , , , True
        Case "cmbScrollWheelDirection"
            sTitle = "Help on the Scroll Wheel Direction"
            sText = pCmbScrollWheelDirectionBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbScrollWheelDirection.hWnd, sText, , sTitle, , , , True
        Case "cmbWindowLevel"
            sTitle = "Help on the Window Level"
            sText = pCmbWindowLevelBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbWindowLevel.hWnd, sText, , sTitle, , , , True
        Case "cmbHidingTime"
            sTitle = "Help on the Hiding Time"
            sText = pCmbHidingTimeBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbHidingTime.hWnd, sText, , sTitle, , , , True
            
        Case "cmbAspectHidden"
            sTitle = "Help on Hiding in Landscape/Portrait Mode"
            sText = pCmbAspectHiddenBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbAspectHidden.hWnd, sText, , sTitle, , , , True
        Case "cmbWidgetPosition"
            sTitle = "Help on Widget Position in Landscape/Portrait Modes"
            sText = pCmbWidgetPositionBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbWidgetPosition.hWnd, sText, , sTitle, , , , True
        Case "cmbLandscapeLocked"
            sTitle = "Help on Widget Locking in Landscape Mode"
            sText = pcmbLandscapeLockedBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbLandscapeLocked.hWnd, sText, , sTitle, , , , True
        Case "cmbPortraitLocked"
            sTitle = "Help on Widget Locking in Portrait Mode"
            sText = pcmbPortraitLockedBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbPortraitLocked.hWnd, sText, , sTitle, , , , True
        Case "cmbDebug"
            ' MsgBox "here " & sComboName & " " & gblPrefsTooltips & " " & pCmbDebugBalloonTooltip
        
            sTitle = "Help on Debug Mode"
            sText = pCmbDebugBalloonTooltip
            If gblPrefsTooltips = "0" Then CreateToolTip cmbDebug.hWnd, sText, , sTitle, , , , True
        

    End Select
    
   On Error GoTo 0
   Exit Sub

MouseMoveOnComboText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseMoveOnComboText of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnDefaultEditor_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnDefaultEditor_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnDefaultEditor_Click_Error

    Call addTargetFile(txtDefaultEditor.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtDefaultEditor.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Default Editor Confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnDefaultEditor_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDefaultEditor_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub




'---------------------------------------------------------------------------------------
' Procedure : chkIgnoreMouse_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkIgnoreMouse_Click()
   On Error GoTo chkIgnoreMouse_Click_Error

    If chkIgnoreMouse.Value = 0 Then
        gblIgnoreMouse = "0"
    Else
        gblIgnoreMouse = "1"
    End If

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkIgnoreMouse_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkIgnoreMouse_Click of Form widgetPrefs"
End Sub
Private Sub chkIgnoreMouse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkIgnoreMouse.hWnd, "Checking this box causes the program to ignore all mouse events. A strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Ignore Mouse optClockTooltips", , , , True
End Sub
'---------------------------------------------------------------------------------------
' Procedure : chkPreventDragging_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkPreventDragging_Click()
    On Error GoTo chkPreventDragging_Click_Error

    btnSave.Enabled = True ' enable the save button
'    If chkPreventDragging.Value = 0 Then
'        cmbLandscapeLocked.ListIndex = 0
'    Else
'        cmbLandscapeLocked.ListIndex = 1
'    End If
    
    If cmbGaugeType.ListIndex = 0 Then ' temperature
        ' immediately make the widget locked in place
        If chkPreventDragging.Value = 0 Then
            overlayTemperatureWidget.Locked = False
            gblPreventDraggingTemperature = "0"
            menuForm.mnuLockTemperatureGauge.Checked = False

            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = vbNullString
                txtLandscapeVoffset.Text = vbNullString
            Else
                txtPortraitHoffset.Text = vbNullString
                txtPortraitVoffset.Text = vbNullString
            End If
        Else
            overlayTemperatureWidget.Locked = True
            gblPreventDraggingTemperature = "1"
            menuForm.mnuLockTemperatureGauge.Checked = True

            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = fTemperature.temperatureGaugeForm.Left
                txtLandscapeVoffset.Text = fTemperature.temperatureGaugeForm.Top
            Else
                txtPortraitHoffset.Text = fTemperature.temperatureGaugeForm.Left
                txtPortraitVoffset.Text = fTemperature.temperatureGaugeForm.Top
            End If
        End If
    End If
    
        
    If cmbGaugeType.ListIndex = 1 Then ' anemometer
        ' immediately make the widget locked in place
        If chkPreventDragging.Value = 0 Then
            overlayAnemoWidget.Locked = False
            gblPreventDraggingAnemometer = "0"
            menuForm.mnuLockAnemometerGauge.Checked = False
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = vbNullString
                txtLandscapeVoffset.Text = vbNullString
            Else
                txtPortraitHoffset.Text = vbNullString
                txtPortraitVoffset.Text = vbNullString
            End If
        Else
            overlayAnemoWidget.Locked = True
            gblPreventDraggingAnemometer = "1"
            menuForm.mnuLockAnemometerGauge.Checked = True
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = fAnemometer.anemometerGaugeForm.Left
                txtLandscapeVoffset.Text = fAnemometer.anemometerGaugeForm.Top
            Else
                txtPortraitHoffset.Text = fAnemometer.anemometerGaugeForm.Left
                txtPortraitVoffset.Text = fAnemometer.anemometerGaugeForm.Top
            End If
        End If
    End If
    
            
    If cmbGaugeType.ListIndex = 2 Then ' Humidity
        ' immediately make the widget locked in place
        If chkPreventDragging.Value = 0 Then
            overlayHumidWidget.Locked = False
            gblPreventDraggingHumidity = "0"
            menuForm.mnuLockHumidityGauge.Checked = False
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = vbNullString
                txtLandscapeVoffset.Text = vbNullString
            Else
                txtPortraitHoffset.Text = vbNullString
                txtPortraitVoffset.Text = vbNullString
            End If
        Else
            overlayHumidWidget.Locked = True
            gblPreventDraggingHumidity = "1"
            menuForm.mnuLockHumidityGauge.Checked = True
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = fHumidity.humidityGaugeForm.Left
                txtLandscapeVoffset.Text = fHumidity.humidityGaugeForm.Top
            Else
                txtPortraitHoffset.Text = fHumidity.humidityGaugeForm.Left
                txtPortraitVoffset.Text = fHumidity.humidityGaugeForm.Top
            End If
        End If
    End If
            
    If cmbGaugeType.ListIndex = 3 Then ' Barometer
        ' immediately make the widget locked in place
        If chkPreventDragging.Value = 0 Then
            overlayBaromWidget.Locked = False
            gblPreventDraggingBarometer = "0"
            menuForm.mnuLockBarometerGauge.Checked = False
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = vbNullString
                txtLandscapeVoffset.Text = vbNullString
            Else
                txtPortraitHoffset.Text = vbNullString
                txtPortraitVoffset.Text = vbNullString
            End If
        Else
            overlayBaromWidget.Locked = True
            gblPreventDraggingBarometer = "1"
            menuForm.mnuLockBarometerGauge.Checked = True
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = fBarometer.barometerGaugeForm.Left
                txtLandscapeVoffset.Text = fBarometer.barometerGaugeForm.Top
            Else
                txtPortraitHoffset.Text = fBarometer.barometerGaugeForm.Left
                txtPortraitVoffset.Text = fBarometer.barometerGaugeForm.Top
            End If
        End If
    End If
    
            
    If cmbGaugeType.ListIndex = 4 Then ' Pictorial
        ' immediately make the widget locked in place
        If chkPreventDragging.Value = 0 Then
            overlayPictorialWidget.Locked = False
            gblPreventDraggingPictorial = "0"
            menuForm.mnuLockPictorialGauge.Checked = False
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = vbNullString
                txtLandscapeVoffset.Text = vbNullString
            Else
                txtPortraitHoffset.Text = vbNullString
                txtPortraitVoffset.Text = vbNullString
            End If
        Else
            overlayPictorialWidget.Locked = True
            gblPreventDraggingPictorial = "1"
            menuForm.mnuLockPictorialGauge.Checked = True
            If aspectRatio = "landscape" Then
                txtLandscapeHoffset.Text = fPictorial.pictorialGaugeForm.Left
                txtLandscapeVoffset.Text = fPictorial.pictorialGaugeForm.Top
            Else
                txtPortraitHoffset.Text = fPictorial.pictorialGaugeForm.Left
                txtPortraitVoffset.Text = fPictorial.pictorialGaugeForm.Top
            End If
        End If
    End If

    On Error GoTo 0
    Exit Sub

chkPreventDragging_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkPreventDragging_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkWidgetHidden_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkWidgetHidden_Click()
   On Error GoTo chkWidgetHidden_Click_Error

    If chkWidgetHidden.Value = 0 Then
        'overlayTemperatureWidget.Hidden = False
        fTemperature.temperatureGaugeForm.Visible = True

        frmTimer.revealWidgetTimer.Enabled = False
        gblWidgetHidden = "0"
    Else
        'overlayTemperatureWidget.Hidden = True
        fTemperature.temperatureGaugeForm.Visible = False


        frmTimer.revealWidgetTimer.Enabled = True
        gblWidgetHidden = "1"
    End If
    
    sPutINISetting "Software\PzTemperatureGauge", "widgetHidden", gblWidgetHidden, gblSettingsFile
    
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkWidgetHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWidgetHidden_Click of Form widgetPrefs"

End Sub

Private Sub chkWidgetHidden_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkWidgetHidden.hWnd, "Checking this box causes the program to hide for a certain number of minutes. More useful from the widget's right click menu where you can hide the widget at will. Seemingly, a strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Hidden option", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbAspectHidden_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbAspectHidden_Click()

   On Error GoTo cmbAspectHidden_Click_Error

    If cmbAspectHidden.ListIndex = 1 And aspectRatio = "portrait" Then
        'overlayTemperatureWidget.Hidden = True
        fTemperature.temperatureGaugeForm.Visible = False
    ElseIf cmbAspectHidden.ListIndex = 2 And aspectRatio = "landscape" Then
        'overlayTemperatureWidget.Hidden = True
        fTemperature.temperatureGaugeForm.Visible = False
    Else
        'overlayTemperatureWidget.Hidden = False
        fTemperature.temperatureGaugeForm.Visible = True
    End If

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbAspectHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbAspectHidden_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbDebug_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbDebug_Click()
    On Error GoTo cmbDebug_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbDebug.ListIndex = 0 Then
        txtDefaultEditor.Text = "eg. E:\vb6\Panzer Earth gauge VB6\Panzer Earth Gauge.vbp"
        txtDefaultEditor.Enabled = False
        
        lblDebug(7).Enabled = False
        btnDefaultEditor.Enabled = False
        
        lblDebug(9).Enabled = False
    Else
        #If TWINBASIC Then
            txtDefaultEditor.Text = gblDefaultTBEditor
        #Else
            txtDefaultEditor.Text = gblDefaultVB6Editor
        #End If
        txtDefaultEditor.Enabled = True
        
        lblDebug(7).Enabled = True
        btnDefaultEditor.Enabled = True
        
        lblDebug(9).Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbDebug_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDebug_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



Private Sub cmbHidingTime_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbScrollWheelDirection_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbScrollWheelDirection_Click()
   On Error GoTo cmbScrollWheelDirection_Click_Error

    btnSave.Enabled = True ' enable the save button
    'overlayTemperatureWidget.ZoomDirection = cmbScrollWheelDirection.List(cmbScrollWheelDirection.ListIndex)

   On Error GoTo 0
   Exit Sub

cmbScrollWheelDirection_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbScrollWheelDirection_Click of Form widgetPrefs"
End Sub



Private Sub cmbLandscapeLocked_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

Private Sub cmbPortraitLocked_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetPosition_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetPosition_Click()
    On Error GoTo cmbWidgetPosition_Click_Error

    btnSave.Enabled = True ' enable the save button
    If cmbWidgetPosition.ListIndex = 1 Then
        cmbLandscapeLocked.ListIndex = 0
        cmbPortraitLocked.ListIndex = 0
        cmbLandscapeLocked.Enabled = False
        cmbPortraitLocked.Enabled = False
        txtLandscapeHoffset.Enabled = False
        txtLandscapeVoffset.Enabled = False
        txtPortraitHoffset.Enabled = False
        txtPortraitVoffset.Enabled = False
        
    Else
        cmbLandscapeLocked.Enabled = True
        cmbPortraitLocked.Enabled = True
        txtLandscapeHoffset.Enabled = True
        txtLandscapeVoffset.Enabled = True
        txtPortraitHoffset.Enabled = True
        txtPortraitVoffset.Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbWidgetPosition_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetPosition_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Get IsLoaded() As Boolean
 
   On Error GoTo IsLoaded_Error

    IsLoaded = mIsLoaded

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Let IsLoaded(ByVal newValue As Boolean)
 
   On Error GoTo IsLoaded_Error

   If mIsLoaded <> newValue Then mIsLoaded = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property

'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 08/05/2023
' Purpose   : calling a manual property to a form allows external checks to the form to
'             determine whether it is loaded, without also activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If gblPrefsLoadedFlg Then
        If Me.WindowState = vbNormal Then
            IsVisible = Me.Visible
        Else
            IsVisible = False
        End If
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form widgetPrefs"
            Resume Next
          End If
    End With
End Property


'---------------------------------------------------------------------------------------
' Procedure : showLastTab
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : make the last used tab appear on startup
'---------------------------------------------------------------------------------------
'
Private Sub showLastTab()

   On Error GoTo showLastTab_Error
    
    If gblLastSelectedTab = "general" Then Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton)  ' was imgGeneralMouseUpEvent
    If gblLastSelectedTab = "config" Then Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton)     ' was imgConfigMouseUpEvent
    If gblLastSelectedTab = "position" Then Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    If gblLastSelectedTab = "development" Then Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    If gblLastSelectedTab = "fonts" Then Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    If gblLastSelectedTab = "sounds" Then Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    If gblLastSelectedTab = "window" Then Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    If gblLastSelectedTab = "about" Then Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)

   On Error GoTo 0
   Exit Sub

showLastTab_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showLastTab of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : positionPrefsFramesButtons
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 01/05/2023
' Purpose   : size and position the frames and buttons. Note we are NOT using control
'             arrays so the form can be converted to Cairo forms later.
'---------------------------------------------------------------------------------------
'
' for the future, when reading multiple buttons from XML config.
' read the XML prefs group and identify prefgroups - <prefGroup name="general" and count them.
'
' for each group read all the controls and identify those in the group - ie. preference group =
' for each specific group, identify the group image, title and order
' read those into an array
' use a for-loop (can't use foreach unless you read the results into a collection, foreach requires use of variant
'   elements as foreach needs an object or variant type to operate.
' create a group, fraHiding, image and text element and order in a class of yWidgetGroup
' create a button of yWidgetGroup for each group
' loop through each line and identify the controls belonging to the group

' for the moment though, we will do it manually
'
Private Sub positionPrefsFramesButtons()
    On Error GoTo positionPrefsFramesButtons_Error

    
    Dim frameWidth As Integer: frameWidth = 0
    
    Dim frameTop As Integer: frameTop = 0
    
    Dim frameLeft As Integer: frameLeft = 0
    
    Dim buttonTop As Integer:    buttonTop = 0
    'Dim currentFrameHeight As Integer: currentFrameHeight = 0
    
    Dim rightHandAlignment As Long: rightHandAlignment = 0
    
    Dim leftHandGutterWidth As Long: leftHandGutterWidth = 0
    
    ' align frames rightmost and leftmost to the buttons at the top
    buttonTop = -15
    frameTop = 1150
    leftHandGutterWidth = 240
    frameLeft = leftHandGutterWidth ' use the first frame leftmost as reference
    rightHandAlignment = fraAboutButton.Left + fraAboutButton.Width ' use final button rightmost as reference
    frameWidth = rightHandAlignment - frameLeft
    fraScrollbarCover.Left = rightHandAlignment - 690
    Me.Width = rightHandAlignment + leftHandGutterWidth + 75 ' (not quite sure why we need the 75 twips padding)
    
    ' align the top buttons
    fraGeneralButton.Top = buttonTop
    fraConfigButton.Top = buttonTop
    fraFontsButton.Top = buttonTop
    fraSoundsButton.Top = buttonTop
    fraPositionButton.Top = buttonTop
    fraDevelopmentButton.Top = buttonTop
    fraWindowButton.Top = buttonTop
    fraAboutButton.Top = buttonTop
    
    ' align the frames
    fraGeneral.Top = frameTop
    fraConfig.Top = frameTop
    fraFonts.Top = frameTop
    fraSounds.Top = frameTop
    fraPosition.Top = frameTop
    fraDevelopment.Top = frameTop
    fraWindow.Top = frameTop
    fraAbout.Top = frameTop
    
    fraGeneral.Left = frameLeft
    fraConfig.Left = frameLeft
    fraSounds.Left = frameLeft
    fraPosition.Left = frameLeft
    fraFonts.Left = frameLeft
    fraDevelopment.Left = frameLeft
    fraWindow.Left = frameLeft
    fraAbout.Left = frameLeft
    
    fraGeneral.Width = frameWidth
    fraConfig.Width = frameWidth
    fraSounds.Width = frameWidth
    fraPosition.Width = frameWidth
    fraFonts.Width = frameWidth
    fraWindow.Width = frameWidth
    fraDevelopment.Width = frameWidth
    fraAbout.Width = frameWidth
    
    ' set the base visibility of the frames
    fraGeneral.Visible = True
    fraConfig.Visible = False
    fraSounds.Visible = False
    fraPosition.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraDevelopment.Visible = False
    fraAbout.Visible = False
    
    fraGeneralButton.BorderStyle = 1
    
    #If TWINBASIC Then
        fraGeneralButton.Refresh
    #End If

    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - 50
    btnHelp.Left = frameLeft
    

   On Error GoTo 0
   Exit Sub

positionPrefsFramesButtons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsFramesButtons of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error

    btnSave.Enabled = False ' disable the save button
    Me.Hide
    Me.themeTimer.Enabled = False
    
    Call writePrefsPosition

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form widgetPrefs"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : display the help file
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    
    On Error GoTo btnHelp_Click_Error
    
        If fFExists(App.path & "\help\Help.chm") Then
            Call ShellExecute(Me.hWnd, "Open", App.path & "\help\Help.chm", vbNullString, App.path, 1)
        Else
            MsgBox ("%Err-I-ErrorNumber 11 - The help file - Help.chm - is missing from the help folder.")
        End If

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form widgetPrefs"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : btnSave_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : save the values from all the tabs
'---------------------------------------------------------------------------------------
'
Private Sub btnSave_Click()

'    Dim btnCnt As Integer: btnCnt = 0
'    Dim msgCnt As Integer: msgCnt = 0
'    Dim useloop As Integer: useloop = 0
'    Dim thisText As String: thisText = vbNullString
    
    On Error GoTo btnSave_Click_Error

    ' configuration
    gblGaugeTooltips = CStr(optClockTooltips(0).Tag)
    gblPrefsTooltips = CStr(optPrefsTooltips(0).Tag)
    
    'gblEnableBalloonTooltips = LTrim$(Str$(chkEnableBalloonTooltips.Value))
    gblShowTaskbar = LTrim$(Str$(chkShowTaskbar.Value))
    gblDpiAwareness = LTrim$(Str$(chkDpiAwareness.Value))
    
    'gblTemperatureGaugeSize = LTrim$(Str$(sliGaugeSize.Value))
    'gblAnemometerGaugeSize = LTrim$(Str$(sliAnemometerGaugeSize.Value))
    
'    gblClipbBSize = LTrim$(Str$(sliGaugeSize.Value))
'    gblSelectorSize = LTrim$(Str$(sliGaugeSize.Value))
    
    gblScrollWheelDirection = LTrim$(Str$(cmbScrollWheelDirection.ListIndex))
    
    ' general
    gblGaugeFunctions = LTrim$(Str$(chkGaugeFunctions.Value))
    gblStartup = LTrim$(Str$(chkGenStartup.Value))
    
    gblIcao = txtIcao.Text

    
    'gblPointerAnimate = cmbTickSwitchPref.ListIndex
    gblSamplingInterval = LTrim$(Str$(sliSamplingInterval.Value))
    gblStormTestInterval = LTrim$(Str$(sliStormTestInterval.Value))
    gblErrorInterval = LTrim$(Str$(sliErrorInterval.Value))
    
    gblAirportsURL = txtAirportsURL.Text
    
    gblTemperatureScale = cmbTemperatureScale.ListIndex
    gblPressureScale = cmbPressureScale.ListIndex
    gblMetricImperial = cmbMetricImperial.ListIndex
    gblWindSpeedScale = cmbWindSpeedScale.ListIndex
    
    ' sounds
    gblEnableSounds = LTrim$(Str$(chkEnableSounds.Value))
    
    'development
    gblDebug = LTrim$(Str$(cmbDebug.ListIndex))
    gblDblClickCommand = txtDblClickCommand.Text
    gblOpenFile = txtOpenFile.Text
    #If TWINBASIC Then
        gblDefaultTBEditor = txtDefaultEditor.Text
    #Else
        gblDefaultVB6Editor = txtDefaultEditor.Text
    #End If
    
    
    ' position
    gblAspectHidden = LTrim$(Str$(cmbAspectHidden.ListIndex))
    gblGaugeType = LTrim$(Str$(cmbGaugeType.ListIndex))
    
    gblWidgetPosition = LTrim$(Str$(cmbWidgetPosition.ListIndex))
        
    If cmbGaugeType.ListIndex = 0 Then ' temperature
        gblTemperatureLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        gblTemperaturePortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        gblTemperatureLandscapeLockedHoffset = txtLandscapeHoffset.Text
        gblTemperatureLandscapeLockedVoffset = txtLandscapeVoffset.Text
        gblTemperaturePortraitLockedHoffset = txtPortraitHoffset.Text
        gblTemperaturePortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    If cmbGaugeType.ListIndex = 1 Then ' anemometer
        gblAnemometerLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        gblAnemometerPortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        gblAnemometerLandscapeLockedHoffset = txtLandscapeHoffset.Text
        gblAnemometerLandscapeLockedVoffset = txtLandscapeVoffset.Text
        gblAnemometerPortraitLockedHoffset = txtPortraitHoffset.Text
        gblAnemometerPortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    If cmbGaugeType.ListIndex = 2 Then ' humidity
        gblHumidityLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        gblHumidityPortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        gblHumidityLandscapeLockedHoffset = txtLandscapeHoffset.Text
        gblHumidityLandscapeLockedVoffset = txtLandscapeVoffset.Text
        gblHumidityPortraitLockedHoffset = txtPortraitHoffset.Text
        gblHumidityPortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    If cmbGaugeType.ListIndex = 3 Then ' Barometer
        gblBarometerLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        gblBarometerPortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        gblBarometerLandscapeLockedHoffset = txtLandscapeHoffset.Text
        gblBarometerLandscapeLockedVoffset = txtLandscapeVoffset.Text
        gblBarometerPortraitLockedHoffset = txtPortraitHoffset.Text
        gblBarometerPortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    If cmbGaugeType.ListIndex = 4 Then ' Pictorial
        gblPictorialLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        gblPictorialPortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        gblPictorialLandscapeLockedHoffset = txtLandscapeHoffset.Text
        gblPictorialLandscapeLockedVoffset = txtLandscapeVoffset.Text
        gblPictorialPortraitLockedHoffset = txtPortraitHoffset.Text
        gblPictorialPortraitLockedVoffset = txtPortraitVoffset.Text
    End If
        
'    gblTemperatureVLocationPerc
'    gblTemperatureHLocationPerc

    ' fonts
    gblPrefsFont = txtPrefsFont.Text
    gblTempFormFont = gblPrefsFont
    
    ' the sizing is not saved here again as it saved during the setting phase.
    
'    If gblDpiAwareness = "1" Then
'        gblPrefsFontSizeHighDPI = txtPrefsFontSize.Text
'    Else
'        gblPrefsFontSizeLowDPI = txtPrefsFontSize.Text
'    End If
    'gblPrefsFontItalics = txtFontSize.Text

    ' Windows
    gblWindowLevel = LTrim$(Str$(cmbWindowLevel.ListIndex))
    gblOpacity = LTrim$(Str$(sliOpacity.Value))
    gblWidgetHidden = LTrim$(Str$(chkWidgetHidden.Value))
    gblHidingTime = LTrim$(Str$(cmbHidingTime.ListIndex))
    gblIgnoreMouse = LTrim$(Str$(chkIgnoreMouse.Value))
            

            
    If gblStartup = "1" Then
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PzTemperatureGaugeWidget", """" & App.path & "\" & "Panzer Temperature Gauge.exe""")
    Else
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PzTemperatureGaugeWidget", vbNullString)
    End If

    ' save the values from the general tab
    If fFExists(gblSettingsFile) Then
        sPutINISetting "Software\PzTemperatureGauge", "gaugeTooltips", gblGaugeTooltips, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsTooltips", gblPrefsTooltips, gblSettingsFile
        'sPutINISetting "Software\PzTemperatureGauge", "enableBalloonTooltips", gblEnableBalloonTooltips, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "showTaskbar", gblShowTaskbar, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "dpiAwareness", gblDpiAwareness, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "temperatureGaugeSize", gblTemperatureGaugeSize, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureLandscapeLocked", gblTemperatureLandscapeLocked, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperaturePortraitLocked", gblTemperaturePortraitLocked, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureVLocationPerc", gblTemperatureVLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureHLocationPerc", gblTemperatureHLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormHighDpiXPos", gblTemperatureFormHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormHighDpiYPos", gblTemperatureFormHighDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormLowDpiXPos", gblTemperatureFormLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormLowDpiYPos", gblTemperatureFormLowDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "preventDraggingTemperature", gblPreventDraggingTemperature, gblSettingsFile
        
        sPutINISetting "Software\PzAnemometerGauge", "anemometerGaugeSize", gblAnemometerGaugeSize, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerLandscapeLocked", gblAnemometerLandscapeLocked, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerPortraitLocked", gblAnemometerPortraitLocked, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerVLocationPerc", gblAnemometerVLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerHLocationPerc", gblAnemometerHLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormHighDpiXPos", gblAnemometerFormHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormHighDpiYPos", gblAnemometerFormHighDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormLowDpiXPos", gblAnemometerFormLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormLowDpiYPos", gblAnemometerFormLowDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "preventDraggingAnemometer", gblPreventDraggingAnemometer, gblSettingsFile
               
        sPutINISetting "Software\PzHumidityGauge", "humidityGaugeSize", gblHumidityGaugeSize, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityLandscapeLocked", gblHumidityLandscapeLocked, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityPortraitLocked", gblHumidityPortraitLocked, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityVLocationPerc", gblHumidityVLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityHLocationPerc", gblHumidityHLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormHighDpiXPos", gblHumidityFormHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormHighDpiYPos", gblHumidityFormHighDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormLowDpiXPos", gblHumidityFormLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormLowDpiYPos", gblHumidityFormLowDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "preventDraggingHumidity", gblPreventDraggingHumidity, gblSettingsFile
               
        sPutINISetting "Software\PzBarometerGauge", "barometerGaugeSize", gblBarometerGaugeSize, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerLandscapeLocked", gblBarometerLandscapeLocked, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerPortraitLocked", gblBarometerPortraitLocked, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerVLocationPerc", gblBarometerVLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerHLocationPerc", gblBarometerHLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormHighDpiXPos", gblBarometerFormHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormHighDpiYPos", gblBarometerFormHighDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormLowDpiXPos", gblBarometerFormLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormLowDpiYPos", gblBarometerFormLowDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "preventDraggingBarometer", gblPreventDraggingBarometer, gblSettingsFile
               
        sPutINISetting "Software\PzPictorialGauge", "pictorialGaugeSize", gblPictorialGaugeSize, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialLandscapeLocked", gblPictorialLandscapeLocked, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialPortraitLocked", gblPictorialPortraitLocked, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialVLocationPerc", gblPictorialVLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialHLocationPerc", gblPictorialHLocationPerc, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialFormHighDpiXPos", gblPictorialFormHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialFormHighDpiYPos", gblPictorialFormHighDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialFormLowDpiXPos", gblPictorialFormLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "pictorialFormLowDpiYPos", gblPictorialFormLowDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzPictorialGauge", "preventDraggingPictorial", gblPreventDraggingPictorial, gblSettingsFile


        sPutINISetting "Software\PzClipB", "clipBSize", gblClipBSize, gblSettingsFile
        sPutINISetting "Software\PzSelector", "selectorSize", gblSelectorSize, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "scrollWheelDirection", gblScrollWheelDirection, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "gaugeFunctions", gblGaugeFunctions, gblSettingsFile
'        sPutINISetting "Software\PzTemperatureGauge", "pointerAnimate", gblPointerAnimate, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "samplingInterval", gblSamplingInterval, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "stormTestInterval", gblStormTestInterval, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "errorInterval", gblErrorInterval, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "airportsURL", gblAirportsURL, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "temperatureScale", gblTemperatureScale, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "pressureScale", gblPressureScale, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "windSpeedScale", gblWindSpeedScale, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "metricImperial", gblMetricImperial, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "icao", gblIcao, gblSettingsFile

        
        sPutINISetting "Software\PzTemperatureGauge", "aspectHidden", gblAspectHidden, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "gaugeType", gblGaugeType, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "widgetPosition", gblWidgetPosition, gblSettingsFile
        

        sPutINISetting "Software\PzTemperatureGauge", "prefsFont", gblPrefsFont, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "tempFormFont", gblTempFormFont, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeHighDPI", gblPrefsFontSizeHighDPI, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeLowDPI", gblPrefsFontSizeLowDPI, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontItalics", gblPrefsFontItalics, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontColour", gblPrefsFontColour, gblSettingsFile

        'save the values from the Windows Config Items
        sPutINISetting "Software\PzTemperatureGauge", "windowLevel", gblWindowLevel, gblSettingsFile
        
        
        sPutINISetting "Software\PzTemperatureGauge", "opacity", gblOpacity, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "widgetHidden", gblWidgetHidden, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "hidingTime", gblHidingTime, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "ignoreMouse", gblIgnoreMouse, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "startup", gblStartup, gblSettingsFile

        sPutINISetting "Software\PzTemperatureGauge", "enableSounds", gblEnableSounds, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "lastSelectedTab", gblLastSelectedTab, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "debug", gblDebug, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "dblClickCommand", gblDblClickCommand, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "openFile", gblOpenFile, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "defaultVB6Editor", gblDefaultVB6Editor, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "defaultTBEditor", gblDefaultTBEditor, gblSettingsFile
        
        sPutINISetting "Software\PzClipB", "clipBFormHighDpiXPos", gblClipBFormHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzClipB", "clipBFormHighDpiYPos", gblClipBFormHighDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzClipB", "clipBFormLowDpiXPos", gblClipBFormLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzClipB", "clipBFormLowDpiYPos", gblClipBFormLowDpiYPos, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormHighDpiXPos", gblSelectorFormHighDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormHighDpiYPos", gblSelectorFormHighDpiYPos, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormLowDpiXPos", gblSelectorFormLowDpiXPos, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormLowDpiYPos", gblSelectorFormLowDpiYPos, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "lastUpdated", gblLastUpdated, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "metarPref", gblMetarPref, gblSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "oldPressureStorage", gblOldPressureStorage, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "pressureStorageDate", gblPressureStorageDate, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "currentPressureValue", gblCurrentPressureValue, gblSettingsFile
        
        'save the values from the Text Items

'        btnCnt = 0
'        msgCnt = 0
    End If
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

    ' sets the characteristics of the gauge and menus immediately after saving
    Call adjustTempMainControls
    
    Me.SetFocus
    btnSave.Enabled = False ' disable the save button showing it has successfully saved
    
    ' reload here if the gblWindowLevel Was Changed
    If gblWindowLevelWasChanged = True Then
        gblWindowLevelWasChanged = False
        Call reloadProgram
    Else
         'WeatherMeteo.GetMetar = True
        Call WeatherMeteo.getData
    End If
    
   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSave_Click of Form widgetPrefs"

End Sub

Private Sub chkEnableSounds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkEnableSounds.hWnd, "Check this box to enable or disable all of the sounds used during any animation on the main steampunk GUI, as well as all other chimes, tick sounds &c.", _
                  TTIconInfo, "Help on Enabling/Disabling Sounds", , , , True
End Sub

Private Sub chkEnableSounds_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'Private Sub cmbRefreshInterval_Click()
'    btnSave.Enabled = True ' enable the save button
'End Sub

Private Sub cmbWindowLevel_Click()
    btnSave.Enabled = True ' enable the save button
    If pvtPrefsStartupFlg = False Then gblWindowLevelWasChanged = True
End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnPrefsFont_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnPrefsFont_Click()

    
    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    
    Dim fntColour As Long: fntColour = 0
    
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    On Error GoTo btnPrefsFont_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    ' set the preliminary vars to feed and populate the changefont routine
    fntFont = gblPrefsFont
    ' gblTempFormFont
    
    If gblDpiAwareness = "1" Then
        fntSize = Val(gblPrefsFontSizeHighDPI)
    Else
        fntSize = Val(gblPrefsFontSizeLowDPI)
    End If
    
    If fntSize = 0 Then fntSize = 8
    fntItalics = CBool(gblPrefsFontItalics)
    fntColour = CLng(gblPrefsFontColour)
        
    
    Call changeFont(widgetPrefs, True, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
    
    gblPrefsFont = CStr(fntFont)
    gblTempFormFont = gblPrefsFont
    
    If gblDpiAwareness = "1" Then
        gblPrefsFontSizeHighDPI = CStr(fntSize)
        Call Form_Resize
    Else
        gblPrefsFontSizeLowDPI = CStr(fntSize)
    End If
    
    gblPrefsFontItalics = CStr(fntItalics)
    gblPrefsFontColour = CStr(fntColour)

    If fFExists(gblSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\PzTemperatureGauge", "prefsFont", gblPrefsFont, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "tempFormFont", gblTempFormFont, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeHighDPI", gblPrefsFontSizeHighDPI, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeLowDPI", gblPrefsFontSizeLowDPI, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontItalics", gblPrefsFontItalics, gblSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "PrefsFontColour", gblPrefsFontColour, gblSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "arial"
    txtPrefsFont.Text = fntFont
    txtPrefsFont.Font.Name = fntFont
    'txtPrefsFont.Font.Size = fntSize
    txtPrefsFont.Font.Italic = fntItalics
    txtPrefsFont.ForeColor = fntColour
    
    txtPrefsFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnPrefsFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrefsFont_Click of Form widgetPrefs"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsControls()
    
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    
    Dim sliTemperatureGaugeSizeOldValue As Long: sliTemperatureGaugeSizeOldValue = 0
    
    'Dim testDST As Boolean: testDST = False
    
    On Error GoTo adjustPrefsControls_Error
            
    ' general tab
    chkGaugeFunctions.Value = Val(gblGaugeFunctions)
    chkGenStartup.Value = Val(gblStartup)
'
'    txtBias.Visible = False
'    lblGeneral(1).Visible = False

    'If testDST = True Then cmbMainDaylightSaving.ListIndex = 1
    
    'set the choice for four timezone comboboxes that were populated from file.
'    cmbMainGaugeTimeZone.ListIndex = Val(gblMainGaugeTimeZone)
'    cmbMainDaylightSaving.ListIndex = Val(gblMainDaylightSaving)
        
    'txtBias.Text = tzDelta

'    cmbTickSwitchPref.ListIndex = Val(gblPointerAnimate)
    
    cmbTemperatureScale.ListIndex = Val(gblTemperatureScale)
    cmbPressureScale.ListIndex = Val(gblPressureScale)
    cmbWindSpeedScale.ListIndex = Val(gblWindSpeedScale)
    cmbMetricImperial.ListIndex = Val(gblMetricImperial)
 
    sliSamplingInterval.Value = Val(gblSamplingInterval)
    sliStormTestInterval.Value = Val(gblStormTestInterval)
    sliErrorInterval.Value = Val(gblErrorInterval)
     
    txtAirportsURL.Text = gblAirportsURL
    
    txtIcao.Text = gblIcao
    
    ' configuration tab
   
    ' check whether the size has been previously altered via ctrl+mousewheel on the widget
    sliTemperatureGaugeSizeOldValue = sliGaugeSize.Value
    'sliGaugeSize.Value = Val(gblTemperatureGaugeSize) 'deaniebabe
    'sliAnemometerGaugeSize.Value = val(gblAnemometerGaugeSize)
    
'    sliClipBSize.Value = val(gblClipBSize)
'    sliSelectorSize.Value = val(gblSelectorSize)

    If sliGaugeSize.Value <> sliTemperatureGaugeSizeOldValue Then
        btnSave.Visible = True
    End If
    
    cmbScrollWheelDirection.ListIndex = Val(gblScrollWheelDirection)
    
    optClockTooltips(CStr(gblGaugeTooltips)).Value = True
    optClockTooltips(0).Tag = CStr(gblGaugeTooltips)
    optClockTooltips(1).Tag = CStr(gblGaugeTooltips)
    optClockTooltips(2).Tag = CStr(gblGaugeTooltips)
        
    optPrefsTooltips(CStr(gblPrefsTooltips)).Value = True
    optPrefsTooltips(0).Tag = CStr(gblPrefsTooltips)
    optPrefsTooltips(1).Tag = CStr(gblPrefsTooltips)
    optPrefsTooltips(2).Tag = CStr(gblPrefsTooltips)
        
    'chkEnableTooltips.Value = Val(gblGaugeTooltips)
    'chkEnableBalloonTooltips.Value = Val(gblEnableBalloonTooltips)
    chkShowTaskbar.Value = Val(gblShowTaskbar)
    chkDpiAwareness.Value = Val(gblDpiAwareness)
    
    ' chkPrefsTooltips .Value = Val(gblPrefsTooltips)
    
    ' sounds tab
    chkEnableSounds.Value = Val(gblEnableSounds)
    
    ' development
    cmbDebug.ListIndex = Val(gblDebug)
    txtDblClickCommand.Text = gblDblClickCommand
    txtOpenFile.Text = gblOpenFile
    #If TWINBASIC Then
        txtDefaultEditor.Text = gblDefaultTBEditor
    #Else
        txtDefaultEditor.Text = gblDefaultVB6Editor
    #End If
    lblGitHub.Caption = "You can find the code for the Panzer Weather Gauges on github, visit by double-clicking this link https://github.com/yereverluvinunclebert/Panzer-Weather-Temperature-Gauge-VB6"
  
     ' fonts tab
    If gblPrefsFont <> vbNullString Then
        txtPrefsFont.Text = gblPrefsFont
        If gblDpiAwareness = "1" Then
            Call changeFormFont(widgetPrefs, gblPrefsFont, Val(gblPrefsFontSizeHighDPI), fntWeight, fntStyle, gblPrefsFontItalics, gblPrefsFontColour)
            txtPrefsFontSize.Text = gblPrefsFontSizeHighDPI
        Else
            Call changeFormFont(widgetPrefs, gblPrefsFont, Val(gblPrefsFontSizeLowDPI), fntWeight, fntStyle, gblPrefsFontItalics, gblPrefsFontColour)
            txtPrefsFontSize.Text = gblPrefsFontSizeLowDPI
        End If
    End If
    
    
    ' position tab
    cmbAspectHidden.ListIndex = Val(gblAspectHidden)
    cmbGaugeType.ListIndex = Val(gblGaugeType)
    
    cmbWidgetPosition.ListIndex = Val(gblWidgetPosition)
        
    If gblPreventDraggingTemperature = "1" Then
        If aspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = fTemperature.temperatureGaugeForm.Left
            txtLandscapeVoffset.Text = fTemperature.temperatureGaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fTemperature.temperatureGaugeForm.Left
            txtPortraitVoffset.Text = fTemperature.temperatureGaugeForm.Top
            If gblDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gblTemperatureFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gblTemperatureFormLowDpiYPos & "px"
            End If
        End If
    End If
    
    'cmbLandscapeLocked
    
    If cmbGaugeType.ListIndex = 0 Then ' temperature
        cmbLandscapeLocked.ListIndex = Val(gblTemperatureLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(gblTemperaturePortraitLocked)
        txtLandscapeHoffset.Text = gblTemperatureLandscapeLockedHoffset
        txtLandscapeVoffset.Text = gblTemperatureLandscapeLockedVoffset
        txtPortraitHoffset.Text = gblTemperaturePortraitLockedHoffset
        txtPortraitVoffset.Text = gblTemperaturePortraitLockedVoffset
    End If
    
    If cmbGaugeType.ListIndex = 1 Then ' Anemometer
        cmbLandscapeLocked.ListIndex = Val(gblAnemometerLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(gblAnemometerPortraitLocked)
        txtLandscapeHoffset.Text = gblAnemometerLandscapeLockedHoffset
        txtLandscapeVoffset.Text = gblAnemometerLandscapeLockedVoffset
        txtPortraitHoffset.Text = gblAnemometerPortraitLockedHoffset
        txtPortraitVoffset.Text = gblAnemometerPortraitLockedVoffset
    End If

    If cmbGaugeType.ListIndex = 2 Then ' Humidity
        cmbLandscapeLocked.ListIndex = Val(gblHumidityLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(gblHumidityPortraitLocked)
        txtLandscapeHoffset.Text = gblHumidityLandscapeLockedHoffset
        txtLandscapeVoffset.Text = gblHumidityLandscapeLockedVoffset
        txtPortraitHoffset.Text = gblHumidityPortraitLockedHoffset
        txtPortraitVoffset.Text = gblHumidityPortraitLockedVoffset
    End If
      
    If cmbGaugeType.ListIndex = 3 Then ' Barometer
        cmbLandscapeLocked.ListIndex = Val(gblBarometerLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(gblBarometerPortraitLocked)
        txtLandscapeHoffset.Text = gblBarometerLandscapeLockedHoffset
        txtLandscapeVoffset.Text = gblBarometerLandscapeLockedVoffset
        txtPortraitHoffset.Text = gblBarometerPortraitLockedHoffset
        txtPortraitVoffset.Text = gblBarometerPortraitLockedVoffset
    End If
    
    If cmbGaugeType.ListIndex = 4 Then ' Pictorial
        cmbLandscapeLocked.ListIndex = Val(gblPictorialLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(gblPictorialPortraitLocked)
        txtLandscapeHoffset.Text = gblPictorialLandscapeLockedHoffset
        txtLandscapeVoffset.Text = gblPictorialLandscapeLockedVoffset
        txtPortraitHoffset.Text = gblPictorialPortraitLockedHoffset
        txtPortraitVoffset.Text = gblPictorialPortraitLockedVoffset
    End If
        
    ' Windows tab
    cmbWindowLevel.ListIndex = Val(gblWindowLevel)
    chkIgnoreMouse.Value = Val(gblIgnoreMouse)
    
    'chkPreventDragging.Value = Val(gblPreventDraggingTemperature)
    'chkPreventDragging.Value = val(gblPreventDraggingAnemometer)
    'gblPreventDraggingHumidity
    
    cmbGaugeType.ListIndex = 0
    
    sliOpacity.Value = Val(gblOpacity)
    chkWidgetHidden.Value = Val(gblWidgetHidden)
    cmbHidingTime.ListIndex = Val(gblHidingTime)
    
        
   On Error GoTo 0
   Exit Sub

adjustPrefsControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsControls of Form widgetPrefs on line " & Erl

End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : populatePrefsComboBoxes
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/09/2022
' Purpose   : all combo boxes in the prefs are populated here with default values
'           : done by preference here rather than in the IDE
'---------------------------------------------------------------------------------------

Private Sub populatePrefsComboBoxes()
    'Dim ret As Boolean: ret = False
    
    On Error GoTo populatePrefsComboBoxes_Error
    
    ' obtain the daylight savings time data from the system
'    ret = fGetTimeZoneArray
'    If ret = False Then MsgBox "Problem getting the Daylight Savings Time data from the system."

    cmbScrollWheelDirection.AddItem "up", 0
    cmbScrollWheelDirection.ItemData(0) = 0
    cmbScrollWheelDirection.AddItem "down", 1
    cmbScrollWheelDirection.ItemData(1) = 1
    
    cmbAspectHidden.AddItem "none", 0
    cmbAspectHidden.ItemData(0) = 0
    cmbAspectHidden.AddItem "portrait", 1
    cmbAspectHidden.ItemData(1) = 1
    cmbAspectHidden.AddItem "landscape", 2
    cmbAspectHidden.ItemData(2) = 2
    
    cmbGaugeType.AddItem "Temperature Gauge", 0
    cmbGaugeType.ItemData(0) = 0
    cmbGaugeType.AddItem "Anemometer Gauge", 1
    cmbGaugeType.ItemData(1) = 1
    cmbGaugeType.AddItem "Humidity Gauge", 2
    cmbGaugeType.ItemData(2) = 2
    cmbGaugeType.AddItem "Barometer Gauge", 3
    cmbGaugeType.ItemData(3) = 3
    cmbGaugeType.AddItem "Pictorial Gauge", 4
    cmbGaugeType.ItemData(4) = 4

    cmbWidgetPosition.AddItem "disabled", 0
    cmbWidgetPosition.ItemData(0) = 0
    cmbWidgetPosition.AddItem "enabled", 1
    cmbWidgetPosition.ItemData(1) = 1
    
    cmbLandscapeLocked.AddItem "disabled", 0
    cmbLandscapeLocked.ItemData(0) = 0
    cmbLandscapeLocked.AddItem "enabled", 1
    cmbLandscapeLocked.ItemData(1) = 1
    
    cmbPortraitLocked.AddItem "disabled", 0
    cmbPortraitLocked.ItemData(0) = 0
    cmbPortraitLocked.AddItem "enabled", 1
    cmbPortraitLocked.ItemData(1) = 1
    
    cmbDebug.AddItem "Debug OFF", 0
    cmbDebug.ItemData(0) = 0
    cmbDebug.AddItem "Debug ON", 1
    cmbDebug.ItemData(1) = 1
    
    ' populate comboboxes in the windows tab
    cmbWindowLevel.AddItem "Keep on top of other windows", 0
    cmbWindowLevel.ItemData(0) = 0
    cmbWindowLevel.AddItem "Normal", 0
    cmbWindowLevel.ItemData(1) = 1
    cmbWindowLevel.AddItem "Keep below all other windows", 0
    cmbWindowLevel.ItemData(2) = 2

    ' populate the hiding timer combobox
    cmbHidingTime.AddItem "1 minute", 0
    cmbHidingTime.ItemData(0) = 1
    cmbHidingTime.AddItem "5 minutes", 1
    cmbHidingTime.ItemData(1) = 5
    cmbHidingTime.AddItem "10 minutes", 2
    cmbHidingTime.ItemData(2) = 10
    cmbHidingTime.AddItem "20 minutes", 3
    cmbHidingTime.ItemData(3) = 20
    cmbHidingTime.AddItem "30 minutes", 4
    cmbHidingTime.ItemData(4) = 30
    cmbHidingTime.AddItem "I hour", 5
    cmbHidingTime.ItemData(5) = 60
    
'    cmbTickSwitchPref.AddItem "Flick", 0
'    cmbTickSwitchPref.ItemData(0) = 0
'    cmbTickSwitchPref.AddItem "Smooth", 1
'    cmbTickSwitchPref.ItemData(1) = 1
    
    cmbTemperatureScale.AddItem "centigrade/celsius", 0
    cmbTemperatureScale.ItemData(0) = 0
    cmbTemperatureScale.AddItem "fahrenheit", 1
    cmbTemperatureScale.ItemData(1) = 1
    cmbTemperatureScale.AddItem "kelvin", 2
    cmbTemperatureScale.ItemData(2) = 2

    cmbPressureScale.AddItem "millibars", 0
    cmbPressureScale.ItemData(0) = 0
    cmbPressureScale.AddItem "inches of mercury (hg)", 1
    cmbPressureScale.ItemData(1) = 1
    cmbPressureScale.AddItem "mm of mercury (mmhg)", 2
    cmbPressureScale.ItemData(2) = 2
    cmbPressureScale.AddItem "hectoPascals", 3
    cmbPressureScale.ItemData(2) = 3

    cmbWindSpeedScale.AddItem "knots", 0
    cmbWindSpeedScale.ItemData(0) = 0
    cmbWindSpeedScale.AddItem "metres", 1
    cmbWindSpeedScale.ItemData(1) = 1
    
    cmbMetricImperial.AddItem "Imperial", 0
    cmbMetricImperial.ItemData(0) = 0
    cmbMetricImperial.AddItem "Metric", 1
    cmbMetricImperial.ItemData(1) = 1

    On Error GoTo 0
    Exit Sub

populatePrefsComboBoxes_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populatePrefsComboBoxes of Form widgetPrefs"
            Resume Next
          End If
    End With
                
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readFileWriteComboBox
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 28/07/2023
' Purpose   : Open and load the Array with the timezones text File
'---------------------------------------------------------------------------------------
'
'Private Sub readFileWriteComboBox(ByRef thisComboBox As Control, ByVal thisFileName As String)
'    Dim strArr() As String
'    Dim lngCount As Long: lngCount = 0
'    Dim lngIdx As Long: lngIdx = 0
'
'    On Error GoTo readFileWriteComboBox_Error
'
'    If fFExists(thisFileName) = True Then
'       ' the files must be DOS CRLF delineated
'       Open thisFileName For Input As #1
'
'           strArr() = Split(Input(LOF(1), 1), vbCrLf)
'       Close #1
'
'       lngCount = UBound(strArr)
'
'       '@Ignore MemberNotOnInterface
'       thisComboBox.Clear
'       For lngIdx = 0 To lngCount
'           '@Ignore MemberNotOnInterface
'           thisComboBox.AddItem strArr(lngIdx)
'       Next lngIdx
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'readFileWriteComboBox_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readFileWriteComboBox of Form widgetPrefs"
'
'End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : clearBorderStyle
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : removes all styling from the icon frames and makes the major frames below invisible too, not using control arrays.
'---------------------------------------------------------------------------------------
'
Private Sub clearBorderStyle()

   On Error GoTo clearBorderStyle_Error

    fraGeneral.Visible = False
    fraConfig.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraPosition.Visible = False
    fraDevelopment.Visible = False
    fraSounds.Visible = False
    fraAbout.Visible = False

    fraGeneralButton.BorderStyle = 0
    fraConfigButton.BorderStyle = 0
    fraDevelopmentButton.BorderStyle = 0
    fraPositionButton.BorderStyle = 0
    fraFontsButton.BorderStyle = 0
    fraWindowButton.BorderStyle = 0
    fraSoundsButton.BorderStyle = 0
    fraAboutButton.BorderStyle = 0
    
    #If TWINBASIC Then
        fraGeneralButton.Refresh
        fraConfigButton.Refresh
        fraDevelopmentButton.Refresh
        fraPositionButton.Refresh
        fraFontsButton.Refresh
        fraWindowButton.Refresh
        fraSoundsButton.Refresh
        fraAboutButton.Refresh
    #End If

   On Error GoTo 0
   Exit Sub

clearBorderStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearBorderStyle of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/05/2023
' Purpose   : If the form is NOT to be resized then restrain the height/width. Otherwise,
'             maintain the aspect ratio. When minimised and a resize is called then simply exit.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()
    
    On Error GoTo Form_Resize_Error
    
    Call PrefsForm_Resize_Event
    
    On Error GoTo 0
    Exit Sub

Form_Resize_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrefsForm_Resize_Event
' Author    : beededea
' Date      : 10/10/2024
' Purpose   :
'
'---------------------------------------------------------------------------------------
Public Sub PrefsForm_Resize_Event()

    Dim currentFontSize As Long: currentFontSize = 0
    Dim ratio As Double: ratio = 0
    Dim currentFont As Long: currentFont = 0
    
    On Error GoTo PrefsForm_Resize_Event_Error
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' move the drag corner label along with the form's bottom right corner
    lblDragCorner.Move Me.ScaleLeft + Me.ScaleWidth - (lblDragCorner.Width + 40), _
               Me.ScaleTop + Me.ScaleHeight - (lblDragCorner.Height + 40)

    ratio = cPrefsFormHeight / cPrefsFormWidth
    
    If pvtPrefsDynamicSizingFlg = True Then
        
        If gblDpiAwareness = "1" Then
            currentFont = gblPrefsFontSizeHighDPI
        Else
            currentFont = gblPrefsFontSizeLowDPI
        End If
        
        Call resizeControls(Me, prefsControlPositions(), prefsCurrentWidth, prefsCurrentHeight, currentFont)
        
        Call tweakPrefsControlPositions(Me, prefsCurrentWidth, prefsCurrentHeight)
        
        Me.Width = Me.Height / ratio ' maintain the aspect ratio

    Else
        If Me.WindowState = 0 Then ' normal
            If Me.Width > 9090 Then Me.Width = 9090
            If Me.Width < 9085 Then Me.Width = 9090
            If pvtLastFormHeight <> 0 Then Me.Height = pvtLastFormHeight
        End If
    End If
    'lblSize.Caption = "topIconWidth = " & topIconWidth & " imgGeneral width = " & imgGeneral.Width
    
   On Error GoTo 0
   Exit Sub

PrefsForm_Resize_Event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrefsForm_Resize_Event of Form widgetPrefs"

End Sub
        

'---------------------------------------------------------------------------------------
' Procedure : tweakPrefsControlPositions
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 22/09/2023
' Purpose   : final tweak the bottom frame top and left positions
'---------------------------------------------------------------------------------------
'
Private Sub tweakPrefsControlPositions(ByVal thisForm As Form, ByVal m_FormWid As Single, ByVal m_FormHgt As Single)

    ' not sure why but the resizeControls routine can lead to incorrect positioning of frames and buttons
    
    Dim x_scale As Single: x_scale = 0
    
    Dim y_scale As Single: y_scale = 0
    
    On Error GoTo tweakPrefsControlPositions_Error

    ' Get the form's current scale factors.
    x_scale = thisForm.ScaleWidth / m_FormWid
    y_scale = thisForm.ScaleHeight / m_FormHgt

    fraGeneral.Left = fraGeneralButton.Left
    fraConfig.Left = fraGeneralButton.Left
    fraSounds.Left = fraGeneralButton.Left
    fraPosition.Left = fraGeneralButton.Left
    fraFonts.Left = fraGeneralButton.Left
    fraDevelopment.Left = fraGeneralButton.Left
    fraWindow.Left = fraGeneralButton.Left
    fraAbout.Left = fraGeneralButton.Left
         
    'fraGeneral.Top = fraGeneralButton.Top
    fraConfig.Top = fraGeneral.Top
    fraSounds.Top = fraGeneral.Top
    fraPosition.Top = fraGeneral.Top
    fraFonts.Top = fraGeneral.Top
    fraDevelopment.Top = fraGeneral.Top
    fraWindow.Top = fraGeneral.Top
    fraAbout.Top = fraGeneral.Top
    
    ' final tweak the bottom button positions
    
    btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (200 * y_scale)
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnHelp.Top
    
    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - (150 * x_scale)
    btnHelp.Left = fraGeneral.Left

    txtPrefsFontCurrentSize.Text = y_scale * txtPrefsFontCurrentSize.FontSize
    
   On Error GoTo 0
   Exit Sub

tweakPrefsControlPositions_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tweakPrefsControlPositions of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

    gblPrefsLoadedFlg = False
    
    Call writePrefsPosition
    
    Call DestroyToolTip

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form widgetPrefs"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraScrollbarCover.Visible = True

End Sub
Private Sub fraAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraScrollbarCover.Visible = True
    If gblPrefsTooltips = "0" Then CreateToolTip fraAbout.hWnd, "The About tab tells you all about this program and its creation using " & gblCodingEnvironment & ".", _
                  TTIconInfo, "Help on the About Tab", , , , True
End Sub
Private Sub fraConfigInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfigInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraConfigInner.hWnd, "The configuration panel is the location for optional configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True

End Sub
Private Sub fraConfig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraConfig.hWnd, "The configuration panel is the location for optional configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True

End Sub

Private Sub fraDefaultEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblGitHub.ForeColor = &H80000012
End Sub

Private Sub fraDevelopment_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraDevelopment.hWnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True
End Sub


Private Sub fraDevelopmentInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopmentInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraDevelopmentInner.hWnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True

End Sub
Private Sub fraFonts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraFonts.hWnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True

End Sub
Private Sub fraFontsInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraFontsInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraFontsInner.hWnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gbl program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True
End Sub



' ----------------------------------------------------------------
' Procedure Name: fraGaugePosition_MouseMove
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter Button (Integer):
' Parameter Shift (Integer):
' Parameter X (Single):
' Parameter Y (Single):
' Author: beededea
' Date: 08/05/2024
' ----------------------------------------------------------------
Private Sub fraGaugePosition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo fraGaugePosition_MouseMove_Error
    
    If gblPrefsTooltips = "0" Then CreateToolTip fraGaugePosition.hWnd, "Select the gauge type first - then this section allows you to determine " _
        & "the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet. " _
        & "This functionality is a hangover from the Yahoo/Konfabulator widget of the same name that was created when Windows tablets were briefly " _
        & "a 'thing'. Who uses Windows Tablets nowadays anyway?" & vbCrLf _
        & "Note: Each gauge can be locked in place using the gauge's locked button (top left), in this case the X and Y " _
        & "values are populated automatically.", _
                  TTIconInfo, "Help on Gauge Positioning", , , , True

    On Error GoTo 0
    Exit Sub

fraGaugePosition_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fraGaugePosition_MouseMove, line " & Erl & "."

End Sub

'Private Sub fraConfigurationButtonInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 2 Then
'        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
'    End If
'End Sub
Private Sub fraGeneral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

End Sub
Private Sub fraGeneral_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraGeneral.hWnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraGeneralInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraGeneralInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraGeneralInner.hWnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraPosition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraPosition.hWnd, "This section allows you to determine size, lockability and positioning of your widget in various ways on different screen aspect ratios. ", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub



Private Sub fraPositionBalloonBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip fraPositionBalloonBox.hWnd, "Aspect ratio is for tablets only. Don't fiddle with this unless you really know what you are doing. Here you can choose whether this widget is hidden by default in either landscape or portrait mode or not at all. This option allows you to have certain widgets that do not obscure the screen in either landscape or portrait. If you accidentally set it so you can't find your widget on screen then change the setting here to NONE.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub

Private Sub fraPositionInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraScrollbarCover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraScrollbarCover.Visible = False

End Sub
Private Sub fraSounds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSounds_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If gblPrefsTooltips = "0" Then CreateToolTip fraSounds.hWnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraSoundsInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSoundsInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraSoundsInner.hWnd, "The sound panel allows you to configure the sounds that occur within gbl. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub

Private Sub fraWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraWindow.hWnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub
Private Sub fraWindowInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindowInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If gblPrefsTooltips = "0" Then CreateToolTip fraWindowInner.hWnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub




Private Sub imgGeneral_Click()
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lblGitHub_dblClick
' Author    : beededea
' Date      : 14/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblGitHub_dblClick()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo lblGitHub_dblClick_Error

    answer = vbYes
    answerMsg = "This option opens a browser window and take you straight to Github. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Proceed to Github? ", True, "lblGitHubDblClick")
    If answer = vbYes Then
       Call ShellExecute(Me.hWnd, "Open", "https://github.com/yereverluvinunclebert/Panzer-Weather-Temperature-Gauge-" & gblCodingEnvironment, vbNullString, App.path, 1)
    End If

   On Error GoTo 0
   Exit Sub

lblGitHub_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGitHub_dblClick of Form widgetPrefs"
End Sub


Private Sub lblGitHub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblGitHub.ForeColor = &H8000000D
End Sub

'
''---------------------------------------------------------------------------------------
'' Procedure : chkEnableTooltips_Click
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 19/08/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub chkEnableTooltips_Click()
'    Dim answer As VbMsgBoxResult: answer = vbNo
'    Dim answerMsg As String: answerMsg = vbNullString
'    On Error GoTo chkEnableTooltips_Click_Error
'
'    btnSave.Enabled = True ' enable the save button
'
'    If pvtPrefsStartupFlg = False Then
'        If chkEnableTooltips.Value = 1 Then
'            gblGaugeTooltips = "1"
'        Else
'            gblPrefsTooltips = "0"
'        End If
'
'        sPutINISetting "Software\PzTemperatureGauge", "enableTooltips", gblGaugeTooltips, gblSettingsFile
'
'        answer = vbYes
'        answerMsg = "You must soft reload this widget, in order to change the tooltip setting, do you want me to reload this widget? I can do it now for you."
'        answer = msgBoxA(answerMsg, vbYesNo, "Request to Enable Tooltips", True, "chkEnableTooltipsClick")
'        If answer = vbNo Then
'            Exit Sub
'        Else
'            Call reloadProgram
'        End If
'    End If
'
'
'   On Error GoTo 0
'   Exit Sub
'
'chkEnableTooltips_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableTooltips_Click of Form widgetPrefs"
'
'End Sub

''---------------------------------------------------------------------------------------
'' Procedure : chkEnableBalloonTooltips_Click
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 09/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub chkEnableBalloonTooltips_Click()
'   On Error GoTo chkEnableBalloonTooltips_Click_Error
'
'    btnSave.Enabled = True ' enable the save button
'    If chkEnableBalloonTooltips.Value = 1 Then
'        gblPrefsTooltips = "0"
'    Else
'        gblEnableBalloonTooltips = "0"
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'chkEnableBalloonTooltips_Click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableBalloonTooltips_Click of Form widgetPrefs"
'End Sub



'Private Sub chkEnableBalloonTooltips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If gblPrefsTooltips = "0" Then CreateToolTip chkEnableBalloonTooltips.hwnd, "Best to enable the balloon tooltips rather than the standard ones. Do that here. The balloon tooltips are much prettier, have more space for pertinent information so I can fill them up with useful text to assist you.", _
'                  TTIconInfo, "Help on the Balloon Tooltips", , , , True
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : chkEnableTooltips_MouseMove
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 05/10/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub chkEnableTooltips_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    On Error GoTo chkEnableTooltips_MouseMove_Error
'
'    If gblPrefsTooltips = "0" Then CreateToolTip chkEnableTooltips.hwnd, "There is a problem with the 'standard' tooltips on the gauge elements as they resize along with the program graphical elements, meaning that they cannot be seen, there is also a problem with tooltip handling different fonts, hoping to get Olaf to fix these soon. My suggestion is to turn them off for the moment.", _
'                  TTIconInfo, "Help on the Program tooltip problem", , , , True
'
'    On Error GoTo 0
'    Exit Sub
'
'chkEnableTooltips_MouseMove_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableTooltips_MouseMove of Form widgetPrefs"
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : optClockTooltips_Click
' Author    : beededea
' Date      : 19/08/2023
' Purpose   : three options radio buttons for selecting the clock/cal tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optClockTooltips_Click(Index As Integer)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo optClockTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button

    If pvtPrefsStartupFlg = False Then
        gblGaugeTooltips = CStr(Index)
    
        optClockTooltips(0).Tag = CStr(Index)
        optClockTooltips(1).Tag = CStr(Index)
        optClockTooltips(2).Tag = CStr(Index)
        
        sPutINISetting "Software\PzTemperatureGauge", "gaugeTooltips", gblGaugeTooltips, gblSettingsFile

        answer = vbYes
        answerMsg = "You must soft reload this widget, in order to change the tooltip setting, do you want me to reload this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "Request to Enable Tooltips", True, "optClockTooltipsClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call reloadProgram
        End If
    End If

   On Error GoTo 0
   Exit Sub

optClockTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optClockTooltips_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : optClockTooltips_MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : setting the tooltip text for the specific radio button for selecting the clock/cal tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optClockTooltips_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim thisToolTip As String: thisToolTip = vbNullString
    On Error GoTo optClockTooltips_MouseMove_Error

    If gblPrefsTooltips = "0" Then
        If Index = 0 Then
            thisToolTip = "This setting enables the balloon tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips, note that their font size will match the Windows system font size."
            CreateToolTip optClockTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Balloon Tooltips on the GUI", , , , True
        ElseIf Index = 1 Then
            thisToolTip = "This setting enables the RichClient square tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips."
            CreateToolTip optClockTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on RichClient Tooltips on the GUI", , , , True
        ElseIf Index = 2 Then
            thisToolTip = "This setting disables the balloon tooltips for elements within the Steampunk GUI."
            CreateToolTip optClockTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Disabling Tooltips on the GUI", , , , True
        End If
    
    End If

   On Error GoTo 0
   Exit Sub

optClockTooltips_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optClockTooltips_MouseMove of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : optPrefsTooltips_Click
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : three options radio buttons for selecting the VB6 preference form tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optPrefsTooltips_Click(Index As Integer)

   On Error GoTo optPrefsTooltips_Click_Error

    If pvtPrefsStartupFlg = False Then
    
        btnSave.Enabled = True ' enable the save button
        gblPrefsTooltips = CStr(Index)
        optPrefsTooltips(0).Tag = CStr(Index)
        optPrefsTooltips(1).Tag = CStr(Index)
        optPrefsTooltips(2).Tag = CStr(Index)
        
        sPutINISetting "Software\PzTemperatureGauge", "prefsTooltips", gblPrefsTooltips, gblSettingsFile
        
        ' set the tooltips on the prefs screen
        Call setPrefsTooltips
    End If
     
   On Error GoTo 0
   Exit Sub

optPrefsTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optPrefsTooltips_Click of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optPrefsTooltips_MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : series of radio buttons to set the tooltip type for the prefs utility
'---------------------------------------------------------------------------------------
'
Private Sub optPrefsTooltips_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim thisToolTip As String: thisToolTip = vbNullString

    On Error GoTo optPrefsTooltips_MouseMove_Error

    If gblPrefsTooltips = "0" Then
        If Index = 0 Then
            thisToolTip = "This setting enables the balloon tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips, note that their font size will match the Windows system font size."
            CreateToolTip optPrefsTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Balloon Tooltips on the Preference Utility", , , , True
        ElseIf Index = 1 Then
            thisToolTip = "This setting enables the standard Windows-style square tooltips for elements within the Steampunk GUI. These tooltips are single-line and the font size is limited to the Windows font size."
            CreateToolTip optPrefsTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on VB6 Native Tooltips on the Preference Utility", , , , True
        ElseIf Index = 2 Then
            thisToolTip = "This setting disables the balloon tooltips for elements within the Steampunk GUI."
            CreateToolTip optPrefsTooltips(Index).hWnd, thisToolTip, _
                  TTIconInfo, "Help on Disabling Tooltips on the Preference Utility", , , , True
        End If
    End If

   On Error GoTo 0
   Exit Sub

optPrefsTooltips_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optPrefsTooltips_MouseMove of Form widgetPrefs"
End Sub



' ----------------------------------------------------------------
' Procedure Name: sliErrorInterval_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 19/04/2024
' ----------------------------------------------------------------
Private Sub sliErrorInterval_Click()
    On Error GoTo sliErrorInterval_Click_Error
    btnSave.Enabled = True ' enable the save button
 
    If pvtPrefsStartupFlg = False Then
        gblErrorInterval = LTrim$(Str$(sliErrorInterval.Value))
        'overlayTemperatureWidget.samplingInterval = sliSamplingInterval.Value
        sPutINISetting "Software\PzTemperatureGauge", "sliErrorInterval", gblErrorInterval, gblSettingsFile
    End If
    
    On Error GoTo 0
    Exit Sub

sliErrorInterval_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliErrorInterval_Click, line " & Erl & "."

End Sub

Private Sub sliErrorInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip sliErrorInterval.hWnd, "Adjust to error reporting frequency (seconds). This is the interval by which the program determines if a feed is in error by failing to supply data. If the interval is reached and the feed provides no data, then an error message is displayed. A value of zero means that no error messages will be displayed.", _
                  TTIconInfo, "Help on the Error Interval", , , , True
End Sub

Private Sub sliGaugeSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip sliGaugeSize.hWnd, "Changing this slider will change the size of the chosen gauge selected above", _
                  TTIconInfo, "Help on the Gauge Size Slider", , , , True
End Sub

Private Sub sliOpacity_Change()
    btnSave.Enabled = True ' enable the save button
End Sub
Private Sub sliOpacity_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip sliOpacity.hWnd, "Sliding this causes the program's opacity to change from solidly opaque to fully transparent or some way in-between. Seemingly, a strange option for a windows program, a useful left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Opacity Slider", , , , True

End Sub
' ----------------------------------------------------------------
' Procedure Name: sliSamplingInterval_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 10/01/2024
' ----------------------------------------------------------------
Private Sub sliSamplingInterval_Click()
    On Error GoTo sliSamplingInterval_Click_Error
    btnSave.Enabled = True ' enable the save button

 
    If pvtPrefsStartupFlg = False Then
        gblSamplingInterval = LTrim$(Str$(sliSamplingInterval.Value))
        WeatherMeteo.samplingInterval = sliSamplingInterval.Value
        sPutINISetting "Software\PzTemperatureGauge", "samplingInterval", gblSamplingInterval, gblSettingsFile
        
    End If
    
    On Error GoTo 0
    Exit Sub

sliSamplingInterval_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliSamplingInterval_Click, line " & Erl & "."

End Sub

Private Sub sliSamplingInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip sliSamplingInterval.hWnd, "Adjust to determine gauge sampling frequency (seconds). This is the polling interval by which the widget attempts to get the data from the source (default 600 seconds or ten minutes). The metar source provider and the location itself determines when the data is actually provided - there are weather sensors that have to provide rea-time data and a real person somewhere is probably responsible for providing an actual forecast.*", _
                  TTIconInfo, "Help on Sampling Interval", , , , True

End Sub


' ----------------------------------------------------------------
' Procedure Name: sliStormTestInterval_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 19/04/2024
' ----------------------------------------------------------------
Private Sub sliStormTestInterval_Click()
    On Error GoTo sliStormTestInterval_Click_Error
    btnSave.Enabled = True ' enable the save button
 
    If pvtPrefsStartupFlg = False Then
        gblStormTestInterval = LTrim$(Str$(sliStormTestInterval.Value))
        'overlayTemperatureWidget.samplingInterval = sliStormTestInterval.Value
        sPutINISetting "Software\PzTemperatureGauge", "stormTestInterval", gblStormTestInterval, gblSettingsFile
        
    End If

    
    On Error GoTo 0
    Exit Sub

sliStormTestInterval_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStormTestInterval_Click, line " & Erl & "."

End Sub

Private Sub sliStormTestInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip sliStormTestInterval.hWnd, "Adjust to determine storm checking frequency (seconds). This is the interval by which the widget compares pressure drops of 1 millibar (default 3600 seconds or one hour) indicating the increased chance of a storm. If this condition is detected, it will light a red lamp on the barometer gauge", _
                  TTIconInfo, "Help on Storm Test Interval", , , , True

End Sub

'---------------------------------------------------------------------------------------
' Procedure : txtAboutText_MouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub txtAboutText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo txtAboutText_MouseDown_Error

    If Button = vbRightButton Then
        txtAboutText.Enabled = False
        txtAboutText.Enabled = True
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

txtAboutText_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtAboutText_MouseDown of Form widgetPrefs"
End Sub

Private Sub txtAboutText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraScrollbarCover.Visible = False
End Sub

Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgAbout.Visible = False
    imgAboutClicked.Visible = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : General _MouseUp events to generate menu pop-ups across the form
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : due to a bug/difference with TwinBasic versus VB6
'---------------------------------------------------------------------------------------
#If TWINBASIC Then
    Private Sub imgAboutClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
    End Sub
#Else
    Private Sub imgAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgDevelopmentClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    End Sub
#Else
    Private Sub imgDevelopment_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgFontsClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    End Sub
#Else
    Private Sub imgFonts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgConfigClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
    End Sub
#Else
    Private Sub imgConfig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgPositionClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    End Sub
#Else
    Private Sub imgPosition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgSoundsClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    End Sub
#Else
    Private Sub imgSounds_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgWindowClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    End Sub
#Else
    Private Sub imgWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    End Sub
#End If

#If TWINBASIC Then
    Private Sub imgGeneralClicked_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
    End Sub
#Else
    Private Sub imgGeneral_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
    End Sub
#End If




Private Sub imgDevelopment_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDevelopment.Visible = False
    imgDevelopmentClicked.Visible = True
End Sub


Private Sub imgFonts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFonts.Visible = False
    imgFontsClicked.Visible = True
End Sub



Private Sub imgConfig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgConfig.Visible = False
    imgConfigClicked.Visible = True
End Sub



Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub


Private Sub imgPosition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPosition.Visible = False
    imgPositionClicked.Visible = True
End Sub



Private Sub imgSounds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    imgSounds.Visible = False
    imgSoundsClicked.Visible = True
End Sub



Private Sub imgWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgWindow.Visible = False
    imgWindowClicked.Visible = True
End Sub



'Private Sub sliAnimationInterval_Change()
'    'overlayTemperatureWidget.RotationSpeed = sliAnimationInterval.Value
'    btnSave.Enabled = True ' enable the save button
'
'End Sub



Private Sub sliGaugeSize_GotFocus()
    gblAllowSizeChangeFlg = True
End Sub

Private Sub sliGaugeSize_LostFocus()
    gblAllowSizeChangeFlg = False
End Sub
'---------------------------------------------------------------------------------------
' Procedure : sliGaugeSize_Change
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliGaugeSize_Change()
    On Error GoTo sliGaugeSize_Change_Error

    btnSave.Enabled = True ' enable the save button
    
    If cmbGaugeType.ListIndex = 0 Then ' fTemperature gauge
        If gblAllowSizeChangeFlg = True Then Call fTemperature.tempAdjustZoom(sliGaugeSize.Value / 100)
    End If
    
    If cmbGaugeType.ListIndex = 1 Then ' fAnemometer gauge
        If gblAllowSizeChangeFlg = True Then Call fAnemometer.anemoAdjustZoom(sliGaugeSize.Value / 100)
    End If
    
    If cmbGaugeType.ListIndex = 2 Then ' fHumidity gauge
        If gblAllowSizeChangeFlg = True Then Call fHumidity.humidAdjustZoom(sliGaugeSize.Value / 100)
    End If

    If cmbGaugeType.ListIndex = 3 Then ' barometer gauge
        If gblAllowSizeChangeFlg = True Then Call fBarometer.baromAdjustZoom(sliGaugeSize.Value / 100)
    End If

    If cmbGaugeType.ListIndex = 4 Then ' pictorial gauge
        If gblAllowSizeChangeFlg = True Then Call fPictorial.pictAdjustZoom(sliGaugeSize.Value / 100)
    End If

    On Error GoTo 0
    Exit Sub

sliGaugeSize_Change_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliGaugeSize_Change of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 15/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo sliOpacity_Change_Error

    btnSave.Enabled = True ' enable the save button

    If pvtPrefsStartupFlg = False Then
        gblOpacity = LTrim$(Str$(sliOpacity.Value))
    
        sPutINISetting "Software\PzTemperatureGauge", "opacity", gblOpacity, gblSettingsFile
        
        'Call setOpacity(sliOpacity.Value) ' this works but reveals the background form itself
        
        answer = vbYes
        answerMsg = "You must perform a hard reload on this widget in order to change the widget's opacity, do you want me to do it for you now?"
        answer = msgBoxA(answerMsg, vbYesNo, "Hard Reload Request", True, "sliOpacityClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call hardRestart
        End If
    End If

   On Error GoTo 0
   Exit Sub

sliOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliOpacity_Change of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 14/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then
        gblOriginatingForm = "widgetPrefsForm"
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
    

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form widgetPrefs"
End Sub

'Private Sub fraEmail_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
'    If Button = 2 Then
'        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
'    End If
'End Sub

'Private Sub fraEmojis_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
'    If Button = 2 Then
'        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
'    End If
'End Sub

Private Sub fraFonts_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub








Private Sub txtAirportsURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtAirportsURL.hWnd, "Don't change this unless your alternative has an airports file with the " _
        & "EXACT same format and layout. this is the full URL giving the location of the airports.dat file containing all the airfields in the world " _
        & "that ought to be providing a valid METAR data feed. This URL will be used when you select the menu option to download a new Airports.dat file.", _
                  TTIconInfo, "Airports ICAO Data Download", , , , True
End Sub

Private Sub txtDblClickCommand_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDblClickCommand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtDblClickCommand.hWnd, "Field to hold the any double click command that you have assigned to this widget. For example: taskmgr or %systemroot%\syswow64\ncpa.cpl", _
                  TTIconInfo, "Help on the Double Click Command", , , , True
End Sub

Private Sub txtDefaultEditor_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDefaultEditor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtDefaultEditor.hWnd, "Field to hold the path to a Visual Basic Project (VBP) file you would like to execute on a right click menu, edit option, if you select the adjacent button a file explorer will appear allowing you to select the VBP file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the Default Editor Field", , , , True
End Sub
Private Sub txtIcao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtIcao.hWnd, "This shows the current ICAO code used to identify the weather feed source data. To change this field use the button to the right", _
                  TTIconInfo, "Select the current ICAO code", , , , True
End Sub

Private Sub txtLandscapeHoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeVoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtOpenFile_Change()
    btnSave.Enabled = True
    ' enable the save button

End Sub

Private Sub txtOpenFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtOpenFile.hWnd, "Field to hold the path to a file you would like to execute on a shift+DBlClick, if you select the adjacent button a file explorer will appear allowing you to select any file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the shift+DBlClick Field", , , , True
End Sub

Private Sub txtPortraitHoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPortraitVoffset_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPrefsFont_Change()
    btnSave.Enabled = True ' enable the save button
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 13/02/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click()
    
    On Error GoTo mnuAbout_Click_Error

    Call aboutClickEvent

    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsTooltips
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setPrefsTooltips()

   On Error GoTo setPrefsTooltips_Error
   
    ' here we set the variables used for the comboboxes, each combobox has to be sub classed and these variables are used during that process
 
    If gblPrefsTooltips = "0" Then
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        pCmbMultiMonitorResizeBalloonTooltip = "This option will only appear on multi-monitor systems. This dropdown has three choices that affect the automatic sizing of both the main clock and the preference utility. " & vbCrLf & vbCrLf & _
            "For monitors of different sizes, this allows you to resize the widget to suit the monitor it is currently sitting on. The automatic option resizes according to the relative proportions of the two screens.  " & vbCrLf & vbCrLf & _
            "The manual option resizes according to sizes that you set manually. Just resize the clock on the monitor of your choice and the program will store it. This option only works for no more than TWO monitors."
   
        pCmbTemperatureScaleBalloonTooltip = "Select the temperature unit. The default is the celsius scale, the alternatives are fahrenheit and kelvin."
        pCmbPressureScaleBalloonTooltip = "Select the pressure scale you are familiar with."
        pCmbWindSpeedScaleBalloonTooltip = "Select the anemometer unit of scale. The default is the knots scale, the alternative is metres/sec."
        pCmbMetricImperialBalloonTooltip = "Select metric or imperial with regard to cloud cover ONLY."
        
        pCmbGaugeTypeBalloonTooltip = "Select the weather gauge for which you wish to amend the details below"
    
           
        pCmbScrollWheelDirectionBalloonTooltip = "This option will allow you to change the direction of the mouse scroll wheel when resizing the clock gauge. IF you want to resize the clock on your desktop, hold the CTRL key along with moving the scroll wheel UP/DOWN. Some prefer scrolling UP rather than DOWN. You configure that here."
        pCmbWindowLevelBalloonTooltip = "You can determine the window level here. You can keep it above all other windows or you can set it to bottom to keep the widget below all other windows."
        pCmbHidingTimeBalloonTooltip = "The hiding time that you can set here determines how long the widget will disappear when you click the menu option to hide the widget."
        
        pcmbLandscapeLockedBalloonTooltip = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        pcmbPortraitLockedBalloonTooltip = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        pCmbWidgetPositionBalloonTooltip = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        pCmbAspectHiddenBalloonTooltip = "Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        
        pCmbDebugBalloonTooltip = "Here you can set debug mode. This will enable the editor field and allow you to assign a VBP/TwinProj file for the " & gblCodingEnvironment & " IDE editor"
        
        pCmbAlarmDayBalloonTooltip = "Enter a valid day of the month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        pCmbAlarmMonthBalloonTooltip = "Enter a valid month here. When you have entered both a date here and a time in the adjacent field, then click the > key to validate."
        pCmbAlarmYearBalloonTooltip = "Enter a valid year here. When you have entered both a valid year here and a time in the adjacent fields, then click the > key to validate."
        pCmbAlarmHoursBalloonTooltip = "Enter a valid hour here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here in these two fields, then click the > key to validate."
        pCmbAlarmMinutesBalloonTooltip = "Enter valid minutes here by typing a future time in 24hr military format, eg: 23:45. When you have entered both a date in the previous fields and a time here, then click the > key to validate."
    Else
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        
        pCmbTemperatureScaleBalloonTooltip = vbNullString
        pCmbPressureScaleBalloonTooltip = vbNullString
        pCmbWindSpeedScaleBalloonTooltip = vbNullString
        pCmbMetricImperialBalloonTooltip = vbNullString
        
        pCmbMultiMonitorResizeBalloonTooltip = vbNullString
        pCmbScrollWheelDirectionBalloonTooltip = vbNullString
        pCmbWindowLevelBalloonTooltip = vbNullString
        pCmbHidingTimeBalloonTooltip = vbNullString
        
        pcmbLandscapeLockedBalloonTooltip = vbNullString
        pcmbPortraitLockedBalloonTooltip = vbNullString
        pCmbWidgetPositionBalloonTooltip = vbNullString
        pCmbAspectHiddenBalloonTooltip = vbNullString
        pCmbDebugBalloonTooltip = vbNullString
        pCmbAlarmDayBalloonTooltip = vbNullString
        pCmbAlarmMonthBalloonTooltip = vbNullString
        pCmbAlarmYearBalloonTooltip = vbNullString
        
        pCmbAlarmHoursBalloonTooltip = vbNullString
        pCmbAlarmMinutesBalloonTooltip = vbNullString
        
        ' for some reason, the balloon tooltip on the checkbox used to dismiss the balloon tooltips does not disappear, this forces it go away.
        CreateToolTip optPrefsTooltips(0).hWnd, "", _
                  TTIconInfo, "Help", , , , True
        CreateToolTip optPrefsTooltips(1).hWnd, "", _
                  TTIconInfo, "Help", , , , True
        CreateToolTip optPrefsTooltips(2).hWnd, "", _
                  TTIconInfo, "Help", , , , True
                  
    End If

     ' next we just do the native VB6 tooltips
    If gblPrefsTooltips = "1" Then
        imgConfig.ToolTipText = "Opens the configuration tab"
        imgConfigClicked.ToolTipText = "Opens the configuration tab"
        imgDevelopment.ToolTipText = "Opens the Development tab"
        imgDevelopmentClicked.ToolTipText = "Opens the Development tab"
        imgPosition.ToolTipText = "Opens the Position tab"
        imgPositionClicked.ToolTipText = "Opens the Position tab"
        btnSave.ToolTipText = "Save the changes you have made to the preferences"
        btnHelp.ToolTipText = "Open the help utility"
        imgSounds.ToolTipText = "Opens the Sounds tab"
        imgSoundsClicked.ToolTipText = "Opens the Sounds tab"
        btnClose.ToolTipText = "Close the utility"
        imgWindow.ToolTipText = "Opens the Window tab"
        imgWindowClicked.ToolTipText = "Opens the Window tab"
        lblWindow.ToolTipText = "Opens the Window tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFontsClicked.ToolTipText = "Opens the Fonts tab"
        imgGeneral.ToolTipText = "Opens the general tab"
        imgGeneralClicked.ToolTipText = "Opens the general tab"
        
        lblPosition(6).ToolTipText = "Tablets only. Don't fiddle with this unless you really know what you are doing. Here you can choose whether this the widget widget is hidden by default in either landscape or portrait mode or not at all. This option allows you to have certain widgets that do not obscure the screen in either landscape or portrait. If you accidentally set it so you can't find your widget on screen then change the setting here to NONE."
        chkGenStartup.ToolTipText = "Check this box to enable the automatic start of the program when Windows is started."
        chkGaugeFunctions.ToolTipText = "When checked this box enables the pointer. Any adjustment takes place instantly. "
'        sliAnimationInterval.ToolTipText = "Adjust to make the animation smooth or choppy. Any adjustment in the interval takes place instantly. Lower values are smoother but the smoother it runs the more Temperature it uses."
        txtPortraitVoffset.ToolTipText = "Field to hold the vertical offset for the widget position in portrait mode."
        txtPortraitHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in portrait mode."
        txtLandscapeVoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        txtLandscapeHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        cmbLandscapeLocked.ToolTipText = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        cmbPortraitLocked.ToolTipText = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        cmbWidgetPosition.ToolTipText = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        cmbAspectHidden.ToolTipText = " Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        chkEnableSounds.ToolTipText = "Check this box to enable or disable all of the sounds used during any animation on the main screen."
        btnDefaultEditor.ToolTipText = "Click to select the .vbp file to edit the program - You need to have access to the source!"
        txtDblClickCommand.ToolTipText = "Enter a Windows command for the gauge to operate when double-clicked."
        btnOpenFile.ToolTipText = "Click to select a particular file for the gauge to run or open when double-clicked."
        txtOpenFile.ToolTipText = "Enter a particular file for the gauge to run or open when double-clicked."
        cmbDebug.ToolTipText = "Choose to set debug mode."
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
        btnPrefsFont.ToolTipText = "The Font Selector."
        txtPrefsFont.ToolTipText = "Disabled for manual input. Choose a font via the font selector to be used only for this preferences window"
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size via the font selector that fits the text boxes"
        cmbWindowLevel.ToolTipText = "You can determine the window position here. Set to bottom to keep the widget below other windows."
        cmbHidingTime.ToolTipText = "."
        chkEnableResizing.ToolTipText = "Provides an alternative method of supporting high DPI screens."
        chkPreventDragging.ToolTipText = "Checking this box turns off the ability to drag the program with the mouse. The locking in position effect takes place instantly."
        chkIgnoreMouse.ToolTipText = "Checking this box causes the program to ignore all mouse events."
        sliOpacity.ToolTipText = "Set the transparency of the program. Any change in opacity takes place instantly."
        sliSamplingInterval.ToolTipText = "Setting the sampling interval affects the frequency of the pointer updates."
        sliStormTestInterval.ToolTipText = "Adjust to determine storm checking frequency (seconds). This is the polling interval by which the widget compares pressure drops of 1 millibar (default 3600 seconds or one hour) indicating the increased chance of a storm."
        sliErrorInterval.ToolTipText = "Adjust to error reporting frequency (seconds). This is the interval by which the program determines if a feed is in error by failing to supply data. If the interval is reached and the feed provides no data, then an error message is displayed."
        cmbScrollWheelDirection.ToolTipText = "To change the direction of the mouse scroll wheel when resizing the gauge."
                        
        optClockTooltips(0).ToolTipText = "Check the box to enable larger balloon tooltips for all controls on the main program"
        optClockTooltips(1).ToolTipText = "Check the box to enable RichClient square tooltips for all controls on the main program"
        optClockTooltips(2).ToolTipText = "Check the box to disable tooltips for all controls on the main program"
                
        optPrefsTooltips(0).ToolTipText = "Check the box to enable larger balloon tooltips for all controls within this Preference Utility. These tooltips are multi-line and in general more attractive, note that their font size will match the Windows system font size."
        optPrefsTooltips(1).ToolTipText = "Check the box to enable Windows-style square tooltips for all controls within this Preference Utility. Note that their font size will match the Windows system font size."
        optPrefsTooltips(2).ToolTipText = "This setting enables/disables the tooltips for all elements within this Preference Utility."
        
        'chkEnableBalloonTooltips.ToolTipText = "Check the box to enable larger balloon tooltips for all controls on the main program"
        chkShowTaskbar.ToolTipText = "Check the box to show the widget in the taskbar"
        'chkEnableTooltips.ToolTipText = "Check the box to enable tooltips for all controls on the main program"
        sliGaugeSize.ToolTipText = "Adjust to a percentage of the original size. Any adjustment in size takes place instantly (you can also use Ctrl+Mousewheel hovering over the gauge itself)."
        'sliWidgetSkew.ToolTipText = "Adjust to a degree skew of the original position. Any adjustment in direction takes place instantly (you can also use the Mousewheel hovering over the gauge itself."
        btnFacebook.ToolTipText = "This will link you to the our Steampunk/Dieselpunk program users Group."
        imgAbout.ToolTipText = "Opens the About tab"
        btnAboutDebugInfo.ToolTipText = "This gives access to the debugging tool"
        btnDonate.ToolTipText = "Buy me a Kofi! This button opens a browser window and connects to Kofi donation page"
        btnUpdate.ToolTipText = "Here you can visit the update location where you can download new versions of the programs."
        
        lblFontsTab(0).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
'        lblFontsTab(1).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
'        lblFontsTab(2).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        
        lblFontsTab(6).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        
        lblFontsTab(7).ToolTipText = "Choose a font size that fits the text boxes"
        txtPrefsFontCurrentSize.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        lblCurrentFontsTab.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
'        cmbMainGaugeTimeZone.ToolTipText = "Select the timezone of your choice."
'        cmbMainDaylightSaving.ToolTipText = "Select and activate Daylight Savings Time for your area."
        
'        cmbTickSwitchPref.ToolTipText = "The movement of the hand can be set to smooth or regular flicks, the smooth movement uses slightly more Temperature."
        
        'lstTimezoneRegions.ToolTipText = "These are the regions associated with the chosen timezone."
        chkDpiAwareness.ToolTipText = " Check the box to make the program DPI aware. RESTART required."
        'chkPrefsTooltips.ToolTipText = "Check the box to enable tooltips for all controls in the preferences utility"
        btnResetMessages.ToolTipText = "This button restores the pop-up messages to their original visible state."
        
        cmbTemperatureScale.ToolTipText = "Select the temperature unit. The default is the celsius scale, the alternatives are fahrenheit and kelvin."
        cmbPressureScale.ToolTipText = "Select the scale you are familiar with."
        cmbWindSpeedScale.ToolTipText = "Select the anemometer unit of scale. The default is the knots scale, the alternative is metres/sec."
        cmbMetricImperial.ToolTipText = "Select metric or imperial with regard to cloud cover ONLY."
        
        txtIcao.ToolTipText = "This is the current ICAO code used to identify the weather feed. You can change it using the Select button to the right."
        btnLocation.ToolTipText = "Press to select the current ICAO code used to identify the weather feed."
    Else
        lblPosition(6).ToolTipText = vbNullString
        imgConfig.ToolTipText = vbNullString
        imgConfigClicked.ToolTipText = vbNullString
        imgDevelopment.ToolTipText = vbNullString
        imgDevelopmentClicked.ToolTipText = vbNullString
        imgPosition.ToolTipText = vbNullString
        imgPositionClicked.ToolTipText = vbNullString
        btnSave.ToolTipText = vbNullString
        btnHelp.ToolTipText = vbNullString
        imgSounds.ToolTipText = vbNullString
        imgSoundsClicked.ToolTipText = vbNullString
        btnClose.ToolTipText = vbNullString
        imgWindow.ToolTipText = vbNullString
        imgWindowClicked.ToolTipText = vbNullString
        imgFonts.ToolTipText = vbNullString
        imgFontsClicked.ToolTipText = vbNullString
        imgGeneral.ToolTipText = vbNullString
        imgGeneralClicked.ToolTipText = vbNullString
        chkGenStartup.ToolTipText = vbNullString
        chkGaugeFunctions.ToolTipText = vbNullString
'        sliAnimationInterval.ToolTipText = vbNullString
        txtPortraitVoffset.ToolTipText = vbNullString
        txtPortraitHoffset.ToolTipText = vbNullString
        txtLandscapeVoffset.ToolTipText = vbNullString
        txtLandscapeHoffset.ToolTipText = vbNullString
        cmbLandscapeLocked.ToolTipText = vbNullString
        cmbPortraitLocked.ToolTipText = vbNullString
        cmbWidgetPosition.ToolTipText = vbNullString
        cmbAspectHidden.ToolTipText = vbNullString
        chkEnableSounds.ToolTipText = vbNullString
        btnDefaultEditor.ToolTipText = vbNullString
        txtDblClickCommand.ToolTipText = vbNullString
        btnOpenFile.ToolTipText = vbNullString
        txtOpenFile.ToolTipText = vbNullString
        cmbDebug.ToolTipText = vbNullString
        txtPrefsFontSize.ToolTipText = vbNullString
        btnPrefsFont.ToolTipText = vbNullString
        txtPrefsFont.ToolTipText = vbNullString
        cmbWindowLevel.ToolTipText = vbNullString
        cmbHidingTime.ToolTipText = vbNullString
        chkEnableResizing.ToolTipText = vbNullString
        chkPreventDragging.ToolTipText = vbNullString
        chkIgnoreMouse.ToolTipText = vbNullString
        sliOpacity.ToolTipText = vbNullString
        sliSamplingInterval.ToolTipText = vbNullString
        sliStormTestInterval.ToolTipText = vbNullString
        sliErrorInterval.ToolTipText = vbNullString
        cmbScrollWheelDirection.ToolTipText = vbNullString
               
        optClockTooltips(0).ToolTipText = vbNullString
        optClockTooltips(1).ToolTipText = vbNullString
        optClockTooltips(2).ToolTipText = vbNullString
        
        optPrefsTooltips(0).ToolTipText = vbNullString
        optPrefsTooltips(1).ToolTipText = vbNullString
        optPrefsTooltips(2).ToolTipText = vbNullString
         
        'chkEnableBalloonTooltips.ToolTipText = vbNullString
        chkShowTaskbar.ToolTipText = vbNullString
        'chkEnableTooltips.ToolTipText = vbNullString
        sliGaugeSize.ToolTipText = vbNullString
        'sliWidgetSkew.ToolTipText = vbNullString
        btnFacebook.ToolTipText = vbNullString
        imgAbout.ToolTipText = vbNullString
        btnAboutDebugInfo.ToolTipText = vbNullString
        btnDonate.ToolTipText = vbNullString
        btnUpdate.ToolTipText = vbNullString
        
        lblFontsTab(0).ToolTipText = vbNullString
'        lblFontsTab(1).ToolTipText = vbNullString
'        lblFontsTab(2).ToolTipText = vbNullString
        
        lblFontsTab(6).ToolTipText = vbNullString
        
        lblFontsTab(7).ToolTipText = vbNullString
        txtPrefsFontCurrentSize.ToolTipText = vbNullString
        lblCurrentFontsTab.ToolTipText = vbNullString
'        cmbMainGaugeTimeZone.ToolTipText = vbNullString
'        cmbMainDaylightSaving.ToolTipText = vbNullString
        
'        cmbTickSwitchPref.ToolTipText = vbNullString
        
'        lstTimezoneRegions.ToolTipText = vbNullString
        chkDpiAwareness.ToolTipText = vbNullString
        'chkPrefsTooltips.ToolTipText = vbNullString
        btnResetMessages.ToolTipText = vbNullString
        
        cmbMetricImperial.ToolTipText = vbNullString
        cmbTemperatureScale.ToolTipText = vbNullString
        cmbPressureScale.ToolTipText = vbNullString
        cmbWindSpeedScale.ToolTipText = vbNullString
        
        txtIcao.ToolTipText = vbNullString
        btnLocation.ToolTipText = vbNullString
        
    End If

   On Error GoTo 0
   Exit Sub

setPrefsTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsTooltips of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setPrefsLabels
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/09/2023
' Purpose   : set the text in any labels that need a vbCrLf to space the text
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsLabels()

    On Error GoTo setPrefsLabels_Error

    
    lblFontsTab(0).Caption = "When resizing the form (drag bottom right) the font size will in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change." & vbCrLf & vbCrLf & _
        "Next time you open the prefs it will revert to the default." & vbCrLf & vbCrLf & _
        "My preferred font for this utility is Centurion Light SF at 8pt size."
        
    lblPosition(12).Caption = "Selecting a particular gauge and checking the 'This Gauge Locked' box, turns " & _
        "off the ability to drag the program with the mouse. The gauge can be locked into a certain position in either landscape/portrait mode, " & _
        "it ensures that the gauge always appears exactly where you want it to. Using the fields adjacent, you can assign a default x/y position " & _
        "for both Landscape or Portrait mode. Each gauge is locked in place using the gauge's locked button (top left) - this " & _
        "value is set automatically."

    On Error GoTo 0
    Exit Sub

setPrefsLabels_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsLabels of Form widgetPrefs"
        
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DestroyToolTip
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : It's not a bad idea to put this in the Form_Unload event just to make sure.
'---------------------------------------------------------------------------------------
'
Public Sub DestroyToolTip()
    '
   On Error GoTo DestroyToolTip_Error

    If hwndTT <> 0& Then DestroyWindow hwndTT
    hwndTT = 0&

   On Error GoTo 0
   Exit Sub

DestroyToolTip_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DestroyToolTip of Form widgetPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : loadPrefsAboutText
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 12/03/2020
' Purpose   : The text for the about page is stored here
'---------------------------------------------------------------------------------------
'
Private Sub loadPrefsAboutText()
    On Error GoTo loadPrefsAboutText_Error
    'If debugflg = 1 Then Debug.Print "%loadPrefsAboutText"
    
    lblMajorVersion.Caption = App.Major
    lblMinorVersion.Caption = App.Minor
    lblRevisionNum.Caption = App.Revision
    
    lblAbout(1).Caption = "(32bit WoW64 using " & gblCodingEnvironment & ")"
    
    Call LoadFileToTB(txtAboutText, App.path & "\resources\txt\about.txt", False)

   On Error GoTo 0
   Exit Sub

loadPrefsAboutText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPrefsAboutText of Form widgetPrefs"
    
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : picButtonMouseUpEvent
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : capture the icon button clicks avoiding creating a control array
'---------------------------------------------------------------------------------------
'
Private Sub picButtonMouseUpEvent(ByVal thisTabName As String, ByRef thisPicName As Image, ByRef thisPicNameClicked As Image, ByRef thisFraName As Frame, Optional ByRef thisFraButtonName As Frame)
    
    On Error GoTo picButtonMouseUpEvent_Error
    
    Dim padding As Long: padding = 0
    
    Dim borderWidth As Long: borderWidth = 0
    
    Dim captionHeight As Long: captionHeight = 0
    
    Dim y_scale As Single: y_scale = 0
    
    thisPicNameClicked.Visible = False
    thisPicName.Visible = True
      
    btnSave.Visible = False
    btnClose.Visible = False
    btnHelp.Visible = False
    
    Call clearBorderStyle

    gblLastSelectedTab = thisTabName
    sPutINISetting "Software\PzTemperatureGauge", "lastSelectedTab", gblLastSelectedTab, gblSettingsFile

    thisFraName.Visible = True
    thisFraButtonName.BorderStyle = 1
    
    #If TWINBASIC Then
        thisFraButtonName.Refresh
    #End If
    
    ' Get the form's current scale factors.
    y_scale = Me.ScaleHeight / prefsCurrentHeight
    
    If gblDpiAwareness = "1" Then
        btnHelp.Top = fraGeneral.Top + fraGeneral.Height + (200 * y_scale)
    Else
        btnHelp.Top = thisFraName.Top + thisFraName.Height + (200 * y_scale)
    End If
    
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnSave.Top
    
    btnSave.Visible = True
    btnClose.Visible = True
    btnHelp.Visible = True
    
    lblAsterix.Top = btnSave.Top + 50
    lblSize.Top = lblAsterix.Top - 300
    
    chkEnableResizing.Top = btnSave.Top + 50
    'chkEnableResizing.Left = lblAsterix.Left
    
    borderWidth = (Me.Width - Me.ScaleWidth) / 2
    captionHeight = Me.Height - Me.ScaleHeight - borderWidth
        
    ' under windows 10+ the internal window calcs are all wrong due to the bigger title bars
    If pvtPrefsDynamicSizingFlg = False Then
        padding = 200 ' add normal padding below the help button to position the bottom of the form

        pvtLastFormHeight = btnHelp.Top + btnHelp.Height + captionHeight + borderWidth + padding
        Me.Height = pvtLastFormHeight
    End If
    
    If gblDpiAwareness = "0" Then
        If thisTabName = "about" Then
            lblAsterix.Visible = False
            chkEnableResizing.Visible = True
        Else
            lblAsterix.Visible = True
            chkEnableResizing.Visible = False
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

picButtonMouseUpEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picButtonMouseUpEvent of Form widgetPrefs"

End Sub





''---------------------------------------------------------------------------------------
'' Procedure : scrollFrameDownward
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 02/05/2023
'' Purpose   : unused as the scrolling causes blinking, will reduce the interval and re-test
''---------------------------------------------------------------------------------------
''
'Private Sub scrollFrameDownward(ByVal frameToextend As Frame, ByVal fromPosition As Integer, ByVal toPosition As Integer)
'
'    Dim useloop As Integer: useloop = 0
'    Dim currentHeight As Long: currentHeight = 0
'    Dim loopEnd As Long: loopEnd = 0
'    Dim frmCount  As Integer: frmCount = 0
'    Dim frameCount  As Integer: frameCount = 0
'    Dim stepAmount  As Integer: stepAmount = 0
'
'   On Error GoTo scrollFrameDownward_Error
'
'    currentHeight = fromPosition
'    If toPosition > fromPosition Then
'            loopEnd = toPosition - fromPosition
'            stepAmount = 1
'    Else
'            loopEnd = fromPosition - toPosition
'            stepAmount = -1
'    End If
'    For useloop = 1 To loopEnd
'        frameToextend.Height = currentHeight
'        If stepAmount = 1 Then
'            currentHeight = currentHeight + 1
'            If currentHeight >= toPosition Then
'                currentHeight = toPosition
'                Exit For
'            End If
'        End If
'        If stepAmount = -1 Then
'            currentHeight = currentHeight - 1
'            If currentHeight <= toPosition Then
'                currentHeight = toPosition
'                Exit For
'            End If
'        End If
'
'        frameCount = frameCount + 1
'        If frameCount >= 50 Then
'            frameCount = 0
'            frameToextend.Refresh
'        End If
'
'        frmCount = frmCount + 1
'        If frmCount >= 500 Then
'            frmCount = 0
'            widgetPrefs.Refresh
'        End If
'    Next useloop
'
'   On Error GoTo 0
'   Exit Sub
'
'scrollFrameDownward_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure scrollFrameDownward of Form widgetPrefs"
'
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 13/06/2020
' Purpose   : a timer to apply a theme automatically
'---------------------------------------------------------------------------------------
'
Private Sub themeTimer_Timer()
        
    
    Dim SysClr As Long: SysClr = 0

    On Error GoTo themeTimer_Timer_Error

    SysClr = GetSysColor(COLOR_BTNFACE)

    If SysClr <> storeThemeColour Then
        Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click()
    On Error GoTo mnuCoffee_Click_Error
    
    Call mnuCoffee_ClickEvent

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form widgetPrefs"
End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : mnuLicenceA_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/08/2022
' Purpose   : menu option to show licence
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicenceA_Click()
    On Error GoTo mnuLicenceA_Click_Error

    Call mnuLicence_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuLicenceA_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicenceA_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 13/02/2019
' Purpose   : menu option to open support page
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()
    
    On Error GoTo mnuSupport_Click_Error

    Call mnuSupport_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of Form widgetPrefs"
End Sub




Private Sub mnuClosePreferences_Click()
    Call btnClose_Click
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAuto_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuAuto_Click()
    
   On Error GoTo mnuAuto_Click_Error

    If themeTimer.Enabled = True Then
            MsgBox "Automatic Theme Selection is now Disabled"
            mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
            mnuAuto.Checked = False
            
            themeTimer.Enabled = False
    Else
            MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            mnuAuto.Checked = True
            
            themeTimer.Enabled = True
            Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

mnuAuto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuDark_Click()
   On Error GoTo mnuDark_Click_Error

    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enabled"
    mnuLight.Caption = "Light Theme Enable"
    themeTimer.Enabled = False
    
    gblSkinTheme = "dark"

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLight_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 19/05/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLight_Click()
    'MsgBox "Auto Theme Selection Manually Disabled"
   On Error GoTo mnuLight_Click_Error
    
    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enable"
    mnuLight.Caption = "Light Theme Enabled"
    themeTimer.Enabled = False
    
    gblSkinTheme = "light"

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_Click of Form widgetPrefs"
End Sub




'
'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 06/05/2023
' Purpose   : set the theme shade, Windows classic dark/new lighter theme colours
'---------------------------------------------------------------------------------------
'
Private Sub setThemeShade(ByVal redC As Integer, ByVal greenC As Integer, ByVal blueC As Integer)
    
    Dim Ctrl As Control
    
    On Error GoTo setThemeShade_Error

    ' RGB(redC, greenC, blueC) is the background colour used by the lighter themes
    
    Me.BackColor = RGB(redC, greenC, blueC)
    
    ' all buttons must be set to graphical
    For Each Ctrl In Me.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          '@Ignore MemberNotOnInterface
          Ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    If redC = 212 Then
        'classicTheme = True
        mnuLight.Checked = False
        mnuDark.Checked = True
        
        Call setPrefsIconImagesDark
        
    Else
        'classicTheme = False
        mnuLight.Checked = True
        mnuDark.Checked = False
        
        Call setPrefsIconImagesLight
                
    End If
    
    'now change the color of the sliders.
'    widgetPrefs.sliAnimationInterval.BackColor = RGB(redC, greenC, blueC)
    'widgetPrefs.'sliWidgetSkew.BackColor = RGB(redC, greenC, blueC)
    sliGaugeSize.BackColor = RGB(redC, greenC, blueC)
    sliOpacity.BackColor = RGB(redC, greenC, blueC)
    txtAboutText.BackColor = RGB(redC, greenC, blueC)
    sliSamplingInterval.BackColor = RGB(redC, greenC, blueC)
    sliStormTestInterval.BackColor = RGB(redC, greenC, blueC)
    sliErrorInterval.BackColor = RGB(redC, greenC, blueC)
    
    sPutINISetting "Software\PzTemperatureGauge", "skinTheme", gblSkinTheme, gblSettingsFile

    On Error GoTo 0
    Exit Sub

setThemeShade_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeShade of Module Module1"
            Resume Next
          End If
    End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Private Sub setThemeColour()
    
    Dim SysClr As Long: SysClr = 0
    
   On Error GoTo setThemeColour_Error
   'If debugflg = 1  Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        gblSkinTheme = "dark"
        
        mnuDark.Caption = "Dark Theme Enabled"
        mnuLight.Caption = "Light Theme Enable"

    Else
        Call setModernThemeColours
        mnuDark.Caption = "Dark Theme Enable"
        mnuLight.Caption = "Light Theme Enabled"
    End If

    storeThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeColour of module module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsTheme
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 25/04/2023
' Purpose   : adjust the theme used by the prefs alone
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsTheme()
   On Error GoTo adjustPrefsTheme_Error

    If gblSkinTheme <> vbNullString Then
        If gblSkinTheme = "dark" Then
            Call setThemeShade(212, 208, 199)
        Else
            Call setThemeShade(240, 240, 240)
        End If
    Else
        If classicThemeCapable = True Then
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            themeTimer.Enabled = True
        Else
            gblSkinTheme = "light"
            Call setModernThemeColours
        End If
    End If

   On Error GoTo 0
   Exit Sub

adjustPrefsTheme_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsTheme of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setModernThemeColours
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setModernThemeColours()
         
    
    Dim SysClr As Long: SysClr = 0
    
    On Error GoTo setModernThemeColours_Error
    
    'Pz EarthPrefs.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"

    'MsgBox "Windows Alternate Theme detected"
    SysClr = GetSysColor(COLOR_BTNFACE)
    If SysClr = 13160660 Then
        Call setThemeShade(212, 208, 199)
        gblSkinTheme = "dark"
    Else ' 15790320
        Call setThemeShade(240, 240, 240)
        gblSkinTheme = "light"
    End If

   On Error GoTo 0
   Exit Sub

setModernThemeColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setModernThemeColours of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : loadHigherResPrefsImages
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 18/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub loadHigherResPrefsImages()
    
    On Error GoTo loadHigherResPrefsImages_Error
      
    If Me.WindowState = vbMinimized Then Exit Sub
        
    If mnuDark.Checked = True Then
        Call setPrefsIconImagesDark
    Else
        Call setPrefsIconImagesLight
    End If
    
   On Error GoTo 0
   Exit Sub

loadHigherResPrefsImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResPrefsImages of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : positionTimer_Timer
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub positionTimer_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    On Error GoTo positionTimer_Timer_Error
   
    Call writePrefsPosition

   On Error GoTo 0
   Exit Sub

positionTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionTimer_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkEnableResizing_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableResizing_Click()
   On Error GoTo chkEnableResizing_Click_Error

    If chkEnableResizing.Value = 1 Then
        pvtPrefsDynamicSizingFlg = True
        txtPrefsFontCurrentSize.Visible = True
        lblCurrentFontsTab.Visible = True
        'Call writePrefsPosition
        chkEnableResizing.Caption = "Disable Corner Resizing"
    Else
        pvtPrefsDynamicSizingFlg = False
        txtPrefsFontCurrentSize.Visible = False
        lblCurrentFontsTab.Visible = False
        Unload widgetPrefs
        Me.Show
        Call readPrefsPosition
        chkEnableResizing.Caption = "Enable Corner Resizing"
    End If
    
    Call setframeHeights

   On Error GoTo 0
   Exit Sub

chkEnableResizing_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableResizing_Click of Form widgetPrefs"

End Sub

Private Sub chkEnableResizing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip chkEnableResizing.hWnd, "This allows you to resize the whole prefs window by dragging the bottom right corner of the window. It provides an alternative method of supporting high DPI screens.", _
                  TTIconInfo, "Help on Resizing", , , , True
End Sub
 



'---------------------------------------------------------------------------------------
' Procedure : setframeHeights
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setframeHeights()
   On Error GoTo setframeHeights_Error

    If pvtPrefsDynamicSizingFlg = True Then
        fraGeneral.Height = fraAbout.Height
        fraFonts.Height = fraAbout.Height
        fraConfig.Height = fraAbout.Height
        fraSounds.Height = fraAbout.Height
        fraPosition.Height = fraAbout.Height
        fraDevelopment.Height = fraAbout.Height
        fraWindow.Height = fraAbout.Height
        
        fraGeneral.Width = fraAbout.Width
        fraFonts.Width = fraAbout.Width
        fraConfig.Width = fraAbout.Width
        fraSounds.Width = fraAbout.Width
        fraPosition.Width = fraAbout.Width
        fraDevelopment.Width = fraAbout.Width
        fraWindow.Width = fraAbout.Width
    
        'If gblDpiAwareness = "1" Then
            ' save the initial positions of ALL the controls on the prefs form
            
            Call SaveSizes(Me, prefsControlPositions(), prefsCurrentWidth, prefsCurrentHeight)
        'End If
    Else
        fraGeneral.Height = 9096
        fraConfig.Height = 5354
        fraSounds.Height = 1992
        fraPosition.Height = 8472
        fraFonts.Height = 4481
        fraWindow.Height = 6388
        fraDevelopment.Height = 6297
        fraAbout.Height = 8700
    End If
    
    


   On Error GoTo 0
   Exit Sub

setframeHeights_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setframeHeights of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesDark
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 22/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesDark()
    
    On Error GoTo setPrefsIconImagesDark_Error
    
    Set imgGeneral.Picture = Cairo.ImageList("metar-icon-dark").Picture
    Set imgConfig.Picture = Cairo.ImageList("config-icon-dark").Picture
    Set imgFonts.Picture = Cairo.ImageList("font-icon-dark").Picture
    Set imgSounds.Picture = Cairo.ImageList("sounds-icon-dark").Picture
    Set imgPosition.Picture = Cairo.ImageList("position-icon-dark").Picture
    Set imgDevelopment.Picture = Cairo.ImageList("development-icon-dark").Picture
    Set imgWindow.Picture = Cairo.ImageList("windows-icon-dark").Picture
    Set imgAbout.Picture = Cairo.ImageList("about-icon-dark").Picture
'
    Set imgGeneralClicked.Picture = Cairo.ImageList("metar-icon-dark-clicked").Picture
    Set imgConfigClicked.Picture = Cairo.ImageList("config-icon-dark-clicked").Picture
    Set imgFontsClicked.Picture = Cairo.ImageList("font-icon-dark-clicked").Picture
    Set imgSoundsClicked.Picture = Cairo.ImageList("sounds-icon-dark-clicked").Picture
    Set imgPositionClicked.Picture = Cairo.ImageList("position-icon-dark-clicked").Picture
    Set imgDevelopmentClicked.Picture = Cairo.ImageList("development-icon-dark-clicked").Picture
    Set imgWindowClicked.Picture = Cairo.ImageList("windows-icon-dark-clicked").Picture
    Set imgAboutClicked.Picture = Cairo.ImageList("about-icon-dark-clicked").Picture

   On Error GoTo 0
   Exit Sub

setPrefsIconImagesDark_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesDark of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesLight
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 22/06/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesLight()
    
    'Dim resourcePath As String: resourcePath = vbNullString
    
    On Error GoTo setPrefsIconImagesLight_Error
    
    'resourcePath = App.path & "\resources\images"
    
'    If fFExists(resourcePath & "\config-icon-light-" & thisIconWidth & ".jpg") Then imgConfig.Picture = LoadPicture(resourcePath & "\config-icon-light-" & thisIconWidth & ".jpg")
'    If fFExists(resourcePath & "\general-icon-light-" & thisIconWidth & ".jpg") Then imgGeneral.Picture = LoadPicture(resourcePath & "\general-icon-light-" & thisIconWidth & ".jpg")
'    If fFExists(resourcePath & "\position-icon-light-" & thisIconWidth & ".jpg") Then imgPosition.Picture = LoadPicture(resourcePath & "\position-icon-light-" & thisIconWidth & ".jpg")
'    If fFExists(resourcePath & "\font-icon-light-" & thisIconWidth & ".jpg") Then imgFonts.Picture = LoadPicture(resourcePath & "\font-icon-light-" & thisIconWidth & ".jpg")
'    If fFExists(resourcePath & "\development-icon-light-" & thisIconWidth & ".jpg") Then imgDevelopment.Picture = LoadPicture(resourcePath & "\development-icon-light-" & thisIconWidth & ".jpg")
'    If fFExists(resourcePath & "\sounds-icon-light-" & thisIconWidth & ".jpg") Then imgSounds.Picture = LoadPicture(resourcePath & "\sounds-icon-light-" & thisIconWidth & ".jpg")
'    If fFExists(resourcePath & "\windows-icon-light-" & thisIconWidth & ".jpg") Then imgWindow.Picture = LoadPicture(resourcePath & "\windows-icon-light-" & thisIconWidth & ".jpg")
'    If fFExists(resourcePath & "\about-icon-light-" & thisIconWidth & ".jpg") Then imgAbout.Picture = LoadPicture(resourcePath & "\about-icon-light-" & thisIconWidth & ".jpg")
'
'    ' I may yet create clicked versions of all the icons but not now!
'    If fFExists(resourcePath & "\config-icon-light-600-clicked.jpg") Then imgConfigClicked.Picture = LoadPicture(resourcePath & "\config-icon-light-600-clicked.jpg")
'    If fFExists(resourcePath & "\general-icon-light-600-clicked.jpg") Then imgGeneralClicked.Picture = LoadPicture(resourcePath & "\general-icon-light-600-clicked.jpg")
'    If fFExists(resourcePath & "\position-icon-light-600-clicked.jpg") Then imgPositionClicked.Picture = LoadPicture(resourcePath & "\position-icon-light-600-clicked.jpg")
'    If fFExists(resourcePath & "\font-icon-light-600-clicked.jpg") Then imgFontsClicked.Picture = LoadPicture(resourcePath & "\font-icon-light-600-clicked.jpg")
'    If fFExists(resourcePath & "\development-icon-light-600-clicked.jpg") Then imgDevelopmentClicked.Picture = LoadPicture(resourcePath & "\development-icon-light-600-clicked.jpg")
'    If fFExists(resourcePath & "\sounds-icon-light-600-clicked.jpg") Then imgSoundsClicked.Picture = LoadPicture(resourcePath & "\sounds-icon-light-600-clicked.jpg")
'    If fFExists(resourcePath & "\windows-icon-light-600-clicked.jpg") Then imgWindowClicked.Picture = LoadPicture(resourcePath & "\windows-icon-light-600-clicked.jpg")
'    If fFExists(resourcePath & "\about-icon-light-600-clicked.jpg") Then imgAboutClicked.Picture = LoadPicture(resourcePath & "\about-icon-light-600-clicked.jpg")

    
    Set imgGeneral.Picture = Cairo.ImageList("metar-icon-light").Picture
    Set imgConfig.Picture = Cairo.ImageList("config-icon-light").Picture
    Set imgFonts.Picture = Cairo.ImageList("font-icon-light").Picture
    Set imgSounds.Picture = Cairo.ImageList("sounds-icon-light").Picture
    Set imgPosition.Picture = Cairo.ImageList("position-icon-light").Picture
    Set imgDevelopment.Picture = Cairo.ImageList("development-icon-light").Picture
    Set imgWindow.Picture = Cairo.ImageList("windows-icon-light").Picture
    Set imgAbout.Picture = Cairo.ImageList("about-icon-light").Picture
'
    Set imgGeneralClicked.Picture = Cairo.ImageList("metar-icon-light-clicked").Picture
    Set imgConfigClicked.Picture = Cairo.ImageList("config-icon-light-clicked").Picture
    Set imgFontsClicked.Picture = Cairo.ImageList("font-icon-light-clicked").Picture
    Set imgSoundsClicked.Picture = Cairo.ImageList("sounds-icon-light-clicked").Picture
    Set imgPositionClicked.Picture = Cairo.ImageList("position-icon-light-clicked").Picture
    Set imgDevelopmentClicked.Picture = Cairo.ImageList("development-icon-light-clicked").Picture
    Set imgWindowClicked.Picture = Cairo.ImageList("windows-icon-light-clicked").Picture
    Set imgAboutClicked.Picture = Cairo.ImageList("about-icon-light-clicked").Picture

   On Error GoTo 0
   Exit Sub

setPrefsIconImagesLight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesLight of Form widgetPrefs"

End Sub

Private Sub txtPrefsFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPrefsFont.hWnd, "This is a read-only text box. It displays the current font as set when you click the font selector button. This is in operation for informational purposes only. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change.  My preferred font for this utility is Centurion Light SF at 8pt size.", _
                  TTIconInfo, "Help on the Currently Selected Font", , , , True
End Sub

Private Sub txtPrefsFontCurrentSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPrefsFontCurrentSize.hWnd, "This is a read-only text box. It displays the current font as set when dynamic form resizing is enabled. Drag the right hand corner of the window downward and the form will auto-resize. This text box will display the resized font currently in operation for informational purposes only.", _
                  TTIconInfo, "Help on Setting the Font size Dynamically", , , , True
End Sub

Private Sub btnPrefsFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip btnPrefsFont.hWnd, "This is the font selector button, if you click it the font selection window will pop up for you to select your chosen font. Centurion Light SF is a good one and my personal favourite. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change. ", _
                  TTIconInfo, "Help on Setting the Font Selector Button", , , , True
End Sub

Private Sub txtPrefsFontSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gblPrefsTooltips = "0" Then CreateToolTip txtPrefsFontSize.hWnd, "This is a read-only text box. It displays the current base font size as set when dynamic form resizing is enabled. The adjacent text box will display the automatically resized font currently in operation, for informational purposes only.", _
                  TTIconInfo, "Help on the Base Font Size", , , , True
End Sub



'---------------------------------------------------------------------------------------
' Procedure : populateTimeZoneRegions
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub populateTimeZoneRegions()
'
'   Dim cnt As Long: cnt = 0
'
'  'do a lookup for the Bias entered
'   On Error GoTo populateTimeZoneRegions_Error
'
'   With lstTimezoneRegions
'      .Clear
'
'      For cnt = LBound(tzinfo) To UBound(tzinfo)
'
'         If tzinfo(cnt).bias = txtBias.Text Then
'
'            .AddItem tzinfo(cnt).TimeZoneName
'            'Debug.Print tzinfo(cnt).TimeZoneName
'         End If
'
'      Next
'
'   End With
'
'   On Error GoTo 0
'   Exit Sub
'
'populateTimeZoneRegions_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateTimeZoneRegions of Form widgetPrefs"
'
'End Sub

' Randy Birch for his timezone code - http://vbnet.mvps.org/index.html?code/locale/timezonebiaslookup.htm

'---------------------------------------------------------------------------------------
' Procedure : fGetTimeZoneArray
' Author    : Randy Birch for his timezone code - http://vbnet.mvps.org/index.html?code/locale/timezonebiaslookup.htm
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function fGetTimeZoneArray() As Boolean
'
'   Dim success As Long
'   Dim dwIndex As Long
'   Dim cbName As Long
'   Dim hKey As Long
'   Dim sName As String
'   Dim dwSubKeys As Long
'   Dim dwMaxSubKeyLen As Long
'   Dim ft As FILETIME
'
'  'Win9x and WinNT have a slightly
'  'different registry structure.
'  'Determine the operating system and
'  'set a module variable to the
'  'correct key.
'
'  'assume OS is win9x
'   On Error GoTo fGetTimeZoneArray_Error
'
'   sTzKey = SKEY_9X
'
'  'see if OS is NT, and if so,
'  'use assign the correct key
'   If IsWinNTPlus Then sTzKey = SKEY_NT
'
'  'BiasAdjust is used when calculating the
'  'bias values retrieved from the registry.
'  'If True, the reg value retrieved represents
'  'the location's bias with the bias for
'  'daylight saving time added. If false, the
'  'location's bias is returned with the
'  'standard bias adjustment applied (this
'  'is usually 0). Doing this allows us to
'  'use the bias returned from a TIME_OF_DAY_INFO
'  'call as the correct lookup value dependant
'  'on whether the world is currently on
'  'daylight saving time or not. For those
'  'countries not recognizing daylight saving
'  'time, the registry daylight bias will be 0,
'  'therefore proper lookup will not be affected.
'  'Not considered (nor can such be coded) are those
'  'special areas within a given country that do
'  'not recognize daylight saving time, even
'  'when the rest of the country does (like
'  'Saskatchewan in Canada).
'   BiasAdjust = IsDaylightSavingTime()
'
'  'open the timezone registry key
'   hKey = OpenRegKey(HKEY_LOCAL_MACHINE, sTzKey)
'
'   If hKey <> 0 Then
'
'     'query registry for the number of
'     'entries under that key
'      If RegQueryInfoKey(hKey, _
'                         0&, _
'                         0&, _
'                         0, _
'                         dwSubKeys, _
'                         dwMaxSubKeyLen&, _
'                         0&, _
'                         0&, _
'                         0&, _
'                         0&, _
'                         0&, _
'                         ft) = ERROR_SUCCESS Then
'
'
'        'create a UDT array for the time zone info
'         ReDim tzinfo(0 To dwSubKeys - 1) As TZ_LOOKUP_DATA
'
'         dwIndex = 0
'         cbName = 32
'
'         Do
'
'           'pad a string for the returned value
'            sName = Space$(cbName)
'            success = RegEnumKey(hKey, dwIndex, sName, cbName)
'
'            If success = ERROR_SUCCESS Then
'
'              'add the data to the appropriate
'              'tzinfo UDT array members
'               With tzinfo(dwIndex)
'
'                  .TimeZoneName = TrimNull(sName)
'                  .bias = GetTZBiasByName(.TimeZoneName)
'                  .IsDST = BiasAdjust
'
'                 'is also added to a list
'                  'cmbMainDaylightSaving.AddItem .bias & vbTab & .TimeZoneName
'
'               End With
'
'            End If
'
'           'increment the loop...
'            dwIndex = dwIndex + 1
'
'        '...and continue while the reg
'        'call returns success.
'         Loop While success = ERROR_SUCCESS
'
'        'clean up
'         RegCloseKey hKey
'
'        'return success if, well, successful
'         fGetTimeZoneArray = dwIndex > 0
'
'      End If  'If RegQueryInfoKey
'
'   Else
'
'     'could not open reg key
'      fGetTimeZoneArray = False
'
'   End If  'If hKey
'
'   On Error GoTo 0
'   Exit Function
'
'fGetTimeZoneArray_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetTimeZoneArray of Form widgetPrefs"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDaylightSavingTime
' Author    : Randy Birch for his timezone code - http://vbnet.mvps.org/index.html?code/locale/timezonebiaslookup.htm
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function IsDaylightSavingTime() As Boolean
'
'   Dim tzi As TIME_ZONE_INFORMATION
'
'   On Error GoTo IsDaylightSavingTime_Error
'
'   IsDaylightSavingTime = GetTimeZoneInformation(tzi) = TIME_ZONE_ID_DAYLIGHT
'
'   On Error GoTo 0
'   Exit Function
'
'IsDaylightSavingTime_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsDaylightSavingTime of Form widgetPrefs"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTZBiasByName
' Author    : Randy Birch for his timezone code - http://vbnet.mvps.org/index.html?code/locale/timezonebiaslookup.htm
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function GetTZBiasByName(sTimeZone As String) As Long
'
'   Dim rtzi As REG_TIME_ZONE_INFORMATION
'   Dim hKey As Long
'
'  'open the passed time zone key
'   On Error GoTo GetTZBiasByName_Error
'
'   hKey = OpenRegKey(HKEY_LOCAL_MACHINE, sTzKey & "\" & sTimeZone)
'
'   If hKey <> 0 Then
'
'     'obtain the data from the TZI member
'      If RegQueryValueEx(hKey, _
'                         "TZI", _
'                         0&, _
'                         ByVal 0&, _
'                         rtzi, _
'                         Len(rtzi)) = ERROR_SUCCESS Then
'
'        'tweak the Bias when in Daylight Saving time
'         If BiasAdjust Then
'            GetTZBiasByName = (rtzi.bias + rtzi.DaylightBias)
'         Else
'            GetTZBiasByName = (rtzi.bias + rtzi.StandardBias) 'StandardBias is usually 0
'         End If
'
'      End If
'
'      RegCloseKey hKey
'
'   End If
'
'   On Error GoTo 0
'   Exit Function
'
'GetTZBiasByName_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetTZBiasByName of Form widgetPrefs"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : TrimNull
' Author    : Randy Birch for his timezone code - http://vbnet.mvps.org/index.html?code/locale/timezonebiaslookup.htm
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function TrimNull(startstr As String) As String
'
'   On Error GoTo TrimNull_Error
'
'   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
'
'   On Error GoTo 0
'   Exit Function
'
'TrimNull_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TrimNull of Form widgetPrefs"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenRegKey
' Author    : Randy Birch for his timezone code - http://vbnet.mvps.org/index.html?code/locale/timezonebiaslookup.htm
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function OpenRegKey(ByVal hKey As Long, _
'                            ByVal lpSubKey As String) As Long
'
'  Dim hSubKey As Long
'
'   On Error GoTo OpenRegKey_Error
'
'  If RegOpenKeyEx(hKey, _
'                  lpSubKey, _
'                  0, _
'                  KEY_READ, _
'                  hSubKey) = ERROR_SUCCESS Then
'
'      OpenRegKey = hSubKey
'
'  End If
'
'   On Error GoTo 0
'   Exit Function
'
'OpenRegKey_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenRegKey of Form widgetPrefs"
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : IsWinNTPlus
' Author    : Randy Birch for his timezone code - http://vbnet.mvps.org/index.html?code/locale/timezonebiaslookup.htm
' Date      : 13/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Function IsWinNTPlus() As Boolean
'
'   'returns True if running WinNT or better
'   On Error GoTo IsWinNTPlus_Error
'
'   #If Win32 Then
'
'      Dim OSV As OSVERSIONINFO
'
'      OSV.OSVSize = Len(OSV)
'
'      If GetVersionEx(OSV) = 1 Then
'
'         IsWinNTPlus = (OSV.PlatformID = VER_PLATFORM_WIN32_NT)
'
'      End If
'
'   #End If
'
'   On Error GoTo 0
'   Exit Function
'
'IsWinNTPlus_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsWinNTPlus of Form widgetPrefs"
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo lblDragCorner_MouseDown_Error
    
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
    
    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseDown of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo lblDragCorner_MouseMove_Error

    lblDragCorner.MousePointer = 8

    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseMove of Form widgetPrefs"
   
End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnGithubHome_Click
' Author    : beededea
' Date      : 22/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGithubHome_Click()
   On Error GoTo btnGithubHome_Click_Error

    Call menuForm.mnuGithubHome_Click

   On Error GoTo 0
   Exit Sub

btnGithubHome_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGithubHome_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setPrefsFormZordering
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setPrefsFormZordering()

   On Error GoTo setPrefsFormZordering_Error

'    If Val(gblWindowLevel) = 0 Then
'        Call SetWindowPos(Me.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
'    ElseIf Val(gblWindowLevel) = 1 Then
'        Call SetWindowPos(Me.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
'    ElseIf Val(gblWindowLevel) = 2 Then
'        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
'    End If

   On Error GoTo 0
   Exit Sub

setPrefsFormZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsFormZordering of Module modMain"
End Sub

#If TWINBASIC Then
    '---------------------------------------------------------------------------------------
    ' Procedure : setVisualStyles
    ' Author    : beededea
    ' Date      : 13/01/2025
    ' Purpose   : loop through all the controls and identify the labels and text boxes and disable modern styles
    '             reverts TwinBasic form themeing to that of the earlier classic look and feel.
    '---------------------------------------------------------------------------------------
    '
        Private Sub setVisualStyles()
            Dim Ctrl As Control
          
            On Error GoTo setVisualStyles_Error

            For Each Ctrl In widgetPrefs.Controls
                If (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
                    Ctrl.VisualStyles = False
                End If
            Next

       On Error GoTo 0
       Exit Sub

setVisualStyles_Error:

        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setVisualStyles of Form widgetPrefs"
        End Sub
#End If


'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


