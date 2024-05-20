VERSION 5.00
Object = "{BCE37951-37DF-4D69-A8A3-2CFABEE7B3CC}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form panzerPrefs 
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
      TabIndex        =   42
      Top             =   0
      Width           =   1065
      Begin VB.Label lblDevelopment 
         Caption         =   "Development"
         Height          =   240
         Left            =   45
         TabIndex        =   43
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
      TabIndex        =   97
      Top             =   10125
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Frame fraAboutButton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   7695
      TabIndex        =   67
      Top             =   0
      Width           =   975
      Begin VB.Label lblAbout 
         Caption         =   "About"
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   68
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
      TabIndex        =   44
      Top             =   -15
      Width           =   930
      Begin VB.Label lblConfig 
         Caption         =   "Config."
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   45
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
      TabIndex        =   40
      Top             =   0
      Width           =   930
      Begin VB.Label lblPosition 
         Caption         =   "Position"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   41
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
      TabIndex        =   38
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
      Left            =   660
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
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   8550
      Left            =   75
      TabIndex        =   50
      Top             =   1200
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Frame fraGeneralInner 
         BorderStyle     =   0  'None
         Height          =   8040
         Left            =   450
         TabIndex        =   51
         Top             =   225
         Width           =   7215
         Begin VB.CheckBox chkGaugeFunctions 
            Caption         =   "Enable Gauges and METAR polling *"
            Height          =   225
            Left            =   2010
            TabIndex        =   180
            ToolTipText     =   "When checked this box enables the pointer. That's it!"
            Top             =   225
            Width           =   3405
         End
         Begin VB.TextBox txtAirportsURL 
            Height          =   315
            Left            =   2055
            Locked          =   -1  'True
            TabIndex        =   144
            Text            =   "https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports.dat"
            Top             =   7185
            Width           =   4755
         End
         Begin VB.CommandButton btnLocation 
            Caption         =   "Select ICAO"
            Height          =   315
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   705
            Width           =   1215
         End
         Begin VB.ComboBox cmbMetricImperial 
            Height          =   315
            ItemData        =   "frmPrefs.frx":5A2C
            Left            =   2010
            List            =   "frmPrefs.frx":5A2E
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   2880
            Width           =   1740
         End
         Begin VB.ComboBox cmbWindSpeedScale 
            Height          =   315
            ItemData        =   "frmPrefs.frx":5A30
            Left            =   2010
            List            =   "frmPrefs.frx":5A32
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   2370
            Width           =   1740
         End
         Begin VB.ComboBox cmbPressureScale 
            Height          =   315
            ItemData        =   "frmPrefs.frx":5A34
            Left            =   2010
            List            =   "frmPrefs.frx":5A36
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   1830
            Width           =   1740
         End
         Begin VB.TextBox txtIcao 
            Height          =   315
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   124
            Text            =   "EGSH"
            Top             =   705
            Width           =   1725
         End
         Begin VB.ComboBox cmbTemperatureScale 
            Height          =   315
            ItemData        =   "frmPrefs.frx":5A38
            Left            =   2010
            List            =   "frmPrefs.frx":5A3A
            Style           =   2  'Dropdown List
            TabIndex        =   121
            Top             =   1260
            Width           =   1740
         End
         Begin vb6projectCCRSlider.Slider sliSamplingInterval 
            Height          =   390
            Left            =   1890
            TabIndex        =   115
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
            TabIndex        =   138
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
            TabIndex        =   147
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
            TabIndex        =   181
            Top             =   210
            Width           =   1320
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "hours"
            Height          =   315
            Index           =   23
            Left            =   3615
            TabIndex        =   151
            Top             =   6720
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "This is the full URL giving the location of the airports.dat file "
            Height          =   285
            Index           =   20
            Left            =   2070
            TabIndex        =   146
            Top             =   7650
            Width           =   4635
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Error Message Interval :"
            Height          =   360
            Index           =   24
            Left            =   180
            TabIndex        =   150
            Top             =   6300
            Width           =   1770
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "8hrs"
            Height          =   300
            Index           =   22
            Left            =   5490
            TabIndex        =   149
            Top             =   6720
            Width           =   465
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "0 (disabled)"
            Height          =   300
            Index           =   21
            Left            =   1995
            TabIndex        =   148
            Top             =   6720
            Width           =   945
         End
         Begin VB.Label lblGeneral 
            Caption         =   $"frmPrefs.frx":5A3C
            Height          =   1005
            Index           =   12
            Left            =   3900
            TabIndex        =   132
            Top             =   2850
            Width           =   3270
         End
         Begin VB.Label lblGeneral 
            Caption         =   "ICAO Airports URL :"
            Height          =   255
            Index           =   15
            Left            =   255
            TabIndex        =   145
            Top             =   7230
            Width           =   1545
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "1800"
            Height          =   315
            Index           =   19
            Left            =   2070
            TabIndex        =   143
            Top             =   5475
            Width           =   540
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "7200"
            Height          =   315
            Index           =   18
            Left            =   5385
            TabIndex        =   142
            Top             =   5460
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "secs"
            Height          =   315
            Index           =   17
            Left            =   3615
            TabIndex        =   141
            Top             =   5460
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Storm Test Interval :"
            Height          =   315
            Index           =   16
            Left            =   300
            TabIndex        =   140
            Top             =   5055
            Width           =   1635
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Adjust to determine storm checking frequency."
            Height          =   540
            Index           =   15
            Left            =   2100
            TabIndex        =   139
            Top             =   5760
            Width           =   3810
         End
         Begin VB.Label lblGeneral 
            Caption         =   "The Wind Speed Scale"
            Height          =   480
            Index           =   14
            Left            =   3885
            TabIndex        =   134
            Top             =   2430
            Width           =   2610
         End
         Begin VB.Label lblGeneral 
            Caption         =   "The Air Pressure Scale"
            Height          =   480
            Index           =   13
            Left            =   3885
            TabIndex        =   133
            Top             =   1875
            Width           =   2610
         End
         Begin VB.Label lblGeneral 
            Alignment       =   1  'Right Justify
            Caption         =   "Metric or Imperial :"
            Height          =   480
            Index           =   8
            Left            =   255
            TabIndex        =   131
            Top             =   2910
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Alignment       =   1  'Right Justify
            Caption         =   "Anemometer :"
            Height          =   480
            Index           =   7
            Left            =   255
            TabIndex        =   129
            Top             =   2400
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Alignment       =   1  'Right Justify
            Caption         =   "Barometer :"
            Height          =   345
            Index           =   4
            Left            =   255
            TabIndex        =   127
            Top             =   1860
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Caption         =   "ICAO Station ID :"
            Height          =   255
            Index           =   1
            Left            =   585
            TabIndex        =   125
            Top             =   765
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Temperature :"
            Height          =   480
            Index           =   5
            Left            =   810
            TabIndex        =   123
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Select Celsius / Fahrenheit / Kelvin"
            Height          =   480
            Index           =   10
            Left            =   3855
            TabIndex        =   122
            Top             =   1305
            Width           =   2610
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Adjust to determine gauge sampling frequency.*"
            Height          =   270
            Index           =   14
            Left            =   2100
            TabIndex        =   120
            Top             =   4530
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Sampling Interval :"
            Height          =   315
            Index           =   13
            Left            =   495
            TabIndex        =   119
            Top             =   3825
            Width           =   1410
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "secs"
            Height          =   315
            Index           =   12
            Left            =   3615
            TabIndex        =   118
            Top             =   4230
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "4800"
            Height          =   315
            Index           =   11
            Left            =   5385
            TabIndex        =   117
            Top             =   4230
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "60"
            Height          =   315
            Index           =   10
            Left            =   2070
            TabIndex        =   116
            Top             =   4230
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      Height          =   5280
      Left            =   240
      TabIndex        =   8
      Top             =   1185
      Width           =   7140
      Begin VB.Frame fraConfigInner 
         BorderStyle     =   0  'None
         Height          =   4650
         Left            =   435
         TabIndex        =   34
         Top             =   435
         Width           =   6450
         Begin VB.CheckBox chkGenStartup 
            Caption         =   "Run the Temperature Widget at Windows Startup "
            Height          =   465
            Left            =   2010
            TabIndex        =   136
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   3915
            Width           =   4020
         End
         Begin VB.CheckBox chkEnablePrefsTooltips 
            Caption         =   "Enable Preference Utility Tooltips *"
            Height          =   225
            Left            =   2010
            TabIndex        =   109
            ToolTipText     =   "Check the box to enable tooltips for all controls in the preferences utility"
            Top             =   2145
            Width           =   3345
         End
         Begin VB.CheckBox chkDpiAwareness 
            Caption         =   "DPI Awareness Enable *"
            Height          =   225
            Left            =   2010
            TabIndex        =   107
            ToolTipText     =   "Check the box to make the program DPI aware. RESTART required."
            Top             =   2925
            Width           =   3405
         End
         Begin VB.CheckBox chkShowTaskbar 
            Caption         =   "Show Gauges in Taskbar"
            Height          =   225
            Left            =   2010
            TabIndex        =   105
            ToolTipText     =   "Check the box to show the widget in the taskbar"
            Top             =   2535
            Width           =   3405
         End
         Begin VB.ComboBox cmbScrollWheelDirection 
            Height          =   315
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   63
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   120
            Width           =   2490
         End
         Begin VB.CheckBox chkEnableBalloonTooltips 
            Caption         =   "Enable Balloon Tooltips on all Controls *"
            Height          =   225
            Left            =   2010
            TabIndex        =   39
            ToolTipText     =   "Check the box to enable larger balloon tooltips for all controls on the main program"
            Top             =   1770
            Width           =   4035
         End
         Begin VB.CheckBox chkEnableTooltips 
            Caption         =   "Enable Main Program Tooltips"
            Height          =   225
            Left            =   2010
            TabIndex        =   35
            ToolTipText     =   "Check the box to enable tooltips for all controls on the main program"
            Top             =   1335
            Width           =   3345
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Auto Start :"
            Height          =   375
            Index           =   11
            Left            =   975
            TabIndex        =   137
            Tag             =   "lblRefreshInterval"
            Top             =   4035
            Width           =   1740
         End
         Begin VB.Label lblConfiguration 
            Caption         =   $"frmPrefs.frx":5ADA
            Height          =   930
            Index           =   0
            Left            =   1980
            TabIndex        =   108
            Top             =   3255
            Width           =   4335
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "The scroll-wheel resizing direction can be determined here. The direction chosen causes the gauge to grow. *"
            Height          =   690
            Index           =   6
            Left            =   2025
            TabIndex        =   88
            Top             =   540
            Width           =   3930
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Mouse Wheel Resize :"
            Height          =   345
            Index           =   3
            Left            =   255
            TabIndex        =   64
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
            TabIndex        =   112
            Top             =   3405
            Width           =   885
         End
         Begin VB.TextBox txtPrefsFontCurrentSize 
            Height          =   315
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   98
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
            TabIndex        =   113
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   3345
            Width           =   3360
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Reset Pop ups :"
            Height          =   300
            Index           =   1
            Left            =   435
            TabIndex        =   111
            Tag             =   "lblPrefsFont"
            Top             =   3450
            Width           =   1470
         End
         Begin VB.Label lblCurrentFontsTab 
            Caption         =   "Resized Font"
            Height          =   315
            Left            =   4950
            TabIndex        =   99
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1110
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   $"frmPrefs.frx":5B8E
            Height          =   1710
            Index           =   0
            Left            =   1725
            TabIndex        =   66
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
            TabIndex        =   36
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   285
            Width           =   3405
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Audio :"
            Height          =   255
            Index           =   3
            Left            =   885
            TabIndex        =   65
            Tag             =   "lblSharedInputFile"
            Top             =   285
            Width           =   765
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "When checked, this box enables all the sounds used during any animation/interaction with the main program."
            Height          =   705
            Index           =   4
            Left            =   1515
            TabIndex        =   37
            Tag             =   "lblEnableSoundsDesc"
            Top             =   645
            Width           =   3705
         End
      End
   End
   Begin VB.Frame fraWindow 
      Caption         =   "Window"
      Height          =   6300
      Left            =   405
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
            TabIndex        =   89
            Top             =   2325
            Width           =   5130
            Begin VB.ComboBox cmbHidingTime 
               Height          =   315
               Left            =   825
               Style           =   2  'Dropdown List
               TabIndex        =   92
               Top             =   1680
               Width           =   3720
            End
            Begin VB.CheckBox chkWidgetHidden 
               Caption         =   "Hiding Widget *"
               Height          =   225
               Left            =   855
               TabIndex        =   90
               Top             =   225
               Width           =   2955
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   "Hiding :"
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   93
               Top             =   210
               Width           =   720
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   $"frmPrefs.frx":5CCC
               Height          =   975
               Index           =   1
               Left            =   855
               TabIndex        =   91
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
            TabIndex        =   95
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
   Begin VB.Frame fraPosition 
      Caption         =   "Position && Size"
      Height          =   8355
      Left            =   240
      TabIndex        =   46
      Top             =   1230
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraPositionInner 
         BorderStyle     =   0  'None
         Height          =   7755
         Left            =   150
         TabIndex        =   47
         Top             =   300
         Width           =   7680
         Begin VB.Frame fraGaugePosition 
            Caption         =   "Temperature Gauge "
            Height          =   3435
            Left            =   255
            TabIndex        =   160
            Top             =   4305
            Width           =   4410
            Begin VB.TextBox txtLandscapeHoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   169
               Top             =   1170
               Width           =   960
            End
            Begin VB.TextBox txtLandscapeVoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   168
               Top             =   1575
               Width           =   960
            End
            Begin VB.TextBox txtPortraitHoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   167
               Top             =   2460
               Width           =   960
            End
            Begin VB.TextBox txtPortraitVoffset 
               Height          =   330
               Left            =   1965
               TabIndex        =   166
               Top             =   2865
               Width           =   960
            End
            Begin VB.ComboBox cmbPortraitLocked 
               Height          =   315
               Left            =   1965
               Style           =   2  'Dropdown List
               TabIndex        =   163
               ToolTipText     =   "Choose the alarm sound."
               Top             =   2025
               Width           =   2160
            End
            Begin VB.ComboBox cmbLandscapeLocked 
               Height          =   315
               Left            =   1965
               Style           =   2  'Dropdown List
               TabIndex        =   162
               ToolTipText     =   "Choose the alarm sound."
               Top             =   750
               Width           =   2160
            End
            Begin VB.CheckBox chkPreventDragging 
               Caption         =   "This gauge Locked. *"
               Height          =   225
               Left            =   1980
               TabIndex        =   161
               ToolTipText     =   "Checking this box turns off the ability to drag the program with the mouse, locking it in position."
               Top             =   360
               Width           =   2250
            End
            Begin VB.Label lblPosition 
               Caption         =   "Locked in Portrait :"
               Height          =   375
               Index           =   11
               Left            =   405
               TabIndex        =   165
               Tag             =   "lblAlarmSound"
               Top             =   2070
               Width           =   2040
            End
            Begin VB.Label lblPosition 
               Caption         =   "Locked in Landscape :"
               Height          =   435
               Index           =   13
               Left            =   165
               TabIndex        =   164
               Tag             =   "lblAlarmSound"
               Top             =   795
               Width           =   2115
            End
         End
         Begin VB.ComboBox cmbGaugeType 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   2415
            Width           =   2430
         End
         Begin VB.ComboBox cmbWidgetPosition 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   156
            ToolTipText     =   "Choose the alarm sound."
            Top             =   1035
            Width           =   2430
         End
         Begin VB.Frame fraPositionBalloonBox 
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   120
            TabIndex        =   152
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cmbAspectHidden 
               Height          =   315
               Left            =   2145
               Style           =   2  'Dropdown List
               TabIndex        =   153
               ToolTipText     =   "Choose the alarm sound."
               Top             =   0
               Width           =   2430
            End
            Begin VB.Label lblPosition 
               Caption         =   "Aspect Ratio Hidden Mode :"
               Height          =   375
               Index           =   3
               Left            =   0
               TabIndex        =   155
               Tag             =   "lblAlarmSound"
               Top             =   45
               Width           =   2145
            End
            Begin VB.Label lblPosition 
               Caption         =   "Tablets only. Don't fiddle with this unless you really know what you are doing. Read the help before fiddling!"
               Height          =   765
               Index           =   6
               Left            =   2145
               TabIndex        =   154
               Tag             =   "lblAlarmSoundDesc"
               Top             =   420
               Width           =   5370
            End
         End
         Begin vb6projectCCRSlider.Slider sliGaugeSize 
            Height          =   390
            Left            =   2085
            TabIndex        =   170
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
            TabIndex        =   179
            Tag             =   "lblAlarmSound"
            Top             =   2460
            Width           =   1935
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "5"
            Height          =   315
            Index           =   0
            Left            =   2235
            TabIndex        =   178
            Top             =   3390
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "220 (%)"
            Height          =   315
            Index           =   5
            Left            =   5550
            TabIndex        =   177
            Top             =   3405
            Width           =   735
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "90"
            Height          =   315
            Index           =   2
            Left            =   3465
            TabIndex        =   176
            Top             =   3405
            Width           =   420
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Gauge Size :"
            Height          =   315
            Index           =   1
            Left            =   1125
            TabIndex        =   175
            Top             =   2955
            Width           =   975
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel. Immediate. *"
            Height          =   555
            Index           =   2
            Left            =   2235
            TabIndex        =   174
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   3720
            Width           =   3810
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "50"
            Height          =   315
            Index           =   1
            Left            =   2895
            TabIndex        =   173
            Top             =   3405
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "130"
            Height          =   315
            Index           =   3
            Left            =   4155
            TabIndex        =   172
            Top             =   3405
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "180"
            Height          =   315
            Index           =   4
            Left            =   4935
            TabIndex        =   171
            Top             =   3405
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "Widget Position by Percent :"
            Height          =   375
            Index           =   8
            Left            =   195
            TabIndex        =   158
            Tag             =   "lblAlarmSound"
            Top             =   1080
            Width           =   2355
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":5D6F
            Height          =   705
            Index           =   10
            Left            =   2250
            TabIndex        =   157
            Tag             =   "lblAlarmSoundDesc"
            Top             =   1485
            Width           =   5325
         End
         Begin VB.Label lblPosition 
            Caption         =   "*"
            Height          =   255
            Index           =   1
            Left            =   4545
            TabIndex        =   96
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "this text box is filled in the setPrefsLabels sub routine"
            Height          =   4065
            Index           =   12
            Left            =   5025
            TabIndex        =   62
            Tag             =   "lblAlarmSoundDesc"
            Top             =   4425
            Width           =   2520
         End
      End
   End
   Begin VB.Frame fraDevelopment 
      Caption         =   "Development"
      Height          =   6210
      Left            =   240
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraDevelopmentInner 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   870
         TabIndex        =   49
         Top             =   300
         Width           =   7455
         Begin VB.Frame fraDefaultEditor 
            BorderStyle     =   0  'None
            Height          =   2370
            Left            =   75
            TabIndex        =   100
            Top             =   3165
            Width           =   7290
            Begin VB.CommandButton btnDefaultEditor 
               Caption         =   "..."
               Height          =   300
               Left            =   5115
               Style           =   1  'Graphical
               TabIndex        =   102
               ToolTipText     =   "Click to select the .vbp file to edit the program - You need to have access to the source!"
               Top             =   210
               Width           =   315
            End
            Begin VB.TextBox txtDefaultEditor 
               Height          =   315
               Left            =   1440
               TabIndex        =   101
               Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
               Top             =   195
               Width           =   3660
            End
            Begin VB.Label lblGitHub 
               Caption         =   $"frmPrefs.frx":5E15
               ForeColor       =   &H8000000D&
               Height          =   915
               Left            =   1560
               TabIndex        =   106
               ToolTipText     =   "Double Click to visit github"
               Top             =   1440
               Width           =   4935
            End
            Begin VB.Label lblDebug 
               Caption         =   $"frmPrefs.frx":5EAE
               Height          =   930
               Index           =   9
               Left            =   1545
               TabIndex        =   104
               Top             =   690
               Width           =   4785
            End
            Begin VB.Label lblDebug 
               Caption         =   "Default Editor :"
               Height          =   255
               Index           =   7
               Left            =   285
               TabIndex        =   103
               Tag             =   "lblSharedInputFile"
               Top             =   225
               Width           =   1350
            End
         End
         Begin VB.TextBox txtDblClickCommand 
            Height          =   315
            Left            =   1515
            TabIndex        =   59
            ToolTipText     =   "Enter a Windows command for the gauge to operate when double-clicked."
            Top             =   1095
            Width           =   3660
         End
         Begin VB.CommandButton btnOpenFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5175
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Click to select a particular file for the gauge to run or open when double-clicked."
            Top             =   2250
            Width           =   315
         End
         Begin VB.TextBox txtOpenFile 
            Height          =   315
            Left            =   1515
            TabIndex        =   55
            ToolTipText     =   "Enter a particular file for the gauge to run or open when double-clicked."
            Top             =   2235
            Width           =   3660
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmPrefs.frx":5F52
            Left            =   1530
            List            =   "frmPrefs.frx":5F54
            Style           =   2  'Dropdown List
            TabIndex        =   52
            ToolTipText     =   "Choose to set debug mode."
            Top             =   -15
            Width           =   2160
         End
         Begin VB.Label lblDebug 
            Caption         =   "DblClick Command :"
            Height          =   510
            Index           =   1
            Left            =   -15
            TabIndex        =   61
            Tag             =   "lblPrefixString"
            Top             =   1155
            Width           =   1545
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Shift+double-clicking on the widget image will open this file. "
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   60
            Top             =   2730
            Width           =   3705
         End
         Begin VB.Label lblDebug 
            Caption         =   "Default command to run when the gauge receives a double-click eg %SystemRoot%/system32/ncpa.cpl"
            Height          =   570
            Index           =   5
            Left            =   1590
            TabIndex        =   58
            Tag             =   "lblSharedInputFileDesc"
            Top             =   1605
            Width           =   4410
         End
         Begin VB.Label lblDebug 
            Caption         =   "Open File :"
            Height          =   255
            Index           =   4
            Left            =   645
            TabIndex        =   57
            Tag             =   "lblSharedInputFile"
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label lblDebug 
            Caption         =   "Turning on the debugging will provide extra information in the debug window.  *"
            Height          =   495
            Index           =   2
            Left            =   1545
            TabIndex        =   54
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   4455
         End
         Begin VB.Label lblDebug 
            Caption         =   "Debug :"
            Height          =   375
            Index           =   0
            Left            =   855
            TabIndex        =   53
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraAbout 
      Caption         =   "About"
      Height          =   8580
      Left            =   255
      TabIndex        =   69
      Top             =   1185
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   7950
         TabIndex        =   83
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
         TabIndex        =   82
         Text            =   "frmPrefs.frx":5F56
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
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1110
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
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   735
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
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   360
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
         Height          =   360
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1485
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
         Top             =   510
         Width           =   1245
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
      TabIndex        =   114
      ToolTipText     =   "drag me"
      Top             =   10335
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblSize 
      Caption         =   "Size in twips"
      Height          =   285
      Left            =   1875
      TabIndex        =   110
      Top             =   9780
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Label lblAsterix 
      Caption         =   "All controls marked with a * take effect immediately."
      Height          =   300
      Left            =   1920
      TabIndex        =   94
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
Attribute VB_Name = "panzerPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ArgumentWithIncompatibleObjectType, AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder

' gaugeForm_BubblingEvent ' leaving that here so I can copy/paste to find it

'---------------------------------------------------------------------------------------
' Module    : panzerPrefs
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
  (ByVal hwnd As Long, _
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
Private PzGPrefsLoadedFlg As Boolean
Private prefsDynamicSizingFlg As Boolean
Private lastFormHeight As Long
Private Const cPrefsFormHeight As Long = 11055
Private Const cPrefsFormWidth  As Long = 9090
'------------------------------------------------------ ENDS

Private prefsStartupFlg As Boolean
Private gblAllowSizeChangeFlg As Boolean


Private Sub btnLocation_Click()
    fSelector.SelectorForm.Show
End Sub

Private Sub btnLocation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip btnLocation.hwnd, "Press to select the current ICAO code used to identify the weather feed source data.*", _
                  TTIconInfo, "Select the current ICAO code", , , , True

End Sub

Private Sub chkEnableBalloonTooltips_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If PzGEnableBalloonTooltips = "1" Then CreateToolTip chkEnableBalloonTooltips.hwnd, "Best to enable the balloon tooltips rather than the standard ones. Do that here. The balloon tooltips are much prettier, have more space for pertinent information so I can fill them up with useful text to assist you.", _
                  TTIconInfo, "Help on the Balloon Tooltips", , , , True
End Sub

Private Sub chkPreventDragging_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip chkPreventDragging.hwnd, "" _
    & "Lock the gauges in a certain position in either landscape/portrait mode. This ensures that the widget always appears exactly where you want it to. Drag" _
    & "the gauge into position using the mouse and when the widget is locked in place (using the Widget lock button), this value is set automatically.", _
                  TTIconInfo, "Lock the gauges in position.", , , , True
End Sub

Private Sub chkShowTaskbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip chkShowTaskbar.hwnd, "Checking this box causes" _
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
            If PzGDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = PzGTemperatureFormHighDpiXPos
                txtLandscapeVoffset.Text = PzGTemperatureFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = PzGTemperatureFormLowDpiXPos
                txtLandscapeVoffset.Text = PzGTemperatureFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fTemperature.temperatureGaugeForm.Left
            txtPortraitVoffset.Text = fTemperature.temperatureGaugeForm.Top
            If PzGDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayTemperatureWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        panzerPrefs.sliGaugeSize.Value = Val(PzGTemperatureGaugeSize)
    End If

    'Anemometer gauge
    
    If cmbGaugeType.ListIndex = 1 Then
        If aspectRatio = "landscape" Then
            If PzGDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = PzGAnemometerFormHighDpiXPos
                txtLandscapeVoffset.Text = PzGAnemometerFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGAnemometerFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGAnemometerFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = PzGAnemometerFormLowDpiXPos
                txtLandscapeVoffset.Text = PzGAnemometerFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGAnemometerFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGAnemometerFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fAnemometer.anemometerGaugeForm.Left
            txtPortraitVoffset.Text = fAnemometer.anemometerGaugeForm.Top
            If PzGDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGAnemometerFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGAnemometerFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGAnemometerFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGAnemometerFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayAnemoWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        
        panzerPrefs.sliGaugeSize.Value = Val(PzGAnemometerGaugeSize)

    End If
    
    'humidity gauge
    
    If cmbGaugeType.ListIndex = 2 Then
        If aspectRatio = "landscape" Then
            If PzGDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = PzGHumidityFormHighDpiXPos
                txtLandscapeVoffset.Text = PzGHumidityFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGHumidityFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGHumidityFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = PzGHumidityFormLowDpiXPos
                txtLandscapeVoffset.Text = PzGHumidityFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGHumidityFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGHumidityFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fHumidity.humidityGaugeForm.Left
            txtPortraitVoffset.Text = fHumidity.humidityGaugeForm.Top
            If PzGDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGHumidityFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGHumidityFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGHumidityFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGHumidityFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayHumidWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        
        panzerPrefs.sliGaugeSize.Value = Val(PzGHumidityGaugeSize)

    End If
    
    'barometer gauge
    
    If cmbGaugeType.ListIndex = 3 Then
        If aspectRatio = "landscape" Then
            If PzGDpiAwareness = "1" Then
                txtLandscapeHoffset.Text = PzGBarometerFormHighDpiXPos
                txtLandscapeVoffset.Text = PzGBarometerFormHighDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGBarometerFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGBarometerFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.Text = PzGBarometerFormLowDpiXPos
                txtLandscapeVoffset.Text = PzGBarometerFormLowDpiYPos
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGBarometerFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGBarometerFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fBarometer.barometerGaugeForm.Left
            txtPortraitVoffset.Text = fBarometer.barometerGaugeForm.Top
            If PzGDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGBarometerFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGBarometerFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGBarometerFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGBarometerFormLowDpiYPos & "px"
            End If
        End If
                
        If overlayAnemoWidget.Locked Then
            chkPreventDragging.Value = 1
        Else
            chkPreventDragging.Value = 0
        End If
        
        panzerPrefs.sliGaugeSize.Value = Val(PzGBarometerGaugeSize)

    End If
    
   On Error GoTo 0
   Exit Sub

cmbGaugeType_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbGaugeType_Click of Form panzerPrefs"
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
    PzGWindSpeedScale = LTrim$(Str$(cmbWindSpeedScale.ListIndex))
    sPutINISetting "Software\PzTemperatureGauge", "metricImperial", PzGMetricImperial, PzGSettingsFile
    
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
'    PzGWindSpeedScale = LTrim$(Str$(cmbWindSpeedScale.ListIndex))
'    sPutINISetting "Software\PzTemperatureGauge", "windSpeedScale", PzGWindSpeedScale, PzGSettingsFile
        
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
    
    PzGPrefsLoadedFlg = False
    prefsDynamicSizingFlg = False
    lastFormHeight = 0

    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form panzerPrefs"
    
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
        
    prefsStartupFlg = True ' this is used to prevent some control initialisations from running code at startup
    prefsDynamicSizingFlg = False
    PzGPrefsLoadedFlg = True ' this is a variable tested by an added form property to indicate whether the form is loaded or not
    PzGWindowLevelWasChanged = False
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

    If PzGDpiAwareness = "1" Then
        prefsDynamicSizingFlg = True
        chkEnableResizing.Value = 1
        lblDragCorner.Visible = True
    End If
      
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
    If PzGDpiAwareness = "1" Then
        If prefsFormHeight < screenHeightTwips Then
            Me.Height = prefsFormHeight
        Else
            Me.Height = screenHeightTwips - 1000
        End If
    End If
    
    ' position the prefs on the current monitor
    Call positionPrefsMonitor
    
    ' set the Z order of the prefs form
    Call setPrefsFormZordering
    
    ' start the timer that records the prefs position every 10 seconds
    positionTimer.Enabled = True
    
    ' end the startup by un-setting the start flag
    prefsStartupFlg = False

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form panzerPrefs"

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
    
    If PzGDpiAwareness = "1" Then
        formLeftTwips = Val(PzGPrefsFormHighDpiXPosTwips)
        formTopTwips = Val(PzGPrefsFormHighDpiYPosTwips)
    Else
        formLeftTwips = Val(PzGPrefsFormLowDpiXPosTwips)
        formTopTwips = Val(PzGPrefsFormLowDpiYPosTwips)
    End If
    
    If formLeftTwips = 0 Then
        If ((fTemperature.temperatureGaugeForm.Left + fTemperature.temperatureGaugeForm.Width) * screenTwipsPerPixelX) + 200 + Me.Width > screenWidthTwips Then
            Me.Left = (fTemperature.temperatureGaugeForm.Left * screenTwipsPerPixelX) - (Me.Width + 200)
        End If
    End If

    ' if a current location not stored then position to the middle of the screen
    If formLeftTwips <> 0 Then
        Me.Left = formLeftTwips
    Else
        Me.Left = screenWidthTwips / 2 - Me.Width / 2
    End If
    
    If formTopTwips <> 0 Then
        Me.Top = formTopTwips
    Else
        Me.Top = Screen.Height / 2 - Me.Height / 2
    End If
    
    monitorCount = fGetMonitorCount
    If monitorCount > 1 Then Call adjustFormPositionToCorrectMonitor(Me.hwnd, formLeftTwips, formTopTwips)
    
    ' calculate the on-screen widget position
    If Me.Left < 0 Then
        Me.Left = 10
    End If
    If Me.Top < 0 Then
        Me.Top = 0
    End If
    If Me.Left > screenWidthTwips - 2500 Then
        Me.Left = screenWidthTwips - 2500
    End If
    If Me.Top > screenHeightTwips - 2500 Then
        Me.Top = screenHeightTwips - 2500
    End If
    
    On Error GoTo 0
    Exit Sub

positionPrefsMonitor_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsMonitor of Form panzerPrefs"
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

    If PzGEnableBalloonTooltips = "1" Then CreateToolTip btnResetMessages.hwnd, "The various pop-up messages that this program generates can be manually hidden. This button restores them to their original visible state.", _
                  TTIconInfo, "Help on the message reset button", , , , True

    On Error GoTo 0
    Exit Sub

btnResetMessages_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_MouseMove of Form panzerPrefs"
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
    If prefsStartupFlg = False Then ' don't run this on startup
                    
        answer = vbYes
        answerMsg = "You must close this widget and HARD restart it, in order to change the widget's DPI awareness (a simple soft reload just won't cut it), do you want me to close and restart this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "DpiAwareness Confirmation", True, "chkDpiAwarenessRestart")
        
        If chkDpiAwareness.Value = 0 Then
            PzGDpiAwareness = "0"
        Else
            PzGDpiAwareness = "1"
        End If

        sPutINISetting "Software\PzTemperatureGauge", "dpiAwareness", PzGDpiAwareness, PzGSettingsFile
        
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "OK, the widget is still DPI aware until you restart. Some forms may show abnormally."
            answer = msgBoxA(answerMsg, vbOKOnly, "DpiAwareness Notification", True, "chkDpiAwarenessAbnormal")
        
            Exit Sub
        Else

            sPutINISetting "Software\PzTemperatureGauge", "dpiAwareness", PzGDpiAwareness, PzGSettingsFile
            'Call reloadWidget ' this is insufficient, image controls still fail to resize and autoscale correctly
            Call hardRestart
        End If

    End If

   On Error GoTo 0
   Exit Sub

chkDpiAwareness_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkDpiAwareness_Click of Form panzerPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkEnablePrefsTooltips_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 07/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnablePrefsTooltips_Click()

   On Error GoTo chkEnablePrefsTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If prefsStartupFlg = False Then
        If chkEnablePrefsTooltips.Value = 1 Then
            PzGEnablePrefsTooltips = "1"
        Else
            PzGEnablePrefsTooltips = "0"
        End If
        
        sPutINISetting "Software\PzTemperatureGauge", "enablePrefsTooltips", PzGEnablePrefsTooltips, PzGSettingsFile

    End If
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

   On Error GoTo 0
   Exit Sub

chkEnablePrefsTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnablePrefsTooltips_Click of Form panzerPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkEnableTooltips_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableTooltips_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo chkEnableTooltips_MouseMove_Error

    If PzGEnableBalloonTooltips = "1" Then CreateToolTip chkEnableTooltips.hwnd, "There is a problem with the 'standard' tooltips on the gauge elements as they resize along with the program graphical elements, meaning that they cannot be seen, there is also a problem with tooltip handling different fonts, hoping to get Olaf to fix these soon. My suggestion is to turn them off for the moment.", _
                  TTIconInfo, "Help on the Program tooltip problem", , , , True

    On Error GoTo 0
    Exit Sub

chkEnableTooltips_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableTooltips_MouseMove of Form panzerPrefs"
End Sub

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
        PzGShowTaskbar = "1"
    Else
        PzGShowTaskbar = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowTaskbar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowTaskbar_Click of Form panzerPrefs"
End Sub



'Private Sub cmbTickSwitchPref_Click()
'   btnSave.Enabled = True ' enable the save button
'
'    If cmbTickSwitchPref.ListIndex = 0 Then
'        overlayAnemoWidget.pointerAnimate = False
'        overlayTemperatureWidget.pointerAnimate = False
'        PzGPointerAnimate = "0"
'    Else
'        overlayAnemoWidget.pointerAnimate = True
'        overlayTemperatureWidget.pointerAnimate = True
'        PzGPointerAnimate = "1"
'    End If
'
'    sPutINISetting "Software\PzAnemometerGauge", "pointerAnimate", PzGPointerAnimate, PzGSettingsFile
'    panzerPrefs.cmbTickSwitchPref.ListIndex = Val(PzGPointerAnimate)
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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_Click of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAboutDebugInfo_Click of form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDonate_Click of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFacebook_Click of Form panzerPrefs"
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
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnOpenFile_Click of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdate_Click of Form panzerPrefs"
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
    overlayTemperatureWidget.Ticking = chkGaugeFunctions.Value

    On Error GoTo 0
    Exit Sub

chkGaugeFunctions_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGaugeFunctions_Click of Form panzerPrefs"
End Sub

Private Sub chkGenStartup_Click()
    btnSave.Enabled = True ' enable the save button
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
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDefaultEditor_Click of Form panzerPrefs"
            Resume Next
          End If
    End With
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkEnableBalloonTooltips_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableBalloonTooltips_Click()
   On Error GoTo chkEnableBalloonTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    If chkEnableBalloonTooltips.Value = 1 Then
        PzGEnableBalloonTooltips = "1"
    Else
        PzGEnableBalloonTooltips = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkEnableBalloonTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableBalloonTooltips_Click of Form panzerPrefs"
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
        PzGIgnoreMouse = "0"
    Else
        PzGIgnoreMouse = "1"
    End If

    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkIgnoreMouse_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkIgnoreMouse_Click of Form panzerPrefs"
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
            PzGPreventDraggingTemperature = "0"
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
            PzGPreventDraggingTemperature = "1"
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
            PzGPreventDraggingAnemometer = "0"
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
            PzGPreventDraggingAnemometer = "1"
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
            PzGPreventDraggingHumidity = "0"
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
            PzGPreventDraggingHumidity = "1"
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
            PzGPreventDraggingBarometer = "0"
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
            PzGPreventDraggingBarometer = "1"
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

    On Error GoTo 0
    Exit Sub

chkPreventDragging_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkPreventDragging_Click of Form panzerPrefs"
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
        PzGWidgetHidden = "0"
    Else
        'overlayTemperatureWidget.Hidden = True
        fTemperature.temperatureGaugeForm.Visible = False


        frmTimer.revealWidgetTimer.Enabled = True
        PzGWidgetHidden = "1"
    End If
    
    sPutINISetting "Software\PzTemperatureGauge", "widgetHidden", PzGWidgetHidden, PzGSettingsFile
    
    btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkWidgetHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWidgetHidden_Click of Form panzerPrefs"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbAspectHidden_Click of Form panzerPrefs"
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
        txtDefaultEditor.Text = PzGDefaultEditor
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
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDebug_Click of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbScrollWheelDirection_Click of Form panzerPrefs"
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
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetPosition_Click of Form panzerPrefs"
            Resume Next
          End If
    End With
End Sub




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

    If PzGPrefsLoadedFlg Then
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
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form panzerPrefs"
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
    
    If PzGLastSelectedTab = "general" Then Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton)  ' was imgGeneralMouseUpEvent
    If PzGLastSelectedTab = "config" Then Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton)     ' was imgConfigMouseUpEvent
    If PzGLastSelectedTab = "position" Then Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    If PzGLastSelectedTab = "development" Then Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    If PzGLastSelectedTab = "fonts" Then Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    If PzGLastSelectedTab = "sounds" Then Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    If PzGLastSelectedTab = "window" Then Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    If PzGLastSelectedTab = "about" Then Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)

   On Error GoTo 0
   Exit Sub

showLastTab_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showLastTab of Form panzerPrefs"

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

    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - 50
    btnHelp.Left = frameLeft
    

   On Error GoTo 0
   Exit Sub

positionPrefsFramesButtons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsFramesButtons of Form panzerPrefs"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form panzerPrefs"
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
            Call ShellExecute(Me.hwnd, "Open", App.path & "\help\Help.chm", vbNullString, App.path, 1)
        Else
            MsgBox ("%Err-I-ErrorNumber 11 - The help file - Help.chm - is missing from the help folder.")
        End If

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form panzerPrefs"
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
    PzGEnableTooltips = LTrim$(Str$(chkEnableTooltips.Value))
    PzGEnablePrefsTooltips = LTrim$(Str$(chkEnablePrefsTooltips.Value))
    PzGEnableBalloonTooltips = LTrim$(Str$(chkEnableBalloonTooltips.Value))
    PzGShowTaskbar = LTrim$(Str$(chkShowTaskbar.Value))
    PzGDpiAwareness = LTrim$(Str$(chkDpiAwareness.Value))
    
    'PzGTemperatureGaugeSize = LTrim$(Str$(sliGaugeSize.Value))
    'PzGAnemometerGaugeSize = LTrim$(Str$(sliAnemometerGaugeSize.Value))
    
'    PzGClipbBSize = LTrim$(Str$(sliGaugeSize.Value))
'    PzGSelectorSize = LTrim$(Str$(sliGaugeSize.Value))
    
    PzGScrollWheelDirection = LTrim$(Str$(cmbScrollWheelDirection.ListIndex))
    
    ' general
    PzGGaugeFunctions = LTrim$(Str$(chkGaugeFunctions.Value))
    PzGStartup = LTrim$(Str$(chkGenStartup.Value))
    
    PzGIcao = txtIcao.Text

    
    'PzGPointerAnimate = cmbTickSwitchPref.ListIndex
    PzGSamplingInterval = LTrim$(Str$(sliSamplingInterval.Value))
    PzGStormTestInterval = LTrim$(Str$(sliStormTestInterval.Value))
    PzGErrorInterval = LTrim$(Str$(sliErrorInterval.Value))
    
    PzGAirportsURL = txtAirportsURL.Text
    
    PzGTemperatureScale = cmbTemperatureScale.ListIndex
    PzGPressureScale = cmbPressureScale.ListIndex
    PzGMetricImperial = cmbMetricImperial.ListIndex
    PzGWindSpeedScale = cmbWindSpeedScale.ListIndex
    
    ' sounds
    PzGEnableSounds = LTrim$(Str$(chkEnableSounds.Value))
    
    'development
    PzGDebug = LTrim$(Str$(cmbDebug.ListIndex))
    PzGDblClickCommand = txtDblClickCommand.Text
    PzGOpenFile = txtOpenFile.Text
    PzGDefaultEditor = txtDefaultEditor.Text
    
    ' position
    PzGAspectHidden = LTrim$(Str$(cmbAspectHidden.ListIndex))
    PzGGaugeType = LTrim$(Str$(cmbGaugeType.ListIndex))
    
    PzGWidgetPosition = LTrim$(Str$(cmbWidgetPosition.ListIndex))
        
    If cmbGaugeType.ListIndex = 0 Then ' temperature
        PzGTemperatureLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        PzGTemperaturePortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        PzGTemperatureLandscapeLockedHoffset = txtLandscapeHoffset.Text
        PzGTemperatureLandscapeLockedVoffset = txtLandscapeVoffset.Text
        PzGTemperaturePortraitLockedHoffset = txtPortraitHoffset.Text
        PzGTemperaturePortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    If cmbGaugeType.ListIndex = 1 Then ' anemometer
        PzGAnemometerLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        PzGAnemometerPortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        PzGAnemometerLandscapeLockedHoffset = txtLandscapeHoffset.Text
        PzGAnemometerLandscapeLockedVoffset = txtLandscapeVoffset.Text
        PzGAnemometerPortraitLockedHoffset = txtPortraitHoffset.Text
        PzGAnemometerPortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    If cmbGaugeType.ListIndex = 2 Then ' humidity
        PzGHumidityLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        PzGHumidityPortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        PzGHumidityLandscapeLockedHoffset = txtLandscapeHoffset.Text
        PzGHumidityLandscapeLockedVoffset = txtLandscapeVoffset.Text
        PzGHumidityPortraitLockedHoffset = txtPortraitHoffset.Text
        PzGHumidityPortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    If cmbGaugeType.ListIndex = 3 Then ' Barometer
        PzGBarometerLandscapeLocked = LTrim$(Str$(cmbLandscapeLocked.ListIndex))
        PzGBarometerPortraitLocked = LTrim$(Str$(cmbPortraitLocked.ListIndex))
        PzGBarometerLandscapeLockedHoffset = txtLandscapeHoffset.Text
        PzGBarometerLandscapeLockedVoffset = txtLandscapeVoffset.Text
        PzGBarometerPortraitLockedHoffset = txtPortraitHoffset.Text
        PzGBarometerPortraitLockedVoffset = txtPortraitVoffset.Text
    End If
    
    'deanieboy
        
'    PzGTemperatureVLocationPerc
'    PzGTemperatureHLocationPerc

    ' fonts
    PzGPrefsFont = txtPrefsFont.Text
    PzGTempFormFont = PzGPrefsFont
    
    ' the sizing is not saved here again as it saved during the setting phase.
    
'    If PzGDpiAwareness = "1" Then
'        PzGPrefsFontSizeHighDPI = txtPrefsFontSize.Text
'    Else
'        PzGPrefsFontSizeLowDPI = txtPrefsFontSize.Text
'    End If
    'PzGPrefsFontItalics = txtFontSize.Text

    ' Windows
    PzGWindowLevel = LTrim$(Str$(cmbWindowLevel.ListIndex))
    PzGOpacity = LTrim$(Str$(sliOpacity.Value))
    PzGWidgetHidden = LTrim$(Str$(chkWidgetHidden.Value))
    PzGHidingTime = LTrim$(Str$(cmbHidingTime.ListIndex))
    PzGIgnoreMouse = LTrim$(Str$(chkIgnoreMouse.Value))
            
    'development
    PzGDebug = LTrim$(Str$(cmbDebug.ListIndex))
    PzGDblClickCommand = txtDblClickCommand.Text
    PzGOpenFile = txtOpenFile.Text
    PzGDefaultEditor = txtDefaultEditor.Text
            
    If PzGStartup = "1" Then
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PzTemperatureGaugeWidget", """" & App.path & "\" & "Panzer Temperature Gauge.exe""")
    Else
        Call savestring(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PzTemperatureGaugeWidget", vbNullString)
    End If

    ' save the values from the general tab
    If fFExists(PzGSettingsFile) Then
        sPutINISetting "Software\PzTemperatureGauge", "enableTooltips", PzGEnableTooltips, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "enablePrefsTooltips", PzGEnablePrefsTooltips, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "enableBalloonTooltips", PzGEnableBalloonTooltips, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "showTaskbar", PzGShowTaskbar, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "dpiAwareness", PzGDpiAwareness, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "temperatureGaugeSize", PzGTemperatureGaugeSize, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureLandscapeLocked", PzGTemperatureLandscapeLocked, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperaturePortraitLocked", PzGTemperaturePortraitLocked, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureVLocationPerc", PzGTemperatureVLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureHLocationPerc", PzGTemperatureHLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormHighDpiXPos", PzGTemperatureFormHighDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormHighDpiYPos", PzGTemperatureFormHighDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormLowDpiXPos", PzGTemperatureFormLowDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "temperatureFormLowDpiYPos", PzGTemperatureFormLowDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "preventDraggingTemperature", PzGPreventDraggingTemperature, PzGSettingsFile
        
        sPutINISetting "Software\PzAnemometerGauge", "anemometerGaugeSize", PzGAnemometerGaugeSize, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerLandscapeLocked", PzGAnemometerLandscapeLocked, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerPortraitLocked", PzGAnemometerPortraitLocked, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerVLocationPerc", PzGAnemometerVLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerHLocationPerc", PzGAnemometerHLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormHighDpiXPos", PzGAnemometerFormHighDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormHighDpiYPos", PzGAnemometerFormHighDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormLowDpiXPos", PzGAnemometerFormLowDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "anemometerFormLowDpiYPos", PzGAnemometerFormLowDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzAnemometerGauge", "preventDraggingAnemometer", PzGPreventDraggingAnemometer, PzGSettingsFile
               
        sPutINISetting "Software\PzHumidityGauge", "humidityGaugeSize", PzGHumidityGaugeSize, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityLandscapeLocked", PzGHumidityLandscapeLocked, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityPortraitLocked", PzGHumidityPortraitLocked, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityVLocationPerc", PzGHumidityVLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityHLocationPerc", PzGHumidityHLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormHighDpiXPos", PzGHumidityFormHighDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormHighDpiYPos", PzGHumidityFormHighDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormLowDpiXPos", PzGHumidityFormLowDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "humidityFormLowDpiYPos", PzGHumidityFormLowDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzHumidityGauge", "preventDraggingHumidity", PzGPreventDraggingHumidity, PzGSettingsFile
               
        sPutINISetting "Software\PzBarometerGauge", "barometerGaugeSize", PzGBarometerGaugeSize, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerLandscapeLocked", PzGBarometerLandscapeLocked, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerPortraitLocked", PzGBarometerPortraitLocked, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerVLocationPerc", PzGBarometerVLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerHLocationPerc", PzGBarometerHLocationPerc, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormHighDpiXPos", PzGBarometerFormHighDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormHighDpiYPos", PzGBarometerFormHighDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormLowDpiXPos", PzGBarometerFormLowDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "barometerFormLowDpiYPos", PzGBarometerFormLowDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzBarometerGauge", "preventDraggingBarometer", PzGPreventDraggingBarometer, PzGSettingsFile


        sPutINISetting "Software\PzClipB", "clipBSize", PzGClipBSize, PzGSettingsFile
        sPutINISetting "Software\PzSelector", "selectorSize", PzGSelectorSize, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "scrollWheelDirection", PzGScrollWheelDirection, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "gaugeFunctions", PzGGaugeFunctions, PzGSettingsFile
'        sPutINISetting "Software\PzTemperatureGauge", "pointerAnimate", PzGPointerAnimate, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "samplingInterval", PzGSamplingInterval, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "stormTestInterval", PzGStormTestInterval, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "errorInterval", PzGErrorInterval, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "airportsURL", PzGAirportsURL, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "temperatureScale", PzGTemperatureScale, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "pressureScale", PzGPressureScale, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "windSpeedScale", PzGWindSpeedScale, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "metricImperial", PzGMetricImperial, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "icao", PzGIcao, PzGSettingsFile

        
        sPutINISetting "Software\PzTemperatureGauge", "aspectHidden", PzGAspectHidden, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "gaugeType", PzGGaugeType, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "widgetPosition", PzGWidgetPosition, PzGSettingsFile
        

        sPutINISetting "Software\PzTemperatureGauge", "prefsFont", PzGPrefsFont, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "tempFormFont", PzGTempFormFont, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeHighDPI", PzGPrefsFontSizeHighDPI, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeLowDPI", PzGPrefsFontSizeLowDPI, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontItalics", PzGPrefsFontItalics, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontColour", PzGPrefsFontColour, PzGSettingsFile

        'save the values from the Windows Config Items
        sPutINISetting "Software\PzTemperatureGauge", "windowLevel", PzGWindowLevel, PzGSettingsFile
        
        
        sPutINISetting "Software\PzTemperatureGauge", "opacity", PzGOpacity, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "widgetHidden", PzGWidgetHidden, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "hidingTime", PzGHidingTime, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "ignoreMouse", PzGIgnoreMouse, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "startup", PzGStartup, PzGSettingsFile

        sPutINISetting "Software\PzTemperatureGauge", "enableSounds", PzGEnableSounds, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "lastSelectedTab", PzGLastSelectedTab, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "debug", PzGDebug, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "dblClickCommand", PzGDblClickCommand, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "openFile", PzGOpenFile, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "defaultEditor", PzGDefaultEditor, PzGSettingsFile
        

      
        sPutINISetting "Software\PzClipB", "clipBFormHighDpiXPos", PzGClipBFormHighDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzClipB", "clipBFormHighDpiYPos", PzGClipBFormHighDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzClipB", "clipBFormLowDpiXPos", PzGClipBFormLowDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzClipB", "clipBFormLowDpiYPos", PzGClipBFormLowDpiYPos, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormHighDpiXPos", PzGSelectorFormHighDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormHighDpiYPos", PzGSelectorFormHighDpiYPos, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormLowDpiXPos", PzGSelectorFormLowDpiXPos, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "selectorFormLowDpiYPos", PzGSelectorFormLowDpiYPos, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "lastUpdated", PzGLastUpdated, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "metarPref", PzGMetarPref, PzGSettingsFile
        
        sPutINISetting "Software\PzTemperatureGauge", "oldPressureStorage", PzGOldPressureStorage, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "pressureStorageDate", PzGPressureStorageDate, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "currentPressureValue", PzGCurrentPressureValue, PzGSettingsFile
        
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
    
    ' reload here if the PzGWindowLevel Was Changed
    If PzGWindowLevelWasChanged = True Then
        PzGWindowLevelWasChanged = False
        Call reloadWidget
    End If
    
   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSave_Click of Form panzerPrefs"

End Sub

' set a var on a checkbox tick
'---------------------------------------------------------------------------------------
' Procedure : chkEnableTooltips_Click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 19/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableTooltips_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo chkEnableTooltips_Click_Error

    btnSave.Enabled = True ' enable the save button
    
    If prefsStartupFlg = False Then
        If chkEnableTooltips.Value = 1 Then
            PzGEnableTooltips = "1"
        Else
            PzGEnableTooltips = "0"
        End If
        
        sPutINISetting "Software\PzTemperatureGauge", "enableTooltips", PzGEnableTooltips, PzGSettingsFile

        answer = vbYes
        answerMsg = "You must soft reload this widget, in order to change the tooltip setting, do you want me to reload this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "Request to Enable Tooltips", True, "chkEnableTooltipsClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call reloadWidget
        End If
    End If


   On Error GoTo 0
   Exit Sub

chkEnableTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableTooltips_Click of Form panzerPrefs"

End Sub

Private Sub chkEnableSounds_Click()
    btnSave.Enabled = True ' enable the save button
End Sub

'Private Sub cmbRefreshInterval_Click()
'    btnSave.Enabled = True ' enable the save button
'End Sub

Private Sub cmbWindowLevel_Click()
    btnSave.Enabled = True ' enable the save button
    If prefsStartupFlg = False Then PzGWindowLevelWasChanged = True
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
    fntFont = PzGPrefsFont
    ' PzGTempFormFont
    
    If PzGDpiAwareness = "1" Then
        fntSize = Val(PzGPrefsFontSizeHighDPI)
    Else
        fntSize = Val(PzGPrefsFontSizeLowDPI)
    End If
    
    If fntSize = 0 Then fntSize = 8
    fntItalics = CBool(PzGPrefsFontItalics)
    fntColour = CLng(PzGPrefsFontColour)
        
    
    Call changeFont(panzerPrefs, True, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
    
    PzGPrefsFont = CStr(fntFont)
    PzGTempFormFont = PzGPrefsFont
    
    If PzGDpiAwareness = "1" Then
        PzGPrefsFontSizeHighDPI = CStr(fntSize)
        Call Form_Resize
    Else
        PzGPrefsFontSizeLowDPI = CStr(fntSize)
    End If
    
    PzGPrefsFontItalics = CStr(fntItalics)
    PzGPrefsFontColour = CStr(fntColour)

    If fFExists(PzGSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\PzTemperatureGauge", "prefsFont", PzGPrefsFont, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "tempFormFont", PzGTempFormFont, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeHighDPI", PzGPrefsFontSizeHighDPI, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontSizeLowDPI", PzGPrefsFontSizeLowDPI, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "prefsFontItalics", PzGPrefsFontItalics, PzGSettingsFile
        sPutINISetting "Software\PzTemperatureGauge", "PrefsFontColour", PzGPrefsFontColour, PzGSettingsFile
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrefsFont_Click of Form panzerPrefs"

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
    chkGaugeFunctions.Value = Val(PzGGaugeFunctions)
    chkGenStartup.Value = Val(PzGStartup)
'
'    txtBias.Visible = False
'    lblGeneral(1).Visible = False

    'If testDST = True Then cmbMainDaylightSaving.ListIndex = 1
    
    'set the choice for four timezone comboboxes that were populated from file.
'    cmbMainGaugeTimeZone.ListIndex = Val(PzGMainGaugeTimeZone)
'    cmbMainDaylightSaving.ListIndex = Val(PzGMainDaylightSaving)
        
    'txtBias.Text = tzDelta

'    cmbTickSwitchPref.ListIndex = Val(PzGPointerAnimate)
    
    cmbTemperatureScale.ListIndex = Val(PzGTemperatureScale)
    cmbPressureScale.ListIndex = Val(PzGPressureScale)
    cmbWindSpeedScale.ListIndex = Val(PzGWindSpeedScale)
    cmbMetricImperial.ListIndex = Val(PzGMetricImperial)
 
    sliSamplingInterval.Value = Val(PzGSamplingInterval)
    sliStormTestInterval.Value = Val(PzGStormTestInterval)
    sliErrorInterval.Value = Val(PzGErrorInterval)
     
    txtAirportsURL.Text = PzGAirportsURL
    
    txtIcao.Text = PzGIcao
    
    ' configuration tab
   
    ' check whether the size has been previously altered via ctrl+mousewheel on the widget
    sliTemperatureGaugeSizeOldValue = sliGaugeSize.Value
    'sliGaugeSize.Value = Val(PzGTemperatureGaugeSize) 'deaniebabe
    'sliAnemometerGaugeSize.Value = val(PzGAnemometerGaugeSize)
    
'    sliClipBSize.Value = val(PzGClipBSize)
'    sliSelectorSize.Value = val(PzGSelectorSize)

    If sliGaugeSize.Value <> sliTemperatureGaugeSizeOldValue Then
        btnSave.Visible = True
    End If
    
    cmbScrollWheelDirection.ListIndex = Val(PzGScrollWheelDirection)
    chkEnableTooltips.Value = Val(PzGEnableTooltips)
    chkEnableBalloonTooltips.Value = Val(PzGEnableBalloonTooltips)
    chkShowTaskbar.Value = Val(PzGShowTaskbar)
    chkDpiAwareness.Value = Val(PzGDpiAwareness)
    
    chkEnablePrefsTooltips.Value = Val(PzGEnablePrefsTooltips)
    
    ' sounds tab
    chkEnableSounds.Value = Val(PzGEnableSounds)
    
    ' development
    cmbDebug.ListIndex = Val(PzGDebug)
    txtDblClickCommand.Text = PzGDblClickCommand
    txtOpenFile.Text = PzGOpenFile
    txtDefaultEditor.Text = PzGDefaultEditor
    
     ' fonts tab
    If PzGPrefsFont <> vbNullString Then
        txtPrefsFont.Text = PzGPrefsFont
        If PzGDpiAwareness = "1" Then
            Call changeFormFont(panzerPrefs, PzGPrefsFont, Val(PzGPrefsFontSizeHighDPI), fntWeight, fntStyle, PzGPrefsFontItalics, PzGPrefsFontColour)
            txtPrefsFontSize.Text = PzGPrefsFontSizeHighDPI
        Else
            Call changeFormFont(panzerPrefs, PzGPrefsFont, Val(PzGPrefsFontSizeLowDPI), fntWeight, fntStyle, PzGPrefsFontItalics, PzGPrefsFontColour)
            txtPrefsFontSize.Text = PzGPrefsFontSizeLowDPI
        End If
    End If
    
    
    ' position tab
    cmbAspectHidden.ListIndex = Val(PzGAspectHidden)
    cmbGaugeType.ListIndex = Val(PzGGaugeType)
    
    cmbWidgetPosition.ListIndex = Val(PzGWidgetPosition)
        
    If PzGPreventDraggingTemperature = "1" Then
        If aspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = fTemperature.temperatureGaugeForm.Left
            txtLandscapeVoffset.Text = fTemperature.temperatureGaugeForm.Top
            If PzGDpiAwareness = "1" Then
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormLowDpiYPos & "px"
            End If
        Else
            txtPortraitHoffset.Text = fTemperature.temperatureGaugeForm.Left
            txtPortraitVoffset.Text = fTemperature.temperatureGaugeForm.Top
            If PzGDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormHighDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & PzGTemperatureFormLowDpiXPos & "px"
                txtPortraitVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & PzGTemperatureFormLowDpiYPos & "px"
            End If
        End If
    End If
    
    'cmbLandscapeLocked
    
    If cmbGaugeType.ListIndex = 0 Then ' temperature
        cmbLandscapeLocked.ListIndex = Val(PzGTemperatureLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(PzGTemperaturePortraitLocked)
        txtLandscapeHoffset.Text = PzGTemperatureLandscapeLockedHoffset
        txtLandscapeVoffset.Text = PzGTemperatureLandscapeLockedVoffset
        txtPortraitHoffset.Text = PzGTemperaturePortraitLockedHoffset
        txtPortraitVoffset.Text = PzGTemperaturePortraitLockedVoffset
    End If
    
    If cmbGaugeType.ListIndex = 1 Then ' Anemometer
        cmbLandscapeLocked.ListIndex = Val(PzGAnemometerLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(PzGAnemometerPortraitLocked)
        txtLandscapeHoffset.Text = PzGAnemometerLandscapeLockedHoffset
        txtLandscapeVoffset.Text = PzGAnemometerLandscapeLockedVoffset
        txtPortraitHoffset.Text = PzGAnemometerPortraitLockedHoffset
        txtPortraitVoffset.Text = PzGAnemometerPortraitLockedVoffset
    End If
      
    If cmbGaugeType.ListIndex = 3 Then ' Barometer
        cmbLandscapeLocked.ListIndex = Val(PzGBarometerLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(PzGBarometerPortraitLocked)
        txtLandscapeHoffset.Text = PzGBarometerLandscapeLockedHoffset
        txtLandscapeVoffset.Text = PzGBarometerLandscapeLockedVoffset
        txtPortraitHoffset.Text = PzGBarometerPortraitLockedHoffset
        txtPortraitVoffset.Text = PzGBarometerPortraitLockedVoffset
    End If
      
    If cmbGaugeType.ListIndex = 2 Then ' Humidity
        cmbLandscapeLocked.ListIndex = Val(PzGHumidityLandscapeLocked)
        cmbPortraitLocked.ListIndex = Val(PzGHumidityPortraitLocked)
        txtLandscapeHoffset.Text = PzGHumidityLandscapeLockedHoffset
        txtLandscapeVoffset.Text = PzGHumidityLandscapeLockedVoffset
        txtPortraitHoffset.Text = PzGHumidityPortraitLockedHoffset
        txtPortraitVoffset.Text = PzGHumidityPortraitLockedVoffset
    End If
    
    ' Windows tab
    cmbWindowLevel.ListIndex = Val(PzGWindowLevel)
    chkIgnoreMouse.Value = Val(PzGIgnoreMouse)
    
    'chkPreventDragging.Value = Val(PzGPreventDraggingTemperature)
    'chkPreventDragging.Value = val(PzGPreventDraggingAnemometer)
    'PzGPreventDraggingHumidity
    
    cmbGaugeType.ListIndex = 0
    
    sliOpacity.Value = Val(PzGOpacity)
    chkWidgetHidden.Value = Val(PzGWidgetHidden)
    cmbHidingTime.ListIndex = Val(PzGHidingTime)
    
        
   On Error GoTo 0
   Exit Sub

adjustPrefsControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsControls of Form panzerPrefs on line " & Erl

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
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populatePrefsComboBoxes of Form panzerPrefs"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readFileWriteComboBox of Form panzerPrefs"
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

   On Error GoTo 0
   Exit Sub

clearBorderStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearBorderStyle of Form panzerPrefs"

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
    
    Dim ratio As Double: ratio = 0
    Dim currentFont As Long: currentFont = 0
    
    On Error GoTo Form_Resize_Error
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' move the drag corner label along with the form's bottom right corner
    lblDragCorner.Move Me.ScaleLeft + Me.ScaleWidth - (lblDragCorner.Width + 40), _
               Me.ScaleTop + Me.ScaleHeight - (lblDragCorner.Height + 40)

    ratio = cPrefsFormHeight / cPrefsFormWidth
    
    If prefsDynamicSizingFlg = True Then
        
        If PzGDpiAwareness = "1" Then
            currentFont = PzGPrefsFontSizeHighDPI
        Else
            currentFont = PzGPrefsFontSizeLowDPI
        End If
        
        Call resizeControls(Me, prefsControlPositions(), prefsCurrentWidth, prefsCurrentHeight, currentFont)
        
        Call tweakPrefsControlPositions(Me, prefsCurrentWidth, prefsCurrentHeight)
        
        Me.Width = Me.Height / ratio ' maintain the aspect ratio

    Else
        If Me.WindowState = 0 Then ' normal
            If Me.Width > 9090 Then Me.Width = 9090
            If Me.Width < 9085 Then Me.Width = 9090
            If lastFormHeight <> 0 Then Me.Height = lastFormHeight
        End If
    End If
        
    'lblSize.Caption = "topIconWidth = " & topIconWidth & " imgGeneral width = " & imgGeneral.Width
    
    On Error GoTo 0
    Exit Sub

Form_Resize_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form panzerPrefs"
            Resume Next
          End If
    End With
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tweakPrefsControlPositions of Form panzerPrefs"

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

    PzGPrefsLoadedFlg = False
    
    Call writePrefsPosition
    
    Call DestroyToolTip

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form panzerPrefs"
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
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraAbout.hwnd, "The About tab tells you all about this program and its creation using VB6.", _
                  TTIconInfo, "Help on the About Tab", , , , True
End Sub
Private Sub fraConfigInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfigInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraConfigInner.hwnd, "The configuration panel is the location for optional configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True

End Sub
Private Sub fraConfig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraConfig.hwnd, "The configuration panel is the location for optional configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
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
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraDevelopment.hwnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True
End Sub


Private Sub fraDevelopmentInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopmentInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraDevelopmentInner.hwnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True

End Sub
Private Sub fraFonts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraFonts.hwnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the PzG program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True

End Sub
Private Sub fraFontsInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraFontsInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraFontsInner.hwnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the PzG program folder. Choose a small 8pt font for each.", _
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
    
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraGaugePosition.hwnd, "Select the gauge type first - then this section allows you to determine " _
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
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraGeneral.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraGeneralInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraGeneralInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraGeneralInner.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraPosition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraPosition.hwnd, "This section allows you to determine size, lockability and positioning of your widget in various ways on different screen aspect ratios. ", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub



Private Sub fraPositionBalloonBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraPositionBalloonBox.hwnd, "Aspect ratio is for tablets only. Don't fiddle with this unless you really know what you are doing. Here you can choose whether this widget is hidden by default in either landscape or portrait mode or not at all. This option allows you to have certain widgets that do not obscure the screen in either landscape or portrait. If you accidentally set it so you can't find your widget on screen then change the setting here to NONE.", _
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
 If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraSounds.hwnd, "The sound panel allows you to configure the sounds that occur within PzG. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraSoundsInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSoundsInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraSoundsInner.hwnd, "The sound panel allows you to configure the sounds that occur within PzG. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub

Private Sub fraWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraWindow.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub
Private Sub fraWindowInner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindowInner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If PzGEnableBalloonTooltips = "1" Then CreateToolTip fraWindowInner.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub




Private Sub imgGeneral_Click()
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub

Private Sub imgGeneral_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
End Sub



Private Sub lblGitHub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblGitHub.ForeColor = &H8000000D
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
 
    If prefsStartupFlg = False Then
        PzGErrorInterval = LTrim$(Str$(sliErrorInterval.Value))
        'overlayTemperatureWidget.samplingInterval = sliSamplingInterval.Value
        sPutINISetting "Software\PzTemperatureGauge", "sliErrorInterval", PzGErrorInterval, PzGSettingsFile
    End If
    
    On Error GoTo 0
    Exit Sub

sliErrorInterval_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliErrorInterval_Click, line " & Erl & "."

End Sub

Private Sub sliErrorInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip sliErrorInterval.hwnd, "Adjust to error reporting frequency (seconds). This is the interval by which the program determines if a feed is in error by failing to supply data. If the interval is reached and the feed provides no data, then an error message is displayed. A value of zero means that no error messages will be displayed.", _
                  TTIconInfo, "Help on the Error Interval", , , , True
End Sub

Private Sub sliGaugeSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip sliGaugeSize.hwnd, "Changing this slider will change the size of the chosen gauge selected above", _
                  TTIconInfo, "Help on the Gauge Size Slider", , , , True
End Sub

Private Sub sliOpacity_Change()
    btnSave.Enabled = True ' enable the save button
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

 
    If prefsStartupFlg = False Then
        PzGSamplingInterval = LTrim$(Str$(sliSamplingInterval.Value))
        overlayTemperatureWidget.samplingInterval = sliSamplingInterval.Value
        sPutINISetting "Software\PzTemperatureGauge", "samplingInterval", PzGSamplingInterval, PzGSettingsFile
        
    End If
    
    On Error GoTo 0
    Exit Sub

sliSamplingInterval_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliSamplingInterval_Click, line " & Erl & "."

End Sub

Private Sub sliSamplingInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip sliSamplingInterval.hwnd, "Adjust to determine gauge sampling frequency (seconds). This is the polling interval by which the widget attempts to get the data from the source (default 600 seconds or ten minutes). The metar source provider and the location itself determines when the data is actually provided - there are weather sensors that have to provide rea-time data and a real person somewhere is probably responsible for providing an actual forecast.*", _
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
 
    If prefsStartupFlg = False Then
        PzGStormTestInterval = LTrim$(Str$(sliStormTestInterval.Value))
        'overlayTemperatureWidget.samplingInterval = sliStormTestInterval.Value
        sPutINISetting "Software\PzTemperatureGauge", "stormTestInterval", PzGStormTestInterval, PzGSettingsFile
        
    End If

    
    On Error GoTo 0
    Exit Sub

sliStormTestInterval_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliStormTestInterval_Click, line " & Erl & "."

End Sub

Private Sub sliStormTestInterval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip sliStormTestInterval.hwnd, "Adjust to determine storm checking frequency (seconds). This is the interval by which the widget compares pressure drops of 1 millibar (default 3600 seconds or one hour) indicating the increased chance of a storm. If this condition is detected, it will light a red lamp on the barometer gauge", _
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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtAboutText_MouseDown of Form panzerPrefs"
End Sub

Private Sub txtAboutText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraScrollbarCover.Visible = False
End Sub

Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgAbout.Visible = False
    imgAboutClicked.Visible = True
End Sub
Private Sub imgAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
End Sub

Private Sub imgDevelopment_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDevelopment.Visible = False
    imgDevelopmentClicked.Visible = True
End Sub

Private Sub imgDevelopment_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    

    Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
End Sub

Private Sub imgFonts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFonts.Visible = False
    imgFontsClicked.Visible = True
End Sub

Private Sub imgFonts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
End Sub

Private Sub imgConfig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgConfig.Visible = False
    imgConfigClicked.Visible = True
End Sub

Private Sub imgConfig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
End Sub

Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub


Private Sub imgPosition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgPosition.Visible = False
    imgPositionClicked.Visible = True
End Sub

Private Sub imgPosition_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
End Sub

Private Sub imgSounds_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    imgSounds.Visible = False
    imgSoundsClicked.Visible = True
End Sub

Private Sub imgSounds_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Call imgSoundsMouseUpEvent

    Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
End Sub

Private Sub imgWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgWindow.Visible = False
    imgWindowClicked.Visible = True
End Sub

Private Sub imgWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
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

    On Error GoTo 0
    Exit Sub

sliGaugeSize_Change_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliGaugeSize_Change of Form panzerPrefs"

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

    If prefsStartupFlg = False Then
        PzGOpacity = LTrim$(Str$(sliOpacity.Value))
    
        sPutINISetting "Software\PzTemperatureGauge", "opacity", PzGOpacity, PzGSettingsFile
        
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliOpacity_Change of Form panzerPrefs"
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
        gblOriginatingForm = "panzerPrefsForm"
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
    

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form panzerPrefs"
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
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip txtAirportsURL.hwnd, "Don't change this unless your alternative has an airports file with the " _
        & "EXACT same format and layout. this is the full URL giving the location of the airports.dat file containing all the airfields in the world " _
        & "that ought to be providing a valid METAR data feed. This URL will be used when you select the menu option to download a new Airports.dat file.", _
                  TTIconInfo, "Airports ICAO Data Download", , , , True
End Sub

Private Sub txtDblClickCommand_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDefaultEditor_Change()
    btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtIcao_Change()
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip txtIcao.hwnd, "This shows the current ICAO code used to identify the weather feed source data. To change this field use the button to the right", _
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

    If chkEnablePrefsTooltips.Value = 1 Then
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
        chkEnableBalloonTooltips.ToolTipText = "Check the box to enable larger balloon tooltips for all controls on the main program"
        chkShowTaskbar.ToolTipText = "Check the box to show the widget in the taskbar"
        chkEnableTooltips.ToolTipText = "Check the box to enable tooltips for all controls on the main program"
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
        chkEnablePrefsTooltips.ToolTipText = "Check the box to enable tooltips for all controls in the preferences utility"
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
        chkEnableBalloonTooltips.ToolTipText = vbNullString
        chkShowTaskbar.ToolTipText = vbNullString
        chkEnableTooltips.ToolTipText = vbNullString
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
        chkEnablePrefsTooltips.ToolTipText = vbNullString
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsTooltips of Form panzerPrefs"

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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsLabels of Form panzerPrefs"
        
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DestroyToolTip of Form panzerPrefs"
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
    
    Call LoadFileToTB(txtAboutText, App.path & "\resources\txt\about.txt", False)

   On Error GoTo 0
   Exit Sub

loadPrefsAboutText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPrefsAboutText of Form panzerPrefs"
    
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

    PzGLastSelectedTab = thisTabName
    sPutINISetting "Software\PzTemperatureGauge", "lastSelectedTab", PzGLastSelectedTab, PzGSettingsFile

    thisFraName.Visible = True
    thisFraButtonName.BorderStyle = 1

    ' Get the form's current scale factors.
    y_scale = Me.ScaleHeight / prefsCurrentHeight
    
    If PzGDpiAwareness = "1" Then
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
    If prefsDynamicSizingFlg = False Then
        padding = 200 ' add normal padding below the help button to position the bottom of the form

        lastFormHeight = btnHelp.Top + btnHelp.Height + captionHeight + borderWidth + padding
        Me.Height = lastFormHeight
    End If
    
    If PzGDpiAwareness = "0" Then
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picButtonMouseUpEvent of Form panzerPrefs"

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
'            panzerPrefs.Refresh
'        End If
'    Next useloop
'
'   On Error GoTo 0
'   Exit Sub
'
'scrollFrameDownward_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure scrollFrameDownward of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form panzerPrefs"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form panzerPrefs"
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
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicenceA_Click of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto_Click of Form panzerPrefs"
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
    
    PzGSkinTheme = "dark"

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_Click of Form panzerPrefs"
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
    
    PzGSkinTheme = "light"

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_Click of Form panzerPrefs"
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
'    panzerPrefs.sliAnimationInterval.BackColor = RGB(redC, greenC, blueC)
    'panzerPrefs.'sliWidgetSkew.BackColor = RGB(redC, greenC, blueC)
    sliGaugeSize.BackColor = RGB(redC, greenC, blueC)
    sliOpacity.BackColor = RGB(redC, greenC, blueC)
    txtAboutText.BackColor = RGB(redC, greenC, blueC)
    sliSamplingInterval.BackColor = RGB(redC, greenC, blueC)
    sliStormTestInterval.BackColor = RGB(redC, greenC, blueC)
    sliErrorInterval.BackColor = RGB(redC, greenC, blueC)
    
    sPutINISetting "Software\PzTemperatureGauge", "skinTheme", PzGSkinTheme, PzGSettingsFile

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
        PzGSkinTheme = "dark"
        
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

    If PzGSkinTheme <> vbNullString Then
        If PzGSkinTheme = "dark" Then
            Call setThemeShade(212, 208, 199)
        Else
            Call setThemeShade(240, 240, 240)
        End If
    Else
        If classicThemeCapable = True Then
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            themeTimer.Enabled = True
        Else
            PzGSkinTheme = "light"
            Call setModernThemeColours
        End If
    End If

   On Error GoTo 0
   Exit Sub

adjustPrefsTheme_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsTheme of Form panzerPrefs"
    
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
        PzGSkinTheme = "dark"
    Else ' 15790320
        Call setThemeShade(240, 240, 240)
        PzGSkinTheme = "light"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResPrefsImages of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionTimer_Timer of Form panzerPrefs"

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
        prefsDynamicSizingFlg = True
        txtPrefsFontCurrentSize.Visible = True
        lblCurrentFontsTab.Visible = True
        'Call writePrefsPosition
        chkEnableResizing.Caption = "Disable Corner Resizing"
    Else
        prefsDynamicSizingFlg = False
        txtPrefsFontCurrentSize.Visible = False
        lblCurrentFontsTab.Visible = False
        Unload panzerPrefs
        Me.Show
        Call readPrefsPosition
        chkEnableResizing.Caption = "Enable Corner Resizing"
    End If
    
    Call setframeHeights

   On Error GoTo 0
   Exit Sub

chkEnableResizing_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableResizing_Click of Form panzerPrefs"

End Sub

Private Sub chkEnableResizing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip chkEnableResizing.hwnd, "This allows you to resize the whole prefs window by dragging the bottom right corner of the window. It provides an alternative method of supporting high DPI screens.", _
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

    If prefsDynamicSizingFlg = True Then
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
    
        'If PzGDpiAwareness = "1" Then
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setframeHeights of Form panzerPrefs"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesDark of Form panzerPrefs"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesLight of Form panzerPrefs"

End Sub

Private Sub txtPrefsFontCurrentSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PzGEnableBalloonTooltips = "1" Then CreateToolTip txtPrefsFontCurrentSize.hwnd, "This is a read-only text box. It displays the current font as set when dynamic form resizing is enabled. Drag the right hand corner of the window downward and the form will auto-resize. This text box will display the resized font currently in operation for informational purposes only.", _
                  TTIconInfo, "Help on Setting the Font size Dynamically", , , , True
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populateTimeZoneRegions of Form panzerPrefs"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetTimeZoneArray of Form panzerPrefs"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsDaylightSavingTime of Form panzerPrefs"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetTZBiasByName of Form panzerPrefs"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TrimNull of Form panzerPrefs"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenRegKey of Form panzerPrefs"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsWinNTPlus of Form panzerPrefs"
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
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
    
    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseDown of Form panzerPrefs"

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

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseMove of Form panzerPrefs"
   
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

    If Val(PzGWindowLevel) = 0 Then
        Call SetWindowPos(Me.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PzGWindowLevel) = 1 Then
        Call SetWindowPos(Me.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PzGWindowLevel) = 2 Then
        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
    End If

   On Error GoTo 0
   Exit Sub

setPrefsFormZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsFormZordering of Module modMain"
End Sub


'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\


