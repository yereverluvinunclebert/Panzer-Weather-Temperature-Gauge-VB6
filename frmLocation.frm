VERSION 5.00
Begin VB.Form frmLocation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Location"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton btnGo 
      Caption         =   "GO"
      Height          =   345
      Left            =   4860
      TabIndex        =   8
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   4350
      TabIndex        =   7
      Top             =   1950
      Width           =   1050
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4350
      TabIndex        =   6
      Top             =   1290
      Width           =   1050
   End
   Begin VB.Frame fraOptions 
      Height          =   1110
      Left            =   2145
      TabIndex        =   2
      Top             =   1215
      Width           =   1680
      Begin VB.OptionButton optLocation 
         Caption         =   "Location"
         Height          =   165
         Left            =   255
         TabIndex        =   4
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optICAO 
         Caption         =   "ICAO"
         Height          =   165
         Left            =   255
         TabIndex        =   3
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.TextBox txtICAOInput 
      Height          =   375
      Left            =   2115
      TabIndex        =   0
      Text            =   "EGSH"
      Top             =   105
      Width           =   2595
   End
   Begin VB.Label lblDisplaySelection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EGSH Norwich International Airport"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   645
      Width           =   5175
   End
   Begin VB.Label lblEnterICAO 
      Caption         =   "Enter ICAO code"
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Top             =   150
      Width           =   1770
   End
End
Attribute VB_Name = "frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

