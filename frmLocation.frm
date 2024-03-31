VERSION 5.00
Begin VB.Form frmLocation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Temporary Form For Location Choice"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   Icon            =   "frmLocation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ComboBox cmbMatchingLocations 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   285
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   900
      Visible         =   0   'False
      Width           =   5130
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Search"
      Height          =   345
      Left            =   4590
      TabIndex        =   8
      Top             =   120
      Width           =   810
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   4350
      TabIndex        =   7
      Top             =   1830
      Width           =   1050
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "Select"
      Height          =   465
      Left            =   4350
      TabIndex        =   6
      Top             =   1290
      Width           =   1050
   End
   Begin VB.Frame fraOptions 
      Height          =   1110
      Left            =   2505
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
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox txtICAOInput 
      Height          =   375
      Left            =   2085
      TabIndex        =   0
      Top             =   105
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Airport by Name"
      Height          =   270
      Left            =   315
      TabIndex        =   11
      Top             =   1830
      Width           =   1755
   End
   Begin VB.Label lbl4char 
      Caption         =   "4-letter identifying code"
      Height          =   270
      Left            =   285
      TabIndex        =   10
      Top             =   1455
      Width           =   1755
   End
   Begin VB.Label lblDisplaySelection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   300
      TabIndex        =   5
      Top             =   645
      Width           =   5115
   End
   Begin VB.Label lblEnterICAO 
      Caption         =   "Enter ICAO code"
      Height          =   315
      Left            =   315
      TabIndex        =   1
      Top             =   150
      Width           =   1725
   End
End
Attribute VB_Name = "frmLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private icaoLocation1 As String
Private icaoLocation2 As String
Private icaoLocation3 As String
Private icaoLocation4 As String
Private icaoLocation5 As String

Private Sub btnExit_Click()
    frmLocation.Hide
End Sub

' ----------------------------------------------------------------
' Procedure Name: btnGo_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 29/03/2024
' ----------------------------------------------------------------
Private Sub btnGo_Click()
    On Error GoTo btnGo_Click_Error
    Dim ee As String: ee = vbNullString
    Dim key As String: key = vbNullString
    Dim ff As String: ff = vbNullString
    Dim gg As String: gg = vbNullString
    Dim result As String: result = vbNullString
    Dim answerMsg As String: answerMsg = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    ee = UCase$(txtICAOInput.Text)
    
    ' if the input is an icao then handle it
    If optICAO.Value = True Then '"location"
        result = testICAO(ee)
    End If
    
    ' if the input is an location then handle it
    If optLocation.Value = True Then ' "icao"
        result = testLocation(ee)
    End If
    
    If result <> vbNullString Then
        answerMsg = "Done - Valid code Found. " & result
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Good code", False)
        
        txtICAOInput.Text = result
        lblDisplaySelection.Caption = overlayTemperatureWidget.icaoLocation
        
    End If
    
    txtICAOInput.SetFocus
    
    On Error GoTo 0
    Exit Sub

btnGo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGo_Click, line " & Erl & "."

End Sub





' ----------------------------------------------------------------
' Procedure Name: testLocation
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter location (String):
' Return Type: String
' Author: beededea
' Date: 28/03/2024
' ----------------------------------------------------------------
Private Function testLocation(ByVal location As String) As String
    On Error GoTo testLocation_Error
    
    Dim answer As VbMsgBoxResult
    Dim answerMsg  As String: answerMsg = vbNullString
    
    location = Replace(location, " ", "")

    If location <> "" Then
        If debugFlg = 1 Then
            Debug.Print ("%txtICAOInput - calling searchIcaoFile")
        End If
        
        ' note: it is possible that a named search location could contain a number
        ' call routine to search
        overlayTemperatureWidget.IcaoToTest = location
        If overlayTemperatureWidget.ValidICAO = True Then
            testLocation = overlayTemperatureWidget.IcaoToTest  ' return
        End If
    End If
    
    'if the station id returned is null then assume the weather information is missing for an unknown reason.
    If testLocation = vbNullString Then
        answerMsg = "No matching Location found "
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Error Information", False)
        testLocation = vbNullString
        Exit Function
    End If
    
    On Error GoTo 0
    Exit Function

testLocation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testLocation, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: testICAO
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter icao (String):
' Return Type: String
' Author: beededea
' Date: 28/03/2024
' ----------------------------------------------------------------
Private Function testICAO(ByVal icao As String) As String
    On Error GoTo testICAO_Error
    Dim i As Integer
    Dim testChar As String
    Dim allLetters As Boolean: allLetters = False
    
    Dim answer As VbMsgBoxResult
    Dim answerMsg  As String: answerMsg = vbNullString
    
    'shorten the input to 4 characters if cut /pasted in with too many characters
    If Len(icao) > 4 Then
        icao = Mid$(icao, 0, 4)
        answerMsg = "Valid ICAO codes are only four digits long. Use the top sliding switch to select a city search. "
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Update Information", False)
        If debugFlg = 1 Then
            Debug.Print ("%txtICAOInput - txtICAOInput.data " & icao)
        End If
        testICAO = vbNullString
        Exit Function
    End If
    icao = Replace(icao, " ", "")
    If icao <> "" Then   ' no empty strings
        If debugFlg = 1 Then
            Debug.Print ("%txtICAOInput - calling searchIcaoFile")
        End If
        For i = 1 To 4
            testChar = Mid$(icao, i, 1)
            If IsLetter(testChar) = False Then
                allLetters = False
                Exit For
            End If
            allLetters = True
        Next i
        If allLetters = True Then
            ' set
            overlayTemperatureWidget.IcaoToTest = icao
            ' call routine to search
            If overlayTemperatureWidget.ValidICAO = True Then
                testICAO = icao  ' return
                'lblDisplaySelection.caption = overlayTemperatureWidget.icaoLocation
            End If
        Else
            answerMsg = "The ICAO search string cannot contain numeric or non alpha characters. "
            answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Error Information", False)
            testICAO = vbNullString
            Exit Function
        End If
    End If
    'if the station id returned is null then assume the weather information is missing for an unknown reason.
    If testICAO = vbNullString Then
        answerMsg = "No matching ICAO found "
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Error Information", False)
        testICAO = vbNullString
        Exit Function
    End If

    
    On Error GoTo 0
    Exit Function

testICAO_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testICAO, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: IsLetter
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter character (String):
' Return Type: Boolean
' Author: rpetrich
' Date: 28/03/2024
' ----------------------------------------------------------------
Private Function IsLetter(ByVal character As String) As Boolean
    IsLetter = UCase$(character) <> LCase$(character)
End Function

Private Sub btnOK_Click()
    PzGIcao = overlayTemperatureWidget.IcaoToTest
    sPutINISetting "Software\PzTemperatureGauge", "icao", PzGIcao, PzGSettingsFile
    
    If panzerPrefs.Visible = True Then
        panzerPrefs.txtIcao = PzGIcao
    End If
    
    overlayTemperatureWidget.GetMetar = True
    
    frmLocation.Hide
        
End Sub

Private Sub Form_Load()
    txtICAOInput.Text = PzGIcao
    lblDisplaySelection.Caption = overlayTemperatureWidget.icaoLocation
End Sub


Private Sub optICAO_Click()
    PzGMetarPref = "ICAO"
    sPutINISetting "Software\PzTemperatureGauge", "metarPref", PzGMetarPref, PzGSettingsFile

End Sub

Private Sub optLocation_Click()
    PzGMetarPref = "Location"
    sPutINISetting "Software\PzTemperatureGauge", "metarPref", PzGMetarPref, PzGSettingsFile
End Sub
