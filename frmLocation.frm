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
   Begin VB.CommandButton btnSearch 
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
   Begin VB.CommandButton btnSelect 
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
   Begin VB.ComboBox cmbMatchingLocations 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   285
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   630
      Visible         =   0   'False
      Width           =   5130
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

Private populateCombo As Boolean



' ----------------------------------------------------------------
' Procedure Name: btnSearch_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 29/03/2024
' ----------------------------------------------------------------
Private Sub btnSearch_Click()
    On Error GoTo btnGo_Click_Error
    Dim ee As String: ee = vbNullString
    Dim key As String: key = vbNullString
    Dim ff As String: ff = vbNullString
    Dim gg As String: gg = vbNullString
    Dim result As String: result = vbNullString
    Dim answerMsg As String: answerMsg = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    cmbMatchingLocations.Visible = False
    
    ee = UCase$(txtICAOInput.Text)
    
    ' if the input is an icao then handle it
    If optICAO.Value = True Then '"icao"
        result = testICAO(ee)
        
        If result <> vbNullString Then
            answerMsg = "Done - Valid code Found. " & result
            answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Search Found", True, "btnSearchClick")
            
            txtICAOInput.Text = result
            lblDisplaySelection.Caption = overlayTemperatureWidget.icaoLocation
            
            btnSelect.Enabled = True
            btnSearch.Enabled = False
        End If
    End If
    
    ' if the input is an location then handle it
    If optLocation.Value = True Then ' "location"
        result = testLocation(ee)
        btnSearch.Enabled = False
        btnSelect.Enabled = False
    End If
    
    txtICAOInput.SetFocus
    
    On Error GoTo 0
    Exit Sub

btnGo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSearch_Click, line " & Erl & "."

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
    
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    Dim cnt As Long: cnt = 0
    Dim i As Integer: i = 0
    
    debugFlg = 0
    
    location = Replace(location, " ", "")

    If location <> "" Then
        If debugFlg = 1 Then
            Debug.Print ("%txtICAOInput - calling searchIcaoFile")
        End If
        
        ' note: it is possible that a named search location could contain a number
        ' call routine to search
        overlayTemperatureWidget.StringToTest = location ' load the string to test
        cnt = overlayTemperatureWidget.ValidLocationCount
        If cnt = 1 Then
            If overlayTemperatureWidget.ValidICAO = True Then ' call routine to search all the ICAO codes in airport.dat
                testLocation = overlayTemperatureWidget.IcaoToTest  ' return
                Exit Function
            End If
        Else
        
            ' if the result contains too many matches, request a better search
            If cnt > 200 Then
                answerMsg = "Too many matches, " & cnt & " entries found matching that pattern, please enter a more unique search string "
                answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Error Information", False)
                testLocation = vbNullString  ' return
                Exit Function
            End If
            
            ' populate the combobox with the <=200 items found
            testLocation = "multiple locations found" ' vbNullString means none found
            cmbMatchingLocations.Visible = True
            cmbMatchingLocations.Clear ' remove old from previous usage
                        
            For i = 0 To cnt - 1
                cmbMatchingLocations.AddItem collValidLocations("key" & CStr(i + 1)) ' the cnt is the key
                cmbMatchingLocations.ItemData(i) = i
            Next i
            cmbMatchingLocations.ListIndex = 0 ' the default entry - Causes a click event to fire which is a pain.
            
            ' just a trial
            
            fSelector.sCmbMatchingLocations.SetDataSource collValidLocations, "collValidLocations"
            fSelector.sCmbMatchingLocations.DataSource.Sort = "Col-Add-Order"
            
            'fSelector.sCmbMatchingLocations.DropDown.Caption = fSelector.sCmbMatchingLocations.DataSource!key
            
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
    
    Dim i As Integer: i = 0
    Dim testChar As String: testChar = vbNullString
    Dim allLetters As Boolean: allLetters = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg  As String: answerMsg = vbNullString
    
    debugFlg = 0
    
    'shorten the input to 4 characters if cut /pasted in with too many characters
    If Len(icao) > 4 Then
        icao = Mid$(icao, 0, 4)
        answerMsg = "Valid ICAO codes are only four digits long. Use the options buttons below to select either an ICAO or a city search. "
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Error Information", False)
        If debugFlg = 1 Then
            Debug.Print ("%txtICAOInput - txtICAOInput.data " & icao)
        End If
        testICAO = vbNullString ' return
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

            overlayTemperatureWidget.IcaoToTest = icao ' load the icao to test
            ' call routine to search
            If overlayTemperatureWidget.ValidICAO = True Then ' call routine to search all the ICAO codes in airport.dat
                testICAO = icao  ' return
            End If
        Else
            answerMsg = "The ICAO code search string cannot contain numeric or non alpha characters. "
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
    On Error GoTo IsLetter_Error
    
    IsLetter = UCase$(character) <> LCase$(character)
    
    On Error GoTo 0
    Exit Function

IsLetter_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLetter, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: btnSelect_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 09/04/2024
' ----------------------------------------------------------------
Private Sub btnSelect_Click()
    On Error GoTo btnSelect_Click_Error
    
    btnSearch.Enabled = True
    btnSelect.Enabled = False
    PzGIcao = overlayTemperatureWidget.IcaoToTest
    sPutINISetting "Software\PzTemperatureGauge", "icao", PzGIcao, PzGSettingsFile
    
    If optLocation.Value = True Then
        ' call routine to search all the ICAO codes in airport.dat
        If overlayTemperatureWidget.ValidICAO = True Then
            sPutINISetting "Software\PzTemperatureGauge", "icao", PzGIcao, PzGSettingsFile
        End If
    End If
            
    If panzerPrefs.Visible = True Then
        panzerPrefs.txtIcao = PzGIcao
    End If
    
    ' trigger METAR get with new ICAO code
    overlayTemperatureWidget.GetMetar = True
    
    frmLocation.Hide
        
    On Error GoTo 0
    Exit Sub

btnSelect_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSelect_Click, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: cmbMatchingLocations_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 09/04/2024
' ----------------------------------------------------------------
Private Sub cmbMatchingLocations_Click()

    On Error GoTo cmbMatchingLocations_Click_Error
    
    Dim icaoData As String: icaoData = vbNullString
    Dim splitIcaoData() As String ' array
    Dim icaoLocation1 As String: icaoLocation1 = vbNullString
    Dim icaoLocation5 As String: icaoLocation5 = vbNullString
    
    btnSearch.Enabled = False
    
    icaoData = UCase$(cmbMatchingLocations.List(cmbMatchingLocations.ListIndex))
    If icaoData = vbNullString Then Exit Sub
    
    splitIcaoData = Split(icaoData, ",")
    
    icaoLocation1 = Replace(splitIcaoData(1), """", "") ' location
    icaoLocation5 = Replace(splitIcaoData(5), """", "") ' icao code
    
    If icaoLocation5 = "\N" Then
        icaoLocation5 = Replace(splitIcaoData(4), """", "")
    End If
    
    ' to prevent a click event causing the next few tasks to occur,
    ' happens when populating and selecting the default value for a combobox
    ' so we add a bypass flag to indicate these tasks should not occur at population time
    If populateCombo = True Then
        populateCombo = False
    Else
        txtICAOInput.Text = icaoLocation5
        lblDisplaySelection.Caption = icaoLocation1
        btnSearch.Enabled = False
        btnSelect.Enabled = True
    End If
    
    overlayTemperatureWidget.IcaoToTest = icaoLocation5 ' load the icao to test
    
    On Error GoTo 0
    Exit Sub

cmbMatchingLocations_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbMatchingLocations_Click, line " & Erl & "."

End Sub


' ----------------------------------------------------------------
' Procedure Name: Form_Load
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 09/04/2024
' ----------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo Form_Load_Error
    
    populateCombo = True
    
    btnSelect.Enabled = False
    
    txtICAOInput.Text = PzGIcao
    lblDisplaySelection.Caption = overlayTemperatureWidget.icaoLocation
    
    If PzGMetarPref = "ICAO" Then
        optICAO.Value = True
    Else
        optLocation.Value = True
    End If
    
    On Error GoTo 0
    Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load, line " & Erl & "."

End Sub


' ----------------------------------------------------------------
' Procedure Name: optICAO_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 09/04/2024
' ----------------------------------------------------------------
Private Sub optICAO_Click()
    On Error GoTo optICAO_Click_Error
    
    btnSearch.Enabled = True
    btnSelect.Enabled = False
    cmbMatchingLocations.Visible = False
    
    PzGMetarPref = "ICAO"
    sPutINISetting "Software\PzTemperatureGauge", "metarPref", PzGMetarPref, PzGSettingsFile
    
    On Error GoTo 0
    Exit Sub

optICAO_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optICAO_Click, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: optLocation_Click
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 09/04/2024
' ----------------------------------------------------------------
Private Sub optLocation_Click()
    On Error GoTo optLocation_Click_Error
    
    btnSearch.Enabled = True
    btnSelect.Enabled = False
    
    PzGMetarPref = "Location"
    sPutINISetting "Software\PzTemperatureGauge", "metarPref", PzGMetarPref, PzGSettingsFile
    
    On Error GoTo 0
    Exit Sub

optLocation_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optLocation_Click, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: txtICAOInput_Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 09/04/2024
' ----------------------------------------------------------------
Private Sub txtICAOInput_Change()
    On Error GoTo txtICAOInput_Change_Error
    
    populateCombo = True
    btnSearch.Enabled = True
    cmbMatchingLocations.Visible = False
    
    On Error GoTo 0
    Exit Sub

txtICAOInput_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtICAOInput_Change, line " & Erl & "."

End Sub

Private Sub btnExit_Click()
    frmLocation.Hide
End Sub
