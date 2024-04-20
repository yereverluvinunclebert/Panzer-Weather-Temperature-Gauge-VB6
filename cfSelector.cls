VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfSelector
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents SelectorForm As cWidgetForm
Attribute SelectorForm.VB_VarHelpID = -1

Private WithEvents sTxtICAOInput As cwTextBox
Attribute sTxtICAOInput.VB_VarHelpID = -1
Public WithEvents sCmbMatchingLocations As cwDropDownList
Attribute sCmbMatchingLocations.VB_VarHelpID = -1

Public collSelectorPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long
Private PSDHeight As Long

Private optICAOSelected As Boolean
Private optLocationSelected As Boolean






'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   On Error GoTo Class_Initialize_Error

    optICAOSelected = True
    optLocationSelected = False
    
    Set collSelectorPSDNonUIElements = New_c.Collection(False)
    

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfSelector"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitSelectorFromPSD
' Author    : olaf schmidt & Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :    loads up the Cairo imageList with keyed entries and a surface from each and every PSD layer
'                adds path X,Y and alpha to the excluded collection
'                creates widgets with corresponding keynames, locations and sizes as per PSD for each layer
'                cwOverlayTemp class puts it on the screen
'---------------------------------------------------------------------------------------
'
Public Sub InitSelectorFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim i As Long: i = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitSelectorFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set SelectorForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName2, True, 1, 1)
        SelectorForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        PSDWidth = .Width
        PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
                    
                    If collSelectorPSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlay) to the "excluded paths" previously defined in Sub Main,
                        
                        collSelectorPSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), val(PzGOpacity) / 100)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                    
                        Set W = SelectorForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
                            W.ImageKey = W.key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget
                        If PzGOpacity < 1 Then
                            W.Alpha = val(PzGOpacity) / 100
                        Else
                            W.Alpha = 1
                        End If
                        
                        W.Tag = 0.05
                        
                        ' set the default characteristics of the non-interactive areas
                        ' note: the clickable layers characteristics are set in adjustSelectMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If PzGEnableTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL

                    End If
                End If
            End If
        Next i
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
    
    Set overlaySelectorWidget = SelectorForm.Widgets.Add(New cwOverlaySelect, "Selector", 0, 0, PSDWidth, PSDHeight)
    Set overlaySelectorWidget.cCollOverlaySelectorImages = collSelectorPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget

    Set sTxtICAOInput = fSelector.SelectorForm.Widgets.Add(New_W("cwTextBox"), "sTxtICAOInput", 195, 20, 118, 30)
    sTxtICAOInput.MultiLine = False
    sTxtICAOInput.Widget.FontSize = 15
    sTxtICAOInput.Widget.FontName = "times new roman"
    sTxtICAOInput.Widget.Alpha = 0.5

' Cairo.ImageList.AddIconFromResourceFile "btnIco", "shell32", 167 'load some resource-Icon

    Set sCmbMatchingLocations = fSelector.SelectorForm.Widgets.Add(New_W("cwDropDownList"), "sCmbMatchingLocations", 120, 58, 225, 28)
    sCmbMatchingLocations.Widget.FontSize = 15
    sCmbMatchingLocations.Widget.FontName = "times new roman"
    sCmbMatchingLocations.Widget.Alpha = 0.5
    sCmbMatchingLocations.DropDown.Caption = "ICAO"

   On Error GoTo 0
   Exit Sub

InitSelectorFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitSelectorFromPSD of Class Module cfSelector"
End Sub
Private Sub sCmbMatchingLocations_Click()
  With sCmbMatchingLocations.DataSource
    MsgBox !key & "  (ItemData=" & !Value & ")"
  End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SelectorForm_BubblingEvent
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub SelectorForm_BubblingEvent(ByRef sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   On Error GoTo SelectorForm_BubblingEvent_Error

    If PzGIgnoreMouse = "1" Then Exit Sub
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(sender.Widget.key)
            Case "sbtnselect"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click me to select the currently displayed ICAO location.", _
                     TTIconInfo, widgetName2 & " Select Button Help", , , , True
            Case "sbtnexit"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click me to close the ICAO selection utility.", _
                     TTIconInfo, widgetName2 & " Exit Button Help", , , , True
            Case "optlocationgreen"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "The selector utility is now searching by location name only.", _
                     TTIconInfo, widgetName2 & " Location Button Help", , , , True
            Case "optlocationred"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to enable the option to search by location name (airport).", _
                     TTIconInfo, widgetName2 & " Location Button Help", , , , True
            Case "opticaored"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to enable the option to search by ICAO code.", _
                     TTIconInfo, widgetName2 & " ICAO Button Help", , , , True
            Case "sbtnsearch"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to search using your supplied airport name or ICAO code.", _
                     TTIconInfo, widgetName2 & " Search Button Help", , , , True
            Case "opticaogreen"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "The selector utility is now searching by ICAO code only.", _
                     TTIconInfo, widgetName2 & " ICAO Button Help", , , , True
            Case "radiobody"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, " " _
                    & "This is the Airport Selector utility. Use this tool to select a nearby airport with a valid ICAO code. Any airport of this type will be able to supply " _
                    & "valid and up-to-date METAR weather data. Start by selecting the ICAO or Location option buttons below. This will allow you to search the supplied airports data " _
                    & "file for either an ICAO code or search for the airport by name, for example: HEATHROW." & vbCrLf & vbCrLf _
                    & "When you have a matching ICAO code, press OK to select and the main gauges will start to display weather information derived from that feed." _
                    & vbCrLf & "" _
                    & vbCrLf & "(Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize)" _
                    , TTIconInfo, widgetName2 & " Airport Selection Information.", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(sender.Widget.key)
            Case "sbtnselect"
               Call sbtnselect_mouseDown(sender)
            Case "sbtnexit"
               Call sbtnexit_mouseDown(sender)
            Case "sbtnsearch"
               Call sbtnsearch_mouseDown(sender)
            Case "optlocationred"
               Call optlocationred_mouseDown(sender)
            Case "opticaored"
               Call opticaored_mouseDown(sender)
'            Case "opticaogreen"
'                 Call opticaogreen_mouseDown(sender)
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
    
'    If EventName = "W_MouseUp" Then
'        Select Case LCase$(sender.Widget.key)
'            Case "housing/prefsbutton"
'                sender.Widget.Alpha = val(PzGOpacity) / 100
'                sender.Widget.Refresh
'            Case "housing/switchfacesbutton"
'               '  Call switchfacesbutton_mouseUp(sender)
'            Case "housing/stopbutton"
'               '  Call stopbutton_mouseUp(sender)
'            Case Else
'                Debug.Print sender.Widget.key
'        End Select
'    End If
    
'    If EventName = "W_Click" Then
'        Select Case LCase$(sender.Widget.key)
'            Case "housing/helpbutton"
'                '  Call helpbutton_click
'            Case "housing/startbutton"
'                '  Call startbutton_click
'            Case Else
'                Debug.Print sender.Widget.key
'        End Select
'    End If
   
   On Error GoTo 0
   Exit Sub

SelectorForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SelectorForm_BubblingEvent of Class Module cfSelector"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sbtnselect_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sbtnselect_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo sbtnselect_mouseDown_Error
    
    'sender.Widget.Refresh ' this removes the delay in the lockButton depressing
    
    fileToPlay = "shutdown.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
   On Error GoTo 0
   Exit Sub

sbtnselect_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sbtnselect_mouseDown of Class Module cfSelector"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : sbtnexit_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sbtnexit_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo sbtnexit_mouseDown_Error
    
    SelectorForm.Hide
    
    SelectorForm.Widgets("sbtnselect").Widget.Enabled = True
    SelectorForm.Widgets("sbtnsearch").Widget.Enabled = False
    
    PzGMetarPref = "Location"
    sPutINISetting "Software\PzTemperatureGauge", "metarPref", PzGMetarPref, PzGSettingsFile
    
    fileToPlay = "shutdown.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

sbtnexit_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sbtnexit_mouseDown of Class Module cfSelector"

End Sub



''---------------------------------------------------------------------------------------
'' Procedure : sbtnsearch_mouseDown
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 03/08/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub sbtnsearch_mouseDown(ByRef sender As Object)
'    Dim fileToPlay As String: fileToPlay = vbNullString
'
'    On Error GoTo sbtnsearch_mouseDown_Error
'
'    'SelectorForm.Hide
'
'    fileToPlay = "click.wav"
'    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'sbtnsearch_mouseDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sbtnsearch_mouseDown of Class Module cfSelector"
'
'End Sub


' ----------------------------------------------------------------
' Procedure Name: sbtnsearch_mouseDown
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 29/03/2024
' ----------------------------------------------------------------
Private Sub sbtnsearch_mouseDown(ByRef sender As Object)
    On Error GoTo btnGo_Click_Error
    
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim ee As String: ee = vbNullString
    Dim key As String: key = vbNullString
    Dim ff As String: ff = vbNullString
    Dim gg As String: gg = vbNullString
    Dim result As String: result = vbNullString
    Dim answerMsg As String: answerMsg = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo

    fSelector.sCmbMatchingLocations.Widget.Visible = False ' dropdown is not visible

    ee = UCase$(sTxtICAOInput.Text)
    
    fileToPlay = "click.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    ' if the input is an icao then handle it
    If optICAOSelected = True Then '"location"
        result = testICAO(ee)
        
        If result <> vbNullString Then ' icao
            answerMsg = "Done - Valid code Found. " & result
            answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Search Found", True, "sbtnsearchClick")
            
            sTxtICAOInput.Text = result
            'lblDisplaySelection.Caption = overlayTemperatureWidget.icaoLocation ' there is a lable underneath the combobox
            
            SelectorForm.Widgets("sbtnselect").Widget.Enabled = True
            SelectorForm.Widgets("sbtnsearch").Widget.Enabled = False
        End If
    End If
    
    ' if the input is an location then handle it
    If optLocationSelected = True Then ' "location"
        result = testLocation(ee)
        SelectorForm.Widgets("sbtnselect").Widget.Enabled = False
        SelectorForm.Widgets("sbtnsearch").Widget.Enabled = False
    End If
    
    'SelectorForm.Widgets("sbtnsearch").Widget.sTxtICAOInput.SetFocus ' need to implement this?
    
    On Error GoTo 0
    Exit Sub

btnGo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sbtnsearch_mouseDown, line " & Erl & "."

End Sub


'---------------------------------------------------------------------------------------
' Procedure : optlocationred_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub optlocationred_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo optlocationred_mouseDown_Error
   
    optICAOSelected = False
    optLocationSelected = True
   
    ' switch location red knob to nothing
    With SelectorForm.Widgets("optlocationred").Widget
        .Visible = False
        .Refresh
    End With
    
   ' switch location knob to green
    With SelectorForm.Widgets("optlocationgreen").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set icao green knob to invisible
    With SelectorForm.Widgets("opticaogreen").Widget
        .Visible = False
        .Refresh
    End With
    
    ' set icao knob to red to show icao is disabled
    With SelectorForm.Widgets("opticaored").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set red block to show location is enabled
    With SelectorForm.Widgets("enterlocation").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set red block to show icao is disabled
    With SelectorForm.Widgets("entericao").Widget
        .Visible = False
        .Refresh
    End With
    
    fileToPlay = "click1.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
   On Error GoTo 0
   Exit Sub

optlocationred_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optlocationred_mouseDown of Class Module cfSelector"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : opticaored_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub opticaored_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo opticaored_mouseDown_Error
   
    optICAOSelected = True
    optLocationSelected = False
   
    ' switch icao red knob to nothing
    With SelectorForm.Widgets("opticaored").Widget
        .Visible = False
        .Refresh
    End With
    
    ' switch icao knob to green
    With SelectorForm.Widgets("opticaogreen").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set location green knob to invisible
    With SelectorForm.Widgets("optlocationgreen").Widget
        .Visible = False
        .Refresh
    End With
    
    ' set location knob to red to show location is disabled
    With SelectorForm.Widgets("optlocationred").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set red block to show icao is enabled
    With SelectorForm.Widgets("entericao").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set red block to show location is disabled
    With SelectorForm.Widgets("enterlocation").Widget
        .Visible = False
        .Refresh
    End With
    
    fileToPlay = "click1.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    SelectorForm.Widgets("sbtnselect").Widget.Enabled = False
    SelectorForm.Widgets("sbtnsearch").Widget.Enabled = True
    fSelector.sCmbMatchingLocations.Widget.Visible = False
        
    PzGMetarPref = "ICAO"
    sPutINISetting "Software\PzTemperatureGauge", "metarPref", PzGMetarPref, PzGSettingsFile
    
       
    On Error GoTo 0
   Exit Sub

opticaored_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure opticaored_mouseDown of Class Module cfSelector"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SelectorForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo W_MouseDown_Error

    If Button = vbRightButton Then
        Call menuForm.PopupMenu(menuForm.mnuSelectorMenu)
    Else
    
        ' ensure that the associated forms pop to the front when the main gauge receives focus
        If fTemperature.temperatureGaugeForm.Visible = True Then fTemperature.temperatureGaugeForm.SetFocus
        If panzerPrefs.Visible = True Then panzerPrefs.SetFocus
        If fSelector.SelectorForm.Visible = True Then fSelector.SelectorForm.SetFocus
        
        If PzGIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cfSelector"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : SelectorForm_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SelectorForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = X: y0 = Y 'just store the offsets when no button is down
  
   On Error GoTo SelectorForm_MouseMove_Error
    
    If overlayTemperatureWidget.Locked = True Then Exit Sub
    If PzGIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not SelectorForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (X - x0) * SelectorForm.WidgetRoot.Zoom + SelectorForm.Left
      FY = (Y - y0) * SelectorForm.WidgetRoot.Zoom + SelectorForm.Top
      SelectorForm.Move FX, FY
    End If


   On Error GoTo 0
   Exit Sub

SelectorForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SelectorForm_MouseMove of Class Module cfSelector"
End Sub


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

    location = Replace(location, " ", "")

    If location <> "" Then
        If debugFlg = 1 Then
            Debug.Print ("%txtICAOInput - calling searchIcaoFile")
        End If

        ' note: it is possible that a named search location could contain a number
        ' call routine to search
        overlayTemperatureWidget.IcaoToTest = location ' load the icao to test
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
            fSelector.sCmbMatchingLocations.Widget.Visible = True
            'fSelector.sCmbMatchingLocations.DataSource. .Clear ' remove old from previous usage

'            For i = 0 To cnt - 1
'                fSelector.sCmbMatchingLocations.AddItem collValidLocations("key" & CStr(i + 1)) ' the cnt is the key
'                fSelector.sCmbMatchingLocations.ItemData(i) = i
'            Next i
            fSelector.sCmbMatchingLocations.ListIndex = 0 ' the default entry - Causes a click event to fire which is a pain.

            ' just a trial
            fSelector.sCmbMatchingLocations.SetDataSource collValidLocations, "collValidLocations"
            fSelector.sCmbMatchingLocations.DataSource.Sort = "Col-Add-Order"

        End If
    End If

    'if the station id returned is null then assume the weather information is missing for an unknown reason.
    If testLocation = vbNullString Then
        answerMsg = "No matching Location found "
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Location Error Information", False)
        testLocation = vbNullString
        Exit Function
    End If
'
    On Error GoTo 0
    Exit Function

testLocation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testLocation, line " & Erl & "."

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
