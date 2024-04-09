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

Private WithEvents textBox As cwTextBox
Attribute textBox.VB_VarHelpID = -1

Public collSelectorPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long
Private PSDHeight As Long




'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   On Error GoTo Class_Initialize_Error

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

    Set textBox = fSelector.SelectorForm.Widgets.Add(New_W("cwTextBox"), "textBox", 195, 20, 118, 30)
    textBox.MultiLine = False
    textBox.Widget.FontSize = 15
    textBox.Widget.FontName = "times new roman"
    textBox.Widget.Alpha = 0.5

   On Error GoTo 0
   Exit Sub

InitSelectorFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitSelectorFromPSD of Class Module cfSelector"
End Sub

' enterlocation
' entericao
' radioknobone
' radioknobtwo
'
' okbutton
' exitbutton
' locationknobgreen
' locationknobred
' icaoknobred
' icaoknobgreen

' radiobody

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
            Case "okbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click me to select the currently displayed ICAO location.", _
                     TTIconInfo, widgetName2 & " Select Button Help", , , , True
            Case "exitbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click me to close the ICAO selection utility.", _
                     TTIconInfo, widgetName2 & " Exit Button Help", , , , True
            Case "locationknobgreen"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "The selector utility is now searching by location name only.", _
                     TTIconInfo, widgetName2 & " Location Button Help", , , , True
            Case "locationknobred"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to enable the option to search by location name (airport).", _
                     TTIconInfo, widgetName2 & " Location Button Help", , , , True
            Case "icaoknobred"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to enable the option to search by ICAO code.", _
                     TTIconInfo, widgetName2 & " ICAO Button Help", , , , True
            Case "searchbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to search using your supplied airport name or ICAO code.", _
                     TTIconInfo, widgetName2 & " Search Button Help", , , , True
            Case "icaoknobgreen"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "The selector utility is now searching by ICAO code only.", _
                     TTIconInfo, widgetName2 & " ICAO Button Help", , , , True
            Case "radiobody"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, " " _
                    & "This is the Airport Selector utility. Use this tool to select a nearby airport with a valid ICAO code. Any airport of this type will be able to supply " _
                    & "valid and up-to-date METAR weather data. Start by selecting the ICAO or Location option buttons below. This will allow you to search the supplied airports data " _
                    & "file for either an ICAO code or search for the airport by name, for example: HEATHROW." & vbCrLf & vbCrLf _
                    & "When you have a matching ICAO code, press OK to select and the main gauges will start to display weather information derived from that feed." _
                    & vbCrLf & "" _
                    & vbCrLf & "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize." _
                    , TTIconInfo, widgetName2 & " Airport Selection Information.", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(sender.Widget.key)
            Case "okbutton"
               Call okbutton_mouseDown(sender)
            Case "exitbutton"
               Call exitbutton_mouseDown(sender)
            Case "searchbutton"
               Call searchbutton_mouseDown(sender)
            Case "locationknobred"
               Call locationknobred_mouseDown(sender)
            Case "icaoknobred"
               Call icaoknobred_mouseDown(sender)
            'Case "icaoknobgreen"
               '  Call icaoknobgreen_mouseDown(sender)
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
' Procedure : okbutton_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub okbutton_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    On Error GoTo okbutton_mouseDown_Error
    
    'sender.Widget.Refresh ' this removes the delay in the lockButton depressing
    
    fileToPlay = "shutdown.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
   On Error GoTo 0
   Exit Sub

okbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure okbutton_mouseDown of Class Module cfSelector"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : exitbutton_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub exitbutton_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo exitbutton_mouseDown_Error
    
    SelectorForm.Hide
    
    fileToPlay = "shutdown.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

exitbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure exitbutton_mouseDown of Class Module cfSelector"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : searchbutton_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub searchbutton_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo searchbutton_mouseDown_Error
    
    'SelectorForm.Hide
    
    fileToPlay = "click.wav"
    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

searchbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure searchbutton_mouseDown of Class Module cfSelector"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : locationknobred_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub locationknobred_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo locationknobred_mouseDown_Error
   
    ' switch location red knob to nothing
    With SelectorForm.Widgets("locationknobred").Widget
        .Visible = False
        .Refresh
    End With
    
   ' switch location knob to green
    With SelectorForm.Widgets("locationknobgreen").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set icao green knob to invisible
    With SelectorForm.Widgets("icaoknobgreen").Widget
        .Visible = False
        .Refresh
    End With
    
    ' set icao knob to red to show icao is disabled
    With SelectorForm.Widgets("icaoknobred").Widget
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

locationknobred_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure locationknobred_mouseDown of Class Module cfSelector"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : icaoknobred_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub icaoknobred_mouseDown(ByRef sender As Object)
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo icaoknobred_mouseDown_Error
   
    ' switch icao red knob to nothing
    With SelectorForm.Widgets("icaoknobred").Widget
        .Visible = False
        .Refresh
    End With
    
    ' switch icao knob to green
    With SelectorForm.Widgets("icaoknobgreen").Widget
        .Visible = True
        .Refresh
    End With
    
    ' set location green knob to invisible
    With SelectorForm.Widgets("locationknobgreen").Widget
        .Visible = False
        .Refresh
    End With
    
    ' set location knob to red to show location is disabled
    With SelectorForm.Widgets("locationknobred").Widget
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
       
    On Error GoTo 0
   Exit Sub

icaoknobred_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure icaoknobred_mouseDown of Class Module cfSelector"

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
        menuForm.mnuCloseSelector.Visible = True
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
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
