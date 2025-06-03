VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfTemperature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfTemperature
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents temperatureGaugeForm As cWidgetForm
Attribute temperatureGaugeForm.VB_VarHelpID = -1

Public collTemperaturePSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long
Private PSDHeight As Long
'Private mGlowing As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   On Error GoTo Class_Initialize_Error

    'mGlowing = False ' when the alpha for the overlays arrives this might be usable
    Set collTemperaturePSDNonUIElements = New_c.Collection(False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfTemperature"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitFromPSD
' Author    : olaf schmidt & Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :    loads up the Cairo imageList with keyed entries and a surface from each and every PSD layer
'                adds path X,Y and alpha to the excluded collection
'                creates widgets with corresponding keynames, locations and sizes as per PSD for each layer
'                cwOverlayTemp class puts it on the screen
'---------------------------------------------------------------------------------------
'
Public Sub InitTemperatureFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim i As Long: i = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitTemperatureFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set temperatureGaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName1, True, 1, 1)
        temperatureGaugeForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        PSDWidth = .Width
        PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
                    
                    If collTemperaturePSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlayTemp) to the "excluded paths" previously defined in Sub Main,
                        
                        collTemperaturePSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), Val(gblOpacity) / 100)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                        Set W = temperatureGaugeForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
                            W.ImageKey = W.key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget
                        If gblOpacity < 1 Then
                            W.Alpha = Val(gblOpacity) / 100
                        Else
                            W.Alpha = 1
                        End If
                        
                        W.Tag = 0.05
                        
                        ' set the default characteristics of the non-interactive areas
                        ' note: the clickable layers characteristics are set in adjustTempMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gblEnableTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL

                    End If
                End If
            End If
        Next i
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
    
    Set overlayTemperatureWidget = temperatureGaugeForm.Widgets.Add(New cwOverlayTemp, "Overlay", 0, 0, PSDWidth, PSDHeight)
    Set overlayTemperatureWidget.cCollOverlayTempImages = collTemperaturePSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayTemperatureWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlayTemp, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayTemperatureWidget.cyPerc = 0.4727272

   On Error GoTo 0
   Exit Sub

InitTemperatureFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitTemperatureFromPSD of Class Module cfTemperature"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : temperatureGaugeForm_BubblingEvent
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub temperatureGaugeForm_BubblingEvent(ByRef sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   On Error GoTo temperatureGaugeForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/lockbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the lock button, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName1 & " Lock Button Help", , , , True
            Case "housing/prefsbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, "Click me to open the preferences utility. All advanced configuration of the widget is performed using the preference utility.", _
                     TTIconInfo, widgetName1 & " Prefs Button Help", , , , True
            Case "housing/tickbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, "This button is non-functional in this gauge.", _
                     TTIconInfo, widgetName1 & " Tick Button Help", , , , True
            Case "housing/switchfacesbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, "Click here to select between different faces showing alternative temperature scales, celsius/fahrenheit or kelvin.", _
                     TTIconInfo, widgetName1 & " Switch Faces Button Help", , , , True
            Case "housing/stopbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, "This button stops all animated gauge functions for this widget, pointer movement, digital displays &c .", _
                     TTIconInfo, widgetName1 & " Stop Button Help", , , , True
            Case "housing/startbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, "This button restarts all stopped animated gauge functions for this widget, pointer movement &c .", _
                     TTIconInfo, widgetName1 & " Start Button Help", , , , True
            Case "housing/helpbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget.", _
                     TTIconInfo, widgetName1 & " Help Button Help", , , , True
            Case Else
                If gblEnableBalloonTooltips = "1" Then CreateToolTip temperatureGaugeForm.hwnd, " " & WeatherMeteo.WeatherDetails _
                    & vbCrLf & "" _
                    & vbCrLf & "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize." _
                    , TTIconInfo, widgetName1 & " METAR summary Information.", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/lockbutton"
                Call lockbutton_mouseDown(sender)
            Case "housing/prefsbutton"
                Call prefsbutton_mouseDown(sender)
'            Case "housing/tickbutton"
'                Call tickbutton_mouseDown(sender)
            Case "housing/switchfacesbutton"
                 Call switchfacesbutton_mouseDown(sender)
            Case "housing/stopbutton"
                Call stopbutton_mouseDown(sender)
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/prefsbutton"
                sender.Widget.Alpha = Val(gblOpacity) / 100
                sender.Widget.Refresh
            Case "housing/switchfacesbutton"
                Call switchfacesbutton_mouseUp(sender)
            Case "housing/stopbutton"
                Call stopbutton_mouseUp(sender)
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
    
    If EventName = "W_Click" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/helpbutton"
                 Call helpbutton_click
            Case "housing/startbutton"
                 Call startbutton_click
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
   
   On Error GoTo 0
   Exit Sub

temperatureGaugeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure temperatureGaugeForm_BubblingEvent of Class Module cfTemperature"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : helpbutton_click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helpbutton_click()
   On Error GoTo helpbutton_click_Error

   Call helpSplash

   On Error GoTo 0
   Exit Sub

helpbutton_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_click of Class Module cfTemperature"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : startbutton_click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub startbutton_click()
    On Error GoTo startbutton_click_Error
   
    
    Dim fileToPlay As String: fileToPlay = vbNullString

    Call TurnFunctionsOn

    fileToPlay = "ting.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

startbutton_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_click of Class Module cfTemperature"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseUp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseUp(ByRef sender As Object)
    
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo stopbutton_mouseUp_Error

    sender.Widget.Alpha = Val(gblOpacity) / 100
    sender.Widget.Refresh
           
    Call SwitchOff
    
    fileToPlay = "stop.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Sub

stopbutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseUp of Class Module cfTemperature"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : stopbutton_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopbutton_mouseDown(ByRef sender As Object)
   
    On Error GoTo stopbutton_mouseDown_Error

    sender.Widget.Alpha = 0
    sender.Widget.Refresh
           
   On Error GoTo 0
   Exit Sub

stopbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseDown of Class Module cfTemperature"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : switchfacesbutton_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub switchfacesbutton_mouseDown(ByRef sender As Object)
    On Error GoTo switchfacesbutton_mouseDown_Error

    sender.Widget.Alpha = 0
    sender.Widget.Refresh
    
   On Error GoTo 0
   Exit Sub

switchfacesbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure switchfacesbutton_mouseDown of Class Module cfTemperature"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : switchfacesbutton_mouseUp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub switchfacesbutton_mouseUp(ByRef sender As Object)
    
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    Dim currFace As Integer: currFace = 0
    
    On Error GoTo switchfacesbutton_mouseUp_Error

    sender.Widget.Alpha = Val(gblOpacity) / 100
    sender.Widget.Refresh
        
    fileToPlay = "clunk.wav"
    
    currFace = Val(gblTemperatureScale)
    currFace = currFace + 1
    If currFace > 2 Then currFace = 0
    gblTemperatureScale = LTrim$(Str$(currFace))
    
    overlayTemperatureWidget.thisFace = Val(gblTemperatureScale)
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
   On Error GoTo 0
   Exit Sub

switchfacesbutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure switchfacesbutton_mouseUp of Class Module cfTemperature"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : lockbutton_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockbutton_mouseDown(ByRef sender As Object)
   On Error GoTo lockbutton_mouseDown_Error
    
    If overlayTemperatureWidget.Locked = False Then
        sender.Widget.Alpha = 0
    Else
        sender.Widget.Alpha = Val(gblOpacity) / 100
    End If
    sender.Widget.Refresh ' this removes the delay in the lockButton depressing
    If gblPreventDraggingTemperature = "1" Then
        overlayTemperatureWidget.Locked = False
    Else
        overlayTemperatureWidget.Locked = True
    End If
    
   On Error GoTo 0
   Exit Sub

lockbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_mouseDown of Class Module cfTemperature"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : prefsbutton_mouseDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub prefsbutton_mouseDown(ByRef sender As Object)
    
    Dim fileToPlay As String: fileToPlay = vbNullString
   
    On Error GoTo prefsbutton_mouseDown_Error

    sender.Widget.Alpha = 0
    'Sender.Widget.Visible = False
    sender.Widget.Refresh ' this removes the delay in the Button depressing
    fileToPlay = "winding.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
   'MsgBox "prefsbutton_mouseDown"
   Call makeProgramPreferencesAvailable

   On Error GoTo 0
   Exit Sub

prefsbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseDown of Class Module cfTemperature"

End Sub
''---------------------------------------------------------------------------------------
'' Procedure : tickbutton_mouseDown
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 03/08/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub tickbutton_mouseDown(ByRef sender As Object)
'
'    Dim fileToPlay As String: fileToPlay = vbNullString
'
'    On Error GoTo tickbutton_mouseDown_Error
'
'    If overlayTemperatureWidget.pointerAnimate = True Then
'        overlayTemperatureWidget.pointerAnimate = False
'        sender.Widget.Alpha = Val(gblOpacity) / 100
'        gblPointerAnimate = "0"
'    Else
'        overlayTemperatureWidget.pointerAnimate = True
'        sender.Widget.Alpha = 0 ' sender is the referring control
'        gblPointerAnimate = "1"
'
'    End If
'    sender.Widget.Refresh ' this removes the delay in the Button depressing
'
'    fileToPlay = "lock.wav"
'    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
'        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
'    End If
'
'    sPutINISetting "Software\PzTemperatureGauge", "pointerAnimate", gblPointerAnimate, gblSettingsFile
'    If panzerPrefs.IsVisible = True Then panzerPrefs.cmbTickSwitchPref.ListIndex = Val(gblPointerAnimate)
'
'   On Error GoTo 0
'   Exit Sub
'
'tickbutton_mouseDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tickbutton_mouseDown of Class Module cfTemperature"
'
'End Sub


    


'---------------------------------------------------------------------------------------
' Procedure : temperatureGaugeForm_DblClick
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub temperatureGaugeForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo temperatureGaugeForm_DblClick_Error

    '    If gblIgnoreMouse = "1" Then Exit Sub

    If LTrim$(gblDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If SHIFT_1 = True Then
        SHIFT_1 = False
        Call ShellExecute(fTemperature.temperatureGaugeForm.hwnd, "Open", gblOpenFile, vbNullString, App.path, 1)
    Else
        Call ShellExecute(fTemperature.temperatureGaugeForm.hwnd, "runas", thisCommand, vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub

temperatureGaugeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure temperatureGaugeForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : temperatureGaugeForm_KeyDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub temperatureGaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo temperatureGaugeForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

temperatureGaugeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure temperatureGaugeForm_KeyDown of Class Module cfMain"
End Sub
''
'---------------------------------------------------------------------------------------
' Procedure : temperatureGaugeForm_MouseUp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the gauge/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub temperatureGaugeForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    On Error GoTo temperatureGaugeForm_MouseUp_Error
    
    Call saveTemperatureGaugePosition

    On Error GoTo 0
    Exit Sub

temperatureGaugeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure temperatureGaugeForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : temperatureGaugeForm_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub temperatureGaugeForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = X: y0 = Y 'just store the offsets when no button is down
  
   On Error GoTo temperatureGaugeForm_MouseMove_Error
    
    If overlayTemperatureWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not temperatureGaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (X - x0) * temperatureGaugeForm.WidgetRoot.Zoom + temperatureGaugeForm.Left
      FY = (Y - y0) * temperatureGaugeForm.WidgetRoot.Zoom + temperatureGaugeForm.Top
      temperatureGaugeForm.Move FX, FY
    End If


   On Error GoTo 0
   Exit Sub

temperatureGaugeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure temperatureGaugeForm_MouseMove of Class Module cfTemperature"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : temperatureGaugeForm_MouseWheel
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub temperatureGaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo temperatureGaugeForm_MouseWheel_Error

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   If MouseKeys = 8 Then Call tempAdjustZoom(FZ + RotationDirection / 2400)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

temperatureGaugeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure temperatureGaugeForm_MouseWheel of Class Module cfTemperature"
End Sub

''
''---------------------------------------------------------------------------------------
'' Procedure : temperatureGaugeForm_ResizeWithDimensionsDIP
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 31/07/2023
'' Purpose   : put any other resizing action in here - unused
''---------------------------------------------------------------------------------------
''
'Private Sub temperatureGaugeForm_ResizeWithDimensionsDIP(ByVal dx As Single, ByVal dy As Single)
'   On Error GoTo temperatureGaugeForm_ResizeWithDimensionsDIP_Error
'
'  'If Not CloseWidget Is Nothing Then CloseWidget.Move dx - 32, 0, 32, 32 'place the close-widget at the top-right of the form
'
'   On Error GoTo 0
'   Exit Sub
'
'temperatureGaugeForm_ResizeWithDimensionsDIP_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure temperatureGaugeForm_ResizeWithDimensionsDIP of Class Module cfTemperature"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : tempAdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub tempAdjustZoom(ByVal NewZoom As Single)
   On Error GoTo tempAdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    temperatureGaugeForm.WidgetRoot.Zoom = FZ '* temperatureGaugeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    temperatureGaugeForm.Move temperatureGaugeForm.Left, temperatureGaugeForm.Top, PSDWidth * temperatureGaugeForm.WidgetRoot.Zoom, PSDHeight * temperatureGaugeForm.WidgetRoot.Zoom
  
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblTemperatureGaugeSize = LTrim$(Str$(NewZoom * 100))
    If panzerPrefs.IsVisible = True Then
        
'        If panzerPrefs.cmbGaugeType.ListIndex = 0 Then
            panzerPrefs.sliGaugeSize.Value = Val(gblTemperatureGaugeSize)
            sPutINISetting "Software\PzTemperatureGauge", "temperatureGaugeSize", gblTemperatureGaugeSize, gblSettingsFile
'        End If
'        If panzerPrefs.cmbGaugeType.ListIndex = 1 Then
'            panzerPrefs.sliGaugeSize.Value = Val(gblAnemometerGaugeSize)
'            sPutINISetting "Software\PzAnemometerGauge", "anemometerGaugeSize", gblAnemometerGaugeSize, gblSettingsFile
'        End If
    End If

    
   On Error GoTo 0
   Exit Sub

tempAdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tempAdjustZoom of Class Module cfTemperature"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub temperatureGaugeForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo W_MouseDown_Error

    gblOriginatingForm = "temperatureForm"
    If Button = vbRightButton Then
        menuForm.mnuLockTemperatureGauge.Visible = True
        menuForm.mnuLockHumidityGauge.Visible = False
        menuForm.mnuLockAnemometerGauge.Visible = False
        menuForm.mnuLockBarometerGauge.Visible = False
        
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        ' ensure that the associated forms pop to the front when the main gauge receives focus
        If fSelector.SelectorForm.Visible = True Then fSelector.SelectorForm.SetFocus
        If panzerPrefs.IsVisible = True Then panzerPrefs.SetFocus
        If fTemperature.temperatureGaugeForm.Visible = True Then fTemperature.temperatureGaugeForm.SetFocus
        
        If gblIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwOverlayTemp"
End Sub


