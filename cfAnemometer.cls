VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfAnemometer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfAnemometer
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents anemometerGaugeForm As cWidgetForm
Attribute anemometerGaugeForm.VB_VarHelpID = -1

Public collAnemometerPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

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
    Set collAnemometerPSDNonUIElements = New_c.Collection(False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfAnemometer"
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
Public Sub InitAnemometerFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim i As Long: i = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitAnemometerFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set anemometerGaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName4, True, 1, 1)
        anemometerGaugeForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        PSDWidth = .Width
        PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
                    
                    If collAnemometerPSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlayTemp) to the "excluded paths" previously defined in Sub Main,
                        
                        collAnemometerPSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), Val(gblOpacity) / 100)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                        Set W = anemometerGaugeForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
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
    
    Set overlayAnemoWidget = anemometerGaugeForm.Widgets.Add(New cwOverlayAnemo, "Overlay", 0, 0, PSDWidth, PSDHeight)
    Set overlayAnemoWidget.cCollOverlayAnemoImages = collAnemometerPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayAnemoWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlayTemp, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayAnemoWidget.cyPerc = 0.4727272

   On Error GoTo 0
   Exit Sub

InitAnemometerFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitAnemometerFromPSD of Class Module cfAnemometer"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : anemometerGaugeForm_BubblingEvent
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub anemometerGaugeForm_BubblingEvent(ByRef sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   On Error GoTo anemometerGaugeForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/lockbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the lock button, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName4 & " Lock Button Help", , , , True
            Case "housing/prefsbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, "Click me to open the preferences utility. All advanced configuration of the widget is performed using the preference utility.", _
                     TTIconInfo, widgetName4 & " Prefs Button Help", , , , True
            Case "housing/tickbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, "This button is non-functional in this gauge.", _
                     TTIconInfo, widgetName4 & " Tick Button Help", , , , True
            Case "housing/switchfacesbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, "Click here to select between different faces showing alternative scales, ie. knots or metres per second.", _
                     TTIconInfo, widgetName4 & " Switch Faces Button Help", , , , True
            Case "housing/stopbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, "This button stops all animated gauge functions for this widget, pointer movement, digital displays &c .", _
                     TTIconInfo, widgetName4 & " Stop Button Help", , , , True
            Case "housing/startbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, "This button restarts all stopped animated gauge functions for this widget, pointer movement &c .", _
                     TTIconInfo, widgetName4 & " Start Button Help", , , , True
            Case "housing/helpbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget.", _
                     TTIconInfo, widgetName4 & " Help Button Help", , , , True
            Case Else
                If gblEnableBalloonTooltips = "1" Then CreateToolTip anemometerGaugeForm.hwnd, " " & WeatherMeteo.WeatherDetails _
                    & vbCrLf & "" _
                    & vbCrLf & "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize." _
                    , TTIconInfo, widgetName4 & " METAR summary Information.", , , , True
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
            Case "housing/helpbutton"
                 Call helpbutton_mouseDown(sender)
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
'            Case "housing/helpbutton"
'                 Call helpbutton_click
            Case "housing/startbutton"
                 Call startbutton_click
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
   
   On Error GoTo 0
   Exit Sub

anemometerGaugeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anemometerGaugeForm_BubblingEvent of Class Module cfAnemometer"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : helpbutton_click
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helpbutton_mouseDown(ByRef sender As Object)
   On Error GoTo helpbutton_click_Error

   Call helpSplash

   On Error GoTo 0
   Exit Sub

helpbutton_click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_click of Class Module cfAnemometer"
    
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_click of Class Module cfAnemometer"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseUp of Class Module cfAnemometer"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseDown of Class Module cfAnemometer"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure switchfacesbutton_mouseDown of Class Module cfAnemometer"

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
    
    currFace = Val(gblWindSpeedScale)
    currFace = currFace + 1
    If currFace > 1 Then currFace = 0
    gblWindSpeedScale = LTrim$(Str$(currFace))
    
    overlayAnemoWidget.thisFace = Val(gblWindSpeedScale)
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
   On Error GoTo 0
   Exit Sub

switchfacesbutton_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure switchfacesbutton_mouseUp of Class Module cfAnemometer"

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
    
    If overlayAnemoWidget.Locked = False Then
        sender.Widget.Alpha = 0
    Else
        sender.Widget.Alpha = Val(gblOpacity) / 100
    End If
    sender.Widget.Refresh ' this removes the delay in the lockButton depressing
    If gblPreventDraggingAnemometer = "1" Then
        overlayAnemoWidget.Locked = False
    Else
        overlayAnemoWidget.Locked = True
    End If
    
   On Error GoTo 0
   Exit Sub

lockbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_mouseDown of Class Module cfAnemometer"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseDown of Class Module cfAnemometer"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : anemometerGaugeForm_DblClick
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub anemometerGaugeForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo anemometerGaugeForm_DblClick_Error

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
        Call ShellExecute(fAnemometer.anemometerGaugeForm.hwnd, "Open", gblOpenFile, vbNullString, App.path, 1)
    Else
        Call ShellExecute(fAnemometer.anemometerGaugeForm.hwnd, "runas", thisCommand, vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub

anemometerGaugeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anemometerGaugeForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : anemometerGaugeForm_KeyDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub anemometerGaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo anemometerGaugeForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

anemometerGaugeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anemometerGaugeForm_KeyDown of Class Module cfMain"
End Sub
''
'---------------------------------------------------------------------------------------
' Procedure : anemometerGaugeForm_MouseUp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the gauge/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub anemometerGaugeForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    On Error GoTo anemometerGaugeForm_MouseUp_Error
    
    Call saveAnemometerGaugePosition

    On Error GoTo 0
    Exit Sub

anemometerGaugeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anemometerGaugeForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : anemometerGaugeForm_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub anemometerGaugeForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = X: y0 = Y 'just store the offsets when no button is down
  
   On Error GoTo anemometerGaugeForm_MouseMove_Error
    
    If overlayAnemoWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub
    If gblMenuOccurred = True Then
        gblMenuOccurred = False
        Exit Sub
    End If
    
    If Button = vbLeftButton And Not anemometerGaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (X - x0) * anemometerGaugeForm.WidgetRoot.Zoom + anemometerGaugeForm.Left
      FY = (Y - y0) * anemometerGaugeForm.WidgetRoot.Zoom + anemometerGaugeForm.Top
      anemometerGaugeForm.Move FX, FY
    End If


   On Error GoTo 0
   Exit Sub

anemometerGaugeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anemometerGaugeForm_MouseMove of Class Module cfAnemometer"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : anemometerGaugeForm_MouseWheel
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub anemometerGaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo anemometerGaugeForm_MouseWheel_Error

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   If MouseKeys = 8 Then Call anemoAdjustZoom(FZ + RotationDirection / 2400)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

anemometerGaugeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anemometerGaugeForm_MouseWheel of Class Module cfAnemometer"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : anemoAdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub anemoAdjustZoom(ByVal NewZoom As Single)
   On Error GoTo anemoAdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    anemometerGaugeForm.WidgetRoot.Zoom = FZ '* anemometerGaugeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    anemometerGaugeForm.Move anemometerGaugeForm.Left, anemometerGaugeForm.Top, PSDWidth * anemometerGaugeForm.WidgetRoot.Zoom, PSDHeight * anemometerGaugeForm.WidgetRoot.Zoom
      
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblAnemometerGaugeSize = LTrim$(Str$(NewZoom * 100))
    If panzerPrefs.IsVisible = True Then

        panzerPrefs.sliGaugeSize.Value = Val(gblAnemometerGaugeSize)
        sPutINISetting "Software\PzAnemometerGauge", "anemometerGaugeSize", gblAnemometerGaugeSize, gblSettingsFile

    End If
    
   On Error GoTo 0
   Exit Sub

anemoAdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure anemoAdjustZoom of Class Module cfAnemometer"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub anemometerGaugeForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo W_MouseDown_Error

    gblOriginatingForm = "anemometerForm"
    If Button = vbRightButton Then
        gblMenuOccurred = True
        menuForm.mnuLockTemperatureGauge.Visible = False
        menuForm.mnuLockHumidityGauge.Visible = False
        menuForm.mnuLockAnemometerGauge.Visible = True
        menuForm.mnuLockBarometerGauge.Visible = False
        menuForm.mnuLockPictorialGauge.Visible = False

        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        ' ensure that the associated forms pop to the front when the main gauge receives focus
        If fSelector.SelectorForm.Visible = True Then fSelector.SelectorForm.SetFocus
        If panzerPrefs.IsVisible = True Then panzerPrefs.SetFocus
        If fAnemometer.anemometerGaugeForm.Visible = True Then fAnemometer.anemometerGaugeForm.SetFocus
        
        If gblIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwOverlayTemp"
End Sub


