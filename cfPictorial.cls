VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfPictorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfPictorial
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents pictorialGaugeForm As cWidgetForm
Attribute pictorialGaugeForm.VB_VarHelpID = -1

Public collPictorialPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

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
    Set collPictorialPSDNonUIElements = New_c.Collection(False)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfPictorial"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitFromPSD
' Author    : olaf schmidt & Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :    loads up the Cairo imageList with keyed entries and a surface from each and every PSD layer
'                adds path X,Y and alpha to the excluded collection
'                creates widgets with corresponding keynames, locations and sizes as per PSD for each layer
'                cwOverlayPict class puts it on the screen
'---------------------------------------------------------------------------------------
'
Public Sub InitPictorialFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim i As Long: i = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitPictorialFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set pictorialGaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName7, True, 1, 1)
        pictorialGaugeForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        PSDWidth = .Width
        PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
                    
                    If collPictorialPSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlayPict) to the "excluded paths" previously defined in Sub Main,
                        
                        collPictorialPSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), Val(gblOpacity) / 100)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                        Set W = pictorialGaugeForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
                            W.ImageKey = W.key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget
                        If gblOpacity < 1 Then
                            W.Alpha = Val(gblOpacity) / 100
                        Else
                            W.Alpha = 1
                        End If
                        
                        W.Tag = 0.05
                        
                        ' set the default characteristics of the non-interactive areas
                        ' note: the clickable layers characteristics are set in adjustPictorialMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gblEnableTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL

                    End If
                End If
            End If
        Next i
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
    
    Set overlayPictorialWidget = pictorialGaugeForm.Widgets.Add(New cwOverlayPict, "Overlay", 0, 0, PSDWidth, PSDHeight)
    Set overlayPictorialWidget.cCollOverlayPictImages = collPictorialPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayPictorialWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlayPict, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayPictorialWidget.cyPerc = 0.4727272

   On Error GoTo 0
   Exit Sub

InitPictorialFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitPictorialFromPSD of Class Module cfPictorial"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : pictorialGaugeForm_BubblingEvent
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub pictorialGaugeForm_BubblingEvent(ByRef sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   On Error GoTo pictorialGaugeForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/lockbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the lock button, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName7 & " Lock Button Help", , , , True
            Case "housing/prefsbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, "Click me to open the preferences utility. All advanced configuration of the widget is performed using the preference utility.", _
                     TTIconInfo, widgetName7 & " Prefs Button Help", , , , True
            Case "housing/tickbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, "This button is non-functional in this gauge.", _
                     TTIconInfo, widgetName7 & " Tick Button Help", , , , True
            Case "housing/switchfacesbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, "Click here to select between different faces showing alternative pictorial scales, celsius/fahrenheit or kelvin.", _
                     TTIconInfo, widgetName7 & " Switch Faces Button Help", , , , True
            Case "housing/stopbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, "This button stops all animated gauge functions for this widget, pointer movement, digital displays &c .", _
                     TTIconInfo, widgetName7 & " Stop Button Help", , , , True
            Case "housing/startbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, "This button restarts all stopped animated gauge functions for this widget, pointer movement &c .", _
                     TTIconInfo, widgetName7 & " Start Button Help", , , , True
            Case "housing/helpbutton"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget.", _
                     TTIconInfo, widgetName7 & " Help Button Help", , , , True
            Case Else
                If gblEnableBalloonTooltips = "1" Then CreateToolTip pictorialGaugeForm.hwnd, " " & WeatherMeteo.WeatherDetails _
                    & vbCrLf & "" _
                    & vbCrLf & "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize." _
                    , TTIconInfo, widgetName7 & " METAR summary Information.", , , , True
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

pictorialGaugeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictorialGaugeForm_BubblingEvent of Class Module cfPictorial"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_click of Class Module cfPictorial"
    
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startbutton_click of Class Module cfPictorial"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseUp of Class Module cfPictorial"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopbutton_mouseDown of Class Module cfPictorial"

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
    
    If overlayPictorialWidget.Locked = False Then
        sender.Widget.Alpha = 0
    Else
        sender.Widget.Alpha = Val(gblOpacity) / 100
    End If
    sender.Widget.Refresh ' this removes the delay in the lockButton depressing
    If gblPreventDraggingPictorial = "1" Then
        overlayPictorialWidget.Locked = False
    Else
        overlayPictorialWidget.Locked = True
    End If
    
   On Error GoTo 0
   Exit Sub

lockbutton_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockbutton_mouseDown of Class Module cfPictorial"

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure prefsbutton_mouseDown of Class Module cfPictorial"

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
'    If overlayPictorialWidget.pointerAnimate = True Then
'        overlayPictorialWidget.pointerAnimate = False
'        sender.Widget.Alpha = Val(gblOpacity) / 100
'        gblPointerAnimate = "0"
'    Else
'        overlayPictorialWidget.pointerAnimate = True
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
'    sPutINISetting "Software\PzPictorialGauge", "pointerAnimate", gblPointerAnimate, gblSettingsFile
'    If panzerPrefs.IsVisible = True Then panzerPrefs.cmbTickSwitchPref.ListIndex = Val(gblPointerAnimate)
'
'   On Error GoTo 0
'   Exit Sub
'
'tickbutton_mouseDown_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tickbutton_mouseDown of Class Module cfPictorial"
'
'End Sub


    


'---------------------------------------------------------------------------------------
' Procedure : pictorialGaugeForm_DblClick
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pictorialGaugeForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo pictorialGaugeForm_DblClick_Error

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
        Call ShellExecute(fPictorial.pictorialGaugeForm.hwnd, "Open", gblOpenFile, vbNullString, App.path, 1)
    Else
        Call ShellExecute(fPictorial.pictorialGaugeForm.hwnd, "runas", thisCommand, vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub

pictorialGaugeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictorialGaugeForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : pictorialGaugeForm_KeyDown
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub pictorialGaugeForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo pictorialGaugeForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

pictorialGaugeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictorialGaugeForm_KeyDown of Class Module cfMain"
End Sub
''
'---------------------------------------------------------------------------------------
' Procedure : pictorialGaugeForm_MouseUp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the gauge/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub pictorialGaugeForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
    On Error GoTo pictorialGaugeForm_MouseUp_Error
    
    Call savePictorialGaugePosition

    On Error GoTo 0
    Exit Sub

pictorialGaugeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictorialGaugeForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : pictorialGaugeForm_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pictorialGaugeForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = X: y0 = Y 'just store the offsets when no button is down
  
   On Error GoTo pictorialGaugeForm_MouseMove_Error
    
    If overlayPictorialWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub
    If gblMenuOccurred = True Then
        gblMenuOccurred = False
        Exit Sub
    End If
    If Button = vbLeftButton And Not pictorialGaugeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (X - x0) * pictorialGaugeForm.WidgetRoot.Zoom + pictorialGaugeForm.Left
      FY = (Y - y0) * pictorialGaugeForm.WidgetRoot.Zoom + pictorialGaugeForm.Top
      pictorialGaugeForm.Move FX, FY
    End If


   On Error GoTo 0
   Exit Sub

pictorialGaugeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictorialGaugeForm_MouseMove of Class Module cfPictorial"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : pictorialGaugeForm_MouseWheel
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : rotate forward positive = smalleroverlayPictorialWidgetovrl
'---------------------------------------------------------------------------------------
'
Private Sub pictorialGaugeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo pictorialGaugeForm_MouseWheel_Error

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   If MouseKeys = 8 Then Call pictAdjustZoom(FZ + RotationDirection / 2400)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

pictorialGaugeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictorialGaugeForm_MouseWheel of Class Module cfPictorial"
End Sub

''
''---------------------------------------------------------------------------------------
'' Procedure : pictorialGaugeForm_ResizeWithDimensionsDIP
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 31/07/2023
'' Purpose   : put any other resizing action in here - unused
''---------------------------------------------------------------------------------------
''
'Private Sub pictorialGaugeForm_ResizeWithDimensionsDIP(ByVal dx As Single, ByVal dy As Single)
'   On Error GoTo pictorialGaugeForm_ResizeWithDimensionsDIP_Error
'
'  'If Not CloseWidget Is Nothing Then CloseWidget.Move dx - 32, 0, 32, 32 'place the close-widget at the top-right of the form
'
'   On Error GoTo 0
'   Exit Sub
'
'pictorialGaugeForm_ResizeWithDimensionsDIP_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictorialGaugeForm_ResizeWithDimensionsDIP of Class Module cfPictorial"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : pictAdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub pictAdjustZoom(ByVal NewZoom As Single)
   On Error GoTo pictAdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    pictorialGaugeForm.WidgetRoot.Zoom = FZ '* pictorialGaugeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    pictorialGaugeForm.Move pictorialGaugeForm.Left, pictorialGaugeForm.Top, PSDWidth * pictorialGaugeForm.WidgetRoot.Zoom, PSDHeight * pictorialGaugeForm.WidgetRoot.Zoom
  
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblPictorialGaugeSize = LTrim$(Str$(NewZoom * 100))
    If panzerPrefs.IsVisible = True Then
        
'        If panzerPrefs.cmbGaugeType.ListIndex = 0 Then
            panzerPrefs.sliGaugeSize.Value = Val(gblPictorialGaugeSize)
            sPutINISetting "Software\PzPictorialGauge", "pictorialGaugeSize", gblPictorialGaugeSize, gblSettingsFile
'        End If
'        If panzerPrefs.cmbGaugeType.ListIndex = 1 Then
'            panzerPrefs.sliGaugeSize.Value = Val(gblAnemometerGaugeSize)
'            sPutINISetting "Software\PzAnemometerGauge", "anemometerGaugeSize", gblAnemometerGaugeSize, gblSettingsFile
'        End If
    End If

    
   On Error GoTo 0
   Exit Sub

pictAdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure pictAdjustZoom of Class Module cfPictorial"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub pictorialGaugeForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo W_MouseDown_Error

    gblOriginatingForm = "pictorialForm"
    
    If Button = vbRightButton Then
        gblMenuOccurred = True
        menuForm.mnuLockPictorialGauge.Visible = True
        menuForm.mnuLockHumidityGauge.Visible = False
        menuForm.mnuLockAnemometerGauge.Visible = False
        menuForm.mnuLockBarometerGauge.Visible = False
        
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        ' ensure that the associated forms pop to the front when the main gauge receives focus
        If fSelector.SelectorForm.Visible = True Then fSelector.SelectorForm.SetFocus
        If panzerPrefs.IsVisible = True Then panzerPrefs.SetFocus
        If fPictorial.pictorialGaugeForm.Visible = True Then fPictorial.pictorialGaugeForm.SetFocus
        
        If gblIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwOverlayPict"
End Sub




