VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfClipB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfClipB
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 05/12/2023
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents clipBForm As cWidgetForm
Attribute clipBForm.VB_VarHelpID = -1

Public collClipBPSDNonUIElements As cCollection, ClipBFX As Long, ClipBFY As Long, ClipBFZ As Single  'available on the outside, even after the Form was closed

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
    
    Set collClipBPSDNonUIElements = New_c.Collection(False)
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfClipB"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitClipBFromPSD
' Author    : olaf schmidt & Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :    loads up the Cairo imageList with keyed entries and a surface from each and every PSD layer
'                adds path X,Y and alpha to the excluded collection
'                creates widgets with corresponding keynames, locations and sizes as per PSD for each layer
'                cwOverlayTemp class puts it on the screen
'---------------------------------------------------------------------------------------
'
Public Sub InitClipBFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim i As Long: i = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitClipBFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set clipBForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName3, True, 1, 1)
        clipBForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        PSDWidth = .Width
        PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
                    
                    If collClipBPSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlay) to the "excluded paths" previously defined in Sub Main,
                        
                        collClipBPSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), val(PzGOpacity) / 100)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                    
                        Set W = clipBForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
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
    
    Set overlayClipbWidget = clipBForm.Widgets.Add(New cwOverlayClipb, "ClipB", 0, 0, PSDWidth, PSDHeight)
    Set overlayClipbWidget.cCollOverlayClipBImages = collClipBPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget

'    Set sTxtICAOInput = fClipB.ClipBForm.Widgets.Add(New_W("cwTextBox"), "sTxtICAOInput", 195, 20, 118, 30)
'    sTxtICAOInput.MultiLine = False
'    sTxtICAOInput.Widget.FontSize = 15
'    sTxtICAOInput.Widget.FontName = "times new roman"
'    sTxtICAOInput.Widget.Alpha = 0.5

' Cairo.ImageList.AddIconFromResourceFile "btnIco", "shell32", 167 'load some resource-Icon

'    Set sCmbMatchingLocations = fClipB.ClipBForm.Widgets.Add(New_W("cwDropDownList"), "sCmbMatchingLocations", 110, 58, 235, 28)
'    sCmbMatchingLocations.Widget.FontSize = 15
'    sCmbMatchingLocations.Widget.FontName = "times new roman"
'    sCmbMatchingLocations.Widget.Alpha = 0.5
'    sCmbMatchingLocations.DropDown.Caption = "ICAO"
'    sCmbMatchingLocations.Widget.Visible = False ' dropdown is not visible


   On Error GoTo 0
   Exit Sub

InitClipBFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitClipBFromPSD of Class Module cfClipB"
End Sub
' ----------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : ClipBForm_BubblingEvent
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub ClipBForm_BubblingEvent(ByRef sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   On Error GoTo ClipBForm_BubblingEvent_Error

    If PzGIgnoreMouse = "1" Then Exit Sub
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(sender.Widget.key)
            Case "clock"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip clipBForm.hwnd, "Click me to select the currently displayed ICAO location.", _
                     TTIconInfo, widgetName3 & " Select Button Help", , , , True
            Case "clipboard"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip clipBForm.hwnd, " " _
                    & "This is the Weather ClipB utility.  " _
                    & " " _
                    & "" & vbCrLf & vbCrLf _
                    & "" _
                    & vbCrLf & "" _
                    & vbCrLf & "(Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize)" _
                    , TTIconInfo, widgetName3 & " Airport Selection Information.", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(sender.Widget.key)
'            Case "sbtnselect"
'               Call sbtnselect_mouseDown(sender)
'            Case "sbtnexit"
'               Call sbtnexit_mouseDown(sender)
'            Case "sbtnsearch"
'               Call sbtnsearch_mouseDown(sender)
'            Case "optlocationred"
'               Call optlocationred_mouseDown(sender)
'            Case "opticaored"
'               Call opticaored_mouseDown(sender)
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

ClipBForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClipBForm_BubblingEvent of Class Module cfClipB"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ClipBForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
   On Error GoTo W_MouseDown_Error

    If Button = vbRightButton Then
        gblOriginatingForm = "ClipBForm"
        
        menuForm.mnuCloseSelector.Visible = False
        menuForm.mnuCloseWidget.Visible = True
        Call menuForm.PopupMenu(menuForm.mnuSelectorMenu)
    Else
    
        ' ensure that the associated forms pop to the front when the main gauge receives focus
        If fTemperature.temperatureGaugeForm.Visible = True Then fTemperature.temperatureGaugeForm.SetFocus
        If panzerPrefs.Visible = True Then panzerPrefs.SetFocus
        If fClipB.clipBForm.Visible = True Then fClipB.clipBForm.SetFocus
        
        If PzGIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cfClipB"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : ClipBForm_MouseMove
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ClipBForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = X: y0 = Y 'just store the offsets when no button is down
  
   On Error GoTo ClipBForm_MouseMove_Error
    
    If overlayTemperatureWidget.Locked = True Then Exit Sub
    If PzGIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not clipBForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      ClipBFX = (X - x0) * clipBForm.WidgetRoot.Zoom + clipBForm.Left
      ClipBFY = (Y - y0) * clipBForm.WidgetRoot.Zoom + clipBForm.Top
      clipBForm.Move ClipBFX, ClipBFY
    End If


   On Error GoTo 0
   Exit Sub

ClipBForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClipBForm_MouseMove of Class Module cfClipB"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClipBForm_MouseWheel
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub ClipBForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo ClipBForm_MouseWheel_Error

    If PzGScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   If MouseKeys = 8 Then Call ClipBAdjustZoom(ClipBFZ + RotationDirection / 2400)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

ClipBForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClipBForm_MouseWheel of Class Module cfTemperature"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : ClipBAdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ClipBAdjustZoom(ByVal NewZoom As Single)
   On Error GoTo ClipBAdjustZoom_Error

      ClipBFZ = NewZoom
      If ClipBFZ < 0.5 Then ClipBFZ = 0.5 Else If ClipBFZ > 3 Then ClipBFZ = 3 'keep ClipBFZ within a sane ZoomInterval
      clipBForm.WidgetRoot.Zoom = ClipBFZ '* ClipBForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
      clipBForm.Move clipBForm.Left, clipBForm.Top, PSDWidth * clipBForm.WidgetRoot.Zoom, PSDHeight * clipBForm.WidgetRoot.Zoom
    
      ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
      'PzGTemperatureGaugeSize = LTrim$(Str$(NewZoom * 100))
      'If panzerPrefs.IsVisible = True Then panzerPrefs.sliTemperatureGaugeSize.Value = val(PzGTemperatureGaugeSize)
    
   On Error GoTo 0
   Exit Sub

ClipBAdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClipBAdjustZoom of Class Module cfTemperature"
End Sub
