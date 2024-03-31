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

Public collSelectorPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

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
                        ' note: the clickable layers characteristics are set in adjustMainControls
                        
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

   On Error GoTo 0
   Exit Sub

InitSelectorFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitSelectorFromPSD of Class Module cfSelector"
End Sub
'
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
            Case "housing/lockbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the lock button, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName2 & " Lock Button Help", , , , True
            Case "housing/prefsbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click me to open the preferences utility. All advanced configuration of the widget is performed using the preference utility.", _
                     TTIconInfo, widgetName2 & " Prefs Button Help", , , , True
            Case "housing/tickbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to toggle the pointer movement from a jerky pointer to a more fluid, continuous movement.", _
                     TTIconInfo, widgetName2 & " Tick Button Help", , , , True
            Case "housing/switchfacesbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "Click here to select between different faces showing alternative Selector scales, celsius/fahrenheit or kelvin.", _
                     TTIconInfo, widgetName2 & " Switch Faces Button Help", , , , True
            Case "housing/stopbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "This button stops all animated gauge functions for this widget, pointer movement, digital displays &c .", _
                     TTIconInfo, widgetName2 & " Stop Button Help", , , , True
            Case "housing/startbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "This button restarts all stopped animated gauge functions for this widget, pointer movement &c .", _
                     TTIconInfo, widgetName2 & " Start Button Help", , , , True
            Case "housing/helpbutton"
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget.", _
                     TTIconInfo, widgetName2 & " Help Button Help", , , , True
            Case Else '& overlaySelectorWidget.SelectorDetails _
                If PzGEnableBalloonTooltips = "1" Then CreateToolTip SelectorForm.hwnd, " " _
                    & vbCrLf & "" _
                    & "Double tap on me to get new weather" _
                    & vbCrLf & "" _
                    & vbCrLf & "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize." _
                    , TTIconInfo, WidgetName2 & " METAR summary Information.", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/lockbutton"
               '  Call lockbutton_mouseDown(sender)
            Case "housing/prefsbutton"
               '  Call prefsbutton_mouseDown(sender)
            Case "housing/tickbutton"
               '  Call tickbutton_mouseDown(sender)
            Case "housing/switchfacesbutton"
                '  Call switchfacesbutton_mouseDown(sender)
            Case "housing/stopbutton"
               '  Call stopbutton_mouseDown(sender)
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/prefsbutton"
                sender.Widget.Alpha = val(PzGOpacity) / 100
                sender.Widget.Refresh
            Case "housing/switchfacesbutton"
               '  Call switchfacesbutton_mouseUp(sender)
            Case "housing/stopbutton"
               '  Call stopbutton_mouseUp(sender)
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
    
    If EventName = "W_Click" Then
        Select Case LCase$(sender.Widget.key)
            Case "housing/helpbutton"
                '  Call helpbutton_click
            Case "housing/startbutton"
                '  Call startbutton_click
            Case Else
                Debug.Print sender.Widget.key
        End Select
    End If
   
   On Error GoTo 0
   Exit Sub

SelectorForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SelectorForm_BubblingEvent of Class Module cfSelector"
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
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If PzGIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwOverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SelectorForm_MouseMove of Class Module cfTemperature"
End Sub
