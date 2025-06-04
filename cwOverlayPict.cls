VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlayPict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwoverlaypict
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfPictorial
'---------------------------------------------------------------------------------------

Option Explicit

'these Public Properties are set in the Parent (here, cfPictorial)
Public cCollOverlayPictImages As cCollection

Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1


Private mLocked As Boolean                      ' property Locked
Private mHidden As Boolean                      ' property Hidden
Private mThisOpacity As String                  ' property ThisOpacity



'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Error
    
    Set WeatherMeteo = New cwWeatherMeteo 'initialised declaration

    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
    
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwoverlaypict"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwoverlaypict"
End Property




'---------------------------------------------------------------------------------------
' Property : thisOpacity
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisOpacity() As String
   On Error GoTo thisOpacityGet_Error

    thisOpacity = mThisOpacity

   On Error GoTo 0
   Exit Property

thisOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisOpacity
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisOpacity(ByVal newValue As String)
    On Error GoTo thisOpacityLet_Error

    If mThisOpacity <> newValue Then mThisOpacity = newValue Else Exit Property
        
    On Error GoTo 0
    Exit Property

thisOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwoverlaypict"
End Property





'---------------------------------------------------------------------------------------
' Property : Locked
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayPict"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "lock.wav"

    If mLocked = True Then
        menuForm.mnuLockPictorialGauge.Checked = True
        If panzerPrefs.IsVisible = True Then
            panzerPrefs.cmbGaugeType.ListIndex = 0
            panzerPrefs.chkPreventDragging.Value = 1
        End If
        
        gblPreventDraggingPictorial = "1"
        If aspectRatio = "landscape" Then
            gblPictorialLandscapeLocked = "1"
            panzerPrefs.cmbLandscapeLocked.ListIndex = 1
        Else
            gblPictorialPortraitLocked = "1"
            panzerPrefs.cmbPortraitLocked.ListIndex = 1
       End If
        
        fPictorial.pictorialGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    Else
        menuForm.mnuLockPictorialGauge.Checked = False
        If panzerPrefs.IsVisible = True Then
            panzerPrefs.cmbGaugeType.ListIndex = 0
            panzerPrefs.chkPreventDragging.Value = 0
        End If
        
        gblPreventDraggingPictorial = "0"
        If aspectRatio = "landscape" Then
            gblPictorialLandscapeLocked = "0"
            panzerPrefs.cmbLandscapeLocked.ListIndex = 0
        Else
            gblPictorialPortraitLocked = "0"
            panzerPrefs.cmbPortraitLocked.ListIndex = 0 ' deanieboy
        End If
        
        fPictorial.pictorialGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    End If

    fPictorial.pictorialGaugeForm.Refresh

    sPutINISetting "Software\PzPictorialGauge", "preventDraggingPictorial", gblPreventDraggingPictorial, gblSettingsFile

    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayPict"
End Property



'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

   Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayPict"
End Property
'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        W.Visible = False
        
        W.Refresh
    Else
        W.Visible = True
        
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayPict"
End Property

'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Overlay-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal X As Single, ByVal Y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlayPict"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt & Dean Beedell (yereverluvinunclebert) & Dean Beedell (yereverluvinunclebert)
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim i As Long: i = 0
    Dim arr As Variant
    Dim key As String: key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtThisPictorial As String: txtThisPictorial = vbNullString
    Dim someOpacity As Double: someOpacity = 0
    Dim boxLeft As Integer: boxLeft = 0
    Dim cream As String: cream = vbNullString
    Dim scaleString As String: scaleString = vbNullString
       
    On Error GoTo W_Paint_Error
    
    cream = RGB(235, 217, 179) ' cannot create a constant from an RGB value
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 16
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    someOpacity = (Val(mThisOpacity) / 100) * 0.75
    
    'txtThisPictorial = CStr(Int(mThisPictorial))

    For i = 0 To cCollOverlayPictImages.Count - 1 ' this is the ExcludedPaths collection populated in cfPictorial by duplicating collPictorialPSDNonUIElements.
      arr = cCollOverlayPictImages.ItemByIndex(i)
      key = cCollOverlayPictImages.KeyByIndex(i)
      
      If IsArray(arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case key
                 
'                 Case "centigradeface"
'
'                    If mThisFace = 0 Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
'                        CC.Paint arr(2), Pat ' render the Pattern
'                    End If
'
'                 Case "fahrenheitface"
'
'                    If mThisFace = 1 Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
'                        CC.Paint arr(2), Pat 'render the Pattern
'                    End If
'
'                 Case "kelvinface"
'
'                    If mThisFace = 2 Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
'                        CC.Paint arr(2), Pat 'render the Pattern
'                    End If
'
'                Case "redlamptrue"
'
'                    If mThisPictorial >= 34 Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
'                        CC.Paint arr(2), Pat 'render the Pattern
'                    End If
'
'                 Case "redlampfalse"
'
'                    If mThisPictorial < 34 Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
'                        CC.Paint arr(2), Pat ' render the Pattern
'                    End If
'
'                 Case "bluelamptrue"
'
'                    If mThisPictorial <= 0 Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
'                        CC.Paint arr(2), Pat 'render the Pattern
'                    End If
'
'                 Case "bluelampfalse"
'
'                    If mThisPictorial > 0 Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
'                        CC.Paint arr(2), Pat ' render the Pattern
'                    End If
'
'
'                 Case "secondshadow", "secondhand"
'
'                     If mThisFace = 0 Then
'                         scaleString = "C"
'                         txtThisPictorial = CStr(Int(mThisPictorial))
'
'                     ElseIf mThisFace = 1 Then
'
'                         scaleString = "F"
'                         txtThisPictorial = CStr(convertToFahrenheit(Int(mThisPictorial)))
'
'                     ElseIf mThisFace = "2" Then
'                         scaleString = "K"
'                         txtThisPictorial = CStr(convertToKelvin(Int(mThisPictorial)))
'                     End If
'
'                    If Val(txtThisPictorial) >= 100 Then
'                        boxLeft = 260
'                    Else
'                        boxLeft = 265
'                    End If
'
'                     ' draw the date and day in the windows just once to retain opacity
'                    CC.DrawText boxLeft, 103, 200, 200, txtThisPictorial, , , 2, True, , someOpacity
'                    CC.DrawText 459, 297, 200, 200, scaleString, , , 2, True, , someOpacity
'                    'textDrawn = True
'
'                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
'                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
'
'                    CC.RotateDrawingsDeg pictorialSDeg
'                    CC.Paint arr(2), Pat 'render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
'
                    
                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                    CC.Paint arr(2), Pat 'render the Pattern
                   
             End Select
             
             CC.Restore
      End If
     Next
     
    W.Refresh

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlayPict"
End Sub




''---------------------------------------------------------------------------------------
'' Procedure : tmrAnimator_Timer
'' Author    :
'' Date      : 30/07/2023
'' Purpose   : animate timer, once every few (40) milliseconds
''---------------------------------------------------------------------------------------
''
'Private Sub tmrAnimator_Timer()
'
'    On Error GoTo tmrAnimator_Timer_Error
'    '
'    ' move the pointer using pictorialSDeg
'    '
'    animationAmount = animationAmount + animationQuanta
'
'    If animateCCW = True Then
'        If pictorialSDeg > newSDeg Then
'            pictorialSDeg = oldSDeg - animationAmount
'        Else
'            tmrAnimator.Enabled = False ' when the animation is complete then stop the animator
'            animationAmount = 0
'            Exit Sub
'        End If
'    Else
'        If pictorialSDeg < newSDeg Then
'            pictorialSDeg = oldSDeg + animationAmount
'        Else
'            tmrAnimator.Enabled = False
'            animationAmount = 0
'            Exit Sub
'        End If
'    End If
'
'    W.Parent.Refresh
'
'   On Error GoTo 0
'   Exit Sub
'
'tmrAnimator_Timer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimator_Timer of Class Module cwOverlayPict"
'End Sub

  




