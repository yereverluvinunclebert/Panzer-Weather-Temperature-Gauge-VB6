VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlayAnemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlayAnemo
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfAnemometer
'---------------------------------------------------------------------------------------

Option Explicit

'these Public Properties are set in the Parent (here, cfAnemometer)
Public cCollOverlayAnemoImages As cCollection


Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private WithEvents tmrAnimator As cTimer
Attribute tmrAnimator.VB_VarHelpID = -1

Private windSpeedSDeg As Single
Private windDirectionSDeg As Single

Private thisWindSpeed As Long
Private mWindSpeed  As Integer
Private mWindDirection  As Integer


Private oldWindSpeedSDeg As Double
Private newSDeg As Double
Private animationQuanta As Double
Private animationAmount As Double
Private animationDivisions As Integer
Private animationInterval As Integer
Private animateCCW As Boolean

Private mPointerAnimate As Boolean ' property PointerAnimate
Private mLocked As Boolean ' property Locked
Private mAnemometerDetails As String ' property
Private mHidden As Boolean ' property Hidden
Private mThisOpacity As String ' property ThisOpacity


    

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Error
    
    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    animationDivisions = 12
    animationInterval = 40
        
    Set tmrAnimator = Nothing
    Set tmrAnimator = New_c.Timer(animationInterval, False)
     
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlayAnemo"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlayAnemo"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlayAnemo"
End Property
'---------------------------------------------------------------------------------------
' Property : WindSpeed
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get windSpeed() As Integer
   On Error GoTo WindSpeedGet_Error

   windSpeed = mWindSpeed

   On Error GoTo 0
   Exit Property

WindSpeedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WindSpeed of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : WindSpeed
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let windSpeed(ByVal newValue As Integer)
   On Error GoTo WindSpeedLet_Error

   If mWindSpeed <> newValue Then mWindSpeed = newValue Else Exit Property
   
   Call calcAnimationData

   On Error GoTo 0
   Exit Property

WindSpeedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WindSpeed of Class Module cwOverlayTemp"
End Property

'---------------------------------------------------------------------------------------
' Property : windDirection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get windDirection() As Integer
   On Error GoTo windDirectionGet_Error

   windDirection = mWindDirection

   On Error GoTo 0
   Exit Property

windDirectionGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property windDirection of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : windDirection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let windDirection(ByVal newValue As Integer)
   On Error GoTo windDirectionLet_Error

   If mWindDirection <> newValue Then mWindDirection = newValue Else Exit Property
   
    Call obtainWindDirection

   On Error GoTo 0
   Exit Property

windDirectionLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property windDirection of Class Module cwOverlayTemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlayAnemo"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlayAnemo"
End Property


'---------------------------------------------------------------------------------------
' Property : pointerAnimate
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get pointerAnimate() As Boolean
   On Error GoTo pointerAnimateGet_Error

    pointerAnimate = mPointerAnimate

   On Error GoTo 0
   Exit Property

pointerAnimateGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlayAnemo"
End Property
'---------------------------------------------------------------------------------------
' Property  : pointerAnimate
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let pointerAnimate(ByVal newValue As Boolean)
    On Error GoTo pointerAnimateLet_Error

    If mPointerAnimate <> newValue Then mPointerAnimate = newValue Else Exit Property
    
   On Error GoTo 0
   Exit Property

pointerAnimateLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlayTemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayTemp"
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
        menuForm.mnuLockAnemometerGauge.Checked = True
        panzerPrefs.chkPreventDragging.Value = 1
        panzerPrefs.cmbLandscape.ListIndex = 1

        PzGPreventDraggingAnemometer = "1"
        If aspectRatio = "landscape" Then
            PzGAnemometerLandscape = "1"
        Else
            PzGAnemometerPortrait = "1"
        End If
            
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    Else
        menuForm.mnuLockAnemometerGauge.Checked = False
        panzerPrefs.chkPreventDragging.Value = 0
        panzerPrefs.cmbLandscape.ListIndex = 0

        PzGPreventDraggingAnemometer = "0"
        If aspectRatio = "landscape" Then
            PzGAnemometerLandscape = "0"
        Else
            PzGAnemometerPortrait = "0"
        End If
            
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(PzGOpacity) / 100
    End If

    fAnemometer.anemometerGaugeForm.Refresh

    sPutINISetting "Software\PzAnemometerGauge", "preventDraggingTemperature", PzGPreventDraggingAnemometer, PzGSettingsFile

    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayTemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayTemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayTemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlayTemp"
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
    Dim someOpacity As Double: someOpacity = 0
    Dim boxLeft As Integer: boxLeft = 0
    Dim cream As String: cream = vbNullString
       
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
    
    For i = 0 To cCollOverlayAnemoImages.Count - 1 ' this is the ExcludedPaths collection populated in cfTemperature by duplicating collTemperaturePSDNonUIElements.
      arr = cCollOverlayAnemoImages.ItemByIndex(i)
      key = cCollOverlayAnemoImages.KeyByIndex(i)
      
      If IsArray(arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case key
                 
                 Case "anemometerface"
                    
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
        
                Case "redlamptrue"
                    
                    If mWindSpeed >= 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "redlampfalse"
                    
                    If mWindSpeed < 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                                    
                    
                 Case "directionpointer"
                                         
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(4, 180) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    ' windDirectionSDeg = 180 ' debug
                    
                    CC.RotateDrawingsDeg windDirectionSDeg
                    CC.Paint arr(2), Pat 'render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
                  
                 Case "pointerShadow", "pointer"

                    If Val(mWindSpeed) >= 100 Then
                        boxLeft = 415
                    Else
                        boxLeft = 420
                    End If

                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText boxLeft, 185, 180, 200, CStr(mWindSpeed), , , 2, True, , someOpacity

                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg windSpeedSDeg
                    CC.Paint arr(2), Pat 'render the Pattern with the current PSD-LayerAlpha, present in Arr(2)

                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                    CC.Paint arr(2), Pat 'render the Pattern
                   
             End Select
             
             CC.Restore
      End If
     Next

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlayTemp"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : tmrAnimator_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : animate timer, once every few (40) milliseconds
'---------------------------------------------------------------------------------------
'
Private Sub tmrAnimator_Timer()

    On Error GoTo tmrAnimator_Timer_Error
    '
    ' move the pointer using windSpeedSDeg
    '
    animationAmount = animationAmount + animationQuanta

    If animateCCW = True Then
        If windSpeedSDeg > newSDeg Then
            windSpeedSDeg = oldWindSpeedSDeg - animationAmount
        Else
            tmrAnimator.Enabled = False ' when the animation is complete then stop the animator
            animationAmount = 0
            Exit Sub
        End If
    Else
        If windSpeedSDeg < newSDeg Then
            windSpeedSDeg = oldWindSpeedSDeg + animationAmount
        Else
            tmrAnimator.Enabled = False
            animationAmount = 0
            Exit Sub
        End If
    End If

    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

tmrAnimator_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimator_Timer of Class Module cwOverlayTemp"
End Sub
'
'
'
'
'
'---------------------------------------------------------------------------------------
' Procedure : calcAnimationData
' Author    : set the animation variables and call the routine to move the gauge pointer
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub calcAnimationData()
    Dim degDiff As Double: degDiff = 0
    Dim windSpeed As Double: windSpeed = 0

    On Error GoTo calcAnimationData_Error
    ' windSpeed = -14 ' debug

    If mPointerAnimate = True Then
        If tmrAnimator.Enabled = False Then

            ' oldWindSpeedSDeg stored
            oldWindSpeedSDeg = windSpeedSDeg

            ' obtain the new windSpeed in percentage
            windSpeed = ((mWindSpeed) / 80) * 100

            ' calculate the new cpu amount in degrees
            If mWindSpeed = 0 Then
                windSpeedSDeg = 30
            Else
                windSpeedSDeg = (windSpeed * 3) + 30
            End If

            If oldWindSpeedSDeg = newSDeg Then Exit Sub

            ' calculate the difference
            degDiff = newSDeg - oldWindSpeedSDeg

            ' if the result is negative animate CCW

            animateCCW = degDiff < 0 ' true/false

            Select Case degDiff
                Case Is <= 5
                    animationDivisions = 4
                Case Is <= 10
                    animationDivisions = 8
                Case Is > 10
                    animationDivisions = 16
                Case Is > 20
                    animationDivisions = 24
            End Select

            ' divide that by 24 (one second duration)
            animationQuanta = Abs(degDiff / animationDivisions)

            tmrAnimator.Enabled = True

        End If
    Else
        Call obtainWindSpeed
    End If

    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

calcAnimationData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcAnimationData of Class Module cwOverlayTemp"
End Sub
'
' ----------------------------------------------------------------
' Procedure Name: obtainWindSpeed
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 13/01/2024
' ----------------------------------------------------------------
Private Sub obtainWindSpeed()

    Dim windSpeed As Double: windSpeed = 0

    On Error GoTo obtainWindSpeed_Error
    
    ' obtain the new windSpeed in percentage
    windSpeed = ((mWindSpeed) / 80) * 100

    ' calculate the new cpu amount in degrees
    If windSpeed = 0 Then
        windSpeedSDeg = 30
    Else
        windSpeedSDeg = (windSpeed * 3) + 30
    End If

    W.Parent.Refresh

    On Error GoTo 0
    Exit Sub

obtainWindSpeed_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainWindSpeed, line " & Erl & "."

End Sub

'
' ----------------------------------------------------------------
' Procedure Name: obtainWindDirection
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 13/01/2024
' ----------------------------------------------------------------
Private Sub obtainWindDirection()

    On Error GoTo obtainWindDirection_Error
    
    windDirectionSDeg = mWindDirection

    W.Parent.Refresh

    On Error GoTo 0
    Exit Sub

obtainWindDirection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainWindDirection, line " & Erl & "."

End Sub



