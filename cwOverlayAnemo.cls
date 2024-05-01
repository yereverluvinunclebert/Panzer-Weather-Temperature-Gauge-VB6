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
' Module    : cwOverlayTemp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfAnemometer
'---------------------------------------------------------------------------------------

' Microsoft XML, v3.0
' c:/windows/SysWow64/msxml3.dll

Option Explicit

'Private Type SYSTEMTIME
'   wYear As Integer
'   wMonth As Integer
'   wDayOfWeek As Integer
'   wDay As Integer
'   wHour As Integer
'   wMinute As Integer
'   wSecond As Integer
'   wMilliseconds As Integer
'End Type
'
'Private Type TIME_ZONE_INFORMATION
'    bias                    As Long
'    StandardName(0 To 63)   As Byte
'    StandardDate            As SYSTEMTIME
'    StandardBias            As Long
'    DaylightName(0 To 63)   As Byte
'    DaylightDate            As SYSTEMTIME
'    DaylightBias            As Long
'End Type
'
'Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
'
'Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'these Public Properties are set in the Parent (here, cfAnemometer)
Public cCollOverlayAnemoImages As cCollection


Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private WithEvents tmrAnimator As cTimer
Attribute tmrAnimator.VB_VarHelpID = -1

Private SDeg As Double
Private thisWindSpeed As Long
Private windSpeedValue  As Long

Private oldSDeg As Double
Private newSDeg As Double
Private animationQuanta As Double
Private animationAmount As Double
Private animationDivisions As Integer
Private animationInterval As Integer
Private animateCCW As Boolean

'------------------------------------------------------ ENDS

Private mTicking As Boolean ' property Ticking
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
' Property  : Ticking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Ticking(ByVal newValue As Boolean)
    On Error GoTo TickingLet_Error

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlayAnemo"
End Property ' the Ticking flag


'---------------------------------------------------------------------------------------
' Property : Ticking
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Ticking() As Boolean
   On Error GoTo TickingGet_Error

  Ticking = mTicking

   On Error GoTo 0
   Exit Property

TickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlayAnemo"
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
        'panzerPrefs.chkPreventDragging.Value = 1
        PzGPreventDraggingAnemometer = "1"
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    Else
        menuForm.mnuLockAnemometerGauge.Checked = False
        'panzerPrefs.chkPreventDragging.Value = 0
        PzGPreventDraggingAnemometer = "0"
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = val(PzGOpacity) / 100
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

'---------------------------------------------------------------------------------------





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
    Dim txtThisWindSpeed As String: txtThisWindSpeed = vbNullString
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
    someOpacity = (val(mThisOpacity) / 100) * 0.75
    
    txtThisWindSpeed = CStr(Int(thisWindSpeed))

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
                    
                    If thisWindSpeed >= 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "redlampfalse"
                    
                    If thisWindSpeed < 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                                    

                 Case "pointerShadow", "pointer"
                    
                    scaleString = "C"
                    txtThisWindSpeed = CStr(thisWindSpeed)
                                              
                    If val(txtThisWindSpeed) >= 100 Then
                        boxLeft = 260
                    Else
                        boxLeft = 265
                    End If
                    
                     
                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText boxLeft, 103, 200, 200, txtThisWindSpeed, , , 2, True, , someOpacity
                    CC.DrawText 459, 297, 200, 200, scaleString, , , 2, True, , someOpacity
                    'textDrawn = True
                    
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    CC.RotateDrawingsDeg SDeg
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

'
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
'    ' move the pointer using SDeg
'    '
'    animationAmount = animationAmount + animationQuanta
'
'    If animateCCW = True Then
'        If SDeg > newSDeg Then
'            SDeg = oldSDeg - animationAmount
'        Else
'            tmrAnimator.Enabled = False ' when the animation is complete then stop the animator
'            animationAmount = 0
'            Exit Sub
'        End If
'    Else
'        If SDeg < newSDeg Then
'            SDeg = oldSDeg + animationAmount
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimator_Timer of Class Module cwOverlayTemp"
'End Sub
'
'
'
'
'
''---------------------------------------------------------------------------------------
'' Procedure : calcAnimationData
'' Author    : set the animation variables and call the routine to move the gauge pointer
'' Date      : 30/07/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub calcAnimationData()
'    Dim degDiff As Double: degDiff = 0
''    Dim min  As Integer: min = 0
''    Dim max  As Integer: max = 0
'    Dim tempPercent As Double: tempPercent = 0
'
'    On Error GoTo calcAnimationData_Error
'    ' windSpeedValue = -14 ' debug
'
'    If mPointerAnimate = True Then
'        If tmrAnimator.Enabled = False Then
'
'            ' oldSDeg stored
'            oldSDeg = SDeg
'
'            ' obtain the new windSpeed in percentage
'            tempPercent = ((windSpeedValue + 30) / 80) * 100
'
'            ' calculate the new cpu amount in degrees
'            If weHaveData = True Then
'                newSDeg = (tempPercent * 3) + 30
'            Else
'                newSDeg = 30
'            End If
'
'            If oldSDeg = newSDeg Then Exit Sub
'
'            ' calculate the difference
'            degDiff = newSDeg - oldSDeg
'
'            ' if the result is negative animate CCW
'
'            animateCCW = degDiff < 0 ' true/false
'
'            Select Case degDiff
'                Case Is <= 5
'                    animationDivisions = 4
'                Case Is <= 10
'                    animationDivisions = 8
'                Case Is > 10
'                    animationDivisions = 16
'                Case Is > 20
'                    animationDivisions = 24
'            End Select
'
'            ' divide that by 24 (one second duration)
'            animationQuanta = Abs(degDiff / animationDivisions)
'
'            tmrAnimator.Enabled = True
'
'        End If
'    Else
'        Call obtainWindSpeed
'    End If
'
'    W.Parent.Refresh
'
'   On Error GoTo 0
'   Exit Sub
'
'calcAnimationData_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcAnimationData of Class Module cwOverlayTemp"
'End Sub
'
'' ----------------------------------------------------------------
'' Procedure Name: obtainWindCharacteristics
'' Purpose:
'' Procedure Kind: Sub
'' Procedure Access: Private
'' Author: Dean Beedell (yereverluvinunclebert)
'' Date: 13/01/2024
'' ----------------------------------------------------------------
'Private Sub obtainWindCharacteristics()
'
'    Dim tempPercent As Double: tempPercent = 0
'
'    On Error GoTo obtainWindCharacteristics_Error
'
'    ' obtain the new windSpeed in percentage
'    tempPercent = ((windSpeedValue + 30) / 80) * 100
'
'    ' calculate the new cpu amount in degrees
'    If weHaveData = True Then
'        SDeg = (tempPercent * 3) + 30
'    Else
'        SDeg = 30
'    End If
'
'    W.Parent.Refresh
'
'    On Error GoTo 0
'    Exit Sub
'
'obtainWindCharacteristics_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainWindCharacteristics, line " & Erl & "."
'
'End Sub



