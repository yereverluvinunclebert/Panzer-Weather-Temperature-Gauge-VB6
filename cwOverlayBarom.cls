VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlayBarom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlayBarom
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfBarometer
'---------------------------------------------------------------------------------------

Option Explicit

'these Public Properties are set in the Parent (here, cfBarometer)
Public cCollOverlayBaromImages As cCollection


Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1


Private mThisPressure  As Long
Private mPressureDrop  As Boolean

Private pressureSDeg As Single
Private oldPressureSDeg As Double
Private newSDeg As Double
Private animationQuanta As Double
Private animationAmount As Double
Private animationDivisions As Integer
Private animationInterval As Integer
Private animateCCW As Boolean

Private mThisFace As Integer             ' property mThisFace
Private mPointerAnimate As Boolean ' property PointerAnimate
Private mLocked As Boolean ' property Locked
Private mThisPressureDetails As String ' property
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
                         
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlayBarom"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlayBarom"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlayBarom"
End Property
'---------------------------------------------------------------------------------------
' Property : pressureDrop
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get pressureDrop() As Boolean
   On Error GoTo pressureDropGet_Error

   pressureDrop = mPressureDrop

   On Error GoTo 0
   Exit Property

pressureDropGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pressureDrop of Class Module cwOverlayBarom"
End Property
'---------------------------------------------------------------------------------------
' Property  : pressureDrop
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let pressureDrop(ByVal newValue As Boolean)
   On Error GoTo pressureDropLet_Error

   If mPressureDrop <> newValue Then mPressureDrop = newValue Else Exit Property
   

   On Error GoTo 0
   Exit Property

pressureDropLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pressureDrop of Class Module cwOverlayBarom"
End Property

'---------------------------------------------------------------------------------------
' Property : thisPressure
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisPressure() As Long
   On Error GoTo thisPressureGet_Error

   thisPressure = mThisPressure

   On Error GoTo 0
   Exit Property

thisPressureGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisPressure of Class Module cwOverlayBarom"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisPressure
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisPressure(ByVal newValue As Long)
   On Error GoTo thisPressureLet_Error

   If mThisPressure <> newValue Then mThisPressure = newValue Else Exit Property
   
   Call obtainPressure

   On Error GoTo 0
   Exit Property

thisPressureLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisPressure of Class Module cwOverlayBarom"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlayBarom"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlayBarom"
End Property


''---------------------------------------------------------------------------------------
'' Property : pointerAnimate
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get pointerAnimate() As Boolean
'   On Error GoTo pointerAnimateGet_Error
'
'    pointerAnimate = mPointerAnimate
'
'   On Error GoTo 0
'   Exit Property
'
'pointerAnimateGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlayBarom"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : pointerAnimate
'' Author    :
'' Date      : 09/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let pointerAnimate(ByVal newValue As Boolean)
'    On Error GoTo pointerAnimateLet_Error
'
'    If mPointerAnimate <> newValue Then mPointerAnimate = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'pointerAnimateLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlayBarom"
'End Property

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayBarom"
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
        menuForm.mnuLockBarometerGauge.Checked = True
        If widgetPrefs.IsVisible = True Then
             widgetPrefs.cmbGaugeType.ListIndex = 3
             widgetPrefs.chkPreventDragging.Value = 1
        End If

        gblPreventDraggingBarometer = "1"
        If aspectRatio = "landscape" Then
            gblBarometerLandscapeLocked = "1"
            widgetPrefs.cmbLandscapeLocked.ListIndex = 1
        Else
            gblBarometerPortraitLocked = "1"
            widgetPrefs.cmbPortraitLocked.ListIndex = 1
        End If
            
        fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    Else
        menuForm.mnuLockBarometerGauge.Checked = False
        If widgetPrefs.IsVisible = True Then
             widgetPrefs.cmbGaugeType.ListIndex = 3
             widgetPrefs.chkPreventDragging.Value = 0
        End If

        gblPreventDraggingBarometer = "0"
        If aspectRatio = "landscape" Then
            gblBarometerLandscapeLocked = "0"
            widgetPrefs.cmbLandscapeLocked.ListIndex = 0
       Else
            gblBarometerPortraitLocked = "0"
            widgetPrefs.cmbPortraitLocked.ListIndex = 0
        End If
            
        fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    End If

    fBarometer.barometerGaugeForm.Refresh

    sPutINISetting "Software\PzBarometerGauge", "preventDraggingBarometer", gblPreventDraggingBarometer, gblSettingsFile

    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayBarom"
End Property
'---------------------------------------------------------------------------------------
' Property : thisFace
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisFace() As Integer
   On Error GoTo thisFaceGet_Error

    thisFace = mThisFace

   On Error GoTo 0
   Exit Property

thisFaceGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisFace of Class Module cwoverlaytemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisFace
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisFace(ByVal newValue As Integer)
    On Error GoTo thisFaceLet_Error

    If mThisFace <> newValue Then mThisFace = newValue Else Exit Property
    
    If widgetPrefs.IsVisible = True Then
        widgetPrefs.cmbPressureScale.ListIndex = mThisFace - 1
    End If
    
    gblPressureScale = mThisFace
    sPutINISetting "Software\PzTemperatureGauge", "pressureScale", gblPressureScale, gblSettingsFile
    
    Call obtainPressure
    
    On Error GoTo 0
    Exit Property

thisFaceLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisFace of Class Module cwoverlaytemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayBarom"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayBarom"
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
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlayBarom"
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
    Dim myPressure As Single: myPressure = 0
    Dim scaleString As String: scaleString = vbNullString
       
    On Error GoTo W_Paint_Error
    
    cream = RGB(235, 217, 179) ' cannot create a constant from an RGB value
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 15
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    someOpacity = (Val(mThisOpacity) / 100) * 0.75
    
    For i = 0 To cCollOverlayBaromImages.Count - 1 ' this is the ExcludedPaths collection populated in cfTemperature by duplicating collTemperaturePSDNonUIElements.
      arr = cCollOverlayBaromImages.ItemByIndex(i)
      key = cCollOverlayBaromImages.KeyByIndex(i)
      
      If IsArray(arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case key
    
                Case "barometermb600-800face"
                    
                    If mThisFace = 0 And mThisPressure <= 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                    
                Case "barometermb800-1066face"
                    
                    If mThisFace = 0 And mThisPressure > 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                    
                 Case "barometerinhg18-21face"
                    
                    If mThisFace = 1 And mThisPressure <= 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                    
                 Case "barometerinhg23-31face"
                    
                    If mThisFace = 1 And mThisPressure > 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                     
                  Case "barometermmhg450-600face"

                    If mThisFace = 2 And mThisPressure <= 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                  
                  Case "barometermmhg600-800face"

                    If mThisFace = 2 And mThisPressure > 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
               
                 Case "barometerhpa600-800face"

                    If mThisFace = 3 And mThisPressure <= 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                                            
                 Case "barometerhpa800-1066face"

                    If mThisFace = 3 And mThisPressure > 800 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                 
                 Case "redlamp"
                    
                    If mPressureDrop = True Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "greenlamp"
                    
                    If mPressureDrop = False Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                                            
                 Case "manualpointer"
                 
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                                            
                 Case "pointerShadow", "pointer"
                 
                    myPressure = mThisPressure
                    ' convert for the digital indicators from millibars to mercury
                    If mThisFace = 1 Then
                        myPressure = myPressure * 0.0295 ' inches of mercury
                    ElseIf mThisFace = 2 Then
                        myPressure = myPressure * 0.75006 ' millimetres of mercury
                    End If
                     
                    If Val(myPressure) >= 100 Then
                        boxLeft = 260
                    Else
                        boxLeft = 265
                    End If

                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText boxLeft, 108, 170, 200, CStr(myPressure), , , 2, True, , someOpacity
                    
                    If mThisFace = 0 Then scaleString = "mb"
                    If mThisFace = 1 Then scaleString = "inH"
                    If mThisFace = 2 Then scaleString = "mHg"
                    If mThisFace = 3 Then scaleString = "hpa"
                    CC.DrawText 442, 303, 200, 200, scaleString, , , 2, True, , someOpacity

                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg pressureSDeg
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlayBarom"
End Sub



'
' ----------------------------------------------------------------
' Procedure Name: obtainPressure
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 13/01/2024
' ----------------------------------------------------------------
Private Sub obtainPressure()

    Dim pressurePercent As Double: pressurePercent = 0
    Dim a As Double: a = 0

    On Error GoTo obtainPressure_Error
    
    'mThisPressure = 850  ' debug
    
    ' obtain the new pressure as a percentage
    If mThisPressure <= 800 Then
        a = mThisPressure - 600
        a = a / 2
        pressurePercent = a
    Else
        a = mThisPressure - 800
        a = a / 2
        pressurePercent = a
    End If
        
    ' calculate the new amount in degrees
    If mThisPressure = 0 Then
        pressureSDeg = 30
    Else
        pressureSDeg = (pressurePercent * 3) + 30
    End If

    W.Parent.Refresh

    On Error GoTo 0
    Exit Sub

obtainPressure_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainPressure, line " & Erl & "."

End Sub







