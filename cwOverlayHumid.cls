VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlayHumid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlayHumid
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfHumidity
'---------------------------------------------------------------------------------------

Option Explicit

'these Public Properties are set in the Parent (here, cfHumidity)
Public cCollOverlayHumidImages As cCollection


Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private mThisHumidity  As Integer

Private HumiditySDeg As Single
Private oldHumiditySDeg As Double
Private newSDeg As Double
Private animationQuanta As Double
Private animationAmount As Double
Private animationDivisions As Integer
Private animationInterval As Integer
Private animateCCW As Boolean

Private mPointerAnimate As Boolean ' property PointerAnimate
Private mLocked As Boolean ' property Locked
Private mthisHumidityDetails As String ' property
Private mHidden As Boolean ' property Hidden
Private mThisOpacity As String ' property ThisOpacity
Private mPrecipitationRedLamp As Boolean


    

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlayHumid"
End Property
'---------------------------------------------------------------------------------------
' Property : thisHumidity
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisHumidity() As Integer
   On Error GoTo thisHumidityGet_Error

   thisHumidity = mThisHumidity

   On Error GoTo 0
   Exit Property

thisHumidityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisHumidity of Class Module cwOverlayHumid"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisHumidity
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisHumidity(ByVal newValue As Integer)
   On Error GoTo thisHumidityLet_Error

   If mThisHumidity <> newValue Then mThisHumidity = newValue Else Exit Property
   
   Call obtainHumidity

   On Error GoTo 0
   Exit Property

thisHumidityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisHumidity of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlayHumid"
End Property




'---------------------------------------------------------------------------------------
' Property : precipitationRedLamp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get precipitationRedLamp() As Boolean
   On Error GoTo precipitationRedLampGet_Error

    precipitationRedLamp = mPrecipitationRedLamp

   On Error GoTo 0
   Exit Property

precipitationRedLampGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property precipitationRedLamp of Class Module cwOverlayHumid"
End Property
'---------------------------------------------------------------------------------------
' Property  : precipitationRedLamp
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let precipitationRedLamp(ByVal newValue As Boolean)
    On Error GoTo precipitationRedLampLet_Error

    If mPrecipitationRedLamp <> newValue Then mPrecipitationRedLamp = newValue Else Exit Property
    
    W.Refresh
    
   On Error GoTo 0
   Exit Property

precipitationRedLampLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property precipitationRedLamp of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayHumid"
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
        menuForm.mnuLockHumidityGauge.Checked = True
        If widgetPrefs.IsVisible = True Then
             widgetPrefs.cmbGaugeType.ListIndex = 2
             widgetPrefs.chkPreventDragging.Value = 1
        End If

        gblPreventDraggingHumidity = "1"
        If aspectRatio = "landscape" Then
            gblHumidityLandscapeLocked = "1"
            widgetPrefs.cmbLandscapeLocked.ListIndex = 1
        Else
            gblHumidityPortraitLocked = "1"
            widgetPrefs.cmbPortraitLocked.ListIndex = 1
        End If
            
        fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    Else
        menuForm.mnuLockHumidityGauge.Checked = False
        If widgetPrefs.IsVisible = True Then
             widgetPrefs.cmbGaugeType.ListIndex = 2
             widgetPrefs.chkPreventDragging.Value = 0
        End If

        gblPreventDraggingHumidity = "0"
        If aspectRatio = "landscape" Then
            gblHumidityLandscapeLocked = "0"
             widgetPrefs.cmbLandscapeLocked.ListIndex = 0
       Else
            gblHumidityPortraitLocked = "0"
             widgetPrefs.cmbPortraitLocked.ListIndex = 0
       End If
            
        fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    End If

    fHumidity.humidityGaugeForm.Refresh

    sPutINISetting "Software\PzHumidityGauge", "preventDraggingHumidity", gblPreventDraggingHumidity, gblSettingsFile

    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlayHumid"
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
    
    For i = 0 To cCollOverlayHumidImages.Count - 1 ' this is the ExcludedPaths collection populated in cfTemperature by duplicating collTemperaturePSDNonUIElements.
      arr = cCollOverlayHumidImages.ItemByIndex(i)
      key = cCollOverlayHumidImages.KeyByIndex(i)
      
      If IsArray(arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case key
                 
                 Case "Humidityface"
                    
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                 
                Case "redlamptrue"
                    
                    ' if it is raining
                    If mPrecipitationRedLamp = True Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "redlampfalse"
                    
                    ' if it is not raining
                    If mPrecipitationRedLamp = False Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                                            
                 Case "pointerShadow", "pointer"

                    If Val(mThisHumidity) >= 100 Then
                        boxLeft = 260
                    Else
                        boxLeft = 265
                    End If

                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText boxLeft, 110, 170, 200, CStr(mThisHumidity), , , 2, True, , someOpacity

                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area

                    CC.RotateDrawingsDeg HumiditySDeg
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlayHumid"
End Sub



'
' ----------------------------------------------------------------
' Procedure Name: obtainHumidity
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 13/01/2024
' ----------------------------------------------------------------
Private Sub obtainHumidity()

    On Error GoTo obtainHumidity_Error
    
    ' calculate the new cpu amount in degrees
    If mThisHumidity = 0 Then
        HumiditySDeg = 30
    Else
        HumiditySDeg = (mThisHumidity * 3) + 30
    End If

    W.Parent.Refresh

    On Error GoTo 0
    Exit Sub

obtainHumidity_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainHumidity, line " & Erl & "."

End Sub






