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
Private mthisWeatherImage As String             ' property
Private mthisWindImage As String                ' property
Private mthisShowersImage As String
Private mthisFogImage As String
Private mthisExoticImage As String

Private mvalidData As Boolean




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
' Property : validData
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get validData() As Boolean
   On Error GoTo validDataGet_Error

    validData = mvalidData

   On Error GoTo 0
   Exit Property

validDataGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property validData of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Property  : validData
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let validData(ByVal newValue As Boolean)
    On Error GoTo validDataLet_Error

    If mvalidData <> newValue Then mvalidData = newValue Else Exit Property
    
    fPictorial.pictorialGaugeForm.Refresh
        
    On Error GoTo 0
    Exit Property

validDataLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property validData of Class Module cwoverlaypict"
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
' Property : thisWindImage
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisWindImage() As String
   On Error GoTo thisWindImageGet_Error

    thisWindImage = mthisWindImage

   On Error GoTo 0
   Exit Property

thisWindImageGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWindImage of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisWindImage
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisWindImage(ByVal newValue As String)
    On Error GoTo thisWindImageLet_Error

    If mthisWindImage <> newValue Then mthisWindImage = newValue Else Exit Property
   
    Cairo.ImageList.AddSurface "windiconimage", Cairo.ImageList(mthisWindImage)

    On Error GoTo 0
    Exit Property

thisWindImageLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWindImage of Class Module cwoverlaypict"
End Property




'---------------------------------------------------------------------------------------
' Property : thisWeatherImage
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisWeatherImage() As String
   On Error GoTo thisWeatherImageGet_Error

    thisWeatherImage = mthisWeatherImage

   On Error GoTo 0
   Exit Property

thisWeatherImageGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWeatherImage of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisWeatherImage
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisWeatherImage(ByVal newValue As String)
    On Error GoTo thisWeatherImageLet_Error

    If mthisWeatherImage <> newValue Then mthisWeatherImage = newValue Else Exit Property
    
    Cairo.ImageList.AddSurface "weathericonimage", Cairo.ImageList(mthisWeatherImage)

    On Error GoTo 0
    Exit Property

thisWeatherImageLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisWeatherImage of Class Module cwoverlaypict"
End Property


'---------------------------------------------------------------------------------------
' Property : thisShowersImage
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisShowersImage() As String
   On Error GoTo thisShowersImageGet_Error

    thisShowersImage = mthisShowersImage

   On Error GoTo 0
   Exit Property

thisShowersImageGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisShowersImage of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisShowersImage
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisShowersImage(ByVal newValue As String)
    On Error GoTo thisShowersImageLet_Error

    If mthisShowersImage <> newValue Then mthisShowersImage = newValue Else Exit Property
    
    Cairo.ImageList.AddSurface "showersiconimage", Cairo.ImageList(mthisShowersImage)

    On Error GoTo 0
    Exit Property

thisShowersImageLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisShowersImage of Class Module cwoverlaypict"
End Property


'---------------------------------------------------------------------------------------
' Property : thisFogImage
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisFogImage() As String
   On Error GoTo thisFogImageGet_Error

    thisFogImage = mthisFogImage

   On Error GoTo 0
   Exit Property

thisFogImageGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisFogImage of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisFogImage
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisFogImage(ByVal newValue As String)
    On Error GoTo thisFogImageLet_Error

    If mthisFogImage <> newValue Then mthisFogImage = newValue Else Exit Property
    
    Cairo.ImageList.AddSurface "fogiconimage", Cairo.ImageList(mthisFogImage)

    On Error GoTo 0
    Exit Property

thisFogImageLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisFogImage of Class Module cwoverlaypict"
End Property



'---------------------------------------------------------------------------------------
' Property : thisExoticImage
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisExoticImage() As String
   On Error GoTo thisExoticImageGet_Error

    thisExoticImage = mthisExoticImage

   On Error GoTo 0
   Exit Property

thisExoticImageGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisExoticImage of Class Module cwoverlaypict"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisExoticImage
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisExoticImage(ByVal newValue As String)
    On Error GoTo thisExoticImageLet_Error

    If mthisExoticImage <> newValue Then mthisExoticImage = newValue Else Exit Property
    
    Cairo.ImageList.AddSurface "exoticiconimage", Cairo.ImageList(mthisExoticImage)

    On Error GoTo 0
    Exit Property

thisExoticImageLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisExoticImage of Class Module cwoverlaypict"
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
            panzerPrefs.cmbGaugeType.ListIndex = 4
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
            panzerPrefs.cmbGaugeType.ListIndex = 4
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

                Case "redlamp"

                     ' change the lamp to bright green or red if no valid data
                    If mvalidData = False Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If

                 Case "greenlampfalse"
                 
'                     ' change the lamp to bright or unlit green
'                    If mvalidData = False Then
'                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
'                        CC.Paint arr(2), Pat ' render the Pattern
'                    End If

                 Case "greenlamp"
                 
                    If mvalidData = True Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                                            
                 Case "manualpointer"
                 
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    
                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                    CC.Paint arr(2), Pat 'render the Pattern
                   
             End Select
             
             CC.Restore
      End If
     Next
     
     CC.RenderSurfaceContent "weathericonimage", 295, 250, , , , someOpacity
     CC.RenderSurfaceContent "windiconimage", 295, 250, , , , someOpacity
     CC.RenderSurfaceContent "showersiconimage", 295, 250, , , , someOpacity
     CC.RenderSurfaceContent "fogiconimage", 295, 250, , , , someOpacity
     CC.RenderSurfaceContent "exoticiconimage", 295, 250, , , , someOpacity
          
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

  




