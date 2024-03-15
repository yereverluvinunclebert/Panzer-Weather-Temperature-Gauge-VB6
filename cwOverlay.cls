VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlay
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfTemperature
'---------------------------------------------------------------------------------------

Option Explicit


'these Public Properties are set in the Parent (here, cfTemperature)
Public cCollOverlayImages As cCollection
Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrSampler As cTimer
Attribute tmrSampler.VB_VarHelpID = -1
Private WithEvents tmrAsyncResponse As cTimer
Attribute tmrAsyncResponse.VB_VarHelpID = -1
Private WithEvents tmrAnimator As cTimer
Attribute tmrAnimator.VB_VarHelpID = -1

Private SDeg As Double
Private thisTemperature As Long
Private temperatureTotalTimeLeft As Single
'Private TemperatureDetails As String
Private TemperatureVal  As Long

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
Private mTemperatureDetails As String ' property

Private mHidden As Boolean ' property Hidden
Private mThisOpacity As String ' property ThisOpacity
Private mSamplingInterval As String ' property SamplingInterval
Private mThisFace As Integer             ' property mThisFace

Private myMSXML As MSXML2.XMLHTTP
Private weHaveData As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Error
    
    'Dim sunriseSunset As cwSunriseSunset 'unitialised declaration
    Set sunriseSunset = New cwSunriseSunset

    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    Set tmrSampler = Nothing
    Set tmrSampler = New_c.Timer(100, mTicking) ' first interval is 1/10th of a second, don't change it, triggers an immediate run. Correct sampling interval read later
    
    animationDivisions = 12
    animationInterval = 40
    
    Set tmrAnimator = Nothing
    Set tmrAnimator = New_c.Timer(animationInterval, False)
 
    Set tmrAsyncResponse = Nothing
    Set tmrAsyncResponse = New_c.Timer(100, False)
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : Ticking
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Ticking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Ticking(ByVal newValue As Boolean)
    'Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo TickingLet_Error

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    
    If mTicking Then
        tmrSampler.Enabled = True
    Else
        tmrSampler.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property ' the earth Ticking flag

'---------------------------------------------------------------------------------------
' Property : thisOpacity
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : thisSensor
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisFace of Class Module cwOverlay"
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
    
    Call obtainTemperature
    
    On Error GoTo 0
    Exit Property

thisFaceLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisFace of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : samplingInterval
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get samplingInterval() As String
   On Error GoTo samplingIntervalGet_Error

    samplingInterval = mSamplingInterval

   On Error GoTo 0
   Exit Property

samplingIntervalGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : samplingInterval
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let samplingInterval(ByVal newValue As String)
    On Error GoTo samplingIntervalLet_Error

    If mSamplingInterval <> newValue Then mSamplingInterval = newValue Else Exit Property
    
    If tmrSampler.Interval = 100 Then Call getData ' allows first early run
    'If tmrSampler.Interval = 100 Then Call tmrSampler_Timer ' allows first early run
    
    tmrSampler.Interval = Val(mSamplingInterval) * 1000
'    If mPointerAnimate = True Then
'        If tmrSampler.Interval = 1000 Then tmrSampler.Interval = 1200 ' take into account the animator
'    End If
    
    On Error GoTo 0
    Exit Property

samplingIntervalLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : pointerAnimate
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : Locked
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : TemperatureDetails
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TemperatureDetails() As String
   On Error GoTo TemperatureDetailsGet_Error

   TemperatureDetails = mTemperatureDetails

   On Error GoTo 0
   Exit Property

TemperatureDetailsGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TemperatureDetails of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : TemperatureDetails
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let TemperatureDetails(ByVal newValue As String)
   On Error GoTo TemperatureDetailsLet_Error

   If mTemperatureDetails <> newValue Then mTemperatureDetails = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

TemperatureDetailsLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TemperatureDetails of Class Module cwOverlay"
End Property




'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlay"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtTemperaturetotal As String: txtTemperaturetotal = vbNullString
    Dim txtThisTemperature As String: txtThisTemperature = vbNullString
    Dim textDrawn As Boolean:   textDrawn = False
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
    
    txtThisTemperature = CStr(Int(thisTemperature))
    txtTemperaturetotal = CStr(Int(temperatureTotalTimeLeft))
    



    For I = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfTemperature by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(I)
      Key = cCollOverlayImages.KeyByIndex(I)
      
      If IsArray(Arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case Key
                 
                 Case "centigradeface"
                    
                    If mThisFace = 0 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                        CC.Paint Arr(2), Pat ' render the Pattern
                    End If
                 
                 Case "fahrenheitface"
                    
                    If mThisFace = 1 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                 
                 Case "kelvinface"
                    
                    If mThisFace = 2 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
        
                Case "redlamptrue"
                    
                    If thisTemperature >= 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "redlampfalse"
                    
                    If thisTemperature < 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                        CC.Paint Arr(2), Pat ' render the Pattern
                    End If
                                        
                 Case "bluelamptrue"
                    
                    If thisTemperature <= 0 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "bluelampfalse"
                    
                    If thisTemperature > 0 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                        CC.Paint Arr(2), Pat ' render the Pattern
                    End If
                    

                 Case "secondshadow", "secondhand"
                    
                     If mThisFace = "0" Then
                         scaleString = "C"
                         txtThisTemperature = CStr(thisTemperature)
                         
                     ElseIf mThisFace = "1" Then
                    
                         scaleString = "F"
                         txtThisTemperature = CStr(convertToFahrenheit(thisTemperature))
                    
                     ElseIf mThisFace = "2" Then
                         scaleString = "K"

                         txtThisTemperature = CStr(convertToKelvin(thisTemperature))
                     End If
                     
                    If Val(txtThisTemperature) >= 100 Then
                        boxLeft = 260
                    Else
                        boxLeft = 265
                    End If
                    
                     
                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText boxLeft, 103, 200, 200, txtThisTemperature, , , 2, True, , someOpacity
                    CC.DrawText 459, 297, 200, 200, scaleString, , , 2, True, , someOpacity
                    textDrawn = True
                    
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    CC.RotateDrawingsDeg SDeg
                    CC.Paint Arr(2), Pat 'render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
                    
                    
                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                    CC.Paint Arr(2), Pat 'render the Pattern
                   
             End Select
             
             CC.Restore
      End If
     Next

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
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
    ' move the pointer using SDeg
    '
    animationAmount = animationAmount + animationQuanta

    If animateCCW = True Then
        If SDeg > newSDeg Then
            SDeg = oldSDeg - animationAmount
        Else
            tmrAnimator.Enabled = False ' when the animation is complete then stop the animator
            animationAmount = 0
            Exit Sub
        End If
    Else
        If SDeg < newSDeg Then
            SDeg = oldSDeg + animationAmount
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimator_Timer of Class Module cwOverlay"
End Sub

  


    
'---------------------------------------------------------------------------------------
' Procedure : calcAnimationData
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub calcAnimationData()
    Dim degDiff As Double: degDiff = 0
    Dim min  As Integer: min = -30
    Dim max  As Integer: max = 50
    
    On Error GoTo calcAnimationData_Error

    If mPointerAnimate = True Then
        If tmrAnimator.Enabled = False Then
        
            ' oldSDeg stored
            oldSDeg = SDeg
                    
            ' obtain the new temperature amount
        
            min = -30
            max = 50
            
            ' calculate the new cpu amount in degrees
            If weHaveData = True Then
                newSDeg = (TemperatureVal - min) / (max - min) * 360
            Else
                newSDeg = 30
            End If
                        
            If oldSDeg = newSDeg Then Exit Sub
            
            ' calculate the difference
            degDiff = newSDeg - oldSDeg
            
            ' if the result is negative animate CCW
            If degDiff < 0 Then
                animateCCW = True
            Else
                animateCCW = False
            End If
            
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
        Call obtainTemperature
    End If
    
    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

calcAnimationData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcAnimationData of Class Module cwOverlay"
End Sub

' ----------------------------------------------------------------
' Procedure Name: obtainTemperature
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 13/01/2024
' ----------------------------------------------------------------
Private Sub obtainTemperature()
        
    Dim min  As Integer: min = -30
    Dim max  As Integer: max = 50
    
    On Error GoTo obtainTemperature_Error

    min = -30
    max = 50
    
    If weHaveData = True Then
        SDeg = (TemperatureVal - min) / (max - min) * 360
    Else
        SDeg = 30
    End If
      
    W.Parent.Refresh
    
    On Error GoTo 0
    Exit Sub

obtainTemperature_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainTemperature, line " & Erl & "."

End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : tmrSampler_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrSampler_Timer()
    
    On Error GoTo tmrSampler_Timer_Error
    
    Call getData

   On Error GoTo 0
   Exit Sub

tmrSampler_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampler_Timer of Class Module cwOverlay"
End Sub




' ----------------------------------------------------------------
' Procedure Name: getData
' Purpose: this function obtains data from the chosen location using an httpXMLrequest. Requires a reference to Microsoft XML library.
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 26/01/2024
' ----------------------------------------------------------------
Private Sub getData()

    Dim location As String: location = vbNullString
    Dim thisURL As String: thisURL = vbNullString
    Dim thisFile As String: thisFile = vbNullString
    
    On Error GoTo getData_Error
    
    location = PzGIcao ' eg. "EGSH"
    thisURL = "https://aviationweather.gov/api/data/dataserver?requestType=retrieve&dataSource=metars&stationString=" & location & "&hoursBeforeNow=6&format=xml&mostRecent=true"
    thisFile = "E:\vb6\Panzer-Weather-Gauges-VB6\test.xml"
    Set myMSXML = CreateObject("Microsoft.XmlHttp")
    myMSXML.open "POST", thisURL, True
    'myMSXML.open "POST", thisFile, True
    myMSXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    myMSXML.setRequestHeader "User-Agent", "Firefox 3.6.4"
    myMSXML.send "param1=value2&param2=value2"
    
    tmrAsyncResponse.Enabled = True
    
    On Error GoTo 0
    Exit Sub

getData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getData, line " & Erl & "."

End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : tmrSampler_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrAsyncResponse_Timer()
    
    On Error GoTo tmrSampler_Timer_Error
    
    Debug.Print myMSXML.readyState
    If myMSXML.readyState = 4 Then ' asynchronous xml request complete
        
        If debugFlg = 1 Then Debug.Print myMSXML.responseText
        Call myStatusProc
        
    End If

   On Error GoTo 0
   Exit Sub

tmrSampler_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampler_Timer of Class Module cwOverlay"
End Sub



' ----------------------------------------------------------------
' Procedure Name: myStatusProc
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 27/01/2024
' ----------------------------------------------------------------
Private Sub myStatusProc()


    Dim humidityVal  As Long
    Dim pressureVal As Long
    Dim TheDate As Date: TheDate = #1/1/2000 12:00:00 PM#
    Dim secsDif As Integer: secsDif = 0
    Dim difString As String: difString = vbNullString
    Dim returnedString As String: returnedString = vbNullString
    Dim data1 As String: data1 = vbNullString
    Dim num_results As Integer: num_results = 0
    Dim observation_time As String: observation_time = vbNullString
    Dim station_id As String: station_id = vbNullString
    Dim temp_c As String: temp_c = vbNullString
    Dim altim_in_hg As String: altim_in_hg = vbNullString
    Dim dewpoint_c As String: dewpoint_c = vbNullString
    Dim wind_dir_degrees As String: wind_dir_degrees = vbNullString
    Dim wind_speed_kt As String: wind_speed_kt = vbNullString
    Dim Latitude As String: Latitude = vbNullString
    Dim Longitude As String: Longitude = vbNullString
    Dim wx_string As String: wx_string = vbNullString
    Dim precip_in As String: precip_in = vbNullString
    Dim visibility_statute_mi As String: visibility_statute_mi = vbNullString
    Dim metar_items As String: metar_items = vbNullString
    Dim current_node As String: current_node = vbNullString
    Dim sky_condition_items() As String
    Dim sky_condition As String: sky_condition = vbNullString
    Dim attributeCounter As Integer
    Dim SkyConditionCounter As Integer
    Dim locationDayNight As Boolean
    
    Dim tmpSky_condition As String: tmpSky_condition = vbNullString
    Dim humidity As Long: humidity = 0
    Dim skyClarity As String: skyClarity = vbNullString
    
    Dim cloud_base_mt_agl As Integer
    Dim baseCloud As Integer
    Dim skyClarityString As String: skyClarityString = vbNullString
    
    Dim cloud_base_ft_agl() As String
    Dim sky_cover() As String
    Dim sky_condition_items_length As Integer
    Dim sky_condition_attributes_length As Integer
    Dim skyNodeName As String: skyNodeName = vbNullString
    Dim cloudNodeName As String: cloudNodeName = vbNullString
    Dim skyNodeValue As String: skyNodeValue = vbNullString
    Dim cloudNodeValue As String: cloudNodeValue = vbNullString
    Dim uboundSkyConditions As Integer
    
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Dim objxmldoc As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    
    Dim MetarNode As MSXML2.IXMLDOMNode
    Dim SkyCondition As MSXML2.IXMLDOMNode
    Dim SkyConditions As MSXML2.IXMLDOMNodeList

    Dim answer As VbMsgBoxResult
    Dim answerMsg  As String: answerMsg = vbNullString
    Dim I As Integer: I = 0
    Dim isDay As Boolean: isDay = False
    
    Dim initialObsTime As String: initialObsTime = vbNullString
    Dim usableObsTime As Date: usableObsTime = #1/1/2000 12:00:00 PM#
    
    On Error GoTo myStatusProc_Error
    
    debugFlg = 1
    
    objxmldoc.async = True
    objxmldoc.LoadXML (myMSXML.responseText)
    
    ' get the values from the XML data response, the num results should be non-zero
    Set nodeList = objxmldoc.selectNodes("response/data/METAR")
    num_results = nodeList.length
    
    Set MetarNode = objxmldoc.selectSingleNode("response/data/METAR") ' There's only the one METAR node
    
    If num_results = 0 Then ' And Preferences.msgboxPref.value = "enabled" Then
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "The source weather feed is currently producing no valid data for " & panzerPrefs.sliSamplingInterval.Value & " secs."
            answer = msgBoxA(answerMsg, vbOKOnly, "Poll Warning", True, "myStatusProcPollingWarning")
        End If
        weHaveData = False
        Exit Sub ' Return
    End If

    If Not nodeList Is Nothing Then
         For Each node In nodeList
         
            On Error Resume Next ' prevents errors being generated from optionally present nodes.
           
           ' get the values from the XML data and return strings - the easy stuff first
            
            observation_time = node.selectSingleNode("observation_time").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - observation_time " + observation_time)
        
            station_id = node.selectSingleNode("station_id").Text
            If debugFlg = 1 Then Debug.Print "%myStatusProc - station_id " & station_id
            
            temp_c = Int(node.selectSingleNode("temp_c").Text)
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - temp_c " + temp_c)
        
            altim_in_hg = node.selectSingleNode("altim_in_hg").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - altim_in_hg " + altim_in_hg)
        
            dewpoint_c = Int(node.selectSingleNode("dewpoint_c").Text)
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - dewpoint_c " + dewpoint_c)
        
            wind_dir_degrees = node.selectSingleNode("wind_dir_degrees").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - wind_dir_degrees " + wind_dir_degrees)
        
            wind_speed_kt = node.selectSingleNode("wind_speed_kt").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - wind_speed_kt " + wind_speed_kt)
        
            Latitude = node.selectSingleNode("latitude").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - latitude " + Latitude)
        
            Longitude = node.selectSingleNode("longitude").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - longitude " + Longitude)
            
            visibility_statute_mi = node.selectSingleNode("visibility_statute_mi").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - visibility_statute_mi " + visibility_statute_mi)
        
            ' the On Error Resume Next above is for the next two optional items that may/may not appear in the returned XML
            
            wx_string = node.selectSingleNode("wx_string").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - wx_string " + wx_string)
        
            precip_in = node.selectSingleNode("precip_in").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - precip_in " + precip_in)
            
            'the first and lowest cloudbase is the one that really counts, there could be as many as three sky cover reading
            
            Set SkyConditions = MetarNode.selectNodes("sky_condition") ' Get all the sky_condition nodes under METAR
            If Not SkyConditions Is Nothing Then
                uboundSkyConditions = SkyConditions.length - 1
                ReDim sky_condition_items(uboundSkyConditions)
                ReDim sky_cover(uboundSkyConditions)
                ReDim cloud_base_ft_agl(uboundSkyConditions)
                SkyConditionCounter = 0
                
                For Each SkyCondition In SkyConditions
                    sky_condition_attributes_length = SkyCondition.Attributes.length  ' no of Attributes
                    If Not sky_condition_attributes_length = 0 Then
                        attributeCounter = 0 ' sky_cover
                        skyNodeName = SkyCondition.Attributes(attributeCounter).nodeName
                        skyNodeValue = SkyCondition.Attributes(attributeCounter).nodeValue
                            
                        sky_condition_items(SkyConditionCounter) = skyNodeValue
                        sky_cover(SkyConditionCounter) = skyNodeValue
                        
                        attributeCounter = 1 ' cloud_base_ft_agl
                        cloudNodeName = SkyCondition.Attributes(attributeCounter).nodeName
                        cloudNodeValue = SkyCondition.Attributes(attributeCounter).nodeValue
                        
                        cloud_base_ft_agl(SkyConditionCounter) = cloudNodeValue
                            
                    End If
                    If debugFlg = 1 Then Debug.Print ("%myStatusProc - sky_condition, sky_cover " + sky_condition_items(attributeCounter))
                    If debugFlg = 1 Then Debug.Print ("%myStatusProc - sky_condition, cloud_base_ft_agl " + cloud_base_ft_agl(attributeCounter))
                    SkyConditionCounter = SkyConditionCounter + 1
                Next
            End If
            On Error GoTo myStatusProc_Error
          Next node
   End If
               
   'Cleanup
   Set nodeList = Nothing
   Set SkyConditions = Nothing
    
    'if the station id returned is null then assume the weather information is missing for an unknown reason.
    If (station_id = "") Then
        If debugFlg = 1 Then Debug.Print ("%myStatusProc - station_id does not seem to providing any valid data, please select another.")
        
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "Rather weird - The supplied ICAO code " & panzerPrefs.txtIcao.Text & " does not seem to providing any valid data, please select another."""
            answer = msgBoxA(answerMsg, vbOKOnly, "Polling Warning", True, "myStatusProcStation_id")
            weHaveData = False
            Exit Sub ' return
        End If
    End If
    
    ' get pressure
    pressureVal = fGetPressure(altim_in_hg)

    ' set temperature
    TemperatureVal = Val(temp_c)

    ' calculate humidity
    humidityVal = fGetHumidity(temp_c, dewpoint_c)

    ' decode the visibility when a cloud cover field found
    skyClarity = get_cloud_cover(sky_condition_items())

    ' decode the base cloud from the upperbound of the cloud array but only if the array is populated
    If (Not cloud_base_ft_agl) = -1 Then baseCloud = cloud_base_ft_agl(UBound(cloud_base_ft_agl))
    
    ' add the base cloud to the visibility string
    skyClarityString = fAddBaseCloudToSkyClarity(baseCloud, skyClarity)
    
    ' get the current decimalTime and use it to test for sunrise or sunset
    sunriseSunset.Longitude = Longitude
    sunriseSunset.Latitude = Latitude
    isDay = sunriseSunset.isDayLight
    If debugFlg = 1 Then Debug.Print ("%myStatusProc - daylight " & isDay)

'    'wx_string = "+FG" 'testing string
'
    ' set the general weather conditions into the icon displayed
    Call determineWeatherConditionIcon(skyClarity, wx_string, sky_cover(), isDay)
'
'    ' get the observation time string
    initialObsTime = Left$(observation_time, 10) & " " & Mid$(observation_time, 12, 8)
    usableObsTime = CDate(initialObsTime)

'    'set the mini clock to show the observation time
'    theClock.displayTime(usableObsTime); ' we don't yet have a mini clock

    ' get the current time string and compare
    TheDate = Now
    secsDif = Int(DateDiff("s", TheDate, usableObsTime))

    ' format the interval for display in the tooltip
    difString = nice_format_interval(secsDif)

    ' store the last updated value in the settings
    'PzGLastUpdated = usableObsTime
    If (debugFlg = 1) Then Debug.Print ("%myStatusProc - difString " & difString)
    
'    humidityBackground.Tooltip = " "

'    ' set the hover over tooltips
'    setHoverTooltip(wind_dir_degrees, wind_speed_kt, precip_in, difString, wx_string)
'
'    setIconTooltip(wx_string, sky_cover)
'
'     busyStop()

    thisTemperature = temp_c
    
    weHaveData = True
    
    Call calcAnimationData
    
    On Error GoTo 0
    Exit Sub

myStatusProc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure myStatusProc, line " & Erl & "."

End Sub


' ----------------------------------------------------------------
' Procedure Name: nice_format_interval
' Purpose: this function provides a text version of when the data was last collected
' Procedure Kind: Function
' Procedure Access: Public
' Parameter duration (Integer):
' Return Type: String
' Author: beededea
' Date: 15/03/2024
' ----------------------------------------------------------------
Function nice_format_interval(duration As Integer) As String
    On Error GoTo nice_format_interval_Error
    Dim hours As Integer: hours = 0
    Dim minutes As Integer: minutes = 0
    Dim res As String: res = vbNullString
    
    duration = Abs(duration)
    hours = Int(duration / 3600)
    minutes = Int((duration - hours * 3600) / 60)
    
    If hours > 72 Then
        'print("%Update " & String(hours) & " hours and " & String(minutes) & " minutes ago");
        nice_format_interval = "No recent data available" & "!"
        Exit Function
    End If
    If hours > 0 Then
        res = "Updated"
        If hours = 1 Then
            res = res & " " & "one hour"
        Else
            res = res & " " & Str$(hours) & " " & "hours"
        End If
        If minutes = 1 Then
            res = res & " " & "and one minute"
        ElseIf minutes > 1 Then
            res = res & " " & "and" & " " & Str$(minutes) & " " & "minutes"
        End If
        nice_format_interval = res & " " & "ago"
        Exit Function
    End If
    If minutes > 0 Then
        If minutes = 1 Then
            nice_format_interval = "Updated one minute ago"
            Exit Function
        End If
        nice_format_interval = "Updated" & " " & Str$(minutes) & " " & "minutes ago"
        Exit Function
    End If
    If duration = 1 Then
        nice_format_interval = "Updated one second ago"
        Exit Function
    End If
    nice_format_interval = "Updated" & " " & Str$(duration) & " " & "seconds ago"
    Exit Function
    
    On Error GoTo 0
    Exit Function

nice_format_interval_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure nice_format_interval, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: get_cloud_cover
' Purpose: function to decode cloud cover information into a text form
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sky_cover (String):
' Return Type: String
' Author: beededea
' Date: 29/01/2024
'
' Decodes cloud cover information.
' Format is SKC or CLR for clear skies, or cccnnn where ccc = 3-letter
' code and nnn = altitude of cloud layer in hundreds of feet. 'VV' seems
' to be used for very low cloud layers. (Other conversion factor:
' 1 m = 3.28084 ft)
' FEW = 1 or 2 eighths cover; SCT = 3 or 4 eighths cover; BKN = 5, 6 or 7 eighths cover & OVC = 8/8 cover
' ----------------------------------------------------------------
Function get_cloud_cover(ByRef sky_condition_items() As String) As String
    Dim skyClarity: skyClarity = ""
   
    On Error GoTo get_cloud_cover_Error

    If (sky_condition_items(0) = "SKC") Then skyClarity = "clear_skies"
    If (sky_condition_items(0) = "CLR") Then skyClarity = "clear_skies"
    If (sky_condition_items(0) = "FEW") Then skyClarity = "partly_cloudy"
    If (sky_condition_items(0) = "SCT") Then skyClarity = "scattered_clouds"
    If (sky_condition_items(0) = "BKN") Then skyClarity = "mostly_cloudy"
    If (sky_condition_items(0) = "OVC") Then skyClarity = "overcast"
    If (sky_condition_items(0) = "VV ") Then skyClarity = "vertical_visibility"
    If (sky_condition_items(0) = "CAVOK ") Then skyClarity = "ceiling_and_visibility_OK"

'    (1) there are no clouds below 5000 feet above aerodrome level (AAL) or minimum sector altitude (whichever is higher) and no cumulonimbus or towering cumulus (2) visibility is at least 10 kilometres (6 statute miles) or more and (3) no current or forecast significant weather such as precipitation, thunderstorms, shallow fog or low drifting snow

    If (Not sky_condition_items) = -1 Then
        If (sky_condition_items(1) = "FEW" And Not sky_condition_items(0) = sky_condition_items(1)) Then skyClarity = skyClarity & " + " & "partly cloudy"
        If (sky_condition_items(1) = "SCT" And Not sky_condition_items(0) = sky_condition_items(1)) Then skyClarity = skyClarity & " + " & "scattered clouds"
        If (sky_condition_items(1) = "BKN" And Not sky_condition_items(0) = sky_condition_items(1)) Then skyClarity = skyClarity & " + " & "mostly cloudy"
        If (sky_condition_items(1) = "OVC" And Not sky_condition_items(0) = sky_condition_items(1)) Then skyClarity = skyClarity & " + " & "overcast"
    End If
    If (debugFlg = 1) Then Debug.Print ("%get_cloud_cover - part " & skyClarity)
    get_cloud_cover = skyClarity 'return skyClarity
    
    On Error GoTo 0
    Exit Function

get_cloud_cover_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure get_cloud_cover, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: fGetPressure
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter altim_in_hg (String):
' Return Type: Integer
' Author: beededea
' Date: 29/01/2024
' ----------------------------------------------------------------
Function fGetPressure(altim_in_hg As String) As Integer

    On Error GoTo fGetPressure_Error
    Dim vals0 As Integer
    Dim pressureVal As Double
    Dim strPressureVal As String
    
    pressureVal = Round(CDbl(altim_in_hg) * 25.3999)
    strPressureVal = Str$(pressureVal) & "(" & CStr(Round(1.3333 * pressureVal)) & ")"
    
    If debugFlg = 1 Then Debug.Print ("fGetPressure - Pressure = " & strPressureVal)

    fGetPressure = pressureVal
    
    
    On Error GoTo 0
    Exit Function

fGetPressure_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetPressure, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: fGetHumidity
' Purpose: calculate humidity
' Procedure Kind: Function
' Procedure Access: Public
' Parameter temp_c ():
' Parameter dewpoint_c ():
' Author: beededea
' Date: 29/01/2024
' ----------------------------------------------------------------
Function fGetHumidity(temp_c, dewpoint_c)

    Dim humidity As Long

    On Error GoTo fGetHumidity_Error
    
    humidity = Round(100 * (((112 - (0.1 * CInt(temp_c)) + CInt(dewpoint_c)) / (112 + (0.9 * CInt(temp_c)))) ^ 8))
    If debugFlg = 1 Then Debug.Print ("fGetHumidity - humidity " + CStr(humidity))
    
    fGetHumidity = humidity
    
    On Error GoTo 0
    Exit Function

fGetHumidity_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetHumidity, line " & Erl & "."

End Function



    '
' ----------------------------------------------------------------
' Procedure Name: fAddBaseCloudToSkyClarity
' Purpose: add the base cloud to the visibility string
' Procedure Kind: Function
' Procedure Access: Public
' Parameter baseCloud (Integer):
' Parameter skyClarity (String):
' Author: beededea
' Date: 30/01/2024
' ----------------------------------------------------------------
Function fAddBaseCloudToSkyClarity(ByVal baseCloud As Integer, ByVal skyClarity As String)
    Dim skyClarityString As String
    
    On Error GoTo fAddBaseCloudToSkyClarity_Error

    If (baseCloud <> 0) Then
'        If (panzerPrefs.imperialMetricPref.Value = "metric") Then
'            ' change the string to metric measurements
'            baseCloud = Int(baseCloud * 0.3048)
'            skyClarityString = skyClarity + " at " + baseCloud + " metres."
'        Else
            skyClarityString = skyClarity + " at " + CStr(baseCloud) + " ft."
        End If
'    End If
    If debugFlg = 1 Then Debug.Print ("fAddBaseCloudToSkyClarity - skyClarityString " + skyClarityString)

    fAddBaseCloudToSkyClarity = skyClarityString
    
    On Error GoTo 0
    Exit Function

fAddBaseCloudToSkyClarity_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fAddBaseCloudToSkyClarity, line " & Erl & "."

End Function
    

' ----------------------------------------------------------------
' Procedure Name: convertToFahrenheit
' Purpose: convert temps from C to F values here to match scale setting
' Procedure Kind: Function
' Procedure Access: Private
' Parameter t ():
' Author: beededea
' Date: 17/01/2024
' ----------------------------------------------------------------
Private Function convertToFahrenheit(T)
    Dim newT As Integer
    
    On Error GoTo convertToFahrenheit_Error
    
    newT = Round(1.8 * T + 32)
    convertToFahrenheit = newT
    
    On Error GoTo 0
    Exit Function

convertToFahrenheit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure convertToFahrenheit, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: convertToKelvin
' Purpose: convert temps from C to F values here to match scale setting
' Procedure Kind: Function
' Procedure Access: Private
' Parameter t ():
' Author: beededea
' Date: 17/01/2024
' ----------------------------------------------------------------
Private Function convertToKelvin(T) As Integer
    Dim newT As Double
    
    On Error GoTo convertToKelvin_Error
    
    Const KELVIN_CELSIUS_DIFF As Double = 273.15
    
    newT = KELVIN_CELSIUS_DIFF + T

    convertToKelvin = Fix(newT)
    
    On Error GoTo 0
    Exit Function

convertToKelvin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure convertToKelvin, line " & Erl & "."

End Function



'================
' this function sets the hover over tooltip
'================
Private Function setHoverTooltip(wind_dir_degrees, wind_speed_kt, precip_in, difString, wx_string)

            Dim force As String
            Dim tooltipString As String
            Dim TemperatureVal As Long
            
            'compassDirection = degToCompass(wind_dir_degrees)

            'force = beaufortConversion(wind_speed_kt)

            'tooltipString = " " & usableObsTime & " " & " " & PzGIcao & " "
            'tooltipString = tooltipString & icaoLocation1 & ", " & icaoLocation2 & ", " & icaoLocation3 & " " & vbCrLf

'            If (Preferences.barometerScalePref.Value = "millibars") Then
'                tooltipString = tooltipString & "Pressure" & ": " & Int(pressureValue * 1.33322) & " mb " & vbCrLf
'             ElseIf (Preferences.barometerScalePref.Value = "hg") Then
'                tooltipString = tooltipString & "Pressure" & ": " & Int(pressureValue / 25.399999704976) & " inHg " & vbCrLf
'             ElseIf (Preferences.barometerScalePref.Value = "hpa") Then
'                tooltipString = tooltipString & "Pressure" & ": " & Int(pressureValue * 1.33322) & " hpa " & vbCrLf
'             ElseIf (Preferences.barometerScalePref.Value = "mmhg") Then
'                tooltipString = tooltipString & "Pressure" & ": " & pressureValue & " mmHg " & vbCrLf
'            End If

'            If (Preferences.tempUnit.Value = "fahrenheit") Then
'                TemperatureVal = temperatureValue * 1.8 & 32
'                tooltipString = tooltipString & "Temperature" & ": " & Round(TemperatureVal).toFixed(2) & " F " & vbCrLf
'            End If
'            If (Preferences.tempUnit.Value = "celsius") Then
'                tooltipString = tooltipString & "Temperature" & ": " & Round(temperatureValue).toFixed(2) & " C " & vbCrLf
'            End If
'            If (Preferences.tempUnit.Value = "kelvin") Then
'                tooltipString = tooltipString & "Temperature" & ": " & Round(Int(temperatureValue & 273.15)).toFixed(2) & " K " & vbCrLf
'                tooltipString = tooltipString & "Humidity" & ": " & humidityValue & " % " & " " & skyClarityString & " " & vbCrLf
'            End If
'            If (Preferences.windUnit.Value = "knots") Then
'                    tooltipString = tooltipString & "wind_speed_is" & " " & wind_speed_kt & " knots, "
'            Else
'                tooltipString = tooltipString & "wind_speed_is" & " " & Int(wind_speed_kt * 0.51444) & " m/sec, "
'                tooltipString = tooltipString & "strength_force" & " " & force & ", " & "direction" & " " & compassDirection
'            End If
'            If (precip_in <> 0) Then
'               tooltipString = tooltipString & " " & "precipitation_is" & " " & precip_in & " inches"
'
'
'            Var wx_str = getMetarDescription(wx_string)    ' get the interpreted wx string data
'            If (wx_str <> "") Then
'               tooltipString = tooltipString & "  " & wx_str & "  " & vbCrLf
'
'            'weatherText.data = String(tooltipString)
'
'            humidityBackground.ToolTip = tooltipString
'
'            ' save the tooltip value so it can be used during the regular timed 60 second tooltip updates.
'            Preferences.lastTooltip.Value = humidityBackground.ToolTip
'            humidityBackground.ToolTip = humidityBackground.ToolTip & " " & difString
'            humidityBackground.ToolTip = humidityBackground.ToolTip & "  " & "Double_tap_on_me_to_get_new_weather"
'            thermometerBackground.ToolTip = humidityBackground.ToolTip
'            anemometerBackground.ToolTip = humidityBackground.ToolTip
'            weatherIconGaugeBackground.ToolTip = humidityBackground.ToolTip
'            barometerBackground.ToolTip = humidityBackground.ToolTip

'========
End Function
'========




' ----------------------------------------------------------------
' Procedure Name: determineWeatherConditionIcon
' Purpose: function to decode_weather when a wx field found and set the icon accordingly
' Procedure Kind: Function
' Procedure Access: Private
' Parameter thisSkyClarity (String):
' Parameter thisWxString (String):
' Parameter sky_cover (String):
' Parameter thisIsDay (Boolean):
' Return Type: String, ByVal thisIsDay As Boolean)
' Author: beededea
' Date: 14/03/2024
' ----------------------------------------------------------------
Private Function determineWeatherConditionIcon(ByVal thisSkyClarity As String, ByVal thisWxString As String, ByRef sky_cover() As String, ByVal thisIsDay As Boolean)
    On Error GoTo determineWeatherConditionIcon_Error
   'thisWxString = "+PRFG" 'testing string
   'sky_cover(0) = "BKN"

   Dim wSeverity As String: wSeverity = vbNullString
   Dim wModifier As String: wModifier = vbNullString
   Dim skycover1 As String: skycover1 = vbNullString
   Dim skycover2 As String: skycover2 = vbNullString
   Dim skycover3 As String: skycover3 = vbNullString
   Dim presentConditions As String: presentConditions = vbNullString
   Dim weatherIconSrc As String: weatherIconSrc = vbNullString
   Dim showersIconSrc As String: showersIconSrc = vbNullString
   Dim fogIconSrc As String: fogIconSrc = vbNullString
   Dim windIconSrc As String: windIconSrc = vbNullString
   Dim exoticIconSrc As String: exoticIconSrc = vbNullString
   Dim iconSrc As String: iconSrc = vbNullString
   Dim weatherType As String: weatherType = vbNullString
   Dim precedingSrc As String: precedingSrc = vbNullString
   Dim followingSrc As String: followingSrc = vbNullString

    
   If UBound(sky_cover()) >= 0 Then
        If (sky_cover(0) = "SKC") Then skycover1 = "SKC"
        If (sky_cover(0) = "CLR") Then skycover1 = "CLR"
        If (sky_cover(0) = "FEW") Then skycover1 = "FEW"
        If (sky_cover(0) = "SCT") Then skycover1 = "SCT"
        If (sky_cover(0) = "BKN") Then skycover1 = "BKN"
        If (sky_cover(0) = "OVC") Then skycover1 = "OVC"
        If (sky_cover(0) = "VV ") Then skycover1 = "VV "
        If (sky_cover(0) = "CAVOK") Then skycover1 = "CAVOK"
    End If
        
   If UBound(sky_cover()) >= 1 Then
        If (sky_cover(1) = "SKC") Then skycover2 = "SKC"
        If (sky_cover(1) = "CLR") Then skycover2 = "CLR"
        If (sky_cover(1) = "FEW") Then skycover2 = "FEW"
        If (sky_cover(1) = "SCT") Then skycover2 = "SCT"
        If (sky_cover(1) = "BKN") Then skycover2 = "BKN"
        If (sky_cover(1) = "OVC") Then skycover2 = "OVC"
        If (sky_cover(1) = "VV ") Then skycover2 = "VV "
        If (sky_cover(1) = "CAVOK") Then skycover2 = "CAVOK"
    End If
    
   If UBound(sky_cover()) >= 2 Then
        If (sky_cover(2) = "SKC") Then skycover3 = "SKC"
        If (sky_cover(2) = "CLR") Then skycover3 = "CLR"
        If (sky_cover(2) = "FEW") Then skycover3 = "FEW"
        If (sky_cover(2) = "SCT") Then skycover3 = "SCT"
        If (sky_cover(2) = "BKN") Then skycover3 = "BKN"
        If (sky_cover(2) = "OVC") Then skycover3 = "OVC"
        If (sky_cover(2) = "VV ") Then skycover3 = "VV "
        If (sky_cover(2) = "CAVOK") Then skycover3 = "CAVOK"
    End If

   If (debugFlg = 1) Then Debug.Print ("%determineWeatherCondition - sky_cover " & sky_cover(0))

    If InStr(thisWxString, "-") <> 0 Then
        wSeverity = "light"
    ElseIf InStr(thisWxString, "+") <> 0 Then
        wSeverity = "heavy"
    Else
        wSeverity = "medium"  ' moderate conditions have no descriptor
    End If

   If (debugFlg = 1) Then Debug.Print ("%determineWeatherCondition - thisWxString " & thisWxString)
   If (debugFlg = 1) Then Debug.Print ("%determineWeatherCondition - wSeverity " & wSeverity)
   If InStr(thisWxString, "UP") <> 0 Then presentConditions = presentConditions & "unknown "

   'in the absence of a weather code assume clear sky
   'var presentConditions = "clear and sunny"
   weatherIconSrc = "0cloud.png"

   'determine the weather type,
   If (InStr(thisWxString, "MI") <> 0) Then
     wModifier = wModifier & "shallow "
   End If
   If (InStr(thisWxString, "PR") <> 0) Then
     wModifier = wModifier & "partial "
   End If
   If (InStr(thisWxString, "DZ") <> 0) Then
     wModifier = wModifier & "drizzle "
   End If
   If (InStr(thisWxString, "BC") <> 0) Then
     wModifier = wModifier & "patches of "
   End If
   If (InStr(thisWxString, "DR") <> 0) Then
     wModifier = wModifier & "drifting "
   End If
   If (InStr(thisWxString, "BL") <> 0) Then
     wModifier = wModifier & "blowing "
   End If
   If (InStr(thisWxString, "SH") <> 0) Then
     wModifier = wModifier & "showers "
   End If
   If (InStr(thisWxString, "VC") <> 0) Then
     '"nearby "
     wModifier = wModifier & "nearby "
   End If
   If (InStr(thisWxString, "FZ") <> 0) Then
     '"freezing "
     wModifier = wModifier & "freezing "
   End If

   ' FEW = 1 or 2 eighths cover SCT = 3 or 4 eighths cover BKN = 5, 6 or 7 eighths cover & OVC = 8/8 cover

   'use the cloud code to determine cloud cover
    If (skycover1 = "SKC" Or skycover2 = "SKC" Or skycover3 = "SKC") Then thisSkyClarity = "Clear skies"       ' default clear
    If (skycover1 = "CLR" Or skycover2 = "CLR" Or skycover3 = "CLR") Then thisSkyClarity = "Clear skies"       ' default
    If (skycover1 = "VV " Or skycover2 = "VV " Or skycover3 = "VV ") Then thisSkyClarity = "Vertical visibility"          ' default
    If (skycover1 = "CAVOK" Or skycover2 = "CAVOK" Or skycover3 = "CAVOK") Then thisSkyClarity = "Ceiling and visibility OK"          ' default

    If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
      'Partly cloudy"
      weatherIconSrc = "1cloud_norain.png"
    End If
    If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
      'Scattered clouds"
      weatherIconSrc = "2cloud_norain.png"
    End If
    If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
      'Mostly cloudy"
      weatherIconSrc = "3cloud_norain.png"
    End If
    If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
      'Overcast"
      weatherIconSrc = "4cloud_norain.png"
    End If


   'type of weather - snow
   If (InStr(thisWxString, "SN") <> 0 Or InStr(thisWxString, "SG") <> 0) Then

       If (wSeverity = "light") Then
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
             'Scattered clouds"
             weatherIconSrc = "1cloud_lightsnow.png" 'default
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "2cloud_lightsnow.png"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Mostly cloudy"
            weatherIconSrc = "4cloud_lightsnow.png"
          End If
       End If
       If (wSeverity = "medium") Then
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_snow.png"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "2cloud_snow.png"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "3cloud_snow.png"
          End If
       End If
       If (wSeverity = "heavy") Then
            weatherIconSrc = "4cloud_heavysnow.png"
       End If
   End If

   ' no need to take the cloud into account, always cloudy with hail
   weatherType = "none"
   If (InStr(thisWxString, "IC") <> 0) Then
     '"ice crystals "
     weatherType = "hail"
   End If
   If (InStr(thisWxString, "PE") <> 0) Then
     '"ice pellets "
     weatherType = "hail"
   End If
   If (InStr(thisWxString, "GR") <> 0) Then
     '"hail "
     weatherType = "hail"
   End If
   If (InStr(thisWxString, "GS") <> 0) Then
     '"small hail "
     weatherType = "hail"
   End If ' and/or snow pellets

   'type of weather - hail
   If (weatherType = "hail") Then
       If (wSeverity = "light") Then
          weatherIconSrc = "2cloud_hail.png" 'default
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Mostly cloudy"
            weatherIconSrc = "4cloud_lighthail.png"
          End If
       End If
       If (wSeverity = "medium") Then
          If ((skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Or (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN")) Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_hail.png"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_hail.png"
          End If
       End If
       If (wSeverity = "heavy") Then
            weatherIconSrc = "4cloud_heavyhail.png"
       End If
   End If

    ' this caters for the event when there are showers indicated but no associated rain code
    ' it is strange but it quite often occurs...
    If (InStr(thisWxString, "SH") <> 0 And InStr(thisWxString, "SN") = 0 And InStr(thisWxString, "SG") = 0 And weatherType <> "hail") Then
         'if no RAin code nor hail or snow but showers indicated then just assume rain...
         showersIconSrc = "Resources/images/icons_metar/day/" & "showers.png"
     Else
        showersIconSrc = vbNullString
    End If

   'type of weather - rain
   If (InStr(thisWxString, "RA") <> 0) Then
       wSeverity = "light" ' default
     '"rain "
       If (wSeverity = "light") Then
          If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
            'Partly cloudy"
            weatherIconSrc = "1cloud_lightrain.png"
          End If
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_lightrain.png"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_lightrain.png"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "4cloud_lightrain.png"
          End If
       End If
       If (wSeverity = "medium") Then
          If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
            'Partly cloudy"
            weatherIconSrc = "1cloud_modrain.png"
          End If
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_modrain.png"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_modrain.png"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "4cloud_modrain.png"
          End If
       End If
       If (wSeverity = "heavy") Then
          If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
            'Partly cloudy"
            weatherIconSrc = "1cloud_heavyrain.png"
          End If
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_heavyrain.png"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_heavyrain.png"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "4cloud_heavyrain.png"
          End If
       End If

       If (InStr(thisWxString, "TS") <> 0) Then
            wModifier = wModifier & "thunderstorm "
            weatherIconSrc = "4cloud_heavyrain.png"
       End If
   End If

   fogIconSrc = vbNullString    'default
   If (InStr(thisWxString, "BR") <> 0 Or InStr(thisWxString, "FG") <> 0) Then
     '"fog " or "mist"
     If (wSeverity = "light") Then
          fogIconSrc = "Resources/images/icons_metar/day/" & "1_fog.png" 'night or day makes no difference
       End If
       If (wSeverity = "medium") Then
          fogIconSrc = "Resources/images/icons_metar/day/" & "2_fog.png"
       End If
       If (wSeverity = "heavy") Then
          fogIconSrc = "Resources/images/icons_metar/day/" & "3_fog.png"
       End If
   End If

   'windIconSrc= "Resources/images/icons_metar/night/windy03.png"   'default
   If (InStr(thisWxString, "SQ") <> 0) Then
     '"strong winds "
     windIconSrc = "Resources/images/icons_metar/day/" & "windy03.png"
   End If

   exoticIconSrc = vbNullString   'default
   'exotic types - we have no icons for any of these
   If (InStr(thisWxString, "FU") <> 0) Then
      exoticIconSrc = "Resources/images/icons_metar/day/" & "smoke.png"
      '"smoke "
   End If
   If (InStr(thisWxString, "VA") <> 0) Then
     '"volcanic ash "
      exoticIconSrc = "Resources/images/icons_metar/day/" & "volcano.png"
   End If
   If (InStr(thisWxString, "DU") <> 0) Then
     '"widespread dust "endif
      exoticIconSrc = "Resources/images/icons_metar/day/" & "dust.png"
   End If
   If (InStr(thisWxString, "SA") <> 0) Then
     '"sand "endif
      exoticIconSrc = "Resources/images/icons_metar/day/" & "sand.png"
   End If
   If (InStr(thisWxString, "HZ") <> 0) Then
     '"haze "endif
      exoticIconSrc = "Resources/images/icons_metar/day/" & "haze.png"
   End If
   If (InStr(thisWxString, "PY") <> 0) Then
     '"spray "endif
      exoticIconSrc = "Resources/images/icons_metar/day/" & "spray.png"
   End If
   If (InStr(thisWxString, "PO") <> 0) Then
     '"dustdevils "endif
      exoticIconSrc = "Resources/images/icons_metar/day/" & "dustdevil.png"
   End If
   If (InStr(thisWxString, "FC") <> 0) Then
     '"tornado "endif
      exoticIconSrc = "Resources/images/icons_metar/day/" & "tornado.png"
   End If
   If (InStr(thisWxString, "SS") <> 0) Then
     '"sandstorm/duststorm "endif
      exoticIconSrc = "Resources/images/icons_metar/day/" & "duststorm.png"
   End If

    If (weatherIconSrc <> vbNullString) Then
      If (thisIsDay = True) Then
         iconSrc = "Resources/images/icons_metar/day/" & weatherIconSrc
         weatherIconSrc = iconSrc
         precedingSrc = iconSrc
         followingSrc = iconSrc
      Else
         iconSrc = "Resources/images/icons_metar/night/n_" & weatherIconSrc
         weatherIconSrc = iconSrc
         precedingSrc = iconSrc
         followingSrc = iconSrc
      End If
     Else
      weatherIconSrc = "Resources/images/globe.png"
    End If
    If (debugFlg = 1) Then Debug.Print ("%determineWeatherCondition - iconSrc " & iconSrc)

    On Error GoTo 0
    Exit Function

determineWeatherConditionIcon_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure determineWeatherConditionIcon, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: getMetarDescription
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter wx_string (String):
' Return Type: String
' Author: beededea
' Date: 15/03/2024
' ----------------------------------------------------------------
Function getMetarDescription(wx_string As String) As String
    On Error GoTo getMetarDescription_Error
    Dim presentConditions As String: presentConditions = ""
    Dim wSeverity As String: wSeverity = ""
    Dim wModifier As String: wModifier = ""
    
    If InStr(wx_string, "UP") <> 0 Then
        presentConditions = presentConditions + "unknown "
    End If
    'in the absence of a weather code assume clear sky
    'Dim presentConditions As String = "clear and sunny"
    'weatherVal = "0cloud.png" <<<<
    'determine the strength of the weather type, light, medium, heavy or very heavy
    
    If InStr(wx_string, "MI") <> 0 Then
        wModifier = wModifier + "shallow"
    End If
    If InStr(wx_string, "PR") <> 0 Then
        wModifier = wModifier + "partial"
    End If
    If InStr(wx_string, "DZ") <> 0 Then
        wModifier = wModifier + "drizzle"
    End If
    If InStr(wx_string, "BC") <> 0 Then
        wModifier = wModifier + "patches of"
    End If
    If InStr(wx_string, "DR") <> 0 Then
        wModifier = wModifier + "drifting"
    End If
    If InStr(wx_string, "BL") <> 0 Then
        wModifier = wModifier + "blowing"
    End If
    If InStr(wx_string, "SH") <> 0 Then
        wModifier = wModifier + "showers"
    End If
    If InStr(wx_string, "VC") <> 0 Then
        'presentConditions = presentConditions +  "nearby "
        wModifier = wModifier + "in_the_vicinity"
    End If
    If InStr(wx_string, "FZ") <> 0 Then
        'presentConditions = presentConditions +  "freezing "
        wModifier = wModifier + "freezing"
    End If
    If InStr(wx_string, "TS") <> 0 Then
        wModifier = wModifier + "thunderstorm"
    End If
    If wModifier <> "" Then
        presentConditions = wModifier
    End If
    
    If InStr(wx_string, "-") <> 0 Then
        wSeverity = "light" & " "
    ElseIf InStr(wx_string, "+") <> 0 Then
        wSeverity = "heavy" & " "
    Else
        wSeverity = "" ' moderate conditions have no descriptor
    End If
    
    If InStr(wx_string, "-") = 0 Or InStr(wx_string, "+") = 0 Then
        presentConditions = presentConditions + wSeverity
        If debugFlg = 1 Then
            Debug.Print ("%getMetarDescription - wSeverity " & wSeverity)
        End If
    End If
    
    If InStr(wx_string, "RA") <> 0 Then
        presentConditions = presentConditions + "rain" & " "
    End If
    
    If InStr(wx_string, "SN") <> 0 Then
        presentConditions = presentConditions + "snow" & " "
    End If
    
    If InStr(wx_string, "SG") <> 0 Then
        presentConditions = presentConditions + "snow_grains" & " "
    End If
    
    If InStr(wx_string, "IC") <> 0 Then
        presentConditions = presentConditions + "ice_crystals" & " "
    End If
    
    If InStr(wx_string, "PE") <> 0 Then
        presentConditions = presentConditions + "ice_pellets" & " "
    End If
    
    If InStr(wx_string, "GR") <> 0 Then
        presentConditions = presentConditions + "hail" & " "
    End If
    
    If InStr(wx_string, "GS") <> 0 Then
        presentConditions = presentConditions + "small_hail" & " "
    End If
    
    If InStr(wx_string, "UP") <> 0 Then
        presentConditions = presentConditions + "unknown" & " "
    End If
    
    If InStr(wx_string, "BR") <> 0 Then
        presentConditions = presentConditions + "mist" & " "
    End If
    
    If InStr(wx_string, "FG") <> 0 Then
        presentConditions = presentConditions + "fog" & " "
    End If
    
    If InStr(wx_string, "FU") <> 0 Then
        presentConditions = presentConditions + "smoke" & " "
    End If
    
    If InStr(wx_string, "VA") <> 0 Then
        presentConditions = presentConditions + "volcanic_ash" & " "
    End If
    
    If InStr(wx_string, "DU") <> 0 Then
        presentConditions = presentConditions + "widespread_dust" & " "
    End If
    
    If InStr(wx_string, "SA") <> 0 Then
        presentConditions = presentConditions + "sand" & " "
    End If
    
    If InStr(wx_string, "HZ") <> 0 Then
        presentConditions = presentConditions + "haze" & " "
    End If
    
    If InStr(wx_string, "PY") <> 0 Then
        presentConditions = presentConditions + "spray" & " "
    End If
    
    If InStr(wx_string, "PO") <> 0 Then
        presentConditions = presentConditions + "dustdevils" & " "
    End If
    
    If InStr(wx_string, "SQ") <> 0 Then
        presentConditions = presentConditions + "strong_winds" & " "
    End If
    
    If InStr(wx_string, "FC") <> 0 Then
        presentConditions = presentConditions + "tornado" & " "
    End If
    
    If InStr(wx_string, "SS") <> 0 Then
        presentConditions = presentConditions + "sandstorm_duststorm" & " "
    End If
    
    If debugFlg = 1 And presentConditions <> "" Then
        Debug.Print ("%getMetarDescription - presentConditions " & presentConditions)
    End If

    
    On Error GoTo 0
    Exit Function

getMetarDescription_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getMetarDescription, line " & Erl & "."

End Function
