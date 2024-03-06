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
' Procedure : postProcessing
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub postProcessing()
    Dim degDiff As Double: degDiff = 0
    Dim min  As Integer: min = -30
    Dim max  As Integer: max = 50
    
    On Error GoTo postProcessing_Error

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

postProcessing_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure postProcessing of Class Module cwOverlay"
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
    Dim secsDif As Integer
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
    Dim sky_condition As String
    Dim attributeCounter As Integer
    Dim SkyConditionCounter As Integer
    Dim locationDayNight As Boolean
    
    Dim tmpSky_condition As String
    Dim humidity As Long
    Dim skyClarity As String
    
    Dim cloud_base_mt_agl As Integer
    Dim baseCloud As Integer
    Dim skyClarityString As String
    
    Dim cloud_base_ft_agl() As String
    Dim sky_cover() As String
    Dim sky_condition_items_length As Integer
    Dim sky_condition_attributes_length As Integer
    Dim skyNodeName As String
    Dim cloudNodeName As String
    Dim skyNodeValue As String
    Dim cloudNodeValue As String
    Dim uboundSkyConditions As Integer
    
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Dim objxmldoc As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    
    Dim MetarNode As MSXML2.IXMLDOMNode
    Dim SkyCondition As MSXML2.IXMLDOMNode
    Dim SkyConditions As MSXML2.IXMLDOMNodeList

    Dim answer As VbMsgBoxResult
    Dim answerMsg  As String
    Dim I As Integer
    
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
                ReDim cloud_base_ft_agl(uboundSkyConditions)
                SkyConditionCounter = 0
                
                For Each SkyCondition In SkyConditions
                    sky_condition_attributes_length = SkyCondition.Attributes.length  ' no of Attributes
                    If Not sky_condition_attributes_length = 0 Then
                        attributeCounter = 0 ' sky_cover
                        skyNodeName = SkyCondition.Attributes(attributeCounter).nodeName
                        skyNodeValue = SkyCondition.Attributes(attributeCounter).nodeValue
                            
                        sky_condition_items(SkyConditionCounter) = skyNodeValue
                        
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

    
'    ' get the current decimalTime and use it to test for sunrise or sunset
    'locationDayNight = SunriseSunset(TheDate.getYear(), TheDate.getMonth(), TheDate.getDay(), Latitude, Longitude)
    
    'mLongitude = 1.283
    'mLatitude = 52.676
    ' we will obtain these from the feed
    sunriseSunset.Longitude = Longitude
    sunriseSunset.Latitude = Latitude
    Dim isDay As Boolean
    isDay = sunriseSunset.isDayLight
    ', Longitude)
    
'    dim decimalTime = theDate.getHours() + ( theDate.getMinutes()/ 60)
'    if debugFlg = 1 then debug.print ("%myStatusProc - decimalTime "+ decimalTime)

'    dim isDay = locationDayNight.isDaylight(decimalTime)
'    if debugFlg = 1 then debug.print("%myStatusProc - daylight " + isDay)
'
'    'wx_string = "+FG" 'testing string
'
'    ' set the general weather conditions into the icon displayed
'    determineWeatherConditionIcon(wx_string, sky_cover, isDay)
'
'    ' get the observation time string
'    dim initialObsTime = observation_time.substr(0,10) + " " + observation_time.substr(11,8)
'    usableObsTime = stringToDate(initialObsTime)


    thisTemperature = temp_c
    
    weHaveData = True
    
    Call postProcessing
    
    On Error GoTo 0
    Exit Sub

myStatusProc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure myStatusProc, line " & Erl & "."

End Sub


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
    
    newT = Math.Round(1.8 * T + 32)
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
