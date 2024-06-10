VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlayTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwoverlaytemp
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfTemperature
'---------------------------------------------------------------------------------------

' Microsoft XML, v3.0
' c:/windows/SysWow64/msxml3.dll

Option Explicit

Public WithEvents weatherM As cwWeatherMeteo
Attribute weatherM.VB_VarHelpID = -1

'these Public Properties are set in the Parent (here, cfTemperature)
Public cCollOverlayTempImages As cCollection

Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrAnimator As cTimer
Attribute tmrAnimator.VB_VarHelpID = -1

Private SDeg As Double
'Private thisTemperature As Long

Private oldSDeg As Double
Private newSDeg As Double
Private animationQuanta As Double
Private animationAmount As Double
Private animationDivisions As Integer
Private animationInterval As Integer
Private animateCCW As Boolean

Private mPointerAnimate As Boolean              ' property PointerAnimate
Private mLocked As Boolean                      ' property Locked
Private mHidden As Boolean                      ' property Hidden
Private mThisOpacity As String                  ' property ThisOpacity
Private mThisTemperature As Integer             ' property mThisTemperature
Private mThisFace As Integer                    ' property mThisFace


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
                    
'    Set tmrSampler = Nothing
'    Set tmrSampler = New_c.Timer(100, mTicking) ' first interval is 1/10th of a second, don't change it, triggers an immediate run. Correct sampling interval read later
    
    animationDivisions = 12
    animationInterval = 40
    
    Set tmrAnimator = Nothing
    Set tmrAnimator = New_c.Timer(animationInterval, False)
 
'    Set tmrAsyncResponse = Nothing
'    Set tmrAsyncResponse = New_c.Timer(100, False)
    
'    Set tmrPressureDrop = Nothing
'    Set tmrPressureDrop = New_c.Timer(Val(PzGStormTestInterval) * 1000, False) ' one hour
    
'    Set tmrSleepTester = Nothing
'    Set tmrSleepTester = New_c.Timer(2500, True)  ' two seconds, should start automatically but it doesn't
'
    
'    tmrPressureDrop.Enabled = True  '
'    tmrSleepTester.Enabled = True   ' for some reason I have to explicitly manually start the sleep timer

'    icaoLocation1 = vbNullString
'    icaoLocation2 = vbNullString
'    icaoLocation3 = vbNullString
'    icaoLocation4 = vbNullString
'    icaoLocation5 = vbNullString
'
'    icaoToFind = PzGIcao ' eg. "EGSH"
'
'    fullIcaoLocation = searchIcaoFile(icaoToFind, "ICAO", icaoLocation1, icaoLocation2, icaoLocation3, icaoLocation4, icaoLocation5, searchCount)
'    mIcaoLocation = icaoLocation1
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwoverlaytemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwoverlaytemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwoverlaytemp"
End Property


'---------------------------------------------------------------------------------------
' Property : thisTemperature
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisTemperature() As Integer
   On Error GoTo thisTemperatureGet_Error

   thisTemperature = mThisTemperature

   On Error GoTo 0
   Exit Property

thisTemperatureGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisTemperature of Class Module cwoverlaytemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisTemperature
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisTemperature(ByVal newValue As Integer)
   On Error GoTo thisTemperatureLet_Error

   If mThisTemperature <> newValue Then mThisTemperature = newValue Else Exit Property
   
   Call obtainTemperature

   On Error GoTo 0
   Exit Property

thisTemperatureLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisTemperature of Class Module cwOverlayHumid"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwoverlaytemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwoverlaytemp"
End Property

'---------------------------------------------------------------------------------------
' Property : thisSensor
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
    
    If panzerPrefs.IsVisible = True Then
        panzerPrefs.cmbTemperatureScale.ListIndex = newValue
    End If
    
    PzGTemperatureScale = mThisFace
    sPutINISetting "Software\PzTemperatureGauge", "temperatureScale", PzGTemperatureScale, PzGSettingsFile
    
    Call obtainTemperature
    
    On Error GoTo 0
    Exit Property

thisFaceLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisFace of Class Module cwoverlaytemp"
End Property

''---------------------------------------------------------------------------------------
'' Property : samplingInterval
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get samplingInterval() As Long
'   On Error GoTo samplingIntervalGet_Error
'
'    samplingInterval = mSamplingInterval
'
'   On Error GoTo 0
'   Exit Property
'
'samplingIntervalGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwoverlaytemp"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : samplingInterval
'' Author    :
'' Date      : 09/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let samplingInterval(ByVal newValue As Long)
'    On Error GoTo samplingIntervalLet_Error
'
'    If mSamplingInterval <> newValue Then mSamplingInterval = newValue Else Exit Property
'
'    If tmrSampler.Interval = 100 Then WeatherMeteo.GetMetar = True ' allows first early run, interval is then reset
'
'    tmrSampler.Interval = Val(mSamplingInterval) * 1000
'
'    On Error GoTo 0
'    Exit Property
'
'samplingIntervalLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwoverlaytemp"
'End Property
'---------------------------------------------------------------------------------------
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwoverlaytemp"
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
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlayTemp"
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
        menuForm.mnuLockTemperatureGauge.Checked = True
        If panzerPrefs.IsVisible = True Then
            panzerPrefs.cmbGaugeType.ListIndex = 0
            panzerPrefs.chkPreventDragging.Value = 1
        End If
        
        PzGPreventDraggingTemperature = "1"
        If aspectRatio = "landscape" Then
            PzGTemperatureLandscapeLocked = "1"
            panzerPrefs.cmbLandscapeLocked.ListIndex = 1
        Else
            PzGTemperaturePortraitLocked = "1"
            panzerPrefs.cmbPortraitLocked.ListIndex = 1
       End If
        
        fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    Else
        menuForm.mnuLockTemperatureGauge.Checked = False
        If panzerPrefs.IsVisible = True Then
            panzerPrefs.cmbGaugeType.ListIndex = 0
            panzerPrefs.chkPreventDragging.Value = 0
        End If
        
        PzGPreventDraggingTemperature = "0"
        If aspectRatio = "landscape" Then
            PzGTemperatureLandscapeLocked = "0"
            panzerPrefs.cmbLandscapeLocked.ListIndex = 0
        Else
            PzGTemperaturePortraitLocked = "0"
            panzerPrefs.cmbPortraitLocked.ListIndex = 0
        End If
        
        fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(PzGOpacity) / 100
    End If

    fTemperature.temperatureGaugeForm.Refresh

    sPutINISetting "Software\PzTemperatureGauge", "preventDraggingTemperature", PzGPreventDraggingTemperature, PzGSettingsFile

    If PzGEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayTemp"
End Property

'
''---------------------------------------------------------------------------------------
'' Property : WeatherDetails
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get WeatherDetails() As String
'   On Error GoTo WeatherDetailsGet_Error
'
'   WeatherDetails = mWeatherDetails
'
'   On Error GoTo 0
'   Exit Property
'
'WeatherDetailsGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WeatherDetails of Class Module cwOverlayTemp"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : WeatherDetails
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let WeatherDetails(ByVal newValue As String)
'   On Error GoTo WeatherDetailsLet_Error
'
'   If mWeatherDetails <> newValue Then mWeatherDetails = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'WeatherDetailsLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WeatherDetails of Class Module cwOverlayTemp"
'End Property
'
''---------------------------------------------------------------------------------------
'' Property : WindSpeed
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get windSpeed() As String
'   On Error GoTo WindSpeedGet_Error
'
'   windSpeed = mWindSpeed
'
'   On Error GoTo 0
'   Exit Property
'
'WindSpeedGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WindSpeed of Class Module cwOverlayTemp"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : WindSpeed
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let windSpeed(ByVal newValue As String)
'   On Error GoTo WindSpeedLet_Error
'
'   If mWindSpeed <> newValue Then mWindSpeed = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'WindSpeedLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WindSpeed of Class Module cwOverlayTemp"
'End Property





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

''---------------------------------------------------------------------------------------
'' Property  : IcaoToTest
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get IcaoToTest() As String
'
'    On Error GoTo IcaoToTestGet_Error
'
'    IcaoToTest = mIcaoToTest
'
'   On Error GoTo 0
'   Exit Property
'
'IcaoToTestGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IcaoToTest of Class Module cwOverlayTemp"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : IcaoToTest
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 10/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let IcaoToTest(ByVal newValue As String)
'
'    On Error GoTo IcaoToTestLet_Error
'
'    If mIcaoToTest <> newValue Then mIcaoToTest = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'IcaoToTestLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IcaoToTest of Class Module cwOverlayTemp"
'End Property
'
''---------------------------------------------------------------------------------------
'' Property  : StringToTest
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get StringToTest() As String
'
'    On Error GoTo StringToTestGet_Error
'
'    StringToTest = mStringToTest
'
'   On Error GoTo 0
'   Exit Property
'
'StringToTestGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property StringToTest of Class Module cwOverlayTemp"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : StringToTest
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 10/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let StringToTest(ByVal newValue As String)
'
'    On Error GoTo StringToTestLet_Error
'
'    If mStringToTest <> newValue Then mStringToTest = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'StringToTestLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property StringToTest of Class Module cwOverlayTemp"
'End Property

''---------------------------------------------------------------------------------------
'' Property  : ValidICAO
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get ValidICAO() As Boolean
'    Dim icaoToFind As String: icaoToFind = vbNullString
'    Dim fullIcaoLocation As String: fullIcaoLocation = vbNullString
'    Dim searchArray() As String
'    Dim searchCount As Long: searchCount = 0
'
'    On Error GoTo ValidICAOGet_Error
'
'    icaoToFind = mIcaoToTest ' eg. "EGSH"
'    fullIcaoLocation = searchIcaoFile(icaoToFind, "ICAO", icaoLocation1, icaoLocation2, icaoLocation3, icaoLocation4, icaoLocation5, searchCount)
'
'    mIcaoLocation = icaoLocation1 ' the airport name
'
'    If icaoLocation5 <> vbNullString Then ' icao code
'        mValidICAO = True
'        ValidICAO = mValidICAO
'        mIcaoToTest = icaoLocation5 ' find the icao code and set the private var so it can be accessed via a property
'        fSelector.SelectorForm.Widgets("sbtnselect").Widget.Enabled = True
'    Else
'        ValidICAO = False
'        mValidICAO = False
'        fSelector.SelectorForm.Widgets("sbtnselect").Widget.Enabled = False
'
'    End If
'
'   On Error GoTo 0
'   Exit Property
'
'ValidICAOGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ValidICAO of Class Module cwOverlayTemp"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : ValidICAO
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 10/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let ValidICAO(ByVal newValue As Boolean)
'
'    On Error GoTo ValidICAOLet_Error
'
'    If mValidICAO <> newValue Then mValidICAO = newValue Else Exit Property
'
'
'   On Error GoTo 0
'   Exit Property
'
'ValidICAOLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ValidICAO of Class Module cwOverlayTemp"
'End Property
'
'
''---------------------------------------------------------------------------------------
'' Property  : ValidLocationCount
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get ValidLocationCount() As Long ' returns the number of matching locations
'    Dim icaoToFind As String: icaoToFind = vbNullString
'    Dim fullIcaoLocation As String: fullIcaoLocation = vbNullString
'    Dim searchCount As Long: searchCount = 0
'
'    On Error GoTo ValidLocationCountGet_Error
'
'    icaoToFind = mStringToTest ' eg. "EGSH"
'    fullIcaoLocation = searchIcaoFile(icaoToFind, "Location", icaoLocation1, icaoLocation2, icaoLocation3, icaoLocation4, icaoLocation5, searchCount)
'
'    'mIcaoLocation = icaoLocation1 ' the airport name
'
' '   If icaoLocation5 <> vbNullString Then ' icao code
'
'        mValidLocationCount = searchCount ' returns a count only
'        ValidLocationCount = mValidLocationCount
'
'        If mValidLocationCount = 0 Then fSelector.SelectorForm.Widgets("sbtnselect").Widget.Enabled = False
'
'
'        'mIcaoToTest = icaoLocation5 ' find the icao code and set the private var so it can be accessed via a property
''    Else
''        ValidLocationCount = False
''        mValidLocationCount = False
''    End If
'
'   On Error GoTo 0
'   Exit Property
'
'ValidLocationCountGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ValidLocationCount of Class Module cwOverlayTemp"
'End Property
''
''---------------------------------------------------------------------------------------
'' Property  : icaoLocation
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 17/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Get icaoLocation() As String
'
'    On Error GoTo icaoLocationGet_Error
'
'    icaoLocation = mIcaoLocation
'
'   On Error GoTo 0
'   Exit Property
'
'icaoLocationGet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property icaoLocation of Class Module cwOverlayTemp"
'End Property
''---------------------------------------------------------------------------------------
'' Property  : icaoLocation
'' Author    : Dean Beedell (yereverluvinunclebert)
'' Date      : 10/05/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Property Let icaoLocation(ByVal newValue As String)
'
'    On Error GoTo icaoLocationLet_Error
'
'    If mIcaoLocation <> newValue Then mIcaoLocation = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'icaoLocationLet_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property icaoLocation of Class Module cwOverlayTemp"
'End Property
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
    Dim txtThisTemperature As String: txtThisTemperature = vbNullString
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
    
    txtThisTemperature = CStr(Int(mThisTemperature))

    For i = 0 To cCollOverlayTempImages.Count - 1 ' this is the ExcludedPaths collection populated in cfTemperature by duplicating collTemperaturePSDNonUIElements.
      arr = cCollOverlayTempImages.ItemByIndex(i)
      key = cCollOverlayTempImages.KeyByIndex(i)
      
      If IsArray(arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case key
                 
                 Case "centigradeface"
                    
                    If mThisFace = 0 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                 
                 Case "fahrenheitface"
                    
                    If mThisFace = 1 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                 
                 Case "kelvinface"
                    
                    If mThisFace = 2 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
        
                Case "redlamptrue"
                    
                    If mThisTemperature >= 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "redlampfalse"
                    
                    If mThisTemperature < 34 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                                        
                 Case "bluelamptrue"
                    
                    If mThisTemperature <= 0 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "bluelampfalse"
                    
                    If mThisTemperature > 0 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                        CC.Paint arr(2), Pat ' render the Pattern
                    End If
                    

                 Case "secondshadow", "secondhand"
                    
                     If mThisFace = 0 Then
                         scaleString = "C"
                         txtThisTemperature = CStr(Int(mThisTemperature))
                         
                     ElseIf mThisFace = 1 Then
                    
                         scaleString = "F"
                         txtThisTemperature = CStr(convertToFahrenheit(Int(mThisTemperature)))
                    
                     ElseIf mThisFace = "2" Then
                         scaleString = "K"
                         txtThisTemperature = CStr(convertToKelvin(Int(mThisTemperature)))
                     End If
                     
                    If Val(txtThisTemperature) >= 100 Then
                        boxLeft = 260
                    Else
                        boxLeft = 265
                    End If
                    
                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText boxLeft, 103, 200, 200, txtThisTemperature, , , 2, True, , someOpacity
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
     
    W.Refresh

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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimator_Timer of Class Module cwOverlayTemp"
End Sub

  



' ----------------------------------------------------------------
' Procedure Name: obtainTemperature
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 13/01/2024
' ----------------------------------------------------------------
Private Sub obtainTemperature()
    
    Dim tempPercent As Double: tempPercent = 0
    
    On Error GoTo obtainTemperature_Error

    ' obtain the new temperature in percentage
    tempPercent = ((WeatherMeteo.TemperatureValue + 30) / 80) * 100
    
    ' calculate the new cpu amount in degrees
    If WeatherMeteo.wehavedata = True Then
        SDeg = (tempPercent * 3) + 30
    Else
        SDeg = 30
    End If
      
    W.Parent.Refresh
    
    On Error GoTo 0
    Exit Sub

obtainTemperature_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainTemperature, line " & Erl & "."

End Sub

    

'
' ----------------------------------------------------------------
' Procedure Name: convertToFahrenheit
' Purpose: convert temps from C to F values here to match scale setting
' Procedure Kind: Function
' Procedure Access: Private
' Parameter t ():
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 17/01/2024
' ----------------------------------------------------------------
Private Function convertToFahrenheit(ByVal T As Integer) As Integer
    Dim newT As Integer

    On Error GoTo convertToFahrenheit_Error

    newT = Round(1.8 * T + 32)

    convertToFahrenheit = newT ' return

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
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 17/01/2024
' ----------------------------------------------------------------
Private Function convertToKelvin(ByVal T As Integer) As Integer
    Dim newT As Double

    On Error GoTo convertToKelvin_Error

    Const KELVIN_CELSIUS_DIFF As Double = 273.15

    newT = KELVIN_CELSIUS_DIFF + T

    convertToKelvin = Fix(newT) ' return

    On Error GoTo 0
    Exit Function

convertToKelvin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure convertToKelvin, line " & Erl & "."

End Function
'
'


Private Sub instance_onChange(ByVal wehavedata As Boolean)
    MsgBox "hello"
End Sub


