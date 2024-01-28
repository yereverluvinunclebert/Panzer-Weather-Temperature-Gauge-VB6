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
Private WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1
Private WithEvents tmrXML As cTimer
Attribute tmrXML.VB_VarHelpID = -1

Private SDeg As Double
Private temperaturePercent As Long
Private temperatureTotalTimeLeft As Single
'Private TemperatureDetails As String
Private temperatureVal  As Long

'------------------------------------------------------ ENDS

Private mTicking As Boolean ' property Ticking
Private mSmoothSecondHand As Boolean ' property SmoothSecondHand
Private mLocked As Boolean ' property Locked
Private mTemperatureDetails As String ' property
Private mOnACPower As Boolean ' property
Private mHidden As Boolean ' property Hidden
Private mThisOpacity As String ' property ThisOpacity
Private mSamplingInterval As String ' property SamplingInterval
Private mThisFace As Integer             ' property mThisFace

Private myMSXML As MSXML2.XMLHTTP

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Dim intervalValue As Integer: intervalValue = 0
    
    On Error GoTo Class_Initialize_Error

    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    intervalValue = 5000

    Set tmrClock = Nothing
    Set tmrClock = New_c.Timer(100, mTicking)

    Set tmrXML = Nothing
    Set tmrXML = New_c.Timer(100, False)
    
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
'        If mSmoothSecondHand Then
'            'intervalValue = 100
'        Else
'            'intervalValue = 500
'        End If

        tmrClock.Enabled = True
    Else
        tmrClock.Enabled = False
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
    
    Call obtainTemperature
    
    On Error GoTo 0
    Exit Property

samplingIntervalLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : SmoothSecondHand
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SmoothSecondHand() As Boolean
   On Error GoTo SmoothSecondHandGet_Error

    SmoothSecondHand = mSmoothSecondHand

   On Error GoTo 0
   Exit Property

SmoothSecondHandGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SmoothSecondHand of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : SmoothSecondHand
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SmoothSecondHand(ByVal newValue As Boolean)
    Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo SmoothSecondHandLet_Error

    If mSmoothSecondHand <> newValue Then mSmoothSecondHand = newValue Else Exit Property
    
    If mSmoothSecondHand Then
        intervalValue = 100
    Else
        intervalValue = 500
    End If
            
    tmrClock.Interval = intervalValue
    
   On Error GoTo 0
   Exit Property

SmoothSecondHandLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SmoothSecondHand of Class Module cwOverlay"
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
' Property : OnACpower
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get OnACpower() As String
   On Error GoTo OnACpowerGet_Error

   OnACpower = mOnACPower

   On Error GoTo 0
   Exit Property

OnACpowerGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property OnACpower of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : OnACpower
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let OnACpower(ByVal newValue As String)
   On Error GoTo OnACpowerLet_Error

   If mOnACPower <> newValue Then mOnACPower = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

OnACpowerLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property OnACpower of Class Module cwOverlay"
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
    Dim i As Long: i = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtTemperaturetotal As String: txtTemperaturetotal = vbNullString
    Dim txtTemperaturePercent As String: txtTemperaturePercent = vbNullString
    Dim textDrawn As Boolean:   textDrawn = False
    Dim someOpacity As Double: someOpacity = 0
    Dim cream As String: cream = vbNullString
    
    cream = RGB(235, 217, 179) ' cannot create a constant from an RGB value
   
    On Error GoTo W_Paint_Error
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 17
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    someOpacity = (Val(mThisOpacity) / 100) * 0.75
    
    txtTemperaturePercent = CStr(Int(temperaturePercent))
    txtTemperaturetotal = CStr(Int(temperatureTotalTimeLeft))

    For i = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfTemperature by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(i)
      Key = cCollOverlayImages.KeyByIndex(i)
      
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
                    
                    If temperaturePercent <= 20 And mOnACPower = False Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "redlampfalse"
                    
                    If temperaturePercent > 20 Or mOnACPower = True Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                        CC.Paint Arr(2), Pat ' render the Pattern
                    End If
                                        
                 Case "greenlamptrue"
                    
                    If temperaturePercent <= 20 And mOnACPower = False Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "greenlampfalse"
                    
                    If temperaturePercent > 20 Or mOnACPower = True Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                        CC.Paint Arr(2), Pat ' render the Pattern
                    End If
                    

                 Case "secondshadow", "secondhand"
                    
                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText 265, 103, 200, 200, txtTemperaturePercent, , , 2, True, , someOpacity
                    CC.DrawText 455, 297, 200, 200, txtTemperaturetotal, , , 2, True, , someOpacity
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
' Procedure : tmrClock_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrClock_Timer()
    
    On Error GoTo tmrClock_Timer_Error
    
    tmrClock.Interval = Val(PzGSamplingInterval) * 1000

'    If mSmoothSecondHand = True Then

'    Else

        Call obtainTemperature
        
'    End If
    


   On Error GoTo 0
   Exit Sub

tmrClock_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
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
        
    Dim min  As Integer: min = -50
    Dim max  As Integer: max = 70
    Dim angle As Integer: angle = 0
    
    On Error GoTo obtainTemperature_Error

    Call getData
    
    If (PzGTemperatureScale = "1") Then
        temperatureVal = temperatureVal * 1.8 + 32
        min = -45
        max = 135

        If debugFlg = 1 Then Debug.Print "%setTemp - " & CStr(temperatureVal) & " F"
        
    Else

        If debugFlg = 1 Then Debug.Print ("%setTemp - " & CStr(temperatureVal) & " C")
    
    End If

    SDeg = (temperatureVal - min) / (max - min) * 360
    
    'SDeg = (temperaturePercent * 3) + 30

    mTemperatureDetails = mTemperatureDetails & vbCrLf
        
    
    W.Parent.Refresh
    
    On Error GoTo 0
    Exit Sub

obtainTemperature_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainTemperature, line " & Erl & "."

End Sub







' ----------------------------------------------------------------
' Procedure Name: getData
' Purpose: this function gets data from the chosen location  Requires a reference to Microsoft XML library
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 26/01/2024
' ----------------------------------------------------------------
Private Sub getData()

    Dim location As String: location = vbNullString
    Dim thisURL As String: thisURL = vbNullString
    
    On Error GoTo getData_Error
    
    location = PzGIcao '"EGSH"
    thisURL = "https://aviationweather.gov/api/data/dataserver?requestType=retrieve&dataSource=metars&stationString=" & location & "&hoursBeforeNow=6&format=xml&mostRecent=true"

    Set myMSXML = CreateObject("Microsoft.XmlHttp")
    myMSXML.open "POST", thisURL, True
    myMSXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    myMSXML.setRequestHeader "User-Agent", "Firefox 3.6.4"
    myMSXML.send "param1=value2&param2=value2"
    'MsgBox myMSXML.responseText
    tmrXML.Enabled = True
    
    On Error GoTo 0
    Exit Sub

getData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getData, line " & Erl & "."

End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : tmrXML_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrXML_Timer()
    
    On Error GoTo tmrXML_Timer_Error
    
    Debug.Print myMSXML.readyState
    If myMSXML.readyState = 4 Then ' asynchronous xml request complete

        tmrXML.Enabled = False
        
        'If debugFlg = 1 Then
'        Debug.Print myMSXML.responseText
        Call myStatusProc
        
    End If

   On Error GoTo 0
   Exit Sub

tmrXML_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
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

    Dim vals0 As Integer
    Dim pressureVal As Long
    Dim humidityVal  As Long
    
    Dim theDate As Date: theDate = #1/1/2000 12:00:00 PM#
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
    Dim latitude As String: latitude = vbNullString
    Dim longitude As String: longitude = vbNullString
    Dim wx_string As String: wx_string = vbNullString
    Dim precip_in As String: precip_in = vbNullString
    Dim visibility_statute_mi As String: visibility_statute_mi = vbNullString
    Dim metar_items As String: metar_items = vbNullString
    Dim current_node As String: current_node = vbNullString
    Dim sky_condition_items(3) As String
    Dim sky_condition As String
    Dim counter As Integer
    Dim tmpSky_condition As String
    Dim strPressureVal As String
    Dim humidity As Long
    
    Dim cloud_base_ft_agl() As String
    Dim sky_cover() As String
    
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Dim objxmldoc As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    
    Dim answer As VbMsgBoxResult
    Dim answerMsg  As String
    
    
    On Error GoTo myStatusProc_Error
    
    debugFlg = 1
    
    objxmldoc.async = True
    objxmldoc.LoadXML (myMSXML.responseText)
    
    ' get the values from the XML data response, the num results should be non-zero
    Set nodeList = objxmldoc.selectNodes("response/data/METAR")
    num_results = nodeList.length
    
    If num_results = 0 Then ' And Preferences.msgboxPref.value = "enabled" Then
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "The source weather feed is currently producing no valid data for " & panzerPrefs.sliSamplingInterval.Value & " secs."
            answer = msgBoxA(answerMsg, vbOKOnly, "Poll Warning", True, "myStatusProcPollingWarning")
        End If
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
        
            latitude = node.selectSingleNode("latitude").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - latitude " + latitude)
        
            longitude = node.selectSingleNode("longitude").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - longitude " + longitude)
            
            visibility_statute_mi = node.selectSingleNode("visibility_statute_mi").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - visibility_statute_mi " + visibility_statute_mi)
        
            ' the On Error Resume Next above is for the next two optional items that may/may not appear in the returned XML
            
            wx_string = node.selectSingleNode("wx_string").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - wx_string " + wx_string)
        
            precip_in = node.selectSingleNode("precip_in").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - precip_in " + precip_in)
            
            'the first and lowest cloudbase is the one that really counts, there could be as many as three sky cover readings
            
            sky_condition_items(counter) = node.selectSingleNode("sky_condition").Attributes.getNamedItem("sky_cover").Text
            If debugFlg = 1 Then Debug.Print ("%myStatusProc - sky_condition " + sky_condition_items(counter))
        
            'first attempt to test for another sky_cover attribute
            tmpSky_condition = node.selectSingleNode("sky_condition").Attributes.getNamedItem("sky_cover").Text
            If tmpSky_condition <> sky_condition_items(counter) Then
                If debugFlg = 1 Then Debug.Print ("%myStatusProc - sky_condition " + tmpSky_condition)
                counter = counter + 1
                sky_condition_items(counter) = tmpSky_condition
                cloud_base_ft_agl(counter) = node.selectSingleNode("sky_condition").Attributes.getNamedItem("cloud_base_ft_agl").Text
                
            End If
            
            'second attempt to test for another sky_cover attribute
            tmpSky_condition = node.selectSingleNode("sky_condition").Attributes.getNamedItem("sky_cover").Text
            If tmpSky_condition <> sky_condition_items(counter) Then
                If debugFlg = 1 Then Debug.Print ("%myStatusProc - sky_condition " + tmpSky_condition)
                counter = counter + 1
                sky_condition_items(counter) = tmpSky_condition
                cloud_base_ft_agl(counter) = node.selectSingleNode("sky_condition").Attributes.getNamedItem("cloud_base_ft_agl").Text
            End If
            
            tmpSky_condition = node.selectSingleNode("sky_condition").Attributes.getNamedItem("sky_cover").Text
            If tmpSky_condition <> sky_condition_items(counter) Then
                If debugFlg = 1 Then Debug.Print ("%myStatusProc - sky_condition " + tmpSky_condition)
                counter = counter + 1
                sky_condition_items(counter) = tmpSky_condition
                cloud_base_ft_agl(counter) = node.selectSingleNode("sky_condition").Attributes.getNamedItem("cloud_base_ft_agl").Text
            End If
            
            On Error GoTo myStatusProc_Error
            
         Next node
    End If
      
    'if the station id returned is null then assume the weather information is missing for an unknown reason.
    If (station_id = "") Then
        If debugFlg = 1 Then Debug.Print ("%myStatusProc - station_id does not seem to providing any valid data, please select another.")
        
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "Rather weird - The supplied ICAO code " & panzerPrefs.txtIcao.Text & " does not seem to providing any valid data, please select another."""
            answer = msgBoxA(answerMsg, vbOKOnly, "Polling Warning", True, "myStatusProcStation_id")
        End If
        
    End If

    ' get pressure
    pressureVal = altim_in_hg * 25.3999
    vals0 = Val(pressureVal)
    pressureVal = Math.Round(vals0)
    strPressureVal = Str$(pressureVal) & "(" & CStr(Round(1.3333 * vals0)) & ")"
    If debugFlg = 1 Then Debug.Print ("%myStatusProc - Pressure = " & strPressureVal)
'
'    ' set temperature
    temperatureVal = Val(temp_c)
'
'    'calculate humidity
    humidity = Round(100 * (((112 - (0.1 * temp_c) + dewpoint_c) / (112 + (0.9 * temp_c))) ^ 8))
    If debugFlg = 1 Then Debug.Print ("%myStatusProc - humidity " + CStr(humidity))
    humidityVal = humidity
'
'    'decode the visibility when a cloud cover field found
'    skyClarity = get_cloud_cover(sky_cover)
'    If (cloud_base_ft_agl <> 0) Then
'        if (preferences.imperialMetricPref.value = "metric") then
'            ' change the string to metric measurements
'            dim cloud_base_mt_agl = int(cloud_base_ft_agl * 0.3048 )
'            skyClarityString = skyClarity + " at " + cloud_base_mt_agl + " metres."
'         else {
'            skyClarityString = skyClarity + " at " + cloud_base_ft_agl + " ft."
'        End If
'    End If
'    if debugFlg = 1 then debug.print("%myStatusProc - skyClarityString " + skyClarityString)
'
'    ' get the current decimalTime and use it to test for sunrise or sunset
'    dim locationDayNight = new SunriseSunset(theDate.getYear(), theDate.getMonth(), theDate.getDay(), latitude , longitude  )
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


    
    On Error GoTo 0
    Exit Sub

myStatusProc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure myStatusProc, line " & Erl & "."

End Sub
