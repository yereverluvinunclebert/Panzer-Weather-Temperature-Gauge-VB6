VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwWeatherMeteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


Public WithEvents tmrPressureDrop As cTimer
Attribute tmrPressureDrop.VB_VarHelpID = -1
Public WithEvents tmrAsyncResponse As cTimer
Attribute tmrAsyncResponse.VB_VarHelpID = -1
Public WithEvents tmrSleepTester As cTimer
Attribute tmrSleepTester.VB_VarHelpID = -1
Public WithEvents tmrSampler As cTimer
Attribute tmrSampler.VB_VarHelpID = -1

Private mTestPressureDrop As Long       ' property
Private mProcessMetar As Boolean            ' property ProcessMetar
Private mSamplerEnabled As Boolean      ' property GetMetar
Private mSamplingInterval As Long    ' property GetMetar
Private mValidICAO As Boolean             ' property
Private mIcaoToTest As String             ' property
Private mStringToTest As String
Private mIcaoLocation As String             ' property
Private mValidLocationCount As Long             ' property
Private mWindSpeed As Long             ' property
Private mWeatherDetails As String ' property
Private mWeHaveData As Boolean ' property
Private mTicking As Boolean ' property Ticking
Private mTemperatureValue As Single ' property Ticking
Private mPrecipitationWasFalling As Boolean

Public Event OnChange(ByVal wehavedata As Boolean)


Private XMLHttpRequest As MSXML2.XMLHTTP
'Private weHaveData As Boolean

Private icaoLocation1 As String
Private icaoLocation2 As String
Private icaoLocation3 As String
Private icaoLocation4 As String
Private icaoLocation5 As String

Private strTimeThen As Date

'Private PrecipitationWasFalling As Boolean ' was there any precipitation prior to current poll


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    
    Dim icaoToFind As String: icaoToFind = vbNullString
    Dim fullIcaoLocation As String: fullIcaoLocation = vbNullString
    Dim searchArray() As String
    Dim searchCount As Long: searchCount = 0
    
    On Error GoTo Class_Initialize_Error

    strTimeThen = Now()
    'PrecipitationWasFalling = False
    
    ' these initialisations need to be moved elsewhere
    'Dim sunriseSunset As cwSunriseSunset   'uninitialised declaration X
    Set sunriseSunset = New cwSunriseSunset 'initialised declaration
    
    Set tmrSampler = Nothing
    Set tmrSampler = New_c.Timer(5000, True) ' first interval is 1 second, don't change it, triggers an immediate run. Correct sampling interval read later
    tmrSampler.Enabled = True
   
    Set tmrPressureDrop = Nothing
    Set tmrPressureDrop = New_c.Timer(Val(gblStormTestInterval) * 1000, False) ' one hour
    tmrPressureDrop.Enabled = True
    
    Set tmrAsyncResponse = Nothing
    Set tmrAsyncResponse = New_c.Timer(1000, False)
    
    Set tmrSleepTester = Nothing
    Set tmrSleepTester = New_c.Timer(2500, True)  ' two seconds, should start automatically but it doesn't
    
    icaoLocation1 = vbNullString
    icaoLocation2 = vbNullString
    icaoLocation3 = vbNullString
    icaoLocation4 = vbNullString
    icaoLocation5 = vbNullString
    
    icaoToFind = gblIcao ' eg. "EGSH"
    
    fullIcaoLocation = searchIcaoFile(icaoToFind, "ICAO", icaoLocation1, icaoLocation2, icaoLocation3, icaoLocation4, icaoLocation5, searchCount)
    mIcaoLocation = icaoLocation1
        
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwoverlaytemp"
End Sub

'---------------------------------------------------------------------------------------
' Property : samplingInterval
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get samplingInterval() As Long
   On Error GoTo samplingIntervalGet_Error

    samplingInterval = mSamplingInterval

   On Error GoTo 0
   Exit Property

samplingIntervalGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwoverlaytemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : samplingInterval
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let samplingInterval(ByVal newValue As Long)
    On Error GoTo samplingIntervalLet_Error

    If mSamplingInterval <> newValue Then mSamplingInterval = newValue Else Exit Property
    
   ' If tmrSampler.Interval = 5000 Then WeatherMeteo.GetMetar = True ' allows first early run, interval is then reset
   
   If mSamplingInterval = 0 Then mSamplingInterval = 5
    tmrSampler.Interval = Val(mSamplingInterval) * 1000
    
    On Error GoTo 0
    Exit Property

samplingIntervalLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwoverlaytemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : SamplerEnabled
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SamplerEnabled() As Boolean
   On Error GoTo SamplerEnabledGet_Error

   SamplerEnabled = mSamplerEnabled

   On Error GoTo 0
   Exit Property

SamplerEnabledGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SamplerEnabled of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : SamplerEnabled
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SamplerEnabled(ByVal newValue As Boolean)
   On Error GoTo SamplerEnabledLet_Error

    If mSamplerEnabled <> newValue Then mSamplerEnabled = newValue Else Exit Property
    
    If mSamplerEnabled = True Then
        tmrSampler.Enabled = True
    Else
        tmrSampler.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

SamplerEnabledLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SamplerEnabled of Class Module cwOverlayTemp"
End Property

'---------------------------------------------------------------------------------------
' Property  : PrecipitationWasFalling
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get PrecipitationWasFalling() As Boolean
   On Error GoTo PrecipitationWasFallingGet_Error

   PrecipitationWasFalling = mPrecipitationWasFalling

   On Error GoTo 0
   Exit Property

PrecipitationWasFallingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property PrecipitationWasFalling of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : PrecipitationWasFalling
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let PrecipitationWasFalling(ByVal newValue As Boolean)
   On Error GoTo PrecipitationWasFallingLet_Error

    If mPrecipitationWasFalling <> newValue Then mPrecipitationWasFalling = newValue Else Exit Property
    
    If mPrecipitationWasFalling = True Then
        tmrSampler.Enabled = True
    Else
        tmrSampler.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

PrecipitationWasFallingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property PrecipitationWasFalling of Class Module cwOverlayTemp"
End Property


'---------------------------------------------------------------------------------------
' Property  : TestPressureDrop
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TestPressureDrop() As Boolean
   On Error GoTo TestPressureDropGet_Error

   TestPressureDrop = mTestPressureDrop

   On Error GoTo 0
   Exit Property

TestPressureDropGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TestPressureDrop of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : TestPressureDrop
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let TestPressureDrop(ByVal newValue As Boolean)
   On Error GoTo TestPressureDropLet_Error

    If mTestPressureDrop <> newValue Then mTestPressureDrop = newValue Else Exit Property
    
    If mTestPressureDrop = True Then
        tmrPressureDrop.Enabled = True
    Else
        tmrPressureDrop.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

TestPressureDropLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TestPressureDrop of Class Module cwOverlayTemp"
End Property


'---------------------------------------------------------------------------------------
' Property  : ProcessMetar
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ProcessMetar(ByVal newValue As Boolean)

    On Error GoTo ProcessMetarLet_Error

    If mProcessMetar <> newValue Then mProcessMetar = newValue Else Exit Property

    Call getData

    mProcessMetar = False

   On Error GoTo 0
   Exit Property

ProcessMetarLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ProcessMetar of Class Module cwoverlaytemp"
End Property


'---------------------------------------------------------------------------------------
' Property : ProcessMetar
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ProcessMetar() As Boolean
   On Error GoTo ProcessMetarGet_Error

   ProcessMetar = mProcessMetar

   On Error GoTo 0
   Exit Property

ProcessMetarGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ProcessMetar of Class Module cwoverlaytemp"
End Property

'---------------------------------------------------------------------------------------
' Procedure : tmrPressureDrop_Timer
' Author    : beededea
' Date      : 23/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrPressureDrop_Timer()
   On Error GoTo tmrPressureDrop_Timer_Error

    Call TestPressureDropRoutine

   On Error GoTo 0
   Exit Sub

tmrPressureDrop_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPressureDrop_Timer of Class Module cwWeatherMeteo"
End Sub

    

' ----------------------------------------------------------------
' Procedure Name: TestPressureDropRoutine
' Purpose: checked via timer to light the barometer storm warning bulb
' Procedure Kind: Sub
' Procedure Access: Public
' Author: beededea
' Date: 22/04/2024
' ----------------------------------------------------------------
Sub TestPressureDropRoutine()

    On Error GoTo TestPressureDropRoutine_Error
    
    ' dangerous pressure drop leading to volatile weather
    '    store the last pressure from the previous poll
    '    store the time of the last poll
    '    compare the pressure drop, of great compare the two times
    '    if the timeframe is short then light lamp

    Dim oldPressureStorage As Double: oldPressureStorage = 0
    Dim currentPressureValue As Double: currentPressureValue = 0
    Dim oldPressureStorageDate As Date: oldPressureStorageDate = #1/1/2000 12:00:00 PM#
    Dim currentDate As Date: currentDate = #1/1/2000 12:00:00 PM#
    
    debugFlg = 0

    oldPressureStorage = gblOldPressureStorage
    currentPressureValue = gblCurrentPressureValue ' we read this from the prefs as it will be regularly written by the main getData timer
    oldPressureStorageDate = gblPressureStorageDate
    
    currentDate = Now
    
    ' 24 hectopascals/millibars in 24 hours.
    ' A storm system often falls at a rate of more than 1 millibar per hour.
    ' If the air pressure falls 24 mb (or more) in 24 hours, the system is called a bomb cyclone.

    If debugFlg = 1 Then
        Debug.Print ("oldPressureStorageDate " & oldPressureStorageDate)
        'Debug.Print ("currentdate - 3960 " & (currentDate - 3960))
    End If

    ' if the pressure drops by 1 millibar and it is a recent reading
    ' this caters for the computer waking up from a sleep or a restart

    If oldPressureStorage = 0 Then
        gblOldPressureStorage = currentPressureValue
        gblPressureStorageDate = Now    ' time in seconds
        Exit Sub
    End If
    
    If oldPressureStorage - currentPressureValue >= 1.33322 Then    ' one milibar drop < 1 hour - a storm is brewing
        If Int(DateDiff("s", oldPressureStorageDate, currentDate)) < 3600 Then

            Debug.Print ("pressure reading is a recent one, certainly within the last 60 mins")
            Debug.Print ("The lamp is lit as the pressure is dropping fast enough for a serious storm. The old pressure value was " & oldPressureStorage & " the new pressure value is " & currentPressureValue)
            
            overlayBaromWidget.pressureDrop = True
        Else
            Debug.Print ("Note: The pressure reading was taken more than 60 mins ago so it can be ignored")
            
            overlayBaromWidget.pressureDrop = False
        End If
    Else
        overlayBaromWidget.pressureDrop = False
    End If
    gblOldPressureStorage = currentPressureValue
    gblPressureStorageDate = Now    ' time in seconds
    
    On Error GoTo 0
    Exit Sub

TestPressureDropRoutine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TestPressureDropRoutine, line " & Erl & "."

End Sub
'========================
' function ends
'========================



' ----------------------------------------------------------------
' Procedure Name: getData
' Purpose: this function obtains data from the chosen icaoLocation using an httpXMLrequest. Requires a Project Reference to Microsoft XML library.
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 26/01/2024
' ----------------------------------------------------------------
Public Sub getData()

    Dim icaoToFind As String: icaoToFind = vbNullString
    Dim thisURL As String: thisURL = vbNullString
    
    On Error GoTo getData_Error
    
    icaoToFind = gblIcao ' eg. "EGSH"
    thisURL = "https://aviationweather.gov/api/data/dataserver?requestType=retrieve&dataSource=metars&stationString=" & icaoToFind & "&hoursBeforeNow=6&format=xml&mostRecent=true"
    
    Set XMLHttpRequest = CreateObject("Microsoft.XmlHttp")
    XMLHttpRequest.open "POST", thisURL, True
    XMLHttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XMLHttpRequest.setRequestHeader "User-Agent", "Firefox 3.6.4"
    XMLHttpRequest.send "param1=value2&param2=value2"
    
    tmrAsyncResponse.Enabled = True
    
    On Error GoTo 0
    Exit Sub

getData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getData, line " & Erl & "."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrAsyncResponse_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : to replicate VB6's missing asynch/wait functionality, this timer only runs subsequent to a data request being initiated.
'             running on a one second interval
'---------------------------------------------------------------------------------------
'
Private Sub tmrAsyncResponse_Timer()
    
    On Error GoTo tmrAsyncResponse_Timer_Error
    
    Debug.Print XMLHttpRequest.readyState

    If XMLHttpRequest.readyState = 4 Then ' asynchronous xml request complete
    
        ' as soon as the the ready state is complete we stop the response timer
        tmrAsyncResponse.Enabled = False
        
        ' now we process the found data
        Call processMetarData
        
        ' data processed, we remove the XML request and that will reset the readystate
        If Not XMLHttpRequest Is Nothing Then Set XMLHttpRequest = Nothing

    End If

   On Error GoTo 0
   Exit Sub

tmrAsyncResponse_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAsyncResponse_Timer of Class Module cwOverlayTemp"
End Sub

' ----------------------------------------------------------------
' Procedure Name: processMetarData
' Purpose: Processes the XML METAR data and assigns to vars
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 27/01/2024
' ----------------------------------------------------------------
Private Sub processMetarData()

    Dim HumidityValue  As Long: HumidityValue = 0
    Dim pressureValueMillibars As Double: pressureValueMillibars = 0
    Dim TheDate As Date: TheDate = #1/1/2000 12:00:00 PM#
    Dim secsDif As Double: secsDif = 0
    Dim lastUpdateString As String: lastUpdateString = vbNullString
    Dim num_results As Integer: num_results = 0
    Dim observation_time As String: observation_time = vbNullString
    
    Dim raw_text As String: raw_text = vbNullString
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
    Dim sky_condition_items() As String ' array
    Dim attributeCounter As Integer: attributeCounter = 0
    Dim SkyConditionCounter As Integer: SkyConditionCounter = 0
    Dim skyClarity As String: skyClarity = vbNullString
    Dim baseCloud As Integer: baseCloud = 0
    Dim skyClarityString As String: skyClarityString = vbNullString
    Dim cloud_base_ft_agl() As String ' array
    Dim sky_cover() As String ' array
    Dim sky_condition_attributes_length As Integer: sky_condition_attributes_length = 0
    Dim skyNodeValue As String: skyNodeValue = vbNullString
    Dim cloudNodeValue As String: cloudNodeValue = vbNullString
    Dim uboundSkyConditions As Integer: uboundSkyConditions = 0
    Dim visibilityDistance As String: visibilityDistance = vbNullString
    Dim precipitationDepth As String: precipitationDepth = vbNullString
    
    Dim nodeList As MSXML2.IXMLDOMNodeList
    
    Dim objxmldoc As MSXML2.DOMDocument
    Set objxmldoc = New MSXML2.DOMDocument
    
    Dim node As MSXML2.IXMLDOMNode
    
    Dim MetarNode As MSXML2.IXMLDOMNode
    Dim SkyCondition As MSXML2.IXMLDOMNode
    Dim SkyConditions As MSXML2.IXMLDOMNodeList

    Dim answer As VbMsgBoxResult
    Dim answerMsg  As String: answerMsg = vbNullString
    
    Dim isday As Boolean: isday = False
    Dim initialObsTime As String: initialObsTime = vbNullString
    Dim usableObsTime As Date: usableObsTime = #1/1/2000 12:00:00 PM#
    Dim weekDayDate As String: weekDayDate = vbNullString
    Dim thisToolTip As String: thisToolTip = vbNullString
    'Dim weatherIconToolTip As String: weatherIconToolTip = vbNullString
    
    Dim force As Integer: force = 0
    Dim windIconSrc As String: windIconSrc = vbNullString
    
    On Error GoTo processMetarData_Error
    
    debugFlg = 0
    
    ' pause the sampling timer whilst the data is processed
    SamplerEnabled = False ' property
    
    pressureValueMillibars = 0
    
    objxmldoc.async = True
    objxmldoc.LoadXML (XMLHttpRequest.responseText)
    
    ' get the values from the XML data response, the num results should be non-zero
    Set nodeList = objxmldoc.selectNodes("response/data/METAR")
    num_results = nodeList.length
    
    Set MetarNode = objxmldoc.selectSingleNode("response/data/METAR") ' There's only the one METAR node
    
    ' no results found, go on as normal using the sampling interval
    If num_results = 0 Then
        If Val(gblErrorInterval) <> 0 Then
            ' compare last gblLastUpdated to current date time and if that exceeds the error interval then raise an error.
            TheDate = Now
            secsDif = Int(DateDiff("s", CDate(gblLastUpdated), TheDate))
            If secsDif >= (Val(gblErrorInterval) * 3600) And firstPoll = False Then
                'if has just awoken from sleep then suppress the no data error message
                tmrSleepTester.Enabled = False
                Call tmrSleepTester_TimerCode
                If gblJustAwoken = True Then
                    gblJustAwoken = False
                Else
                    answerMsg = "1. The source weather feed has been producing no valid data for " & secsDif & " secs. gblJustAwoken = " & gblJustAwoken
                    answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Poll Warning", True, "processMetarDataPollingWarning")
                End If
            End If
        End If
        mWeHaveData = False
            
        ' set the valid data lamp to display dull unlit version
        overlayPictorialWidget.validData = True

        ' restart the sampling timer
        SamplerEnabled = True

        Exit Sub ' Return
    End If

    If Not nodeList Is Nothing Then
         For Each node In nodeList
         
            On Error Resume Next ' prevents errors being generated from 'optional' nodes not present.
           
           ' get the values from the XML data and return strings - the easy stuff first
            
            observation_time = node.selectSingleNode("observation_time").Text
            If debugFlg = 1 Then Debug.Print ("%processMetarData - observation_time " + observation_time)
            
            raw_text = node.selectSingleNode("raw_text").Text
            If debugFlg = 1 Then Debug.Print "%processMetarData - raw_text " & raw_text
            
            station_id = node.selectSingleNode("station_id").Text
            If debugFlg = 1 Then Debug.Print "%processMetarData - station_id " & station_id
            
            temp_c = Int(node.selectSingleNode("temp_c").Text)
            If debugFlg = 1 Then Debug.Print ("%processMetarData - temp_c " + temp_c)
        
            altim_in_hg = node.selectSingleNode("altim_in_hg").Text
            If debugFlg = 1 Then Debug.Print ("%processMetarData - altim_in_hg " + altim_in_hg)
        
            dewpoint_c = Int(node.selectSingleNode("dewpoint_c").Text)
            If debugFlg = 1 Then Debug.Print ("%processMetarData - dewpoint_c " + dewpoint_c)
        
            wind_dir_degrees = node.selectSingleNode("wind_dir_degrees").Text
            overlayAnemoWidget.windDirection = wind_dir_degrees
            If debugFlg = 1 Then Debug.Print ("%processMetarData - wind_dir_degrees " + wind_dir_degrees)
        
            wind_speed_kt = node.selectSingleNode("wind_speed_kt").Text
            overlayAnemoWidget.windSpeed = wind_speed_kt
            If debugFlg = 1 Then Debug.Print ("%processMetarData - wind_speed_kt " + wind_speed_kt)
        
            Latitude = node.selectSingleNode("latitude").Text
            If debugFlg = 1 Then Debug.Print ("%processMetarData - latitude " + Latitude)
        
            Longitude = node.selectSingleNode("longitude").Text
            If debugFlg = 1 Then Debug.Print ("%processMetarData - longitude " + Longitude)
            
            visibility_statute_mi = node.selectSingleNode("visibility_statute_mi").Text
            If debugFlg = 1 Then Debug.Print ("%processMetarData - visibility_statute_mi " + visibility_statute_mi)
        
            ' the On Error Resume Next above is for the next two optional items that may/may not appear in the returned XML
            
            wx_string = node.selectSingleNode("wx_string").Text
            If debugFlg = 1 Then Debug.Print ("%processMetarData - wx_string " + wx_string)
        
            precip_in = node.selectSingleNode("precip_in").Text
            If debugFlg = 1 Then Debug.Print ("%processMetarData - precip_in " + precip_in)
            
            'the first and lowest height cloudbase is the one that really counts but there could be as many as three sky cover reading
            'the highest and lowest need to be reported.
            
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
                        'skyNodeName = SkyCondition.Attributes(attributeCounter).nodeName
                        skyNodeValue = SkyCondition.Attributes(attributeCounter).nodeValue
                            
                        sky_condition_items(SkyConditionCounter) = skyNodeValue
                        sky_cover(SkyConditionCounter) = skyNodeValue
                        
                        attributeCounter = 1 ' cloud_base_ft_agl
                        'cloudNodeName = SkyCondition.Attributes(attributeCounter).nodeName
                        cloudNodeValue = SkyCondition.Attributes(attributeCounter).nodeValue
                        
                        cloud_base_ft_agl(SkyConditionCounter) = cloudNodeValue
                            
                    End If
                    If debugFlg = 1 Then Debug.Print ("%processMetarData - sky_condition, sky_cover " + sky_condition_items(SkyConditionCounter))
                    If debugFlg = 1 Then Debug.Print ("%processMetarData - sky_condition, cloud_base_ft_agl " + cloud_base_ft_agl(SkyConditionCounter))
                    SkyConditionCounter = SkyConditionCounter + 1
                Next
            End If
          Next node
   End If
   
   On Error GoTo processMetarData_Error ' restart error trapping
   
   'Cleanup
   Set nodeList = Nothing
   Set SkyConditions = Nothing
       
    'if the station id returned is null then assume the weather information is missing for an unknown reason.
    If (station_id = vbNullString) Then
        If debugFlg = 1 Then Debug.Print ("%processMetarData - station_id does not seem to providing any valid data, please select another.")
        
        If answer = vbNo Then
            answer = vbYes ' deanieboy prefs
            answerMsg = "Rather weird - The supplied ICAO code " & widgetPrefs.txtIcao.Text & " does not seem to providing any valid data, please select another."""
            answer = msgBoxA(answerMsg, vbOKOnly, "Polling Warning", True, "processMetarDataStation_id")
            mWeHaveData = False
            Exit Sub ' return
        End If
    End If
    
    ' get pressure
    pressureValueMillibars = fPressureMillibars(altim_in_hg)
    
    'write the current pressure value so that it can be made available to the storm checker timer, straight away.
    gblCurrentPressureValue = pressureValueMillibars
    overlayBaromWidget.thisPressure = pressureValueMillibars

    sPutINISetting "Software\PzTemperatureGauge", "currentPressureValue", gblCurrentPressureValue, gblSettingsFile

    ' set temperature
    mTemperatureValue = Val(temp_c)
    
    ' testing value
    'mTemperatureValue = 40

    ' calculate Humidity
    HumidityValue = fGetHumidity(temp_c, dewpoint_c)
    overlayHumidWidget.thisHumidity = HumidityValue
    
    ' determine the precipitation
    precipitationDepth = fGetPrecipitation(precip_in)

    ' decode the visibility types when a cloud cover field found
    skyClarity = getCloudCover(sky_condition_items(), uboundSkyConditions)

    ' add two of the base cloud values to the visibility string
    skyClarityString = fAddBaseCloudToSkyClarity(baseCloud, skyClarity, cloud_base_ft_agl())
    
    ' get the current decimalTime and use it to test for sunrise or sunset
    sunriseSunset.Longitude = Longitude
    sunriseSunset.Latitude = Latitude
    isday = sunriseSunset.isDayLight
    If debugFlg = 1 Then Debug.Print ("%processMetarData - daylight " & isday)
        
    ' determine the visibility distance
    visibilityDistance = fGetDistance(visibility_statute_mi, isday)

'    'wx_string = "+FG" 'testing string
'
    ' set the general weather conditions into the icon displayed
    Call determineWeatherConditionIcon(skyClarity, wx_string, sky_cover(), isday, uboundSkyConditions)

    ' get the observation time string, eg. 2024-01-28T23:50:00Z and make a standard VB6 date var
    initialObsTime = Left$(observation_time, 10) & " " & Mid$(observation_time, 12, 8)
    usableObsTime = CDate(initialObsTime)
    
    ' make a pretty date for the balloon tooltip
    weekDayDate = humanFriendlyDate(usableObsTime)

    'set the mini clock on the clipboard to show the observation time
    overlayClipbWidget.thisTime = usableObsTime
    
    'refresh the clipboard to show the just-set time
    fClipB.clipBForm.Refresh

    ' get the current time string and compare
    TheDate = Now
    secsDif = CDbl(DateDiff("s", TheDate, usableObsTime))

    ' format the interval for display in the tooltip
    lastUpdateString = niceFormatInterval(secsDif, TheDate)

    ' store the last updated value in the settings
    gblLastUpdated = usableObsTime
    sPutINISetting "Software\PzTemperatureGauge", "lastUpdated", gblLastUpdated, gblSettingsFile

    If (debugFlg = 1) Then Debug.Print ("%processMetarData - lastUpdateString " & lastUpdateString)
    
    ' convert the windspeed to the beaufort scale
    force = beaufortConversion(wind_speed_kt)
    
    ' determine the wind icon image using the beaufort scale
    windIconSrc = fWindIconSrc(force, isday)

    ' set the hover-over tooltips on the gauges
    thisToolTip = setHoverTooltip(altim_in_hg, force, precipitationDepth, visibilityDistance, skyClarityString, HumidityValue, mTemperatureValue, pressureValueMillibars, weekDayDate, wind_dir_degrees, wind_speed_kt, precip_in, wx_string)
    thisToolTip = thisToolTip & vbCrLf & raw_text
    
    ' pass the current tooltip to the clipboard widget
    fClipB.sLblMETARText.Caption = thisToolTip
    
    ' add the last updated time to the end of the gauge tooltip
    thisToolTip = thisToolTip & vbCrLf & vbCrLf & lastUpdateString

    ' ?????
'    weatherIconToolTip = setIconTooltip(wx_string, sky_condition_items(), uboundSkyConditions)
'    If (weatherIconToolTip = vbNullString) Then
'        weatherIconToolTip = thisTooltip
'    End If
    
    ' for setting the balloon tooltip using the WeatherDetails property
    mWeatherDetails = thisToolTip

    mTemperatureValue = temp_c
    mWeHaveData = True
    wehavedata = True
    firstPoll = False
    
    ' set the animation variables and call the routine to move the temperature gauge pointer
    overlayTemperatureWidget.thisTemperature = TemperatureValue
    
    ' restart the sampling timer now the data has been processed
    SamplerEnabled = True
        
    ' pass the wind image to the pictorial gauge
    overlayPictorialWidget.thisWindImage = windIconSrc
    
    ' refresh the pictorial weather gauge so that it displays the provided weather icon (the pic gauge has no timer refresh)
    fPictorial.pictorialGaugeForm.Refresh
    
    ' set the valid data lamp to green
    overlayPictorialWidget.validData = True
        
    On Error GoTo 0
    Exit Sub

processMetarData_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure processMetarData, line " & Erl & "."

End Sub


'---------------------------------------------------------------------------------------
' Procedure : tmrSleepTester_Timer
' Author    : beededea
' Date      : 21/04/2021
' Purpose   : timer that stores the last time to see if the system has just emerged from sleep
' if the current time is greater than the last time stored by more than 30 seconds we can assume the system
' has been sent to sleep, if the two are significantly different then we reorganise the dock
'---------------------------------------------------------------------------------------
'
Private Sub tmrSleepTester_Timer()

    On Error GoTo tmrSleepTester_Timer_Error

    Call tmrSleepTester_TimerCode
    
    On Error GoTo 0
    Exit Sub

tmrSleepTester_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSleepTester_Timer of Form dock"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrSleepTester_TimerCode
' Author    : beededea
' Date      : 21/04/2021
' Purpose   : timer that stores the last time to see if the system has just emerged from sleep
' if the current time is greater than the last time stored by more than 30 seconds we can assume the system
' has been sent to sleep, if the two are significantly different then we reorganise the dock
'---------------------------------------------------------------------------------------
'
Private Sub tmrSleepTester_TimerCode()
    Dim strTimeNow As Date: strTimeNow = #1/1/2000 12:00:00 PM#  'set a variable to compare for the NOW time
    Dim lngSecondsGap As Double: lngSecondsGap = 0  ' set a variable for the difference in time
    
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo tmrSleepTester_TimerCode_Error
    
    tmrSleepTester.Enabled = False
    strTimeNow = Now()
    
    lngSecondsGap = DateDiff("s", strTimeThen, strTimeNow)
            
    If lngSecondsGap > 60 Then
        gblJustAwoken = True
        
'        answerMsg = "2. Weather widget has awoken from sleep after " & lngSecondsGap & " seconds. gblJustAwoken = " & gblJustAwoken
'        answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Awoken from sleep", False)

        ' Set the current pressure storage dates to now to prevent barometric drop alarms after a long sleep and a short storm warning interval
        gblPressureStorageDate = CStr(Now())

        Call WeatherMeteo.getData
        
        fAnemometer.anemometerGaugeForm.Widgets("housing/startbutton").Widget.Alpha = 0 ' visual indicator that a sleep has been handled
        
        fTemperature.temperatureGaugeForm.Refresh
        fAnemometer.anemometerGaugeForm.Refresh
        fHumidity.humidityGaugeForm.Refresh
        fBarometer.barometerGaugeForm.Refresh
        fClipB.clipBForm.Refresh
        fPictorial.pictorialGaugeForm.Refresh
        
        ' trigger another poll 2.5 seconds after the system has awoken to give it time to wake up and re-poll!
        tmrSampler.Interval = 2500 ' this abnormal interval will be reset to the sampling interval by the timer itself
        tmrSleepTester.Enabled = True
    End If
    
    strTimeThen = Now()
    tmrSleepTester.Enabled = True
    
    On Error GoTo 0
    Exit Sub

tmrSleepTester_TimerCode_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSleepTester_TimerCode of Form dock"

End Sub


' ----------------------------------------------------------------
' Procedure Name: setIconTooltip
' Purpose: function to set the mainweather icon tooltip according to two sources
' Procedure Kind: function
' Procedure Access: Private
' Parameter wx_string ():
' Parameter sky_cover ():
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 21/03/2024
' ----------------------------------------------------------------
Private Function setIconTooltip(ByVal wx_string As String, ByRef sky_condition_items() As String, ByVal uboundSkyConditions As Integer) As String
    Dim thisToolTip As String: thisToolTip = vbNullString
    
    On Error GoTo setIconTooltip_Error
    
    'firstly, try the wx_string
    If wx_string <> vbNullString Then thisToolTip = fGetMetarDescription(wx_string)
    
    If (thisToolTip = vbNullString) Then
      ' if no wx_string use the cloud cover or lack of it
      thisToolTip = getCloudCover(sky_condition_items(), uboundSkyConditions)
    End If

    If (debugFlg = 1) Then Debug.Print ("%setIconTooltip - thisTooltip " & thisToolTip)
    
    setIconTooltip = thisToolTip ' return
    
    On Error GoTo 0
    Exit Function

setIconTooltip_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setIconTooltip, line " & Erl & "."

End Function



' ----------------------------------------------------------------
' Procedure Name: niceFormatInterval
' Purpose: this function provides a text version of when the data was last collected
' Procedure Kind: Function
' Procedure Access: Public
' Parameter duration (Integer):
' Return Type: String
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 15/03/2024
' ----------------------------------------------------------------
Private Function niceFormatInterval(ByVal duration As Double, ByVal dateTimeNow As Date) As String
    On Error GoTo niceFormatInterval_Error
    
    Dim hours As Long: hours = 0
    Dim addString As String: addString = vbNullString
    Dim minutes As Long: minutes = 0
    Dim res As String: res = vbNullString
    
    duration = Abs(duration)
    hours = Int(duration / 3600)
    minutes = CLng((duration - hours * 3600) / 60)
    
    addString = "prior to last poll at " & dateTimeNow
    
    If hours > 72 Then
        'print("%Update " & String(hours) & " hours and " & String(minutes) & " minutes ago")
        niceFormatInterval = "No recent data available!" & vbCrLf
        'Exit Function ' return
    End If
    If hours > 0 Then
        res = "Updated"
        If hours = 1 Then
            res = res & " " & "one hour"
        Else
            res = res & " " & CStr(hours) & " " & "hours"
        End If
        If minutes = 1 Then
            res = res & " " & "and one minute"
        ElseIf minutes > 1 Then
            res = res & " " & "and" & " " & CStr(minutes) & " " & "minutes"
        End If
        niceFormatInterval = niceFormatInterval + res & " " & addString
        Exit Function ' return
    End If
    If minutes > 0 Then
        If minutes = 1 Then
            niceFormatInterval = niceFormatInterval + "Updated one minute " & addString
            Exit Function ' return
        End If
        niceFormatInterval = niceFormatInterval + "Updated" & " " & CStr(minutes) & " " & "minutes " & addString
        Exit Function ' return
    End If
    If duration = 1 Then
        niceFormatInterval = niceFormatInterval + "Updated one second " & addString
        Exit Function ' return
    End If
    
    niceFormatInterval = "Updated" & " " & CStr(duration) & " " & "seconds " & addString ' return
    
    On Error GoTo 0
    Exit Function

niceFormatInterval_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure niceFormatInterval, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: getCloudCover
' Purpose: function to decode cloud cover information into a text form
' Procedure Kind: Function
' Procedure Access: Public
' Parameter sky_cover (String):
' Return Type: String
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 29/01/2024
'
' Decodes cloud cover information.
' Format is SKC or CLR for clear skies, or cccnnn where ccc = 3-letter
' code and nnn = altitude of cloud layer in hundreds of feet. 'VV' seems
' to be used for very low cloud layers. (Other conversion factor:
' 1 m = 3.28084 ft)
' FEW = 1 or 2 eighths cover SCT = 3 or 4 eighths cover BKN = 5, 6 or 7 eighths cover & OVC = 8/8 cover
' ----------------------------------------------------------------
Private Function getCloudCover(ByRef sky_condition_items() As String, ByVal uboundSkyConditions As Integer) As String
    Dim thisSkyClarity As String: thisSkyClarity = vbNullString
   
    On Error GoTo getCloudCover_Error
    
    thisSkyClarity = "Ceiling and visibility OK"

   If uboundSkyConditions >= 0 Then
        If (sky_condition_items(0) = "SKC") Then thisSkyClarity = "Clear skies"
        If (sky_condition_items(0) = "CLR") Then thisSkyClarity = "Clear skies"
        If (sky_condition_items(0) = "FEW") Then thisSkyClarity = "Partly cloudy"
        If (sky_condition_items(0) = "SCT") Then thisSkyClarity = "Scattered clouds"
        If (sky_condition_items(0) = "BKN") Then thisSkyClarity = "Mostly cloudy"
        If (sky_condition_items(0) = "OVC") Then thisSkyClarity = "Overcast"
        If (sky_condition_items(0) = "VV ") Then thisSkyClarity = "Vertical visibility"
        If (sky_condition_items(0) = "CAVOK ") Then thisSkyClarity = "Ceiling and visibility OK"
    End If
    
    If uboundSkyConditions >= 1 Then
        If sky_condition_items(1) <> vbNullString Then
            If (sky_condition_items(uboundSkyConditions) = "FEW" And Not sky_condition_items(0) = sky_condition_items(uboundSkyConditions)) Then thisSkyClarity = thisSkyClarity & " + " & "Partly cloudy"
            If (sky_condition_items(uboundSkyConditions) = "SCT" And Not sky_condition_items(0) = sky_condition_items(uboundSkyConditions)) Then thisSkyClarity = thisSkyClarity & " + " & "Scattered clouds"
            If (sky_condition_items(uboundSkyConditions) = "BKN" And Not sky_condition_items(0) = sky_condition_items(uboundSkyConditions)) Then thisSkyClarity = thisSkyClarity & " + " & "Mostly cloudy"
            If (sky_condition_items(uboundSkyConditions) = "OVC" And Not sky_condition_items(0) = sky_condition_items(uboundSkyConditions)) Then thisSkyClarity = thisSkyClarity & " + " & "Overcast"
        End If
    End If
    If (debugFlg = 1) Then Debug.Print ("%getCloudCover - part " & thisSkyClarity)
    
    getCloudCover = thisSkyClarity 'return thisSkyClarity
    
    On Error GoTo 0
    Exit Function

getCloudCover_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getCloudCover, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: fPressureMillibars
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter altim_in_hg (String):
' Return Type: Integer
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 29/01/2024
' ----------------------------------------------------------------
Private Function fPressureMillibars(ByVal altim_in_hg As String) As Integer

    On Error GoTo fPressureMillibars_Error
    
    debugFlg = 0

    Dim thisPressureValue As Double: thisPressureValue = 0
    Dim strPressureValue As String: strPressureValue = vbNullString
    
    thisPressureValue = Round(1.3333 * (CDbl(altim_in_hg) * 25.3999))
    strPressureValue = Str$(thisPressureValue) & "(" & CStr(Round(1.3333 * thisPressureValue)) & ")"
    
    If debugFlg = 1 Then Debug.Print ("fPressureMillibars - Pressure = " & strPressureValue)

    fPressureMillibars = thisPressureValue ' return
    
    On Error GoTo 0
    Exit Function

fPressureMillibars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fPressureMillibars, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: fGetHumidity
' Purpose: calculate Humidity
' Procedure Kind: Function
' Procedure Access: Public
' Parameter temp_c ():
' Parameter dewpoint_c ():
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 29/01/2024
' ----------------------------------------------------------------
Private Function fGetHumidity(ByVal temp_c As String, ByVal dewpoint_c As String) As Long

    Dim humidity As Long: humidity = 0
    
    debugFlg = 0

    On Error GoTo fGetHumidity_Error
    
    humidity = Round(100 * (((112 - (0.1 * CInt(temp_c)) + CInt(dewpoint_c)) / (112 + (0.9 * CInt(temp_c)))) ^ 8))
    If debugFlg = 1 Then Debug.Print ("fGetHumidity - Humidity " + CStr(humidity))
    
    fGetHumidity = humidity ' return
    
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
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 30/01/2024
' ----------------------------------------------------------------
Private Function fAddBaseCloudToSkyClarity(ByVal baseCloud As Integer, ByVal skyClarity As String, ByRef cloud_base_ft_agl() As String) As String
    
    Dim skyClarityString As String: skyClarityString = vbNullString
    Dim thisSkyClarity As String: thisSkyClarity = vbNullString
    Dim thisBaseCloud As Integer: thisBaseCloud = 0
    Dim interimString As String: interimString = vbNullString
    
    On Error GoTo fAddBaseCloudToSkyClarity_Error
    
    debugFlg = 0
    
    thisSkyClarity = skyClarity
    thisBaseCloud = baseCloud
    
    If thisSkyClarity = vbNullString Then Exit Function
    
    If UBound(cloud_base_ft_agl()) >= 0 Then
        If (cloud_base_ft_agl(0) <> vbNullString) Then thisBaseCloud = Val(cloud_base_ft_agl(0))
        If (thisBaseCloud <> 0) Then
            If gblMetricImperial = "1" Then ' "metric"
                ' change the value to metric measurements
                thisBaseCloud = Int(thisBaseCloud * 0.3048)
            End If
            interimString = CStr(thisBaseCloud)
        End If
    End If
    
   If UBound(cloud_base_ft_agl()) >= 1 Then
        If (cloud_base_ft_agl(1) <> vbNullString) Then thisBaseCloud = Val(cloud_base_ft_agl(UBound(cloud_base_ft_agl())))
        If (thisBaseCloud <> 0) Then
            If gblMetricImperial = "1" Then ' "metric"
                ' change the value to metric measurements
                thisBaseCloud = Int(thisBaseCloud * 0.3048)
            End If
            interimString = interimString & ", " & CStr(thisBaseCloud)
        End If
    End If
    
    If gblMetricImperial = "1" Then ' "metric"
        If interimString = vbNullString Then
            skyClarityString = thisSkyClarity
        Else
            skyClarityString = thisSkyClarity & " at " & interimString & " metres."
        End If
    Else
        If interimString = vbNullString Then
            skyClarityString = thisSkyClarity
        Else
            skyClarityString = thisSkyClarity & " at " & interimString & " ft."
        End If
    End If
    
    If debugFlg = 1 Then Debug.Print ("fAddBaseCloudToSkyClarity - skyClarityString " + skyClarityString)

    fAddBaseCloudToSkyClarity = skyClarityString ' return
    
    On Error GoTo 0
    Exit Function

fAddBaseCloudToSkyClarity_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fAddBaseCloudToSkyClarity, line " & Erl & "."

End Function
    

'' ----------------------------------------------------------------
'' Procedure Name: convertToFahrenheit
'' Purpose: convert temps from C to F values here to match scale setting
'' Procedure Kind: Function
'' Procedure Access: Private
'' Parameter t ():
'' Author: Dean Beedell (yereverluvinunclebert)
'' Date: 17/01/2024
'' ----------------------------------------------------------------
'Private Function convertToFahrenheit(ByVal T As Integer) As Integer
'    Dim newT As Integer
'
'    On Error GoTo convertToFahrenheit_Error
'
'    newT = Round(1.8 * T + 32)
'
'    convertToFahrenheit = newT ' return
'
'    On Error GoTo 0
'    Exit Function
'
'convertToFahrenheit_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure convertToFahrenheit, line " & Erl & "."
'
'End Function
'
'
'' ----------------------------------------------------------------
'' Procedure Name: convertToKelvin
'' Purpose: convert temps from C to F values here to match scale setting
'' Procedure Kind: Function
'' Procedure Access: Private
'' Parameter t ():
'' Author: Dean Beedell (yereverluvinunclebert)
'' Date: 17/01/2024
'' ----------------------------------------------------------------
'Private Function convertToKelvin(ByVal T As Integer) As Integer
'    Dim newT As Double
'
'    On Error GoTo convertToKelvin_Error
'
'    Const KELVIN_CELSIUS_DIFF As Double = 273.15
'
'    newT = KELVIN_CELSIUS_DIFF + T
'
'    convertToKelvin = Fix(newT) ' return
'
'    On Error GoTo 0
'    Exit Function
'
'convertToKelvin_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure convertToKelvin, line " & Erl & "."
'
'End Function



' ----------------------------------------------------------------
' Procedure Name: degToCompass
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter windDirection (Integer):
' Return Type: String
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 15/03/2024
' ----------------------------------------------------------------
Private Function degToCompass(ByVal windDirection As String) As String
    On Error GoTo degToCompass_Error
    
    Dim thisVal As Integer: thisVal = 0
    Dim thisArray As Variant ' using variant to populate the array all in one go as per .js and TwinBasic, will allow easy conversion later
    
    Dim compassDirection As String: compassDirection = vbNullString
    
    thisArray = Array("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
    
    thisVal = CInt((Val(windDirection) / 22.5) + 0.5)
    compassDirection = thisArray(thisVal Mod 16)
    
    degToCompass = compassDirection ' return compassDirection
    
    On Error GoTo 0
    Exit Function

degToCompass_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure degToCompass, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: beaufortConversion
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter wind_speed_kt (Integer):
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 15/03/2024
' ----------------------------------------------------------------
Private Function beaufortConversion(ByVal wind_speed_kt As Integer) As Integer
    
    Dim force As Integer: force = 0

    On Error GoTo beaufortConversion_Error
        
    ' Beaufort Scales (Wind Speed)
    If wind_speed_kt < 1 Then
        force = 0 ' Calm    Sea like a mirror.
    End If
    If wind_speed_kt >= 1 And wind_speed_kt <= 3 Then
        force = 1 ' 1 1-3   Light air   Ripples only.
    End If
    If wind_speed_kt >= 4 And wind_speed_kt <= 6 Then
        force = 2 ' 4-6 Light breeze    Small wavelets (0.2 m). Crests have a glassy appearance.
    End If
    If wind_speed_kt >= 7 And wind_speed_kt <= 10 Then
        force = 3 ' 7-10  Gentle breeze     Large wavelets (0.6 m), crests begin to break.
    End If
    If wind_speed_kt >= 11 And wind_speed_kt <= 16 Then
        force = 4 ' 11-16  Moderate breeze  Small waves (1 m), some whitecaps
    End If
    If wind_speed_kt >= 17 And wind_speed_kt <= 21 Then
        force = 5 '  17-21 Fresh breeze
    End If
    If wind_speed_kt >= 22 And wind_speed_kt <= 27 Then
        force = 6 ' 22-27 Strong breeze
    End If
    If wind_speed_kt >= 28 And wind_speed_kt <= 33 Then
        force = 7 ' 28-33 Near gale
    End If
    If wind_speed_kt >= 34 And wind_speed_kt <= 40 Then
        force = 8 ' 34-40 Gale
    End If
    If wind_speed_kt >= 41 And wind_speed_kt <= 47 Then
        force = 9 ' 41-47 Strong gale
    End If
    If wind_speed_kt >= 48 And wind_speed_kt <= 55 Then
        force = 10 ' 48-55 Storm
    End If
    If wind_speed_kt >= 56 And wind_speed_kt <= 63 Then
        force = 11 ' 56-63 Violent storm
    End If
    If wind_speed_kt >= 64 And wind_speed_kt <= 79 Then
        force = 12 ' 64+ Hurricane
    End If
    If wind_speed_kt >= 80 And wind_speed_kt <= 88 Then
        force = 13 ' 80+ Typhoon
    End If
    If wind_speed_kt >= 89 And wind_speed_kt <= 98 Then
        force = 14 ' 80+ Typhoon
    End If
    If wind_speed_kt >= 99 And wind_speed_kt <= 107 Then
        force = 15 ' 80+ Typhoon
    End If
    If wind_speed_kt >= 108 And wind_speed_kt <= 117 Then
        force = 16 ' 80+ Severe Typhoon
    End If
    If wind_speed_kt >= 118 Then
       force = 17 ' 80+ Severe Typhoon
     End If
    
     beaufortConversion = force ' return
    
    On Error GoTo 0
    Exit Function

beaufortConversion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure beaufortConversion, line " & Erl & "."

End Function



' ----------------------------------------------------------------
' Procedure Name: fWindIconTooltip
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter wind_speed_kt (Integer):
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 15/03/2024
' ----------------------------------------------------------------
Private Function fWindIconTooltip(ByVal force As Integer) As String

    Dim thisWindIconToolTip As String: thisWindIconToolTip = vbNullString

    On Error GoTo fWindIconTooltip_Error
    
    fWindIconTooltip = vbNullString
    
    ' Beaufort Scales (Wind Speed)
    If force = 0 Then
         ' Calm    Sea like a mirror.
        thisWindIconToolTip = "Calm " & "strength force " & force
    End If
    If force = 1 Then
        ' 1 1-3   Light air   Ripples only.
        thisWindIconToolTip = "Light air ripples " & "strength force " & force
    End If
    If force = 2 Then
        ' 4-6 Light breeze    Small wavelets (0.2 m). Crests have a glassy appearance.
        thisWindIconToolTip = "Light breeze " & "strength force " & force
    End If
    If force = 3 Then
         ' 7-10  Gentle breeze     Large wavelets (0.6 m), crests begin to break.
        thisWindIconToolTip = "Gentle breeze " & "strength force " & force
    End If
    If force = 4 Then
        ' 11-16  Moderate breeze  Small waves (1 m), some whitecaps
        thisWindIconToolTip = "Moderate breeze " & "strength force " & force
    End If
    If force = 5 Then
        '  17-21 Fresh breeze
        thisWindIconToolTip = "Fresh breeze " & "strength force " & force
    End If
    If force = 6 Then
        ' 22-27 Strong breeze
        thisWindIconToolTip = "Strong Wind " & "strength force " & force
    End If
    If force = 7 Then
         ' 28-33 Near gale
        thisWindIconToolTip = "Near gale " & "strength force " & force
    End If
    If force = 8 Then
         ' 34-40 Gale
        thisWindIconToolTip = "Gale " & "strength force " & force
    End If
    If force = 9 Then
         ' 41-47 Strong gale
        thisWindIconToolTip = "Strong gale " & "strength force " & force
    End If
    If force = 10 Then
         ' 48-55 Storm
        thisWindIconToolTip = "Storm " & "strength force " & force
    End If
    If force = 11 Then
         ' 56-63 Violent storm
        thisWindIconToolTip = "Violent storm " & "strength force " & force
    End If
    If force = 12 Then
         ' 64+ Hurricane
        thisWindIconToolTip = "Hurricane " & "strength force " & force
    End If
    If force = 13 Then
        ' 80+ Typhoon
        thisWindIconToolTip = "Typhoon " & "strength force " & force
    End If
    If force = 14 Then
         ' 80+ Typhoon
        thisWindIconToolTip = "Typhoon " & "strength force " & force
    End If
    If force = 15 Then
         ' 80+ Typhoon
        thisWindIconToolTip = "Typhoon " & "strength force " & force
    End If
    If force = 16 Then
        ' 80+ Severe Typhoon
        thisWindIconToolTip = "Severe Typhoon " & "strength force " & force
    End If
    If force = 17 Then
       ' 80+ Severe Typhoon
       thisWindIconToolTip = "Severe Typhoon " & "strength force " & force
     End If
    
     fWindIconTooltip = thisWindIconToolTip ' return
    
    On Error GoTo 0
    Exit Function

fWindIconTooltip_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fWindIconTooltip, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: fWindIconSrc
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 15\03\2024
' ----------------------------------------------------------------
Private Function fWindIconSrc(ByVal force As Integer, ByVal thisIsDay As Boolean) As String

    Dim thisWindIconSrc As String: thisWindIconSrc = vbNullString

    On Error GoTo fWindIconSrc_Error
    
    thisWindIconSrc = "nowind"
    
    ' Beaufort Scales (Wind Speed)
    If force = 0 Then
        ' Calm    Sea like a mirror.
        thisWindIconSrc = "nowind"
    End If
    If force = 1 Then
        ' 1 1-3   Light air   Ripples only.
        thisWindIconSrc = "nowind"
    End If
    If force = 2 Then
        'force = 2 ' 4-6 Light breeze    Small wavelets (0.2 m). Crests have a glassy appearance.
        thisWindIconSrc = "nowind"
    End If
    If force = 3 Then
        'force = 3 ' 7-10  Gentle breeze     Large wavelets (0.6 m), crests begin to break.
        thisWindIconSrc = "nowind"
    End If
    If force = 4 Then
        'force = 4 ' 11-16  Moderate breeze  Small waves (1 m), some whitecaps
        thisWindIconSrc = "windy01"
    End If
    If force = 5 Then
        'force = 5 '  17-21 Fresh breeze
        thisWindIconSrc = "windy02"
    End If
    If force = 6 Then
         ' 22-27 Strong breeze
        thisWindIconSrc = "windy03"
    End If
    If force = 7 Then
         ' 28-33 Near gale
        thisWindIconSrc = "windy04"
    End If
    If force = 8 Then
        ' 34-40 Gale
        thisWindIconSrc = "windy05"
    End If
    If force = 9 Then
         ' 41-47 Strong gale
        thisWindIconSrc = "windy05"
    End If
    If force = 10 Then
         ' 48-55 Storm
        thisWindIconSrc = "windy05"
    End If
    If force = 11 Then
        ' 56-63 Violent storm
        thisWindIconSrc = "windy05"
    End If
    If force = 12 Then
        ' 64+ Hurricane
        thisWindIconSrc = "windy05"
    End If
    If force = 13 Then
         ' 80+ Typhoon
        thisWindIconSrc = "windy05"
    End If
    If force = 14 Then
        'force = 14 ' 80+ Typhoon
        thisWindIconSrc = "windy05"
    End If
    If force = 15 Then
        'force = 15 ' 80+ Typhoon
        thisWindIconSrc = "windy05"
    End If
    If force = 16 Then
        'force = 16 ' 80+ Severe Typhoon
        thisWindIconSrc = "windy05"
    End If
    If force = 17 Then
       'force = 17 ' 80+ Severe Typhoon
       thisWindIconSrc = "windy05"
     End If
     
    ' turn the image from day to night
    If (thisWindIconSrc <> vbNullString) Then
      If (thisIsDay = False) Then
         thisWindIconSrc = "n_" & thisWindIconSrc
      End If
    End If
     
    fWindIconSrc = thisWindIconSrc ' return
    
    On Error GoTo 0
    Exit Function

fWindIconSrc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fWindIconSrc, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: setHoverTooltip
' Purpose: this function sets the hover over tooltip
' Procedure Kind: Function
' Procedure Access: Private
' Parameter weekDayDate (Date):
' Parameter wind_dir_degrees (String):
' Parameter wind_speed_kt (String):
' Parameter precip_in (String):
' Parameter lastUpdateString (String):
' Parameter wx_string (String):
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 17/03/2024
' ----------------------------------------------------------------
Private Function setHoverTooltip(ByVal thisAltim_in_hg As Double, ByVal force As Integer, ByVal precipitationDepth As String, ByVal visibilityDistance As String, ByVal skyClarityString As String, ByVal HumidityValue As Double, ByVal thisTemperatureValue As Single, ByVal pressureValueMillibars As Double, ByVal weekDayDate As String, ByVal wind_dir_degrees As String, ByVal wind_speed_kt As String, ByVal precip_in As String, ByVal wx_string As String) As String
    
    Dim windSpeedTooltip As String: windSpeedTooltip = vbNullString
    Dim tooltipString As String: tooltipString = vbNullString
    Dim compassDirection As String: compassDirection = vbNullString
    Dim wx_str As String: wx_str = vbNullString

    On Error GoTo setHoverTooltip_Error
    
    compassDirection = degToCompass(wind_dir_degrees)
    
    ' obtain the wind related tooltips
    windSpeedTooltip = fWindIconTooltip(force)
    
    tooltipString = CStr(weekDayDate) & vbCrLf
    tooltipString = tooltipString & gblIcao & " " & icaoLocation1 & ", " & vbCrLf & icaoLocation2 & ", " & icaoLocation3 & " " & vbCrLf

    If (gblPressureScale = "0") Then ' "millibars") Then
        tooltipString = tooltipString & "Pressure" & ": " & CStr(Int(pressureValueMillibars)) & " mb " & vbCrLf
     ElseIf (gblPressureScale = "1") Then '"hg") Then
        tooltipString = tooltipString & "Pressure" & ": " & CStr(thisAltim_in_hg) & " inHg " & vbCrLf
     ElseIf (gblPressureScale = "2") Then '"mmhg") Then
        tooltipString = tooltipString & "Pressure" & ": " & CStr(thisAltim_in_hg * 25.4) & " mmHg " & vbCrLf
     ElseIf (gblPressureScale = "3") Then '"hpa") Then
        tooltipString = tooltipString & "Pressure" & ": " & CStr(Int(pressureValueMillibars)) & " hpa " & vbCrLf
    End If

    If gblTemperatureScale = "1" Then '"fahrenheit"
        tooltipString = tooltipString & "Temperature" & ": " & CStr(Round(thisTemperatureValue * 1.8 + 32, 2)) & " F " & vbCrLf
    End If
    If gblTemperatureScale = "0" Then ' "celsius"
        tooltipString = tooltipString & "Temperature" & ": " & CStr(Round(thisTemperatureValue, 2)) & " C " & vbCrLf
    End If
    
    If gblTemperatureScale = "2" Then '"kelvin"
        tooltipString = tooltipString & "Temperature" & ": " & CStr(Round(Int(thisTemperatureValue + 273.15), 2)) & " K " & vbCrLf
    End If
    
    tooltipString = tooltipString & "Humidity : " & CStr(HumidityValue) & " % " & vbCrLf
    tooltipString = tooltipString & "Clouds : " & skyClarityString & " " & vbCrLf
    tooltipString = tooltipString & "Visibility : " & visibilityDistance & vbCrLf
    
    ' when zero wind and direction North then do not display wind details at all.
    If wind_speed_kt = 0 And compassDirection = "N" Then
        tooltipString = tooltipString
    Else
        If gblWindSpeedScale = "0" Then  ' knots
            tooltipString = tooltipString & "Wind : " & windSpeedTooltip & ", speed " & wind_speed_kt & " knots "
        Else
            tooltipString = tooltipString & "Wind : " & windSpeedTooltip & ", speed " & Int(Val(wind_speed_kt) * 0.51444) & " m/sec "
        End If
        tooltipString = tooltipString & vbCrLf & "Direction : " & compassDirection & vbCrLf
    End If
    
    If (wx_string <> vbNullString) Then
      wx_str = fGetMetarDescription(wx_string)    ' get the interpreted wx string data
      wx_str = UCase$(Left$(wx_str, 1)) & Mid$(wx_str, 2, Len(wx_str))
      tooltipString = tooltipString & "Rainfall: " & wx_str & "  "
    End If
    If (precip_in <> vbNullString) Then
        tooltipString = tooltipString & " measured in " & precipitationDepth & vbCrLf
    Else
        If (wx_string <> vbNullString) Then tooltipString = tooltipString & vbCrLf
    End If
    
    ' one millibar drop < 1 hour - a storm is brewing
    If overlayBaromWidget.pressureDrop = True Then
        tooltipString = tooltipString & "WARNING: The red lamp is lit as the pressure is dropping fast enough for a serious storm. " & vbCrLf
    End If
     
    setHoverTooltip = tooltipString ' return

    On Error GoTo 0
    Exit Function

setHoverTooltip_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHoverTooltip, line " & Erl & "."

End Function





' ----------------------------------------------------------------
' Procedure Name: determineWeatherConditionIcon
' Purpose: function to decode_weather when a wx field found and set the icon accordingly
' Procedure Kind: sub
' Procedure Access: Private
' Parameter SkyClarity (String):
' Parameter thisWxString (String):
' Parameter sky_cover (String):
' Parameter thisIsDay (Boolean):
' Return Type: String, ByVal thisIsDay As Boolean)
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 14\03\2024
' ----------------------------------------------------------------
Private Sub determineWeatherConditionIcon(ByRef skyClarity As String, ByVal thisWxString As String, ByRef sky_cover() As String, ByVal thisIsDay As Boolean, ByVal uboundSkyConditions As Integer)
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
   
    Static precipitationIsFalling As Boolean
      
    weatherIconSrc = "globe"
    precipitationIsFalling = False
    
    If uboundSkyConditions >= 0 Then
         If (sky_cover(0) = "SKC") Then skycover1 = "SKC"
         If (sky_cover(0) = "CLR") Then skycover1 = "CLR"
         If (sky_cover(0) = "FEW") Then skycover1 = "FEW"
         If (sky_cover(0) = "SCT") Then skycover1 = "SCT"
         If (sky_cover(0) = "BKN") Then skycover1 = "BKN"
         If (sky_cover(0) = "OVC") Then skycover1 = "OVC"
         If (sky_cover(0) = "VV ") Then skycover1 = "VV "
         If (sky_cover(0) = "CAVOK") Then skycover1 = "CAVOK"
     End If
        
   
   If uboundSkyConditions >= 1 Then
        If (sky_cover(1) = "SKC") Then skycover2 = "SKC"
        If (sky_cover(1) = "CLR") Then skycover2 = "CLR"
        If (sky_cover(1) = "FEW") Then skycover2 = "FEW"
        If (sky_cover(1) = "SCT") Then skycover2 = "SCT"
        If (sky_cover(1) = "BKN") Then skycover2 = "BKN"
        If (sky_cover(1) = "OVC") Then skycover2 = "OVC"
        If (sky_cover(1) = "VV ") Then skycover2 = "VV "
        If (sky_cover(1) = "CAVOK") Then skycover2 = "CAVOK"
    End If
    
   
   If uboundSkyConditions >= 2 Then
        If (sky_cover(2) = "SKC") Then skycover3 = "SKC"
        If (sky_cover(2) = "CLR") Then skycover3 = "CLR"
        If (sky_cover(2) = "FEW") Then skycover3 = "FEW"
        If (sky_cover(2) = "SCT") Then skycover3 = "SCT"
        If (sky_cover(2) = "BKN") Then skycover3 = "BKN"
        If (sky_cover(2) = "OVC") Then skycover3 = "OVC"
        If (sky_cover(2) = "VV ") Then skycover3 = "VV "
        If (sky_cover(2) = "CAVOK") Then skycover3 = "CAVOK"
    End If

   If (debugFlg = 1 And uboundSkyConditions >= 0) Then Debug.Print ("%determineWeatherCondition - sky_cover " & sky_cover(0))

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
   ' "clear and sunny"
   weatherIconSrc = "0cloud"

   'determine the weather type,
   If thisWxString <> vbNullString Then
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
    End If

   ' FEW = 1 or 2 eighths cover SCT = 3 or 4 eighths cover BKN = 5, 6 or 7 eighths cover & OVC = 8\8 cover

   'use the cloud code to determine cloud cover
    'kyClarity = "Uknown Sky Cover"
    If (skycover1 = "SKC" Or skycover2 = "SKC" Or skycover3 = "SKC") Then skyClarity = "Clear skies"       ' default clear
    If (skycover1 = "CLR" Or skycover2 = "CLR" Or skycover3 = "CLR") Then skyClarity = "Clear skies"       ' default
    If (skycover1 = "VV " Or skycover2 = "VV " Or skycover3 = "VV ") Then skyClarity = "Vertical visibility"          ' default
    If (skycover1 = "CAVOK" Or skycover2 = "CAVOK" Or skycover3 = "CAVOK") Then skyClarity = "Ceiling and visibility OK"          ' default

    If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
      'Partly cloudy"
      weatherIconSrc = "1cloud_norain"
    End If
    If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
      'Scattered clouds"
      weatherIconSrc = "2cloud_norain"
    End If
    If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
      'Mostly cloudy"
      weatherIconSrc = "3cloud_norain"
    End If
    If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
      'Overcast"
      weatherIconSrc = "4cloud_norain"
    End If


   'type of weather - snow
   If (InStr(thisWxString, "SN") <> 0 Or InStr(thisWxString, "SG") <> 0) Then

       If (wSeverity = "light") Then
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
             'Scattered clouds"
             weatherIconSrc = "1cloud_lightsnow" 'default
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "2cloud_lightsnow"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Mostly cloudy"
            weatherIconSrc = "4cloud_lightsnow"
          End If
       End If
       If (wSeverity = "medium") Then
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_snow"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "2cloud_snow"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "3cloud_snow"
          End If
       End If
       If (wSeverity = "heavy") Then
            weatherIconSrc = "4cloud_heavysnow"
       End If
       
        precipitationIsFalling = True ' snow
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
   End If ' and\or snow pellets

   'type of weather - hail
   If (weatherType = "hail") Then
       If (wSeverity = "light") Then
          weatherIconSrc = "2cloud_hail" 'default
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Mostly cloudy"
            weatherIconSrc = "4cloud_lighthail"
          End If
       End If
       If (wSeverity = "medium") Then
          If ((skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Or (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN")) Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_hail"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_hail"
          End If
       End If
       If (wSeverity = "heavy") Then
            weatherIconSrc = "4cloud_heavyhail"
       End If
       
        precipitationIsFalling = True ' hail
   End If

    ' this caters for the event when there are showers indicated but no associated rain code
    ' it is strange but it quite often occurs...
    If (InStr(thisWxString, "SH") <> 0 And InStr(thisWxString, "SN") = 0 And InStr(thisWxString, "SG") = 0 And weatherType <> "hail") Then
         'if no RAin code nor hail or snow but showers indicated then just assume rain...
        showersIconSrc = "showers"
    Else
        showersIconSrc = "blank"
    End If

   'type of weather - rain
   If (InStr(thisWxString, "RA") <> 0) Then
       wSeverity = "light" ' default
     '"rain "
       If (wSeverity = "light") Then
          If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
            'Partly cloudy"
            weatherIconSrc = "1cloud_lightrain"
          End If
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_lightrain"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_lightrain"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "4cloud_lightrain"
          End If
       End If
       If (wSeverity = "medium") Then
          If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
            'Partly cloudy"
            weatherIconSrc = "1cloud_modrain"
          End If
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_modrain"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_modrain"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "4cloud_modrain"
          End If
       End If
       If (wSeverity = "heavy") Then
          If (skycover1 = "FEW" Or skycover2 = "FEW" Or skycover3 = "FEW") Then
            'Partly cloudy"
            weatherIconSrc = "1cloud_heavyrain"
          End If
          If (skycover1 = "SCT" Or skycover2 = "SCT" Or skycover3 = "SCT") Then
            'Scattered clouds"
            weatherIconSrc = "2cloud_heavyrain"
          End If
          If (skycover1 = "BKN" Or skycover2 = "BKN" Or skycover3 = "BKN") Then
            'Mostly cloudy"
            weatherIconSrc = "3cloud_heavyrain"
          End If
          If (skycover1 = "OVC" Or skycover2 = "OVC" Or skycover3 = "OVC") Then
            'Overcast"
            weatherIconSrc = "4cloud_heavyrain"
          End If
       End If

       If (InStr(thisWxString, "TS") <> 0) Then
            wModifier = wModifier & "thunderstorm "
            weatherIconSrc = "4cloud_heavyrain"
       End If
        
        precipitationIsFalling = True ' rain
   End If

   
   fogIconSrc = "blank"    'default
    If (InStr(thisWxString, "BR") <> 0 Or InStr(thisWxString, "FG") <> 0) Then
      '"fog " or "mist"
      If (wSeverity = "light") Then
           fogIconSrc = "1_fog" 'night or day makes no difference
        End If
        If (wSeverity = "medium") Then
           fogIconSrc = "2_fog"
        End If
        If (wSeverity = "heavy") Then
           fogIconSrc = "3_fog"
        End If
        
    End If
    
    'windIconSrc= "Resources\images\icons_metar\windy03"   'default
    If (InStr(thisWxString, "SQ") <> 0) Then
      '"strong winds "
      windIconSrc = "windy03"
    End If
    
    
    exoticIconSrc = "blank"   'default
    'exotic types - we have no icons for any of these
    If (InStr(thisWxString, "FU") <> 0) Then
       exoticIconSrc = "smoke"
       '"smoke "
    End If
    If (InStr(thisWxString, "VA") <> 0) Then
      '"volcanic ash "
       exoticIconSrc = "volcano"
    End If
    If (InStr(thisWxString, "DU") <> 0) Then
      '"widespread dust "endif
       exoticIconSrc = "dust"
    End If
    If (InStr(thisWxString, "SA") <> 0) Then
      '"sand "endif
       exoticIconSrc = "sand"
    End If
    If (InStr(thisWxString, "HZ") <> 0) Then
      '"haze "endif
       exoticIconSrc = "haze"
    End If
    If (InStr(thisWxString, "PY") <> 0) Then
      '"spray "endif
       exoticIconSrc = "spray"
    End If
    If (InStr(thisWxString, "PO") <> 0) Then
      '"dustdevils "endif
       exoticIconSrc = "dustdevil"
    End If
    If (InStr(thisWxString, "FC") <> 0) Then
      '"tornado "endif
       exoticIconSrc = "tornado"
    End If
    If (InStr(thisWxString, "SS") <> 0) Then
      '"sandstorm\duststorm "
       exoticIconSrc = "duststorm"
    End If
   
    ' turn the image from day to night
    If (weatherIconSrc <> vbNullString) Then
      If (thisIsDay = True) Then
         iconSrc = weatherIconSrc
         weatherIconSrc = iconSrc
         'precedingSrc = iconSrc ' the preceding icon not  used yet
         'followingSrc = iconSrc ' the following icon not  used yet
      Else
         iconSrc = "n_" & weatherIconSrc
         weatherIconSrc = iconSrc
         'precedingSrc = iconSrc
         'followingSrc = iconSrc
      End If

    End If
    
    
    ' if it is raining now and it is new rain ie. it was not previously raining nor any other form of precipitation
    If precipitationIsFalling = True And mPrecipitationWasFalling = False Then
        
        ' now we set the flag that lets the program know NOT to raise a precipitation alarm again
        PrecipitationWasFalling = True
             
        ' humidity gauge to light the new rain warning lamp
        'overlayHumidWidget.precipitationRedLamp = True
             
        ' pictorial gauge to show the new precipitation print out
        fPictorial.newPrecipitation = True
    
        fPictorial.pictorialGaugeForm.Widgets("sLblPrecipitationText").Caption = "Weather Data Precipitation Alert System " & vbCrLf & "Weather data indicates imminent rainfall"
            
    End If
               
    If (debugFlg = 1) Then Debug.Print ("%determineWeatherCondition - iconSrc " & iconSrc)
    
    ' pass the weather icon to be used to the pictorial overlay where it will be displayed.
    overlayPictorialWidget.thisWeatherImage = weatherIconSrc
    overlayPictorialWidget.thisShowersImage = showersIconSrc
    overlayPictorialWidget.thisFogImage = fogIconSrc
    overlayPictorialWidget.thisExoticImage = exoticIconSrc
    
    On Error GoTo 0
    Exit Sub

determineWeatherConditionIcon_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure determineWeatherConditionIcon, line " & Erl & "."

End Sub

' ----------------------------------------------------------------
' Procedure Name: fGetMetarDescription
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Parameter wx_string (String):
' Return Type: String
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 15/03/2024
' ----------------------------------------------------------------
Private Function fGetMetarDescription(ByVal wx_string As String) As String
    Dim presentConditions As String: presentConditions = vbNullString
    Dim wSeverity As String: wSeverity = vbNullString
    Dim wModifier As String: wModifier = vbNullString

    On Error GoTo fGetMetarDescription_Error
    
    debugFlg = 0
    
    If InStr(wx_string, "UP") <> 0 Then
        presentConditions = presentConditions + "unknown "
    End If
    
    'determine the strength of the weather type, light, medium, heavy or very heavy
    
    If InStr(wx_string, "MI") <> 0 Then
        wModifier = wModifier & "shallow"
    End If
    If InStr(wx_string, "PR") <> 0 Then
        wModifier = wModifier & "partial"
    End If
    If InStr(wx_string, "DZ") <> 0 Then
        wModifier = wModifier & "drizzle"
    End If
    If InStr(wx_string, "BC") <> 0 Then
        wModifier = wModifier & "patches of"
    End If
    If InStr(wx_string, "DR") <> 0 Then
        wModifier = wModifier & "drifting"
    End If
    If InStr(wx_string, "BL") <> 0 Then
        wModifier = wModifier & "blowing"
    End If
    If InStr(wx_string, "SH") <> 0 Then
        wModifier = wModifier & "showers"
    End If
    If InStr(wx_string, "VC") <> 0 Then
        'presentConditions = presentConditions &  "nearby "
        wModifier = wModifier & " in the vicinity "
    End If
    If InStr(wx_string, "FZ") <> 0 Then
        'presentConditions = presentConditions &  "freezing "
        wModifier = wModifier & "freezing "
    End If
    If InStr(wx_string, "TS") <> 0 Then
        wModifier = wModifier & "thunderstorm "
    End If
    If wModifier <> vbNullString Then
        presentConditions = wModifier
    End If
    
    If InStr(wx_string, "-") <> 0 Then
        wSeverity = " light" & " "
    ElseIf InStr(wx_string, "+") <> 0 Then
        wSeverity = " heavy" & " "
    Else
        wSeverity = vbNullString ' moderate conditions have no descriptor
    End If
    
    If InStr(wx_string, "-") = 0 Or InStr(wx_string, "+") = 0 Then
        presentConditions = presentConditions & wSeverity
        If debugFlg = 1 Then
            Debug.Print ("%fGetMetarDescription - wSeverity " & wSeverity)
        End If
    End If
    
    If InStr(wx_string, "RA") <> 0 Then
        presentConditions = presentConditions & "rain" & " "
    End If
    
    If InStr(wx_string, "SN") <> 0 Then
        presentConditions = presentConditions & "snow" & " "
    End If
    
    If InStr(wx_string, "SG") <> 0 Then
        presentConditions = presentConditions & "snow grains" & " "
    End If
    
    If InStr(wx_string, "IC") <> 0 Then
        presentConditions = presentConditions & "ice crystals" & " "
    End If
    
    If InStr(wx_string, "PE") <> 0 Then
        presentConditions = presentConditions & "ice pellets" & " "
    End If
    
    If InStr(wx_string, "GR") <> 0 Then
        presentConditions = presentConditions & "hail" & " "
    End If
    
    If InStr(wx_string, "GS") <> 0 Then
        presentConditions = presentConditions & "small hail" & " "
    End If
    
    If InStr(wx_string, "UP") <> 0 Then
        presentConditions = presentConditions & "unknown" & " "
    End If
    
    If InStr(wx_string, "BR") <> 0 Then
        presentConditions = presentConditions & "mist" & " "
    End If
    
    If InStr(wx_string, "FG") <> 0 Then
        presentConditions = presentConditions & "fog" & " "
    End If
    
    If InStr(wx_string, "FU") <> 0 Then
        presentConditions = presentConditions & "smoke" & " "
    End If
    
    If InStr(wx_string, "VA") <> 0 Then
        presentConditions = presentConditions & "volcanic ash" & " "
    End If
    
    If InStr(wx_string, "DU") <> 0 Then
        presentConditions = presentConditions & "widespread dust" & " "
    End If
    
    If InStr(wx_string, "SA") <> 0 Then
        presentConditions = presentConditions & "sand" & " "
    End If
    
    If InStr(wx_string, "HZ") <> 0 Then
        presentConditions = presentConditions & "haze" & " "
    End If
    
    If InStr(wx_string, "PY") <> 0 Then
        presentConditions = presentConditions & "spray" & " "
    End If
    
    If InStr(wx_string, "PO") <> 0 Then
        presentConditions = presentConditions & "dustdevils" & " "
    End If
    
    If InStr(wx_string, "SQ") <> 0 Then
        presentConditions = presentConditions & "strong winds" & " "
    End If
    
    If InStr(wx_string, "FC") <> 0 Then
        presentConditions = presentConditions & "tornado" & " "
    End If
    
    If InStr(wx_string, "SS") <> 0 Then
        presentConditions = presentConditions & "sandstorm/duststorm" & " "
    End If
    
    If debugFlg = 1 And presentConditions <> vbNullString Then
        Debug.Print ("%fGetMetarDescription - presentConditions " & presentConditions)
    End If
    
    fGetMetarDescription = presentConditions
    
    On Error GoTo 0
    Exit Function

fGetMetarDescription_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetMetarDescription, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: humanFriendlyDate
' Purpose: obtain date in friendly format eg. Sun Mar 24 2024 11:50:00 GMT+0000 (GMT Standard Time) from VB6 date string 24/03/2024 12:20:00
' Procedure Kind: Function
' Procedure Access: Public
' Parameter usableObsTime (Date):
' Return Type: String
' Author: beededea
' Date: 24/03/2024
' ----------------------------------------------------------------
Private Function humanFriendlyDate(ByVal usableObsTime As Date) As String
    On Error GoTo humanFriendlyDate_Error
    
    Dim dayOfWeek As String: dayOfWeek = vbNullString
    Dim formattedString As String: formattedString = vbNullString
    
'    Dim uInfo As TIME_ZONE_INFORMATION
'    Dim lReturn As Long
'    Dim C As Integer
'    Dim thisString As String
'    Dim strData As String
'
'    lReturn = GetTimeZoneInformation(uInfo)
'
'    If lReturn = TIME_ZONE_ID_DAYLIGHT Then
'        'IsDLSavings = True
'    End If
'
'    Debug.Print "The computer's time zone is: ";
'    For C = 0 To 32 ' the array's range is from 0 to 31
'        thisString = thisString & Chr$(uInfo.StandardName(C))
'        'Debug.Print thisString
'    Next C
'
'    strData = Replace(thisString, Chr(0), vbNullString)
    
    formattedString = Format$(usableObsTime, "dd mmm yyyy hh:nn:ss")
           
    Select Case DatePart("w", usableObsTime)
        Case vbSunday
            dayOfWeek = "Sunday"
        Case vbMonday
            dayOfWeek = "Monday"
        Case vbTuesday
            dayOfWeek = "Tuesday"
        Case vbWednesday
            dayOfWeek = "Wednesday"
        Case vbThursday
            dayOfWeek = "Thursday"
        Case vbFriday
            dayOfWeek = "Friday"
        Case vbSaturday
            dayOfWeek = "Saturday"
    End Select

    humanFriendlyDate = dayOfWeek & " " & formattedString & " " & "GMT+0000 (GMT Standard Time)"
        
    On Error GoTo 0
    Exit Function

humanFriendlyDate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure humanFriendlyDate, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: getNewIcaoLocations
' Purpose: function to open a web-page and obtain the new icao locations data file
' Procedure Kind: Sub
' Procedure Access: Public
' Author: beededea
' Date: 25/03/2024
' ----------------------------------------------------------------
Public Sub getNewIcaoLocations()
    On Error GoTo getNewIcaoLocations_Error
    
    Dim webAddress As String: webAddress = vbNullString
    Dim answerMsg As String: answerMsg = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim firstFileDate As Date: firstFileDate = #1/1/2000 12:00:00 PM#
    Dim secondFileDate As Date: secondFileDate = #1/1/2000 12:00:00 PM#

    answerMsg = "This menu option goes online and downloads/updates the airports data file. Do you wish to proceed?"
    answer = msgBoxA(answerMsg, vbYesNo + vbExclamation, "Obtain the new icao locations data file", False)

    If answer = vbYes Then
        
        If gblAirportsURL = vbNullString Then
            webAddress = "https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports.dat"
        Else
            webAddress = gblAirportsURL
        End If

        If Not fDirExists(App.path & "\dat") Then
            answerMsg = "Error: The output data file folder does not exist."
            answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Downloaded file failed", False)
            Exit Sub
        End If
        
        ' sample the file date first
        firstFileDate = FileDateTime(App.path & "\dat\icao_codes.dat")
        
        ' Download the file.
        Call URLDownloadToFile(0, webAddress, App.path & "\dat\icao_codes.dat", 0, 0)
        
        ' test the new file date
        secondFileDate = FileDateTime(App.path & "\dat\icao_codes.dat")
        
         ' confirm the two dates are different
        If firstFileDate <> secondFileDate Then
            answerMsg = "Downloaded ICAO data file successfully."
            answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Downloaded ICAO data file", False)
        Else
            answerMsg = "Error: There was a problem downloading the ICAO data file."
            answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "Downloaded file failed", False)
        End If
        
    End If

    
    On Error GoTo 0
    Exit Sub

getNewIcaoLocations_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getNewIcaoLocations, line " & Erl & "."

End Sub
'=====================
'End function
'=====================

'=======================================================
' searches the icao data file for a specific code
'=======================================================
' ----------------------------------------------------------------
' Procedure Name: searchIcaoFile
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter loc (String):
' Parameter metarPref (String):
' Parameter icaoLocation1 (String):
' Parameter icaoLocation2 (String):
' Parameter icaoLocation3 (String):
' Parameter icaoLocation4 (String):
' Parameter icaoLocation5 (String):
' Parameter searchCount (Long):
' Return Type: String
' Author: beededea
' Date: 14/04/2024
' ----------------------------------------------------------------
Private Function searchIcaoFile(ByVal loc As String, ByVal metarPref As String, ByRef icaoLocation1 As String, ByRef icaoLocation2 As String, ByRef icaoLocation3 As String, ByRef icaoLocation4 As String, ByRef icaoLocation5 As String, ByRef searchCount As Long) As String
    On Error GoTo searchIcaoFile_Error
    Dim icaoDataArray() As String ' array
    Dim i As Long: i = 0
    Dim fnd As Boolean: fnd = False
    Dim thisKey As String: thisKey = vbNullString
    Dim lookFor As String: lookFor = vbNullString
    Dim retStr As String: retStr = vbNullString
    Dim splitIcaoData() As String
    Dim lReturn As Boolean: lReturn = False
    Dim icaoDataFile As String: icaoDataFile = vbNullString
    
    Dim answerMsg As String: answerMsg = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    
    debugFlg = 0
    
    lookFor = loc
    icaoDataFile = App.path & "\dat\icao_codes.dat"
    lReturn = FileSystemReadFile(icaoDataArray(), icaoDataFile)
    If lReturn = False Then Exit Function
    
    ' If metarPref = "Location" Then ReDim gblValidLocations(UBound(icaoDataArray)) ' set the valid list to be able potentially to contain all members of the array, also wipe it.
    
    ' no need to redim the collValidLocations collection we just empty it
    collValidLocations.RemoveAll
    
    If debugFlg = 1 Then Debug.Print ("searchIcaoFile - " & " search started for " & lookFor)
    
    For i = 0 To UBound(icaoDataArray) - 1 Step 1
        icaoDataArray(i) = UCase$(icaoDataArray(i))
        fnd = InStr(icaoDataArray(i), lookFor)
        If fnd = True Then
            splitIcaoData = Split(icaoDataArray(i), ",")
            icaoLocation1 = Replace(splitIcaoData(1), """", "") '
            If debugFlg = 1 Then
                Debug.Print ("%searchIcaoFile - icaoLocation1 " & " = " & icaoLocation1)
            End If
            icaoLocation2 = Replace(splitIcaoData(2), """", "") '  city
            If debugFlg = 1 Then
                Debug.Print ("%searchIcaoFile - icaoLocation2 " & " = " & icaoLocation2)
            End If
            icaoLocation3 = Replace(splitIcaoData(3), """", "") '  country
            If debugFlg = 1 Then
                Debug.Print ("%searchIcaoFile - icaoLocation3 " & " = " & icaoLocation3)
            End If
            icaoLocation4 = Replace(splitIcaoData(4), """", "") ' // airport code LHR
            If debugFlg = 1 Then
                Debug.Print ("%searchIcaoFile - icaoLocation4 " & " = " & icaoLocation4)
            End If
            icaoLocation5 = Replace(splitIcaoData(5), """", "") ' // icao code
            If icaoLocation5 = "\N" Then
                icaoLocation5 = Replace(splitIcaoData(4), """", "")
            End If ' // use the IATA code as the ICAO code is missing}
            If debugFlg = 1 Then
                Debug.Print ("%searchIcaoFile - icaoLocation5 " & " = " & icaoLocation5)
            End If
            retStr = icaoDataArray(i)
            'Preferences.icao.Value = icaoLocation5
            If debugFlg = 1 Then
                Debug.Print ("searchIcaoFile - found " & lookFor & " in " & retStr)
            End If
            
            ' with a scripting dictionary collection the string/key pair is the other way around! ie. thisKey, retStr
            If metarPref = "ICAO" Then ' single search result
                searchIcaoFile = retStr
                collValidLocations.Add retStr, "key" & "1"  ' key is the loop int. returns a single item in a collection
                Exit For
            Else
                'gblValidLocations(searchCount) = retStr ' returns a global array of matching locations, properties cannot return arrays
                searchCount = searchCount + 1
                
                thisKey = "key" & CStr(searchCount)
                collValidLocations.Add retStr, thisKey  ' key is the search count returning a collection of matching locations
            End If
        End If
    Next
    
    If metarPref = "Location" Then
        ' ReDim Preserve gblValidLocations(searchCount - 1) ' redimension to the actual size
        ' no need to redim the collValidLocations collection as it did not need to be dimmed in the first place
        
        ' if the search count is a mere 1 then we have the details
        If searchCount = 1 Then
            searchIcaoFile = retStr ' return full string
            mIcaoLocation = icaoLocation1
            fnd = True
        Else
            searchIcaoFile = "multiple locations found" ' return full string
            mIcaoLocation = vbNullString
            fnd = True
        End If
        
        ' if there is more than one in the search count then populate a combobox with the contents
        
    End If

    If icaoLocation5 = vbNullString Then
        answerMsg = "That city does not have a valid ICAO code assigned. " & vbCrLf & vbCrLf & " Please select another location."
        answer = msgBoxA(answerMsg, vbOKOnly + vbExclamation, "No valid ICAO code", False)
    End If
    If fnd = False Then
        If debugFlg = 1 Then
           Debug.Print ("searchIcaoFile - lookFor not found " & lookFor)
        End If
        'txtSearching.Data = "Failed to find ICAO code " & lookFor
        'Preferences.icao.Value = vbNullString
        loc = vbNullString
        searchIcaoFile = vbNullString
        icaoLocation5 = vbNullString
    End If
    

    On Error GoTo 0
    Exit Function

searchIcaoFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure searchIcaoFile, line " & Erl & "."

End Function
'==================================
'End function
'==================================

' ----------------------------------------------------------------
' Procedure Name: FileSystemReadFile
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter icaoDataArray (String):
' Parameter sFileName (String):
' Return Type: Boolean
' Author: beededea
' Date: 25/03/2024
' ----------------------------------------------------------------
Private Function FileSystemReadFile(ByRef icaoDataArray() As String, ByVal sFileName As String) As Boolean
    Dim iFileNum As Long
    Dim lFileLen As Long
    Dim sBuffer As String
    
    On Error GoTo ErrFailed

    'Open File
    iFileNum = FreeFile
    Open sFileName For Binary Access Read As #iFileNum
    'Get the size of the file
    lFileLen = LOF(iFileNum)
    If lFileLen Then
        'Create output buffer
        sBuffer = String(lFileLen, " ")
        'Read contents of file
        Get iFileNum, 1, sBuffer
        'Split the file contents
        icaoDataArray = Split(sBuffer, vbLf)
    End If

    Close #iFileNum
    'Return success
    FileSystemReadFile = True

    Exit Function

ErrFailed:
    Debug.Assert False
    Debug.Print Err.Description
    FileSystemReadFile = False
    'Close file
    If iFileNum Then
        Close #iFileNum
    End If
End Function



'Private Function StrToObject(Str As String) As Object
'    Dim O As Object
'    'For Each O In Me
'        'If O.Name = Str Then
'            Set StrToObject = O
'            StrToObject.Name = Str
'        'End If
'    'Next
'End Function


' ----------------------------------------------------------------
' Procedure Name: fGetDistance
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter visibility_statute_mi ():
' Author: beededea
' Date: 19/04/2024
' ----------------------------------------------------------------
Private Function fGetDistance(ByVal visibility_statute_mi As String, ByVal isday As Boolean) As String
    Dim metricEquivalent As String: metricEquivalent = vbNullString
    Dim resultingString As String: resultingString = vbNullString
    Dim additionalText As String: additionalText = vbNullString
    Dim startText As String: startText = vbNullString
    Dim findPlus As Integer: findPlus = 0
    
    On Error GoTo fGetDistance_Error
    
    ' remove the + at night
    findPlus = InStr(visibility_statute_mi, "+")
    If findPlus > 0 And isday = False Then
        visibility_statute_mi = Left$(visibility_statute_mi, findPlus - 1)
    End If
        
    If isday = True Then
        startText = "Daylight, "
        additionalText = vbNullString
    Else
        startText = "Night time, "
        additionalText = ", visibility reduced"
    End If
    
    If gblMetricImperial = "1" Then ' "metric"
        ' change the value to metric measurements
        metricEquivalent = CStr(Int(Val(visibility_statute_mi) * 1.609))
        resultingString = startText & metricEquivalent & " km" & additionalText
    Else
        resultingString = startText & visibility_statute_mi & " miles" & additionalText
    End If
    
    fGetDistance = resultingString
    
    On Error GoTo 0
    Exit Function

fGetDistance_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetDistance, line " & Erl & "."

End Function
  
    
' ----------------------------------------------------------------
' Procedure Name: fGetPrecipitation
' Purpose: ' determine the precipitation
' Procedure Kind: Function
' Procedure Access: Private
' Parameter precip_in (String):
' Return Type: String
' Author: beededea
' Date: 19/04/2024
' ----------------------------------------------------------------
Private Function fGetPrecipitation(ByVal precip_in As String) As String
    On Error GoTo fGetPrecipitation_Error
    Dim metricEquivalent As String: metricEquivalent = vbNullString
    Dim resultingString As String: resultingString = vbNullString
    
    If gblMetricImperial = "1" Then ' "metric"
        ' change the value to metric measurements
        metricEquivalent = CStr(Int(Val(precip_in) * 2.54))
        resultingString = metricEquivalent & "cm"
    Else
        resultingString = precip_in & "inches"
    End If
    
    fGetPrecipitation = resultingString
    
    On Error GoTo 0
    Exit Function

fGetPrecipitation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetPrecipitation, line " & Erl & "."

End Function

'---------------------------------------------------------------------------------------
' Procedure : tmrSampler_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrSampler_Timer()
    
    On Error GoTo tmrSampler_Timer_Error
    
    ' any abnormal, triggered (non integer) interval will be reset, this is 2.5 secs.
    
    If mSamplingInterval = 0 Then mSamplingInterval = 5
    If tmrSampler.Interval = 5000 Then tmrSampler.Interval = mSamplingInterval * 5000
    
    'MsgBox " sampling the METAR data, interval " & mSamplingInterval & " secs"
    
    'WeatherMeteo.GetMetar = True
    Call WeatherMeteo.getData

   On Error GoTo 0
   Exit Sub

tmrSampler_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampler_Timer of Class Module cwOverlayTemp"
End Sub
'---------------------------------------------------------------------------------------
' Property  : ValidICAO
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ValidICAO() As Boolean
    Dim icaoToFind As String: icaoToFind = vbNullString
    Dim fullIcaoLocation As String: fullIcaoLocation = vbNullString
    Dim searchArray() As String
    Dim searchCount As Long: searchCount = 0

    On Error GoTo ValidICAOGet_Error

    icaoToFind = mIcaoToTest ' eg. "EGSH"
    fullIcaoLocation = searchIcaoFile(icaoToFind, "ICAO", icaoLocation1, icaoLocation2, icaoLocation3, icaoLocation4, icaoLocation5, searchCount)
    
    mIcaoLocation = icaoLocation1 ' the airport name

    If icaoLocation5 <> vbNullString Then ' icao code
        mValidICAO = True
        ValidICAO = mValidICAO
        mIcaoToTest = icaoLocation5 ' find the icao code and set the private var so it can be accessed via a property
        fSelector.SelectorForm.Widgets("sbtnselect").Widget.Enabled = True
    Else
        ValidICAO = False
        mValidICAO = False
        fSelector.SelectorForm.Widgets("sbtnselect").Widget.Enabled = False

    End If

   On Error GoTo 0
   Exit Property

ValidICAOGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ValidICAO of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : ValidICAO
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ValidICAO(ByVal newValue As Boolean)

    On Error GoTo ValidICAOLet_Error

    If mValidICAO <> newValue Then mValidICAO = newValue Else Exit Property

    
   On Error GoTo 0
   Exit Property

ValidICAOLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ValidICAO of Class Module cwOverlayTemp"
End Property


'---------------------------------------------------------------------------------------
' Property  : ValidLocationCount
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ValidLocationCount() As Long ' returns the number of matching locations
    Dim icaoToFind As String: icaoToFind = vbNullString
    Dim fullIcaoLocation As String: fullIcaoLocation = vbNullString
    Dim searchCount As Long: searchCount = 0
    
    On Error GoTo ValidLocationCountGet_Error

    icaoToFind = mStringToTest ' eg. "EGSH"
    fullIcaoLocation = searchIcaoFile(icaoToFind, "Location", icaoLocation1, icaoLocation2, icaoLocation3, icaoLocation4, icaoLocation5, searchCount)
    
    'mIcaoLocation = icaoLocation1 ' the airport name

 '   If icaoLocation5 <> vbNullString Then ' icao code
                
        mValidLocationCount = searchCount ' returns a count only
        ValidLocationCount = mValidLocationCount
        
        If mValidLocationCount = 0 Then fSelector.SelectorForm.Widgets("sbtnselect").Widget.Enabled = False

        
        'mIcaoToTest = icaoLocation5 ' find the icao code and set the private var so it can be accessed via a property
'    Else
'        ValidLocationCount = False
'        mValidLocationCount = False
'    End If

   On Error GoTo 0
   Exit Property

ValidLocationCountGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property ValidLocationCount of Class Module cwOverlayTemp"
End Property
'
'---------------------------------------------------------------------------------------
' Property  : icaoLocation
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get icaoLocation() As String
   
    On Error GoTo icaoLocationGet_Error

    icaoLocation = mIcaoLocation

   On Error GoTo 0
   Exit Property

icaoLocationGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property icaoLocation of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : icaoLocation
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let icaoLocation(ByVal newValue As String)
    
    On Error GoTo icaoLocationLet_Error

    If mIcaoLocation <> newValue Then mIcaoLocation = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

icaoLocationLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property icaoLocation of Class Module cwOverlayTemp"
End Property


'---------------------------------------------------------------------------------------
' Property : WeatherDetails
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get WeatherDetails() As String
   On Error GoTo WeatherDetailsGet_Error

   WeatherDetails = mWeatherDetails

   On Error GoTo 0
   Exit Property

WeatherDetailsGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WeatherDetails of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : WeatherDetails
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let WeatherDetails(ByVal newValue As String)
   On Error GoTo WeatherDetailsLet_Error

   If mWeatherDetails <> newValue Then mWeatherDetails = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

WeatherDetailsLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WeatherDetails of Class Module cwOverlayTemp"
End Property

'---------------------------------------------------------------------------------------
' Property : WindSpeed
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get windSpeed() As String
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
Public Property Let windSpeed(ByVal newValue As String)
   On Error GoTo WindSpeedLet_Error

   If mWindSpeed <> newValue Then mWindSpeed = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

WindSpeedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WindSpeed of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : IcaoToTest
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get IcaoToTest() As String
   
    On Error GoTo IcaoToTestGet_Error

    IcaoToTest = mIcaoToTest

   On Error GoTo 0
   Exit Property

IcaoToTestGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IcaoToTest of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : IcaoToTest
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let IcaoToTest(ByVal newValue As String)
    
    On Error GoTo IcaoToTestLet_Error

    If mIcaoToTest <> newValue Then mIcaoToTest = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

IcaoToTestLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property IcaoToTest of Class Module cwOverlayTemp"
End Property

'---------------------------------------------------------------------------------------
' Property  : StringToTest
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get StringToTest() As String
   
    On Error GoTo StringToTestGet_Error

    StringToTest = mStringToTest

   On Error GoTo 0
   Exit Property

StringToTestGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property StringToTest of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : StringToTest
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let StringToTest(ByVal newValue As String)
    
    On Error GoTo StringToTestLet_Error

    If mStringToTest <> newValue Then mStringToTest = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

StringToTestLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property StringToTest of Class Module cwOverlayTemp"
End Property



' ----------------------------------------------------------------
' Procedure Name: Change
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Parameter WeHaveData (Boolean):
' Author: beededea
' Date: 06/06/2024
' ----------------------------------------------------------------
Private Sub Change(ByVal wehavedata As Boolean)
    On Error GoTo Change_Error
    
    RaiseEvent OnChange(wehavedata)
    If mWeHaveData = True Then tmrSampler.Enabled = True
    
    On Error GoTo 0
    Exit Sub

Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Change, line " & Erl & "."

End Sub

'---------------------------------------------------------------------------------------
' Property  : WeHaveData
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get wehavedata() As Boolean
   
    On Error GoTo WeHaveDataGet_Error

    wehavedata = mWeHaveData

   On Error GoTo 0
   Exit Property

WeHaveDataGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WeHaveData of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : WeHaveData
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let wehavedata(ByVal newValue As Boolean)
    
    On Error GoTo WeHaveDataLet_Error

    If mWeHaveData <> newValue Then mWeHaveData = newValue Else Exit Property
    
    Call Change(newValue)

   On Error GoTo 0
   Exit Property

WeHaveDataLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WeHaveData of Class Module cwOverlayTemp"
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
    
    WeatherMeteo.SamplerEnabled = mTicking
    'tmrSampler.Enabled = mTicking
    'tmrPressureDrop.Enabled = mTicking
    
    WeatherMeteo.TestPressureDrop = True
    
    menuForm.mnuSwitchOff.Checked = mTicking
    menuForm.mnuTurnFunctionsOn.Checked = Not mTicking
    
    If mTicking = True Then
        gblGaugeFunctions = "1"
    Else
        gblGaugeFunctions = "0"
    End If
    
    sPutINISetting "Software\PzTemperatureGauge", "gaugeFunctions", gblGaugeFunctions, gblSettingsFile


   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwoverlaytemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwoverlaytemp"
End Property



'---------------------------------------------------------------------------------------
' Property  : TemperatureValue
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get TemperatureValue() As Single
   
    On Error GoTo TemperatureValueGet_Error

    TemperatureValue = mTemperatureValue

   On Error GoTo 0
   Exit Property

TemperatureValueGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TemperatureValue of Class Module cwOverlayTemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : TemperatureValue
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let TemperatureValue(ByVal newValue As Single)
    
    On Error GoTo TemperatureValueLet_Error

    If mTemperatureValue <> newValue Then mTemperatureValue = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

TemperatureValueLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property TemperatureValue of Class Module cwOverlayTemp"
End Property
