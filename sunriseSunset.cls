VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwSunriseSunset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed

Option Explicit

'   SunriseSunset Class (2013-04-21)
'
' OVERVIEW
'
'
' DESCRIPTION
'
'   Provides sunrise and sunset times for specified date and position.
'   All dates are UTC.  Year is 4-digit.  Month is 1-12.  Day is 1-31.
'   Longitude is positive for east, negative for west. Latitude is
'   positive for north, negative for south.
'
' SAMPLE USAGE
'
'   dim tokyo = new SunriseSunset( 2011, 1, 19, 35+40/60, 139+45/60)
'   tokyo.sunriseUtcHours()      --> 21.8199 = 21:49 GMT
'   tokyo.sunsetUtcHours()       --> 7.9070  = 07:54 GMT
'   tokyo.sunriseLocalHours(9)   --> 6.8199  = 06:49 at GMT+9
'   tokyo.sunsetLocalHours(9)    --> 16.9070 = 16:54 at GMT+9
'   tokyo.isDaylight(1.5)        --> true
'
'   var losangeles = new SunriseSunset( 2011, 1, 19, 34.05, -118.233333333 )
'   etc.

Private thisMonth As String
Private thisYear As String
Private thisDay As String

Private lnghour As Double
Private sinDeclination As Double
Private cosDeclination As Double
Private rightAscension As Double
Private approximateTime As Double
Private returnSunString As String

Const pi As Double = 3.14159265358979
Const to_rad  As Double = pi / 180
Const zenith  As Double = 90 + 50 / 60 '   offical      = 90 degrees 50'
                          '   civil        = 96 degrees
                          '   nautical     = 102 degrees
                          '   astronomical = 108 degrees
Private mLatitude As Double
Private mLongitude As Double
Private mFullYear As String
Private mUtcMonth As String
Private mUtcDay As String
Private mIsDayLight As Boolean



' ----------------------------------------------------------------
' Procedure Name: Class_Initialize
' Purpose:
' Procedure Kind: Constructor (Initialize)
' Procedure Access: Private
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Sub Class_Initialize()

    
    On Error GoTo Class_Initialize_Error

    mLongitude = 1.283
    mLatitude = 52.676 ' we will obtain these from the feed
    
    debugFlg = 0

    
    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize, line " & Erl & "."

End Sub
'---------------------------------------------------------------------------------------
' Property : isDayLight
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get isDayLight() As Boolean
    Dim ret As Double: ret = 0
    
    Dim sunriseHours As Double: sunriseHours = 0
    Dim sunsetHours As Double: sunsetHours = 0

    On Error GoTo isDayLightGet_Error
    
    debugFlg = 1

    ' set to true for sunrise
    Call calcSunPosition(True, Date, ret)
    sunriseHours = getLocalMeanTime(True, lnghour, sinDeclination, cosDeclination, rightAscension, approximateTime, returnSunString)
    
    ' set to false for sunset
    Call calcSunPosition(False, Date, ret)
    sunsetHours = getLocalMeanTime(False, lnghour, sinDeclination, cosDeclination, rightAscension, approximateTime, returnSunString)

    debugFlg = 1

    If debugFlg = 1 Then Debug.Print (" sunriseHours " & sunriseHours)
    If debugFlg = 1 Then Debug.Print (" sunsetHours " & sunsetHours)
    
    debugFlg = 0
    
    isDayLight = testIsDaylight(sunsetHours, sunriseHours)
    
    mIsDayLight = isDayLight

    
   On Error GoTo 0
   Exit Property

isDayLightGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property isDayLight of Class Module cwSunriseSunset"
End Property
'---------------------------------------------------------------------------------------
' Property  : isDayLight
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let isDayLight(ByVal newValue As Boolean)
   On Error GoTo isDayLightLet_Error

   If mIsDayLight <> newValue Then mIsDayLight = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

isDayLightLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property isDayLight of Class Module cwSunriseSunset"
End Property


'---------------------------------------------------------------------------------------
' Property : Latitude
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Latitude() As String
   On Error GoTo LatitudeGet_Error

   Latitude = mLatitude

   On Error GoTo 0
   Exit Property

LatitudeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Latitude of Class Module cwSunriseSunset"
End Property
'---------------------------------------------------------------------------------------
' Property  : Latitude
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Latitude(ByVal newValue As String)
   On Error GoTo LatitudeLet_Error

   If mLatitude <> newValue Then mLatitude = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LatitudeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Latitude of Class Module cwSunriseSunset"
End Property

'---------------------------------------------------------------------------------------
' Property : Longitude
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Longitude() As String
   On Error GoTo LongitudeGet_Error

   Longitude = mLongitude

   On Error GoTo 0
   Exit Property

LongitudeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Longitude of Class Module cwSunriseSunset"
End Property
'---------------------------------------------------------------------------------------
' Property  : Longitude
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Longitude(ByVal newValue As String)
   On Error GoTo LongitudeLet_Error

   If mLongitude <> newValue Then mLongitude = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LongitudeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Longitude of Class Module cwSunriseSunset"
End Property

'---------------------------------------------------------------------------------------
' Property : FullYear
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get FullYear() As String
   On Error GoTo FullYearGet_Error

   FullYear = mFullYear

   On Error GoTo 0
   Exit Property

FullYearGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FullYear of Class Module cwSunriseSunset"
End Property
'---------------------------------------------------------------------------------------
' Property  : FullYear
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let FullYear(ByVal newValue As String)
   On Error GoTo FullYearLet_Error

   If mFullYear <> newValue Then mFullYear = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

FullYearLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FullYear of Class Module cwSunriseSunset"
End Property

'---------------------------------------------------------------------------------------
' Property : UtcMonth
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get UtcMonth() As String
   On Error GoTo UtcMonthGet_Error

   UtcMonth = mUtcMonth

   On Error GoTo 0
   Exit Property

UtcMonthGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UtcMonth of Class Module cwSunriseSunset"
End Property
'---------------------------------------------------------------------------------------
' Property  : UtcMonth
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let UtcMonth(ByVal newValue As String)
   On Error GoTo UtcMonthLet_Error

   If mUtcMonth <> newValue Then mUtcMonth = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

UtcMonthLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UtcMonth of Class Module cwSunriseSunset"
End Property

'---------------------------------------------------------------------------------------
' Property : utcDay
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get utcDay() As String
   On Error GoTo utcDayGet_Error

   utcDay = mUtcDay

   On Error GoTo 0
   Exit Property

utcDayGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property utcDay of Class Module cwSunriseSunset"
End Property
'---------------------------------------------------------------------------------------
' Property  : utcDay
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let utcDay(ByVal newValue As String)
   On Error GoTo utcDayLet_Error

   If mUtcDay <> newValue Then mUtcDay = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

utcDayLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property utcDay of Class Module cwSunriseSunset"
End Property

Private Function thisSin(ByVal deg As Double) As Double
    thisSin = Sin(deg * pi / 180)
    'Debug.Print ("%SRS-O thisSin " & thisSin)
End Function
Private Function thisCos(ByVal deg As Double) As Double
    thisCos = Cos(deg * pi / 180)
    'Debug.Print ("%SRS-O thisCos " & thisCos)
End Function
Private Function thisTan(ByVal deg As Double) As Double
    thisTan = Tan(deg * pi / 180)
    'Debug.Print ("%SRS-O thisTan " & thisTan)
End Function
Private Function thisAsin(ByVal x As Double) As Double
    thisAsin = Atn(x / Sqr(-x * x + 1))
    'Debug.Print ("%SRS-O thisAsin " & thisAsin)
End Function

Public Function Asn(ByVal x As Double) As Double
' x is double
' Equivalent of spreadsheet Asin(x)
    Asn = 2 * Atn(x / (1 + Sqr(1 - x * x))) ' Asn is double
End Function

Private Function thisAtan(ByVal x As Double) As Double
    thisAtan = (180 / pi) * Atn(x)
    'Debug.Print ("%SRS-O thisAtan " & thisAtan)
End Function
' ----------------------------------------------------------------
' Procedure Name: calcSunPosition
' Purpose: determines the Sun's declination and position, run once for sunrise and again for sunset
' Procedure Kind: Function
' Procedure Access: Private
' Parameter returnHours (Double):
' Return Type: Boolean
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Sub calcSunPosition(ByVal rising As Boolean, ByVal dateInput As Date, ByRef returnHours As Double)

    
    Dim N1 As Integer: N1 = 0
    Dim N2 As Integer: N2 = 0
    Dim N3 As Integer: N3 = 0
    Dim N As Integer: N = 0
    Dim RA As Double
    Dim meanAnomaly As Double: meanAnomaly = 0
    Dim trueLongitude As Double: trueLongitude = 0
    Dim returnSunString  As String: returnSunString = vbNullString
    Dim T As Double: T = 0
    Dim M As Double: M = 0
    Dim Lquadrant As Double
    Dim RAquadrant As Double
    Dim longi As Double
    Dim thisDateInput As Date
    
    On Error GoTo calcSunPosition_Error
    
        
    debugFlg = 0
    
    thisDateInput = dateInput
    thisMonth = Month(thisDateInput) - 1
    thisYear = Year(thisDateInput) - 1900
    thisDay = fUtcDayOfWeek(thisDateInput)
    
    If debugFlg = 1 Then Debug.Print ("************************************************************")
    If debugFlg = 1 Then Debug.Print ("%SFN-O latitude " & Latitude)
    If debugFlg = 1 Then Debug.Print ("%SFN-O longitude " & Longitude)
    If debugFlg = 1 Then Debug.Print ("%SFN-I year " & thisYear)
    If debugFlg = 1 Then Debug.Print ("%SFN-I month " & thisMonth)
    If debugFlg = 1 Then Debug.Print ("%SFN-I day " & thisDay)
    
    ' calculate the day of the year
        
    N1 = Fix(275 * thisMonth / 9)
    If debugFlg = 1 Then Debug.Print ("%SFN-I N1 " & N1)
    
    N2 = Fix((thisMonth + 9) / 12)
    If debugFlg = 1 Then Debug.Print ("%SFN-I N2 " & N2)
    
    N3 = (1 + Fix((thisYear - 4 * Fix(thisYear / 4) + 2) / 3))
    If debugFlg = 1 Then Debug.Print ("%SFN-I N3 " & N3)
    
    N = N1 - (N2 * N3) + val(thisDay) - 30
    If debugFlg = 1 Then Debug.Print ("%SFN-I N " & N)
    
    ' convert the longitude to hour value and calculate an approximate time
    
    lnghour = mLongitude / 15
    If (rising) Then
        approximateTime = N + ((6 - lnghour) / 24) 'return
    Else
        approximateTime = N + ((18 - lnghour) / 24) 'return
    End If
    If debugFlg = 1 Then Debug.Print ("%SFN-I approximateTime " & approximateTime)

    ' calculate the Sun's mean anomaly

    T = approximateTime
    meanAnomaly = (0.9856 * T) - 3.289 'return
    If debugFlg = 1 Then Debug.Print ("%SFN-I meanAnomaly " & meanAnomaly)
    
    ' calculate the Sun's true longitude
    
    M = meanAnomaly
    longi = M + (1.916 * thisSin(M)) + (0.02 * thisSin(2 * M)) + 282.634
    
    If (longi < 0) Then
       longi = longi + 360
    ElseIf (longi >= 360) Then
       longi = longi - 360
    End If
    trueLongitude = longi 'Mod 360 'return

    If debugFlg = 1 Then Debug.Print ("%SFN-I trueLongitude " & trueLongitude)
    
    ' calculate the Sun's right ascension
    
    RA = (1 / to_rad) * Atn(0.91764 * Tan(to_rad * trueLongitude)) 'ra  = double
    If (RA < 0) Then
       RA = RA + 360
    ElseIf (RA >= 360) Then
       RA = RA - 360
    End If

    ' right ascension value needs to be in the same quadrant as trueLongitude
    Lquadrant = (Fix(trueLongitude / 90)) * 90
    RAquadrant = (Fix(RA / 90)) * 90
    RA = RA + (Lquadrant - RAquadrant)
    
    ' right ascension value needs to be converted into hours
    
    RA = RA / 15

    rightAscension = RA 'return
    If debugFlg = 1 Then Debug.Print ("%SFN-I rightAscension " & rightAscension)
    
    'longi = trueLongitude()
    'sinDeclination = 0.39782 * thisSin(trueLongitude)
    
    ' calculate the Sun's declination
    
    sinDeclination = 0.39782 * Sin(to_rad * trueLongitude)
    
    'cosDeclination = Cos(thisAsin(sinDeclination))
    cosDeclination = Cos(Asn(sinDeclination))
        
    If debugFlg = 1 Then Debug.Print ("%SFN-I sinDeclination " & sinDeclination)
    If debugFlg = 1 Then Debug.Print ("%SFN-I cosDeclination " & cosDeclination)
    
    If debugFlg = 1 Then Debug.Print ("************************************************************")

    On Error GoTo 0
    Exit Sub

calcSunPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcSunPosition, line " & Erl & "."

End Sub





' ----------------------------------------------------------------
' Procedure Name: testIsDaylight
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter sunsetHours (Double):
' Parameter sunriseHours (Double):
' Return Type: Boolean
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function testIsDaylight(ByVal sunsetHours As Double, ByRef sunriseHours As Double) As Boolean
    
    Dim utcCurrentHours As Double: utcCurrentHours = 0
    
    On Error GoTo testIsDaylight_Error
    
    debugFlg = 1
    
    ' decimal time
    utcCurrentHours = Hour(Now) + (Minute(Now) / 60)
                
    If (sunsetHours < sunriseHours) Then
        ' Either the sunrise or sunset time is for tomorrow
        If (utcCurrentHours > sunriseHours) Then
            testIsDaylight = True 'return
        ElseIf (utcCurrentHours < sunsetHours) Then
            testIsDaylight = True 'return
        Else
            testIsDaylight = False 'return
        End If
    End If
    
    If (utcCurrentHours >= sunriseHours) Then
        testIsDaylight = utcCurrentHours < sunsetHours ' true/false
    End If

    If debugFlg = 1 Then Debug.Print ("%SFN-I testIsDaylight " & testIsDaylight)
    
    debugFlg = 0
    
    On Error GoTo 0
    Exit Function

testIsDaylight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testIsDaylight, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: getLocalMeanTime (.js version is LocalMeanTime)
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter rising (Boolean):
' Parameter sinDeclination (Double):
' Parameter cosDeclination (Double):
' Parameter rightAscension (Double):
' Parameter approximateTime (Double):
' Parameter returnSunString (String):
' Return Type: Double
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function getLocalMeanTime(ByVal rising As Boolean, ByVal lnghour As Double, ByVal sinDeclination As Double, ByVal cosDeclination As Double, ByVal rightAscension As Double, ByVal approximateTime As Double, ByRef returnSunString As String) As Double
    Dim Tt As Double: Tt = 0
    Dim T As Double: T = 0
    Dim H As Double: H = 0
    Dim UT As Double: UT = 0
    Dim RA As Double: RA = 0
    Dim cosH As Double: cosH = 0
    Dim cosRadZenith As Double: cosRadZenith = 0
    Dim sinDecLatitude As Double: sinDecLatitude = 0
    Dim cosDecLatitude As Double: cosDecLatitude = 0
    Dim zenLat As Double: zenLat = 0
    
    On Error GoTo getLocalMeanTime_Error
    
    debugFlg = 1
    
    If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< STARTS ")
     
    ' calculate the Sun's local hour angle
    ' cosH = Cos(to_rad * zenith) - (sinDeclination * Sin(to_rad * mLatitude)) / (cosDeclination * Cos(to_rad * mLatitude)) ' FAILS
    
    cosRadZenith = Cos(to_rad * zenith)
    sinDecLatitude = sinDeclination * Sin(to_rad * mLatitude)
    cosDecLatitude = cosDeclination * Cos(to_rad * mLatitude)
    zenLat = cosRadZenith - sinDecLatitude
    
    cosH = zenLat / cosDecLatitude
    
    ' the earlier calculation of cosH above has a coercion problem that results in a completely incorrect value...
    ' I have forced the vars to doubles - I am not going to research further - this fix will do
             
     If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime cosH = " & cosH)
     
     If (cosH > 1) Then
         returnSunString = "the sun never rises on this location (on the specified date)"
         getLocalMeanTime = -1
     ElseIf (cosH < -1) Then
         returnSunString = "the sun never sets on this location (on the specified date)"
         getLocalMeanTime = -2
     Else
     
         'finish calculating H and convert into hours
         
         If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime rising = " & rising)
         If rising Then
            Dim e As Double
            e = thisAcos(cosH)
             H = 360 - e
             If (rising And debugFlg = 1) Then Debug.Print ("%SRS-O localMeanTime thisAcos(cosH) " & e)
             'H = 360 - thisAcos(cosH)
             'H = 360 - (1 / to_rad) * thisAcos(cosH)
         Else
             H = thisAcos(cosH)
         End If
         If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime H = " & H)

         H = H / 15
         If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime H = H / 15 = " & H)
         
         RA = rightAscension
         If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime RA = " & RA)
         
         T = approximateTime
         If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime t = " & T)
         
         'calculate local mean time of rising/setting
         
         Tt = H + RA - (0.06571 * T) - 6.622
         
         If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime Tt = " & Tt)
         
         getLocalMeanTime = Tt 'return
    End If
    
    ' adjust back to UTC
    UT = getLocalMeanTime - lnghour
    If (UT < 0) Then
       UT = UT + 24
    ElseIf (UT >= 24) Then
       UT = UT - 24
    End If
    getLocalMeanTime = UT
    
    If debugFlg = 1 And rising Then Debug.Print ("%SFN-I getLocalMeanTime  = " & getLocalMeanTime)
    'Debug.Print ("%SFN-I getLocalMeanTime <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDS ")
    
    debugFlg = 0
    
    On Error GoTo 0
    Exit Function

getLocalMeanTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getLocalMeanTime, line " & Erl & "."

End Function
' ----------------------------------------------------------------
' Procedure Name: thisAcos
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter x (Double):
' Return Type: Double
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function thisAcos(x As Double) As Double

    On Error GoTo thisAcos_Error

    thisAcos = (180 / pi) * Acos(x)

    'Debug.Print ("%SRS-O thisAcos " & thisAcos)

    On Error GoTo 0
    Exit Function

thisAcos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thisAcos, line " & Erl & "."

End Function

Public Function Acos(x As Variant) As Variant

    Select Case x
        Case -1
            Acos = 4 * Atn(1)
             
        Case 0:
            Acos = 2 * Atn(1)
             
        Case 1:
            Acos = 0
             
        Case Else:
            Acos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End Select
End Function

'---------------------------------------------------------------------------------------
' Function   : fUtcMonth
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 07/10/2023
' Purpose   : get the number of the month given a month name
'---------------------------------------------------------------------------------------
'
Private Function fUtcMonth(ByVal thisMonth As String, ByVal utcFlag As Boolean) As Integer
    
    On Error GoTo fUtcMonth_Error
    
    ''Debug.Print ("%SRS func fUtcMonth")
    ''Debug.Print ("%SRS-I thisMonth " & thisMonth)
    
    fUtcMonth = Month(CDate(thisMonth & "/1/2000"))
    If utcFlag = True Then fUtcMonth = fUtcMonth - 1 ' convert 'normal month starting number of 1 to starting with 0 UTC

    If fUtcMonth < 0 Or fUtcMonth > 11 Then
        MsgBox ("fUtcMonth: " & thisMonth & " is not a valid month name")
        fUtcMonth = -1 ' return invalid
        
        ''Debug.Print ("%SRS-O abnormal fUtcMonth " & fUtcMonth)
    End If
    
    Debug.Print ("%SRS-O fUtcMonth " & fUtcMonth)
    
    On Error GoTo 0
    Exit Function

fUtcMonth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   fUtcMonth of Module modDaylightSavings"

End Function


' ----------------------------------------------------------------
' Procedure Name: fUtcDayOfWeek
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter vb6DateTime (Date):
' Return Type: Integer
' Author: Dean Beedell (yereverluvinunclebert)
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function fUtcDayOfWeek(vb6DateTime As Date) As Integer
            
    On Error GoTo fUtcDayOfWeek_Error
     Select Case DatePart("w", vb6DateTime)
         Case vbSunday
             fUtcDayOfWeek = 0 '"Sun"
         Case vbMonday
             fUtcDayOfWeek = 1 '"Mon"
         Case vbTuesday
             fUtcDayOfWeek = 2 '"Tue"
         Case vbWednesday
             fUtcDayOfWeek = 3 '"Wed"
         Case vbThursday
             fUtcDayOfWeek = 4 '"Thu"
         Case vbFriday
             fUtcDayOfWeek = 5 '"Fri"
         Case vbSaturday
             fUtcDayOfWeek = 6 ' "Sat"
     End Select
    
    On Error GoTo 0
    Exit Function

fUtcDayOfWeek_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fUtcDayOfWeek, line " & Erl & "."

End Function
