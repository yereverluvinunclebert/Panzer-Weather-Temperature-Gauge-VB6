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

'private  doy = This.fGetDOY()


Private longi As Double
Private H As Double
Private T As Double
Private UT As Double
Private RA As Double
Private Lquadrant As Double
Private RAquadrant As Double

Private cosH As Double
'private  UT = T - This.lngHour
'private  sunriseHours = this.sunriseUtcHours()
Private sunsetHours
Private zenith As Double
Private rising As Boolean
Private lnghour As Double

Const pi As Double = 3.14159265358979
Const to_rad  As Double = pi / 180

Private mLatitude As Double
Private mLongitude As Double
Private mFullYear As String
Private mUtcMonth As String
Private mUtcDay As String

Private sunDeclination As Double
Private cosDec As Double
Private rightAscension As Double
Private approximateTime As Double
Private returnSunString As String


'    locationDayNight = new SunriseSunset(theDate.getYear(), theDate.getMonth(), theDate.getDay(), latitude , longitude  )

'Private Function SunriseSunset(ByVal FullYear, fUtcMonth, utcDay, latitude, longitude)
Private Sub Class_Initialize()

    Dim decimalTime As Double
    Dim ret As Double
    Dim sunriseHours  As Double
    Dim sunsetHours  As Double
    Dim isDaylight  As Boolean
    
    zenith = 90 + 50 / 60 '   offical      = 90 degrees 50'
                        '   civil        = 96 degrees
                        '   nautical     = 102 degrees
                        '   astronomical = 108 degrees
    mLongitude = 1.283
    mLatitude = 52.676 ' we will obtain these from the feed
    
    rising = True ' set to true for sunrise
    Call calcSunPosition(ret)
    sunriseHours = localmeantime2(sunDeclination, cosDec, rightAscension, approximateTime, returnSunString) '32.333
    
    rising = False ' set to false for sunset
    Call calcSunPosition(ret)
    sunsetHours = localmeantime2(sunDeclination, cosDec, rightAscension, approximateTime, returnSunString)

    Debug.Print (">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> sunriseHours " & sunriseHours)
    Debug.Print (">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> sunsetHours " & sunsetHours)
    
    isDaylight = testIsDaylight(sunsetHours, sunriseHours)
    
    '1. first calculate the day of the year
    'Call fGetDOY(Date)
    '2. convert the longitude to hour value and calculate an approximate time
    'Call approximateTime
    '3. calculate the Sun's mean anomaly
    'Call meanAnomaly
    '4. calculate the Sun's true longitude
    'Call trueLongitude
    '5a. calculate the Sun's right ascension
    'Call rightAscension
    '5b. right ascension value needs to be in the same quadrant as L
    '5c. right ascension value needs to be converted into hours
    '6. calculate the Sun's declination
    'Call fsunDeclination
    '7a. calculate the Sun's local hour angle
    'Call localMeanTime
    '7b. finish calculating H and convert into hours
    '8. calculate local mean time of rising/setting
    '9. adjust back to UTC

End Sub


'---------------------------------------------------------------------------------------
' Property : Latitude
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Latitude of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Latitude
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Latitude of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : Longitude
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Longitude of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Longitude
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Longitude of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : FullYear
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FullYear of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : FullYear
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property FullYear of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : UtcMonth
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UtcMonth of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : UtcMonth
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property UtcMonth of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : utcDay
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property utcDay of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : utcDay
' Author    : beededea
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property utcDay of Class Module cwOverlay"
End Property
'SunriseSunset.prototype =
Private Function thisSin(deg) As Double
    thisSin = Sin(deg * pi / 180)
    'Debug.Print ("%SRS-O thisSin " & thisSin)
End Function
Private Function thisCos(deg) As Double
    thisCos = Cos(deg * pi / 180)
    'Debug.Print ("%SRS-O thisCos " & thisCos)
End Function
Private Function thisTan(deg) As Double
    thisTan = Tan(deg * pi / 180)
    'Debug.Print ("%SRS-O thisTan " & thisTan)
End Function
Private Function thisAsin(x) As Double
    thisAsin = Atn(x / Sqr(-x * x + 1))
    'Debug.Print ("%SRS-O thisAsin " & thisAsin)
End Function

Public Function Asn(x)
' x is double
' Equivalent of spreadsheet Asin(x)
Asn = 2 * Atn(x / (1 + Sqr(1 - x * x))) ' Asn is double
End Function

Private Function thisAtan(x) As Double
    thisAtan = (180 / pi) * Atn(x)
    'Debug.Print ("%SRS-O thisAtan " & thisAtan)
End Function
' ----------------------------------------------------------------
' Procedure Name: calcSunPosition
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter returnHours (Double):
' Return Type: Boolean
' Author: beededea
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function calcSunPosition(ByRef returnHours As Double) As Boolean
    On Error GoTo calcSunPosition_Error

    
    Dim N1 As Integer
    Dim N2 As Integer
    Dim N3 As Integer
    Dim N As Integer
    Dim dateinput As Date
    'Dim approximateTime As Double
    Dim meanAnomaly As Double
    Dim trueLongitude As Double
    'Dim rightAscension As Double
    'Dim sunDeclination As Double
    'Dim cosDec As Double
    Dim returnSunString  As String
    Dim localMeanTime  As Double
    Dim UTCTime  As Double

    dateinput = Date
    
    thisMonth = Month(dateinput) - 1
    thisYear = Year(dateinput) - 1900
    thisDay = fUtcDayOfWeek(dateinput)
    
    Debug.Print ("************************************************************")
    
    Debug.Print ("%SFN-O latitude " & Latitude)
    Debug.Print ("%SFN-O longitude " & Longitude)
    Debug.Print ("%SFN-I year " & thisYear)
    Debug.Print ("%SFN-I month " & thisMonth)
    Debug.Print ("%SFN-I day " & thisDay)
        
    N1 = Fix(275 * thisMonth / 9)
    Debug.Print ("%SFN-I N1 " & N1)
    
    N2 = Fix((thisMonth + 9) / 12)
    Debug.Print ("%SFN-I N2 " & N2)
    
    N3 = (1 + Fix((thisYear - 4 * Fix(thisYear / 4) + 2) / 3))
    Debug.Print ("%SFN-I N3 " & N3)
    
    N = N1 - (N2 * N3) + Val(thisDay) - 30
    Debug.Print ("%SFN-I N " & N)
    
    lnghour = mLongitude / 15
    If (rising) Then
        approximateTime = N + ((6 - lnghour) / 24) 'return
    Else
        approximateTime = N + ((18 - lnghour) / 24) 'return
    End If
    Debug.Print ("%SFN-I approximateTime " & approximateTime)


    Dim T As Double
    T = approximateTime
    meanAnomaly = (0.9856 * T) - 3.289 'return
    Debug.Print ("%SFN-I meanAnomaly " & meanAnomaly)
    
    Dim M As Double
    
    M = meanAnomaly
    longi = M + (1.916 * thisSin(M)) + (0.02 * thisSin(2 * M)) + 282.634
    
    If (longi < 0) Then
       longi = longi + 360
    ElseIf (longi >= 360) Then
       longi = longi - 360
    End If
    trueLongitude = longi 'Mod 360 'return

    Debug.Print ("%SFN-I trueLongitude " & trueLongitude)
    
    RA = (1 / to_rad) * Atn(0.91764 * Tan(to_rad * trueLongitude)) 'ra  = double
    If (RA < 0) Then
       RA = RA + 360
    ElseIf (RA >= 360) Then
       RA = RA - 360
    End If

    ' right ascension value needs to be in the same quadrant as longi
    Lquadrant = (Fix(trueLongitude / 90)) * 90
    RAquadrant = (Fix(RA / 90)) * 90
    RA = RA + (Lquadrant - RAquadrant)
    ' right ascension value needs to be converted into hours
    RA = RA / 15

    rightAscension = RA 'return
    Debug.Print ("%SFN-I rightAscension " & rightAscension)
    
    'longi = trueLongitude()
    'sunDeclination = 0.39782 * thisSin(trueLongitude)
    
    sunDeclination = 0.39782 * Sin(to_rad * trueLongitude)
    
    Debug.Print ("%SFN-I sunDeclination " & sunDeclination)
    
    'cosDec = Cos(thisAsin(sunDeclination))
    
    cosDec = Cos(Asn(sunDeclination))
    
    Debug.Print ("%SFN-I cosDec " & cosDec)
        
   
    Debug.Print ("************************************************************")

    On Error GoTo 0
    Exit Function

calcSunPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcSunPosition, line " & Erl & "."

End Function





' ----------------------------------------------------------------
' Procedure Name: testIsDaylight
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter sunsetHours (Double):
' Parameter sunriseHours (Double):
' Return Type: Boolean
' Author: beededea
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function testIsDaylight(ByVal sunsetHours As Double, ByRef sunriseHours As Double) As Boolean
    Dim utcCurrentHours As Double
    
    On Error GoTo testIsDaylight_Error
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
        If utcCurrentHours < sunsetHours Then
            testIsDaylight = True
        Else
            testIsDaylight = False
        End If
    End If

    Debug.Print ("%SFN-I testIsDaylight " & testIsDaylight)

    
    On Error GoTo 0
    Exit Function

testIsDaylight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testIsDaylight, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: localmeantime2
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter rising (Boolean):
' Parameter sunDeclination (Double):
' Parameter cosDec (Double):
' Parameter rightAscension (Double):
' Parameter approximateTime (Double):
' Parameter returnSunString (String):
' Return Type: Double
' Author: beededea
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function localmeantime2(ByVal sunDeclination As Double, ByVal cosDec As Double, ByVal rightAscension As Double, ByVal approximateTime As Double, ByRef returnSunString As String) As Double
    On Error GoTo localmeantime2_Error
     If rising Then Debug.Print ("%SFN-I localmeantime2 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< STARTS ")

    '7a. calculate the Sun's local hour angle
     Dim Tt As Double
     
     'cosH = (thisCos(zenith) - (sunDeclination() * thisSin(mLatitude)))
     cosH = (Cos(to_rad * zenith) - (sunDeclination * Sin(to_rad * mLatitude)) / (cosDec * Cos(to_rad * mLatitude)))
     '     -1.45438976515809E-02            -0.213797236617717                      0.592568806464951
     If rising Then Debug.Print ("%SFN-I localmeantime2 cosH = " & cosH)
     
     If (cosH > 1) Then
         returnSunString = "the sun never rises on this location (on the specified date)"
         localmeantime2 = 0
     ElseIf (cosH < -1) Then
         returnSunString = "the sun never sets on this location (on the specified date)"
         localmeantime2 = 0
     Else
         '7b. finish calculating H and convert into hours
         If rising Then Debug.Print ("%SFN-I localmeantime2 rising = " & rising)
         If rising Then
             H = 360 - thisAcos(cosH)
             'H = 360 - (1 / to_rad) * thisAcos(cosH)
         Else
             H = thisAcos(cosH)
         End If
         If rising Then Debug.Print ("%SFN-I localmeantime2 H = " & H)

         H = H / 15
         If rising Then Debug.Print ("%SFN-I localmeantime2 H = H / 15 = " & H)
         
         RA = rightAscension
         If rising Then Debug.Print ("%SFN-I localmeantime2 RA = " & RA)
         
         T = approximateTime
         If rising Then Debug.Print ("%SFN-I localmeantime2 t = " & T)
         
         '8. calculate local mean time of rising/setting
         Tt = H + RA - (0.06571 * T) - 6.622
         If rising Then Debug.Print ("%SFN-I localmeantime2 Tt = " & Tt)
         
         localmeantime2 = Tt 'return
    End If
    
    ' adjust back to UTC
    UT = localmeantime2 - lnghour
    If (UT < 0) Then
       UT = UT + 24
    ElseIf (UT >= 24) Then
       UT = UT - 24
    End If
    localmeantime2 = UT
    
    If rising Then Debug.Print ("%SFN-I localmeantime2  = " & localmeantime2)
    'Debug.Print ("%SFN-I localmeantime2 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDS ")
    
    On Error GoTo 0
    Exit Function

localmeantime2_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure localmeantime2, line " & Erl & "."

End Function
' ----------------------------------------------------------------
' Procedure Name: thisAcos
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter x (Double):
' Return Type: Double
' Author: beededea
' Date: 12/02/2024
' ----------------------------------------------------------------
Private Function thisAcos(x As Double) As Double

    On Error GoTo thisAcos_Error
    'thisAcos = (180 / pi) * acos(x)
    'thisAcos = thisAsin(x) + pi / 2
'    If (x >= 0) Then
'        thisAcos = 2 * Atn(Sqr((1 - x) / (1 + x)))
'    Else
'        thisAcos = pi - 2 * Atn(Sqr((1 + x) / (1 - x)))
'    End If
    

    Select Case x
        Case -1
            thisAcos = 4 * Atn(1)
             
        Case 0:
            thisAcos = 2 * Atn(1)
             
        Case 1:
            thisAcos = 0
             
        Case Else:
            thisAcos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End Select
    thisAcos = (180 / pi) * thisAcos
    
    'Debug.Print ("%SRS-O thisAcos " & thisAcos)
    
    On Error GoTo 0
    Exit Function

thisAcos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure thisAcos, line " & Erl & "."

End Function



'---------------------------------------------------------------------------------------
' Function   : fUtcMonth
' Author    : beededea
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
' Author: beededea
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


'Private Function fHoursRange(H As Double) As Double
'    fHoursRange = (H + 24) Mod 24 'return
'    Debug.Print (">>>>>>>>>>>>>>>>> %SRS-O fHoursRange " & fHoursRange)
'End Function

'Private Function fsunriseUtcHours() As Double
'    rising = True
'    fsunriseUtcHours = fUTCTime()
'    Debug.Print ("%SRS-O fSunriseUtcHours " & fsunriseUtcHours)
'End Function
'
'Private Function fsunsetUtcHours() As Double
'    rising = False
'    fsunsetUtcHours = fUTCTime() 'return
'    Debug.Print ("%SRS-O fsunsetUtcHours " & fsunsetUtcHours)
'End Function
'
'' utcCurrentHours is the time that you would like to test for daylight, in hours, at UTC
'' For example, to test if it's daylight in Tokyo (GMT+9) at 10:30am, pass in
'' utcCurrentHours=1.5, which corresponds to 1:30am UTC.
'Private Function isDaylight(ByVal utcCurrentHours As Double, ByRef returnHours As Double) As Boolean
'    Dim sunriseHours As Double
'    Dim sunsetHours As Double
'
'    Debug.Print ("%SRS-O isDaylight (decimal time )" & utcCurrentHours)
'
'    sunriseHours = fSunriseUtcHours()
'    sunsetHours = sunsetUtcHours()
'
'    If (sunsetHours < sunriseHours) Then
'        ' Either the sunrise or sunset time is for tomorrow
'        If (utcCurrentHours > sunriseHours) Then
'            isDaylight = True 'return
'        ElseIf (utcCurrentHours < sunsetHours) Then
'            isDaylight = True 'return
'        Else
'            isDaylight = False 'return
'        End If
'    End If
'
'    If (utcCurrentHours >= sunriseHours) Then
'        returnHours = utcCurrentHours < sunsetHours
'    End If
'    isDaylight = False 'return
'    Debug.Print ("%SRS-O isDaylight " & isDaylight)
'
'End Function





'Public Function fDayOfTheWeek_I(TheDate As Date) As Integer
'  fDayOfTheWeek_I = Format(TheDate, "w") - 1
'  Debug.Print ("%SRS-O fDayOfTheWeek_I " & fDayOfTheWeek_I)
'
'End Function




'Private Function fUTCTime() As Double
'    Dim returnSunString As String
'    Dim T As Double '= This.approximateTime()
'    T = fLocalMeanTime(returnSunString)
'    UT = T - lnghour
'    fUTCTime = fHoursRange(UT) 'return
'    Debug.Print ("%SRS-O fUTCTime " & fUTCTime)
'End Function

'Private Function fSunriseUtcHours() As Double
'    rising = True
'    fSunriseUtcHours = UTCTime()
'    Debug.Print ("%SRS-O fSunriseUtcHours " & fSunriseUtcHours)
'End Function
'
'Private Function sunsetUtcHours() As Double
'    rising = False
'    sunsetUtcHours = UTCTime() 'return
'    Debug.Print ("%SRS-O sunsetUtcHours " & sunsetUtcHours)
'End Function

'Private Function sunriseLocalHours(gmt)
'    sunriseLocalHours = fHoursRange(gmt + fSunriseUtcHours())
'    Debug.Print ("%SRS-O sunriseLocalHours " & sunriseLocalHours)
'End Function
'
'Private Function sunsetLocalHours(gmt)
'    sunsetLocalHours = fHoursRange(gmt + sunsetUtcHours())
'    Debug.Print ("%SRS-O sunsetLocalHours " & sunsetLocalHours)
'End Function
''1. first calculate the day of the year
'    Private Function fGetDOY(dateinput As Date) As Integer
'        Dim N1 As Integer
'        Dim N2 As Integer
'        Dim N3 As Integer
'        Dim N As Integer
'
'        thisMonth = Month(dateinput) - 1
'        thisYear = Year(dateinput) - 1900
'        thisDay = fUtcDayOfWeek(dateinput)
'
'        Debug.Print ("%SRS-O year " & thisYear)
'        Debug.Print ("%SRS-O month " & thisMonth)
'        Debug.Print ("%SRS-O day " & thisDay)
'
'        N1 = Fix(275 * thisMonth / 9)
'        Debug.Print ("%SRS-O N1 " & N1)
'
'        N2 = Fix((thisMonth + 9) / 12)
'        Debug.Print ("%SRS-O N2 " & N2)
'
'        N3 = (1 + Fix((thisYear - 4 * Fix(thisYear / 4) + 2) / 3))
'        Debug.Print ("%SRS-O N3 " & N3)
'
'        N = N1 - (N2 * N3) + Val(thisDay) - 30
'
'        fGetDOY = N 'return
'
'        Debug.Print ("%SRS-O fGetDOY " & fGetDOY)
'    End Function
'''2. convert the longitude to hour value and calculate an approximate time
'Private Function fapproximateTime() As Double
'    Dim doy As Integer
'    doy = fGetDOY(Date)
'
'    lnghour = mLongitude / 15
'    If (rising) Then
'        fapproximateTime = doy + ((6 - lnghour) / 24) 'return
'    Else
'        fapproximateTime = doy + ((18 - lnghour) / 24) 'return
'    End If
'    Debug.Print ("%SRS-O fapproximateTime " & fapproximateTime)
'End Function
''3. calculate the Sun's mean anomaly
'Private Function fmeanAnomaly() As Double
'
'    Dim T As Double
'    T = fapproximateTime()
'    fmeanAnomaly = (0.9856 * T) - 3.289 'return
'    Debug.Print ("%SRS-O fmeanAnomaly " & fmeanAnomaly)
'End Function
''4. calculate the Sun's true longitude
'Private Function ftrueLongitude() As Double
'    Dim M As Double
'
'    M = fmeanAnomaly()
'    longi = M + (1.916 * thisSin(M)) + (0.02 * thisSin(2 * M)) + 282.634
'    ftrueLongitude = longi 'Mod 360 'return
'    Debug.Print ("%SRS-O ftrueLongitude " & ftrueLongitude)
'End Function
''
'''5a. calculate the Sun's right ascension
'Private Function frightAscension() As Double
'    longi = ftrueLongitude()
'    'RA = (1 / to_rad) * thisAtan(0.91764 * thisTan(longi))
'    RA = (1 / to_rad) * Atn(0.91764 * Tan(to_rad * longi)) 'ra  = double
'    If (RA < 0) Then
'       RA = RA + 360
'    ElseIf (RA >= 360) Then
'       RA = RA - 360
'    End If
'
'    ' right ascension value needs to be in the same quadrant as longi
'    Lquadrant = (Fix(longi / 90)) * 90
'    RAquadrant = (Int(RA / 90)) * 90
'    RA = RA + (Lquadrant - RAquadrant)
'    ' right ascension value needs to be converted into hours
'    RA = RA / 15
'
'    frightAscension = RA 'return
'    Debug.Print ("%SRS-O frightAscension " & frightAscension)
'End Function
''
'' 6. calculate the Sun's declination
'Private Function fsunDeclination() As Double
'    longi = ftrueLongitude()
'    fsunDeclination = 0.39782 * thisSin(longi) ' return
'    'Debug.Print ("%SRS-O fsunDeclination " & fsunDeclination)
'End Function
'
'Private Function fcosDec()
'    fcosDec = Cos(thisAsin(fsunDeclination())) 'return
'    'Debug.Print ("%SRS-O fCosDec " & fCosDec)
'End Function
''7a. calculate the Sun's local hour angle
'Private Function fLocalMeanTime(ByRef returnSunString As String) As Double
'    Dim T As Double '= This.approximateTime()
'    'cosH = (thisCos(zenith) - (fsunDeclination() * thisSin(mLatitude)))
'
'    cosH = (Cos(to_rad * zenith) - (fsunDeclination() * Sin(to_rad * mLatitude)) / (fcosDec * Cos(to_rad * mLatitude)))
'    '     -1.45438976515809E-02            -0.213797236617717                      0.592568806464951
'
'    Debug.Print ("%SRS-O fLocalMeanTime cosH " & cosH)
'
'    'cosH = (Cos(to_rad * zenith) - (sinDec * Sin(to_rad * Lat))) / (fCosDec * Cos(to_rad * Lat))
'
'     If (cosH > 1) Then
'         returnSunString = "the sun never rises on this location (on the specified date)"
'         fLocalMeanTime = 0
'     ElseIf (cosH < -1) Then
'         returnSunString = "the sun never sets on this location (on the specified date)"
'         fLocalMeanTime = 0
'     Else
'         '7b. finish calculating H and convert into hours
'         If rising Then
'             H = 360 - (1 / to_rad) * thisAcos(cosH)
'         Else
'             thisAcos (cosH)
'         End If
'         H = H / 15
'         RA = frightAscension()
'         T = fapproximateTime()
'
'         '8. calculate local mean time of rising/setting
'         T = H + RA - (0.06571 * T) - 6.622
'
'         fLocalMeanTime = T 'return
'    End If
'    Debug.Print ("%SRS-O fLocalMeanTime " & fLocalMeanTime)
'End Function
