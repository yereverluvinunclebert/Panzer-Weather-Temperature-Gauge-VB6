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
Private Day As String

'private  doy = This.getDOY()


Private longi As Long
Private H As Double
Private UT As String
Private RA As Long
Private Lquadrant As Long
Private RAquadrant As Long

Private cosH As Long
'private  UT = T - This.lngHour
'private  sunriseHours = this.sunriseUtcHours()
Private sunsetHours
Private zenith As Double
Private rising As Boolean
Private lnghour As Integer

Const pi As Double = 3.14159265358979

Private mLatitude As Double
Private mLongitude As Double
Private mFullYear As String
Private mUtcMonth As String
Private mUtcDay As String

'    locationDayNight = new SunriseSunset(theDate.getYear(), theDate.getMonth(), theDate.getDay(), latitude , longitude  )

'Private Function SunriseSunset(ByVal FullYear, fUtcMonth, utcDay, latitude, longitude)
Private Sub Class_Initialize()
   zenith = 90 + 50 / 60 '   offical      = 90 degrees 50'
                              '   civil        = 96 degrees
                              '   nautical     = 102 degrees
                              '   astronomical = 108 degrees
' 2011, 1, 19, 35+40/60, 139+45/60

   rising = True ' set to true for sunrise, false for sunset
   
    '1. first calculate the day of the year
    Call getDOY
    '2. convert the longitude to hour value and calculate an approximate time
    Call approximateTime
    '3. calculate the Sun's mean anomaly
    Call meanAnomaly
    '4. calculate the Sun's true longitude
    Call trueLongitude
    '5a. calculate the Sun's right ascension
    Call rightAscension
    '5b. right ascension value needs to be in the same quadrant as L
    '5c. right ascension value needs to be converted into hours
    '6. calculate the Sun's declination
    Call sinDec
    '7a. calculate the Sun's local hour angle
    '7b. finish calculating H and convert into hours
    '8. calculate local mean time of rising/setting
    '9. adjust back to UTC

   lnghour = Longitude / 15

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
End Function
Private Function thisCos(deg) As Double
    thisCos = Cos(deg * pi / 180)
End Function
Private Function thisTan(deg) As Double
    thisTan = Tan(deg * pi / 180)
End Function
Private Function thisAsin(x) As Double
    thisAsin = Atn(x / Sqr(-x * x + 1))
End Function
Private Function thisAcos(x) As Double
    'thisAcos = thisAsin(x) + pi / 2
    If (x >= 0) Then
        thisAcos = 2 * Atn(Sqr((1 - x) / (1 + x)))
    Else
        thisAcos = pi - 2 * Atn(Sqr((1 + x) / (1 - x)))
    End If
End Function
Private Function thisAtan(x) As Double
    thisAtan = (180 / pi) * Atn(x)
End Function
'1. first calculate the day of the year
Private Function getDOY() As Integer
    Dim N1 As Long
    Dim N2 As Long
    Dim N3 As Long
    Dim N As Long
    
    thisMonth = fUtcMonth(Month(Now), True)
    thisYear = Year(Now)
    Day = fDayOfTheWeek_I(Now())

    N1 = Fix(275 * thisMonth / 9)
    N2 = Fix((thisMonth + 9) / 12)
    N3 = (1 + Fix((thisYear - 4 * Fix(thisYear / 4) + 2) / 3))
    N = N1 - (N2 * N3) + Day - 30
    
    getDOY = N 'return

End Function
'2. convert the longitude to hour value and calculate an approximate time
Private Function approximateTime() As Integer
    Dim doy As String
    doy = getDOY()
    If (rising) Then
        approximateTime = doy + ((6 - lnghour) / 24) 'return
    Else
        approximateTime = doy + ((18 - lnghour) / 24) 'return
    End If
End Function
'3. calculate the Sun's mean anomaly
Private Function meanAnomaly() As Double

    Dim T As Double '= This.approximateTime()
    T = approximateTime()
    meanAnomaly = (0.9856 * T) - 3.289 'return
End Function
'4. calculate the Sun's true longitude
Private Function trueLongitude() As Double
    Dim M As Double '= This.meanAnomaly()
    
    M = meanAnomaly()
    longi = M + (1.916 * thisSin(M)) + (0.02 * thisSin(2 * M)) + 282.634
    trueLongitude = longi Mod 360 'return
End Function

'5a. calculate the Sun's right ascension
Private Function rightAscension() As Double
    longi = trueLongitude()
    RA = thisAtan(0.91764 * thisTan(longi))
    RA = RA Mod 360

    ' right ascension value needs to be in the same quadrant as longi
    Lquadrant = (Fix(longi / 90)) * 90
    RAquadrant = (Fix(RA / 90)) * 90
    RA = RA + (Lquadrant - RAquadrant)
    ' right ascension value needs to be converted into hours
    RA = RA / 15

    rightAscension = RA 'return
End Function

' 6. calculate the Sun's declination
Private Function sinDec() As Double
    longi = trueLongitude()
    sinDec = 0.39782 * thisSin(longi) ' return
End Function

Private Function cosDec()
    cosDec = thisCos(thisAsin(sinDec())) 'return
End Function
'7a. calculate the Sun's local hour angle
Private Function localMeanTime(ByRef returnSunString As String) As Double
    Dim T As Double '= This.approximateTime()
    cosH = (thisCos(zenith) - (sinDec() * thisSin(Latitude)))

     If (cosH > 1) Then
         returnSunString = "the sun never rises on this location (on the specified date)"
         localMeanTime = 0
     ElseIf (cosH < -1) Then
         returnSunString = "the sun never sets on this location (on the specified date)"
         localMeanTime = 0
     Else
         '7b. finish calculating H and convert into hours
         If rising Then
             H = 360 - thisAcos(cosH)
         Else
             thisAcos (cosH)
         End If
         H = H / 15
         RA = rightAscension()
         T = approximateTime()
         
         '8. calculate local mean time of rising/setting
         T = H + RA - (0.06571 * T) - 6.622

         localMeanTime = T 'return
    End If
End Function

Private Function hoursRange(UT As String) As Integer
    hoursRange = (H + 24) Mod 24 'return
End Function

Private Function UTCTime() As Integer
    Dim returnSunString As String
    Dim T As Double '= This.approximateTime()
    T = localMeanTime(returnSunString)
    UT = T - lnghour
    UTCTime = hoursRange(UT) 'return
End Function

Private Function sunriseUtcHours() As Date
    rising = True
    sunriseUtcHours = UTCTime()
End Function

Private Function sunsetUtcHours() As Date
    rising = False
    sunsetUtcHours = UTCTime() 'return
End Function

Private Function sunriseLocalHours(gmt)
    sunriseLocalHours = hoursRange(gmt + sunriseUtcHours())
End Function

Private Function sunsetLocalHours(gmt)
    sunsetLocalHours = hoursRange(gmt + sunsetUtcHours())
End Function

' utcCurrentHours is the time that you would like to test for daylight, in hours, at UTC
' For example, to test if it's daylight in Tokyo (GMT+9) at 10:30am, pass in
' utcCurrentHours=1.5, which corresponds to 1:30am UTC.
Private Function isDaylight(ByVal utcCurrentHours, ByRef returnHours As Integer) As Boolean
    Dim sunriseHours As Integer
    Dim sunsetHours As Integer
    
    sunriseHours = sunriseUtcHours()
    sunsetHours = sunsetUtcHours()
    
    If (sunsetHours < sunriseHours) Then
        ' Either the sunrise or sunset time is for tomorrow
        If (utcCurrentHours > sunriseHours) Then
            isDaylight = True 'return
        ElseIf (utcCurrentHours < sunsetHours) Then
            isDaylight = True 'return
        Else
            isDaylight = False 'return
        End If
    End If
    
    If (utcCurrentHours >= sunriseHours) Then
        returnHours = utcCurrentHours < sunsetHours
    End If
    isDaylight = False 'return
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
    
    ''Debug.Print ("%DST func fUtcMonth")
    ''Debug.Print ("%DST-I thisMonth " & thisMonth)
    
    fUtcMonth = Month(CDate(thisMonth & "/1/2000"))
    If utcFlag = True Then fUtcMonth = fUtcMonth - 1 ' convert 'normal month starting number of 1 to starting with 0 UTC

    If fUtcMonth < 0 Or fUtcMonth > 11 Then
        MsgBox ("fUtcMonth: " & thisMonth & " is not a valid month name")
        fUtcMonth = -1 ' return invalid
        
        ''Debug.Print ("%DST-O abnormal fUtcMonth " & fUtcMonth)
    End If
    
    ''Debug.Print ("%DST-O fUtcMonth " & fUtcMonth)
    
    On Error GoTo 0
    Exit Function

fUtcMonth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Function   fUtcMonth of Module modDaylightSavings"

End Function


Public Function fDayOfTheWeek_I(TheDate As Date) As Integer
  fDayOfTheWeek_I = Format(TheDate, "w") - 1
End Function
