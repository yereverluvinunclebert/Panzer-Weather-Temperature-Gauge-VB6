Attribute VB_Name = "modMain"
'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
' gaugeForm_BubblingEvent ' leaving that here so I can copy/paste to find it

Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Public Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' to set the full window Opacity
''Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
''Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
''Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Private Const WS_EX_LAYERED  As Long = &H80000
'Private Const GWL_EXSTYLE  As Long = (-20)
'Private Const LWA_COLORKEY  As Long = &H1       'to transparent
'Private Const LWA_ALPHA  As Long = &H2          'to semi transparent
'------------------------------------------------------ ENDS

Public fMain As New cfMain
Public aboutWidget As cwAbout
Public helpWidget As cwHelp
Public licenceWidget As cwLicence

Public revealWidgetTimerCount As Integer
 
Public fTemperature As New cfTemperature
Public overlayTemperatureWidget As cwOverlayTemp

Public fSelector As New cfSelector
Public overlaySelectorWidget As cwOverlaySelect

Public fClipB As New cfClipB
Public overlayClipbWidget As cwOverlayClipb

Public fAnemometer As New cfAnemometer
Public overlayAnemoWidget As cwOverlayAnemo

Public fHumidity As New cfHumidity
Public overlayHumidWidget As cwOverlayHumid

Public fBarometer As New cfBarometer
Public overlayBaromWidget As cwOverlayBarom
 
Public fPictorial As New cfPictorial
Public overlayPictorialWidget As cwOverlayPict

Public sunriseSunset As cwSunriseSunset
Public WeatherMeteo As cwWeatherMeteo

Public widgetName1 As String
Public widgetName2 As String
Public widgetName3 As String
Public widgetName4 As String
Public widgetName5 As String
Public widgetName6 As String
Public widgetName7 As String

'Public startupFlg As Boolean

Public firstPoll As Boolean

    




'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Main()
   On Error GoTo Main_Error
    
   Call mainRoutine(False)

   On Error GoTo 0
   Exit Sub

Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : main_routine
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/06/2023
' Purpose   : called by sub main() to allow this routine to be called elsewhere,
'             a reload for example.
'---------------------------------------------------------------------------------------
'
Public Sub mainRoutine(ByVal restart As Boolean)
    
    Dim extractCommand As String: extractCommand = vbNullString
    Dim temperaturePSDFullPath As String: temperaturePSDFullPath = vbNullString
    Dim selectorPSDFullPath As String: selectorPSDFullPath = vbNullString
    Dim clipBPSDFullPath As String: clipBPSDFullPath = vbNullString
    Dim anemometerPSDFullPath As String: anemometerPSDFullPath = vbNullString
    Dim HumidityPSDFullPath As String: HumidityPSDFullPath = vbNullString
    Dim barometerPSDFullPath As String: barometerPSDFullPath = vbNullString
    Dim pictorialPSDFullPath As String: pictorialPSDFullPath = vbNullString
    Dim licenceState As Integer: licenceState = 0

    On Error GoTo main_routine_Error
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    widgetName1 = "Panzer Temperature Gauge"
    temperaturePSDFullPath = App.path & "\Res\Panzer Weather Gauges VB6.psd"
    
    widgetName2 = "ICAO Selector"
    selectorPSDFullPath = App.path & "\Res\Panzer Weather Selector VB6.psd"
    
    widgetName3 = "Clipboard"
    clipBPSDFullPath = App.path & "\Res\Panzer Weather Clipboard VB6.psd"
    
    widgetName4 = "Anemometer Gauge"
    anemometerPSDFullPath = App.path & "\Res\Panzer Weather Anemometer Gauge VB6.psd"
    
    widgetName5 = "Humidity Gauge"
    HumidityPSDFullPath = App.path & "\Res\Panzer Weather Humidity Gauge VB6.psd"
    
    widgetName6 = "Barometer Gauge"
    barometerPSDFullPath = App.path & "\Res\Panzer Weather Barometer Gauge VB6.psd"
    
    widgetName7 = "Pictorial Gauge"
    pictorialPSDFullPath = App.path & "\Res\Panzer Weather Pictorial Gauge VB6.psd"
    
    prefsCurrentWidth = 9075
    prefsCurrentHeight = 16450
    
    gblOriginatingForm = "temperatureForm"
    
    firstPoll = True
    
    'startupFlg = True ' this is used to prevent some control initialisations from running code at startup

    extractCommand = Command$ ' capture any parameter passed, remove if a soft reload
    If restart = True Then extractCommand = vbNullString
    
    #If TWINBASIC Then
        gblCodingEnvironment = "TwinBasic"
    #Else
        gblCodingEnvironment = "VB6"
    #End If
    
    menuForm.mnuAbout.Caption = "About Panzer Weather Gauge Cairo " & gblCodingEnvironment & " widget"
    
    ' create dictionary collection instead of an array to load dropdown list
    'Set collValidLocations = CreateObject("Scripting.Dictionary") ' tested with all three
    Set collValidLocations = New_c.Collection(False)
    
    'add non-PSD Resources to the global ImageList
    Call addGeneralImagesToImageLists
    Call addDayWeatherImagesToImageLists
    'Call addNightWeatherImagesToImageLists
    
    ' check the Windows version
    classicThemeCapable = fTestClassicThemeCapable
  
    ' get this tool's entry in the trinkets settings file and assign the app.path
    Call getTrinketsFile
  
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile("Software\PzTemperatureGauge", gblSettingsFile)
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' check first usage via licence acceptance value and then set initial DPI awareness
    Call setAutomaticDPIState(licenceState)

    'load the collection for storing the overlay surfaces with its relevant keys direct from each PSD
    If restart = False Then
        Call loadTemperatureExcludePathCollection ' no need to reload the collTemperaturePSDNonUIElements layer name keys on a reload
        Call loadSelectorExcludePathCollection
        Call loadClipBExcludePathCollection
        Call loadAnemometerExcludePathCollection
        Call loadHumidityExcludePathCollection
        Call loadBarometerExcludePathCollection
        Call loadPictorialExcludePathCollection
    End If
    
    ' start the load of the PSD files using the RC6 PSD-Parser.instance
    Call fTemperature.InitTemperatureFromPSD(temperaturePSDFullPath)
    Call fSelector.InitSelectorFromPSD(selectorPSDFullPath)
    Call fClipB.InitClipBFromPSD(clipBPSDFullPath)
    Call fAnemometer.InitAnemometerFromPSD(anemometerPSDFullPath)
    Call fHumidity.InitHumidityFromPSD(HumidityPSDFullPath)
    Call fBarometer.InitBarometerFromPSD(barometerPSDFullPath)
    Call fPictorial.InitPictorialFromPSD(pictorialPSDFullPath)
    
    ' resolve VB6 sizing width bug
    Call determineScreenDimensions
            
    ' initialise and create the three main RC forms on the current display
    Call createRCFormsOnCurrentDisplay
    
    ' check the selected monitor properties
    Call monitorProperties(fTemperature.temperatureGaugeForm)  ' might use RC6 for this?
    
    ' place the form at the saved location
    Call makeVisibleFormElements
    
    ' run the functions that are ALSO called at reload time elsewhere.
    
    ' set menu items for all the gauges
    Call setMenuItems
    
    ' set taskbar entry for all the gauges
    Call setTaskbarEntry
        
    ' set characteristics of widgets on the temperature gauge form
    Call adjustTempMainControls ' this needs to be here after the initialisation of the Cairo forms and widgets

    ' set characteristics of widgets on the anemometer gauge form
    Call adjustAnemometerMainControls

    ' set characteristics of widgets on the Humidity gauge form
    Call adjustHumidityMainControls

    ' set characteristics of widgets on the Humidity gauge form
    Call adjustBarometerMainControls
    
    ' set characteristics of widgets on the selector form
    Call adjustSelectorMainControls
    
    ' set characteristics of widgets on the clipboard form
    Call adjustClipBMainControls
    
    ' set characteristics of widgets on the pictorial form
    Call adjustPictorialMainControls
    
    ' set the z-ordering of the window
    Call setAlphaFormZordering
    
    ' set the tooltips on the main screen
    Call setMainTooltips
    
    ' set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
    Call setHidingTime

    If minutesToHide > 0 Then menuForm.mnuHideWidget.Caption = "Hide Widget for " & minutesToHide & " min."

    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    ' if the parameter states re-open prefs then shows the prefs
    If extractCommand = "prefs" Then
        Call makeProgramPreferencesAvailable
        extractCommand = vbNullString
    End If
    
    'load the preferences form but don't yet show it, speeds up access to the prefs via the menu
    Call loadPreferenceForm
    
    'load the message form but don't yet show it, speeds up access to the message form when needed.
    Load frmMessage
    
    ' display licence screen on first usage
    Call showLicence(fLicenceState)
    
    ' make the prefs appear on the first time running
    Call checkFirstTime
 
    ' configure any global timers here
    Call configureTimers
    
    ' for the first run we are going to call gate data directly, this will attempt to connect and read the METAR data
    Call WeatherMeteo.getData

    'startupFlg = False
        
     ' RC message pump will auto-exit when Cairo Forms > 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload. Do not move this line.
    #If TWINBASIC Then
        Cairo.WidgetForms.EnterMessageLoop
    #Else
        If restart = False Then Cairo.WidgetForms.EnterMessageLoop
    #End If
     
   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain at "
    
End Sub
 
'---------------------------------------------------------------------------------------
' Procedure : checkFirstTime
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 12/05/2023
' Purpose   : check for first time running, first time run shows prefs
'---------------------------------------------------------------------------------------
'
Private Sub checkFirstTime()

   On Error GoTo checkFirstTime_Error

    If gblFirstTimeRun = "true" Then
        'MsgBox "checkFirstTime"

        Call makeProgramPreferencesAvailable
        gblFirstTimeRun = "false"
        sPutINISetting "Software\PzTemperatureGauge", "firstTimeRun", gblFirstTimeRun, gblSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

checkFirstTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkFirstTime of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : initialiseGlobalVars
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 12/05/2023
' Purpose   : initialise global vars
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGlobalVars()
      
    On Error GoTo initialiseGlobalVars_Error

    ' general
    gblStartup = vbNullString
    gblGaugeFunctions = vbNullString
'    gblPointerAnimate = vbNullString
    gblSamplingInterval = vbNullString
    gblStormTestInterval = vbNullString
    gblErrorInterval = vbNullString
    gblAirportsURL = vbNullString
    
    gblIcao = vbNullString

    ' config
    gblEnableTooltips = vbNullString
    gblEnablePrefsTooltips = vbNullString
    gblEnableBalloonTooltips = vbNullString
    gblShowTaskbar = vbNullString
    gblDpiAwareness = vbNullString
    
    
    gblClipBSize = vbNullString
    gblSelectorSize = vbNullString
    
    gblScrollWheelDirection = vbNullString
    
    ' position
    gblAspectHidden = vbNullString
    gblGaugeType = vbNullString
    gblWidgetPosition = vbNullString
    
    gblTemperatureLandscapeLocked = vbNullString
    gblTemperaturePortraitLocked = vbNullString
    gblTemperatureGaugeSize = vbNullString
    gblTemperatureLandscapeLockedHoffset = vbNullString
    gblTemperatureLandscapeLockedVoffset = vbNullString
    gblTemperaturePortraitLockedHoffset = vbNullString
    gblTemperaturePortraitLockedVoffset = vbNullString
    gblTemperatureVLocationPerc = vbNullString
    gblTemperatureHLocationPerc = vbNullString
    gblPreventDraggingTemperature = vbNullString
    gblTemperatureFormHighDpiXPos = vbNullString
    gblTemperatureFormHighDpiYPos = vbNullString
    gblTemperatureFormLowDpiXPos = vbNullString
    gblTemperatureFormLowDpiYPos = vbNullString
    
    gblAnemometerGaugeSize = vbNullString
    gblAnemometerLandscapeLocked = vbNullString
    gblAnemometerPortraitLocked = vbNullString
    gblAnemometerFormHighDpiXPos = vbNullString
    gblAnemometerFormHighDpiYPos = vbNullString
    gblAnemometerFormLowDpiXPos = vbNullString
    gblAnemometerFormLowDpiYPos = vbNullString
    gblAnemometerLandscapeLockedHoffset = vbNullString
    gblAnemometerLandscapeLockedVoffset = vbNullString
    gblAnemometerPortraitLockedHoffset = vbNullString
    gblAnemometerPortraitLockedVoffset = vbNullString
    gblPreventDraggingAnemometer = vbNullString
    
    gblHumidityGaugeSize = vbNullString
    gblHumidityLandscapeLocked = vbNullString
    gblHumidityPortraitLocked = vbNullString
    gblHumidityFormHighDpiXPos = vbNullString
    gblHumidityFormHighDpiYPos = vbNullString
    gblHumidityFormLowDpiXPos = vbNullString
    gblHumidityFormLowDpiYPos = vbNullString
    gblHumidityLandscapeLockedHoffset = vbNullString
    gblHumidityLandscapeLockedVoffset = vbNullString
    gblHumidityPortraitLockedHoffset = vbNullString
    gblHumidityPortraitLockedVoffset = vbNullString
    gblPreventDraggingHumidity = vbNullString
    
    gblBarometerGaugeSize = vbNullString
    gblBarometerLandscapeLocked = vbNullString
    gblBarometerPortraitLocked = vbNullString
    gblBarometerFormHighDpiXPos = vbNullString
    gblBarometerFormHighDpiYPos = vbNullString
    gblBarometerFormLowDpiXPos = vbNullString
    gblBarometerFormLowDpiYPos = vbNullString
    gblBarometerLandscapeLockedHoffset = vbNullString
    gblBarometerLandscapeLockedVoffset = vbNullString
    gblBarometerPortraitLockedHoffset = vbNullString
    gblBarometerPortraitLockedVoffset = vbNullString
    gblPreventDraggingBarometer = vbNullString
    
    gblPictorialGaugeSize = vbNullString
    gblPictorialLandscapeLocked = vbNullString
    gblPictorialPortraitLocked = vbNullString
    gblPictorialFormHighDpiXPos = vbNullString
    gblPictorialFormHighDpiYPos = vbNullString
    gblPictorialFormLowDpiXPos = vbNullString
    gblPictorialFormLowDpiYPos = vbNullString
    gblPictorialLandscapeLockedHoffset = vbNullString
    gblPictorialLandscapeLockedVoffset = vbNullString
    gblPictorialPortraitLockedHoffset = vbNullString
    gblPictorialPortraitLockedVoffset = vbNullString
    gblPreventDraggingPictorial = vbNullString
        
    ' sounds
    gblEnableSounds = vbNullString
    
    ' development
    gblDebug = vbNullString
    gblDblClickCommand = vbNullString
    gblOpenFile = vbNullString
    gblDefaultEditor = vbNullString
         
    ' font
    gblTempFormFont = vbNullString
    gblPrefsFont = vbNullString
    gblPrefsFontSizeHighDPI = vbNullString
    gblPrefsFontSizeLowDPI = vbNullString
    gblPrefsFontItalics = vbNullString
    gblPrefsFontColour = vbNullString
    
    ' window
    gblWindowLevel = vbNullString
    

    gblOpacity = vbNullString

    
    gblWidgetHidden = vbNullString
    gblHidingTime = vbNullString
    gblIgnoreMouse = vbNullString
    gblMenuOccurred = False ' bool
    gblFirstTimeRun = vbNullString
    
    ' general storage variables declared
    gblSettingsDir = vbNullString
    gblSettingsFile = vbNullString
    
    gblTrinketsDir = vbNullString
    gblTrinketsFile = vbNullString
    
    gblClipBFormHighDpiXPos = vbNullString
    gblClipBFormHighDpiYPos = vbNullString
    gblClipBFormLowDpiXPos = vbNullString
    gblClipBFormLowDpiYPos = vbNullString
    
    gblSelectorFormHighDpiXPos = vbNullString
    gblSelectorFormHighDpiYPos = vbNullString
    gblSelectorFormLowDpiXPos = vbNullString
    gblSelectorFormLowDpiYPos = vbNullString
    
    gblLastSelectedTab = vbNullString
    gblSkinTheme = vbNullString
    
    gblLastUpdated = vbNullString
    gblMetarPref = vbNullString
    gblPressureScale = vbNullString
    
    gblOldPressureStorage = vbNullString
    gblPressureStorageDate = vbNullString
    gblCurrentPressureValue = vbNullString
    
    
    ' general variables declared
    'toolSettingsFile = vbNullString
    classicThemeCapable = False
    storeThemeColour = 0
    windowsVer = vbNullString
    
    ' vars to obtain correct screen width (to correct VB6 bug) STARTS
    screenTwipsPerPixelX = 0
    screenTwipsPerPixelY = 0
    screenWidthTwips = 0
    screenHeightTwips = 0
    screenHeightPixels = 0
    screenWidthPixels = 0
    oldScreenHeightPixels = 0
    oldScreenWidthPixels = 0
    
    ' key presses
    CTRL_1 = False
    SHIFT_1 = False
    
    ' other globals
    debugFlg = 0
    minutesToHide = 0
    aspectRatio = vbNullString
    revealWidgetTimerCount = 0
    oldgblSettingsModificationTime = #1/1/2000 12:00:00 PM#
    
    gblJustAwoken = False
    
    gblCodingEnvironment = vbNullString

   On Error GoTo 0
   Exit Sub

initialiseGlobalVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGlobalVars of Module modMain"
    
End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : addGeneralImagesToImageLists
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : add Resources to the global ImageList
'---------------------------------------------------------------------------------------
'
Private Sub addGeneralImagesToImageLists()
    
    On Error GoTo addGeneralImagesToImageLists_Error

'    add Resources to the global ImageList that are not being pulled from the PSD directly
    
    Cairo.ImageList.AddImage "about", App.path & "\Resources\images\about.png"
    Cairo.ImageList.AddImage "helpTemperature", App.path & "\Resources\images\panzerTemperature-help.png"
    Cairo.ImageList.AddImage "helpHumidity", App.path & "\Resources\images\panzerHumidity-help.png"
    Cairo.ImageList.AddImage "helpBarometer", App.path & "\Resources\images\panzerBarometer-help.png"
    Cairo.ImageList.AddImage "helpAnemometer", App.path & "\Resources\images\panzerAnemometer-help.png"
    Cairo.ImageList.AddImage "helpSelector", App.path & "\Resources\images\panzerClipboard-help.png"
    Cairo.ImageList.AddImage "helpClipboard", App.path & "\Resources\images\panzerClipboard-help.png"
    Cairo.ImageList.AddImage "helpWeather", App.path & "\Resources\images\panzerweather-icon-help.png"
    ' deanieboy
    Cairo.ImageList.AddImage "licence", App.path & "\Resources\images\frame.png"
    
    ' prefs icons
    
    Cairo.ImageList.AddImage "about-icon-dark", App.path & "\Resources\images\about-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "about-icon-light", App.path & "\Resources\images\about-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-dark", App.path & "\Resources\images\config-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "config-icon-light", App.path & "\Resources\images\config-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-light", App.path & "\Resources\images\development-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "development-icon-dark", App.path & "\Resources\images\development-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "metar-icon-dark", App.path & "\Resources\images\metar-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "metar-icon-light", App.path & "\Resources\images\metar-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light", App.path & "\Resources\images\sounds-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark", App.path & "\Resources\images\sounds-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-light", App.path & "\Resources\images\windows-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark", App.path & "\Resources\images\windows-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-dark", App.path & "\Resources\images\font-icon-dark-1010.jpg"
    Cairo.ImageList.AddImage "font-icon-light", App.path & "\Resources\images\font-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-light", App.path & "\Resources\images\position-icon-light-1010.jpg"
    Cairo.ImageList.AddImage "position-icon-dark", App.path & "\Resources\images\position-icon-dark-1010.jpg"
    
    Cairo.ImageList.AddImage "metar-icon-dark-clicked", App.path & "\Resources\images\metar-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-dark-clicked", App.path & "\Resources\images\config-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-dark-clicked", App.path & "\Resources\images\font-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-dark-clicked", App.path & "\Resources\images\sounds-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-dark-clicked", App.path & "\Resources\images\position-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-dark-clicked", App.path & "\Resources\images\development-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-dark-clicked", App.path & "\Resources\images\windows-icon-dark-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-dark-clicked", App.path & "\Resources\images\about-icon-dark-600-clicked.jpg"
    
    Cairo.ImageList.AddImage "metar-icon-light-clicked", App.path & "\Resources\images\metar-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "config-icon-light-clicked", App.path & "\Resources\images\config-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "font-icon-light-clicked", App.path & "\Resources\images\font-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "sounds-icon-light-clicked", App.path & "\Resources\images\sounds-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "position-icon-light-clicked", App.path & "\Resources\images\position-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "development-icon-light-clicked", App.path & "\Resources\images\development-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "windows-icon-light-clicked", App.path & "\Resources\images\windows-icon-light-600-clicked.jpg"
    Cairo.ImageList.AddImage "about-icon-light-clicked", App.path & "\Resources\images\about-icon-light-600-clicked.jpg"
    
   On Error GoTo 0
   Exit Sub

addGeneralImagesToImageLists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addGeneralImagesToImageLists of Module modMain - probably a missing file or an incorrect named reference."

End Sub


        
'---------------------------------------------------------------------------------------
' Procedure : addDayWeatherImagesToImageLists
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : add weather icon Resources to the global ImageList that are not being pulled from the PSD directly
'---------------------------------------------------------------------------------------
'
Private Sub addDayWeatherImagesToImageLists()
    
    On Error GoTo addDayWeatherImagesToImageLists_Error
    
    Dim MyPath  As String: MyPath = vbNullString
    Dim thisFullPath  As String: thisFullPath = vbNullString
    Dim match   As String: match = vbNullString
    Dim weatherImagePresent As Boolean: weatherImagePresent = False
    Dim myName  As String: myName = vbNullString
    
    MyPath = App.path & "\Resources\images\icons_metar\"
    weatherImagePresent = False
    
    Cairo.ImageList.AddImage "weathericonimage", MyPath & "globe.png"
    Cairo.ImageList.AddImage "windiconimage", MyPath & "nowind.png"

    If Not fDirExists(MyPath) Then
        MsgBox "WARNING - The Weather Icon folder is not present in the correct location " & App.path
    End If
        
    myName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While myName <> vbNullString   ' Start the loop.
       ' Ignore the current directory and the encompassing directory.
       If myName <> "." And myName <> ".." Then
          ' Use bitwise comparison to make sure MyName is a directory.
          If (GetAttr(MyPath & myName) And vbDirectory) = vbDirectory Then
             'Debug.Print MyName   ' Display entry only if it
          End If   ' it represents a directory.
       End If
       myName = Dir   ' Get next entry.
       If myName <> "." And myName <> ".." And myName <> vbNullString Then
            match = LCase$(Right$(myName, 4))
            If match = ".png" Or match = ".PNG" Then
                thisFullPath = MyPath & myName
                ' add just the name without the .png suffix as the key to the imaagelist
                Cairo.ImageList.AddImage Left$(myName, Len(myName) - 4), thisFullPath
            End If
       End If
    Loop

   On Error GoTo 0
   Exit Sub

addDayWeatherImagesToImageLists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addDayWeatherImagesToImageLists of Module modMain - probably a missing file or an incorrect named reference."

End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : adjustSelectorMainControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the selector, individual controls
'---------------------------------------------------------------------------------------
'
Private Sub adjustSelectorMainControls()
    
    On Error GoTo adjustSelectorMainControls_Error
    
    fSelector.SelectorAdjustZoom Val(gblSelectorSize) / 100
    
    With fSelector.SelectorForm.Widgets("optlocationgreen").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("optlocationred").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("opticaogreen").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("opticaored").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("sbtnexit").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
    
    With fSelector.SelectorForm.Widgets("sbtnselect").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Enabled = False
    End With
        
    With fSelector.SelectorForm.Widgets("entericao").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
    
    With fSelector.SelectorForm.Widgets("enterlocation").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With

    With fSelector.SelectorForm.Widgets("sbtnsearch").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With

    
    With fSelector.SelectorForm.Widgets("radiobody").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
    End With
            
    With fSelector.SelectorForm.Widgets("glassblock").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
    
    fSelector.SelectorForm.Refresh
    
    On Error GoTo 0
    Exit Sub

adjustSelectorMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustSelectorMainControls, line " & Erl & ". " & "Most likely a badly-named layer in the PSD file."

End Sub



'---------------------------------------------------------------------------------------
' Procedure : adjustClipBMainControlsadjustClipBMainControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the Clipboard, individual controls
'---------------------------------------------------------------------------------------
'
Private Sub adjustClipBMainControls()
    
    On Error GoTo adjustClipBMainControls_Error
    
    fClipB.ClipBAdjustZoom Val(gblClipBSize) / 100


'    With fClipB.ClipBForm.Widgets("hourhand").Widget
'        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
'        .MousePointer = IDC_HAND
'        .Alpha = val(gblOpacity) / 100
'        .Tag = 0.25
'    End With
'
'    With fClipB.ClipBForm.Widgets("minhand").Widget
'        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
'        .MousePointer = IDC_HAND
'        .Alpha = val(gblOpacity) / 100
'        .Tag = 0.25
'    End With
'
'    With fClipB.ClipBForm.Widgets("clock").Widget
'        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
'        .MousePointer = IDC_HAND
'        .Alpha = val(gblOpacity) / 100
'        .Tag = 0.25
'    End With
'
    With fClipB.clipBForm.Widgets("clipboard").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With

    overlayClipbWidget.thisOpacity = Val(gblOpacity)
    
    On Error GoTo 0
    Exit Sub

adjustClipBMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustClipBMainControls, line " & Erl & ". " & "Most likely a badly-named layer in the PSD file."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : adjustPictorialMainControls
' Author    : beededea
' Date      : 04/06/2025
' Purpose   : give the interactive controls the hand cursor and set opacity for all elements
'---------------------------------------------------------------------------------------
'
Private Sub adjustPictorialMainControls()

   On Error GoTo adjustPictorialMainControls_Error

    fPictorial.pictAdjustZoom Val(gblPictorialGaugeSize) / 100
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fPictorial.pictorialGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
     
    With fPictorial.pictorialGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fPictorial.pictorialGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fPictorial.pictorialGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fPictorial.pictorialGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fPictorial.pictorialGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fPictorial.pictorialGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fPictorial.pictorialGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
    End With
    
'    If gblPointerAnimate = "0" Then
'        overlaypictorialWidget.pointerAnimate = False
'        fPictorial.PictorialGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        overlaypictorialWidget.pointerAnimate = True
'        fPictorial.PictorialGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
'    End If
        
    If gblPreventDraggingPictorial = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayPictorialWidget.Locked = False
        fPictorial.pictorialGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayPictorialWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fPictorial.pictorialGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayPictorialWidget.thisOpacity = Val(gblOpacity)

   On Error GoTo 0
   Exit Sub

adjustPictorialMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPictorialMainControls of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : adjustTempMainControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustTempMainControls()
   
   On Error GoTo adjustTempMainControls_Error
    
    fTemperature.tempAdjustZoom Val(gblTemperatureGaugeSize) / 100
    
    If gblGaugeFunctions = "1" Then
        WeatherMeteo.Ticking = True
    Else
        WeatherMeteo.Ticking = False
    End If
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fTemperature.temperatureGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
     
    With fTemperature.temperatureGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fTemperature.temperatureGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fTemperature.temperatureGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fTemperature.temperatureGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fTemperature.temperatureGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fTemperature.temperatureGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
    End With
    
'    If gblPointerAnimate = "0" Then
'        overlayTemperatureWidget.pointerAnimate = False
'        fTemperature.temperatureGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        overlayTemperatureWidget.pointerAnimate = True
'        fTemperature.temperatureGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
'    End If
        
    If gblPreventDraggingTemperature = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayTemperatureWidget.Locked = False
        fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayTemperatureWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayTemperatureWidget.thisOpacity = Val(gblOpacity)
    WeatherMeteo.samplingInterval = Val(gblSamplingInterval)
    overlayTemperatureWidget.thisFace = Val(gblTemperatureScale)

    
   On Error GoTo 0
   Exit Sub

adjustTempMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustTempMainControls of Module modMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : adjustAnemometerMainControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustAnemometerMainControls()
   
   On Error GoTo adjustAnemometerMainControls_Error

    ' validate the inputs of any data from the input settings file
    'Call validateInputs
    
    fAnemometer.anemoAdjustZoom Val(gblAnemometerGaugeSize) / 100
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fAnemometer.anemometerGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
     
    With fAnemometer.anemometerGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fAnemometer.anemometerGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fAnemometer.anemometerGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fAnemometer.anemometerGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fAnemometer.anemometerGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fAnemometer.anemometerGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
    End With
    
'    If gblPointerAnimate = "0" Then
'        overlayAnemoWidget.pointerAnimate = False
'        fAnemometer.anemometerGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        overlayAnemoWidget.pointerAnimate = True
'        fAnemometer.anemometerGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
'    End If
        
    If gblPreventDraggingAnemometer = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayAnemoWidget.Locked = False
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayAnemoWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayAnemoWidget.thisOpacity = Val(gblOpacity)
    
    overlayAnemoWidget.thisFace = Val(gblWindSpeedScale)
               
    
   On Error GoTo 0
   Exit Sub

adjustAnemometerMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustAnemometerMainControls of Module modMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : adjustHumidityMainControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustHumidityMainControls()
   
   On Error GoTo adjustHumidityMainControls_Error

    ' validate the inputs of any data from the input settings file
    'Call validateInputs
    
    fHumidity.humidAdjustZoom Val(gblHumidityGaugeSize) / 100
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fHumidity.humidityGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
     
    With fHumidity.humidityGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fHumidity.humidityGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fHumidity.humidityGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fHumidity.humidityGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fHumidity.humidityGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fHumidity.humidityGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
    End With
    
'    If gblPointerAnimate = "0" Then
'        overlayHumidWidget.pointerAnimate = False
'        fHumidity.humidityGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        overlayHumidWidget.pointerAnimate = True
'        fHumidity.humidityGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
'    End If
        
    If gblPreventDraggingHumidity = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayHumidWidget.Locked = False
        fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayHumidWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayHumidWidget.thisOpacity = Val(gblOpacity)

               
    
   On Error GoTo 0
   Exit Sub

adjustHumidityMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustHumidityMainControls of Module modMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : adjustBarometerMainControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the gauge, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustBarometerMainControls()
   
   On Error GoTo adjustBarometerMainControls_Error

    ' validate the inputs of any data from the input settings file
    'Call validateInputs
    
    fBarometer.baromAdjustZoom Val(gblBarometerGaugeSize) / 100
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fBarometer.barometerGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
     
    With fBarometer.barometerGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fBarometer.barometerGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
        .Tag = 0.25
    End With
      
    With fBarometer.barometerGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fBarometer.barometerGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(gblOpacity) / 100
    End With
          
    With fBarometer.barometerGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fBarometer.barometerGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(gblOpacity) / 100
    End With
        
    If gblPreventDraggingBarometer = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayBaromWidget.Locked = False
        fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(gblOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayBaromWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayBaromWidget.thisOpacity = Val(gblOpacity)
               
    overlayBaromWidget.thisFace = Val(gblPressureScale)
    
   On Error GoTo 0
   Exit Sub

adjustBarometerMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustBarometerMainControls of Module modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : setAlphaFormZordering
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setAlphaFormZordering()

   On Error GoTo setAlphaFormZordering_Error

    If Val(gblWindowLevel) = 0 Then
        Call SetWindowPos(fTemperature.temperatureGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fAnemometer.anemometerGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fHumidity.humidityGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fBarometer.barometerGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fPictorial.pictorialGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fClipB.clipBForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblWindowLevel) = 1 Then
        Call SetWindowPos(fTemperature.temperatureGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fAnemometer.anemometerGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fHumidity.humidityGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fBarometer.barometerGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fPictorial.pictorialGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fClipB.clipBForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gblWindowLevel) = 2 Then
        Call SetWindowPos(fTemperature.temperatureGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fAnemometer.anemometerGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fHumidity.humidityGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fBarometer.barometerGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fPictorial.pictorialGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fClipB.clipBForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
    End If

   On Error GoTo 0
   Exit Sub

setAlphaFormZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlphaFormZordering of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal location As String, ByVal gblSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(gblSettingsFile) Then
        
        ' general
        gblStartup = fGetINISetting(location, "startup", gblSettingsFile)
        gblGaugeFunctions = fGetINISetting(location, "gaugeFunctions", gblSettingsFile)
'        gblPointerAnimate = fGetINISetting(location, "pointerAnimate", gblSettingsFile)
        gblSamplingInterval = fGetINISetting(location, "samplingInterval", gblSettingsFile)
        gblStormTestInterval = fGetINISetting(location, "stormTestInterval", gblSettingsFile)
        gblErrorInterval = fGetINISetting(location, "errorInterval", gblSettingsFile)
        
        gblAirportsURL = fGetINISetting(location, "airportsURL", gblSettingsFile)
        
        gblTemperatureScale = fGetINISetting(location, "temperatureScale", gblSettingsFile)
        gblPressureScale = fGetINISetting(location, "pressureScale", gblSettingsFile)
        gblWindSpeedScale = fGetINISetting(location, "windSpeedScale", gblSettingsFile)
        gblMetricImperial = fGetINISetting(location, "metricImperial", gblSettingsFile)
        gblIcao = fGetINISetting(location, "icao", gblSettingsFile)

        ' configuration
        gblEnableTooltips = fGetINISetting(location, "enableTooltips", gblSettingsFile)
        gblEnablePrefsTooltips = fGetINISetting(location, "enablePrefsTooltips", gblSettingsFile)
        gblEnableBalloonTooltips = fGetINISetting(location, "enableBalloonTooltips", gblSettingsFile)
        gblShowTaskbar = fGetINISetting(location, "showTaskbar", gblSettingsFile)
        gblDpiAwareness = fGetINISetting(location, "dpiAwareness", gblSettingsFile)
        
        
        
        gblClipBSize = fGetINISetting("Software\PzClipB", "clipBSize", gblSettingsFile)
        gblSelectorSize = fGetINISetting("Software\PzSelector", "selectorSize", gblSettingsFile)
        
        gblScrollWheelDirection = fGetINISetting(location, "scrollWheelDirection", gblSettingsFile)
        
        ' position
        gblAspectHidden = fGetINISetting(location, "aspectHidden", gblSettingsFile)
        gblGaugeType = fGetINISetting(location, "gaugeType", gblSettingsFile)
        
        gblWidgetPosition = fGetINISetting(location, "widgetPosition", gblSettingsFile)
        
        gblTemperatureGaugeSize = fGetINISetting(location, "temperatureGaugeSize", gblSettingsFile)
        gblTemperatureLandscapeLocked = fGetINISetting(location, "temperatureLandscapeLocked", gblSettingsFile)
        gblTemperaturePortraitLocked = fGetINISetting(location, "temperaturePortraitLocked", gblSettingsFile)
        gblTemperatureLandscapeLockedHoffset = fGetINISetting(location, "temperatureLandscapeHoffset", gblSettingsFile)
        gblTemperatureLandscapeLockedVoffset = fGetINISetting(location, "temperatureLandscapeYoffset", gblSettingsFile)
        gblTemperaturePortraitLockedHoffset = fGetINISetting(location, "temperaturePortraitHoffset", gblSettingsFile)
        gblTemperaturePortraitLockedVoffset = fGetINISetting(location, "temperaturePortraitVoffset", gblSettingsFile)
        gblTemperatureVLocationPerc = fGetINISetting(location, "temperatureVLocationPerc", gblSettingsFile)
        gblTemperatureHLocationPerc = fGetINISetting(location, "temperatureHLocationPerc", gblSettingsFile)
        gblTemperatureFormHighDpiXPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormHighDpiXPos", gblSettingsFile)
        gblTemperatureFormHighDpiYPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormHighDpiYPos", gblSettingsFile)
        gblTemperatureFormLowDpiXPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormLowDpiXPos", gblSettingsFile)
        gblTemperatureFormLowDpiYPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormLowDpiYPos", gblSettingsFile)
        gblPreventDraggingTemperature = fGetINISetting(location, "preventDraggingTemperature", gblSettingsFile)
        
        gblAnemometerLandscapeLocked = fGetINISetting("Software\PzAnemometerGauge", "anemometerLandscapeLocked", gblSettingsFile)
        gblAnemometerPortraitLocked = fGetINISetting("Software\PzAnemometerGauge", "anemometerPortraitLocked", gblSettingsFile)
        gblAnemometerGaugeSize = fGetINISetting("Software\PzAnemometerGauge", "anemometerGaugeSize", gblSettingsFile)
        gblAnemometerLandscapeLockedHoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerLandscapeHoffset", gblSettingsFile)
        gblAnemometerLandscapeLockedVoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerLandscapeVoffset", gblSettingsFile)
        gblAnemometerPortraitLockedHoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerPortraitHoffset", gblSettingsFile)
        gblAnemometerPortraitLockedVoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerPortraitVoffset", gblSettingsFile)
        gblAnemometerVLocationPerc = fGetINISetting("Software\PzAnemometerGauge", "anemometerVLocationPerc", gblSettingsFile)
        gblAnemometerHLocationPerc = fGetINISetting("Software\PzAnemometerGauge", "anemometerHLocationPerc", gblSettingsFile)
        gblAnemometerFormHighDpiXPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormHighDpiXPos", gblSettingsFile)
        gblAnemometerFormHighDpiYPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormHighDpiYPos", gblSettingsFile)
        gblAnemometerFormLowDpiXPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormLowDpiXPos", gblSettingsFile)
        gblAnemometerFormLowDpiYPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormLowDpiYPos", gblSettingsFile)
        gblPreventDraggingAnemometer = fGetINISetting("Software\PzAnemometerGauge", "preventDraggingAnemometer", gblSettingsFile)
        
        gblHumidityLandscapeLocked = fGetINISetting("Software\PzHumidityGauge", "humidityLandscapeLocked", gblSettingsFile)
        gblHumidityPortraitLocked = fGetINISetting("Software\PzHumidityGauge", "humidityPortraitLocked", gblSettingsFile)
        gblHumidityGaugeSize = fGetINISetting("Software\PzHumidityGauge", "humidityGaugeSize", gblSettingsFile)
        gblHumidityLandscapeLockedHoffset = fGetINISetting("Software\PzHumidityGauge", "humidityLandscapeHoffset", gblSettingsFile)
        gblHumidityLandscapeLockedVoffset = fGetINISetting("Software\PzHumidityGauge", "humidityLandscapeVoffset", gblSettingsFile)
        gblHumidityPortraitLockedHoffset = fGetINISetting("Software\PzHumidityGauge", "humidityPortraitHoffset", gblSettingsFile)
        gblHumidityPortraitLockedVoffset = fGetINISetting("Software\PzHumidityGauge", "humidityPortraitVoffset", gblSettingsFile)
        gblHumidityVLocationPerc = fGetINISetting("Software\PzHumidityGauge", "humidityVLocationPerc", gblSettingsFile)
        gblHumidityHLocationPerc = fGetINISetting("Software\PzHumidityGauge", "humidityHLocationPerc", gblSettingsFile)
        gblHumidityFormHighDpiXPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormHighDpiXPos", gblSettingsFile)
        gblHumidityFormHighDpiYPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormHighDpiYPos", gblSettingsFile)
        gblHumidityFormLowDpiXPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormLowDpiXPos", gblSettingsFile)
        gblHumidityFormLowDpiYPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormLowDpiYPos", gblSettingsFile)
        gblPreventDraggingHumidity = fGetINISetting("Software\PzHumidityGauge", "preventDraggingHumidity", gblSettingsFile)
         
        gblBarometerLandscapeLocked = fGetINISetting("Software\PzBarometerGauge", "barometerLandscapeLocked", gblSettingsFile)
        gblBarometerPortraitLocked = fGetINISetting("Software\PzBarometerGauge", "barometerPortraitLocked", gblSettingsFile)
        gblBarometerGaugeSize = fGetINISetting("Software\PzBarometerGauge", "barometerGaugeSize", gblSettingsFile)
        gblBarometerLandscapeLockedHoffset = fGetINISetting("Software\PzBarometerGauge", "barometerLandscapeHoffset", gblSettingsFile)
        gblBarometerLandscapeLockedVoffset = fGetINISetting("Software\PzBarometerGauge", "barometerLandscapeVoffset", gblSettingsFile)
        gblBarometerPortraitLockedHoffset = fGetINISetting("Software\PzBarometerGauge", "barometerPortraitHoffset", gblSettingsFile)
        gblBarometerPortraitLockedVoffset = fGetINISetting("Software\PzBarometerGauge", "barometerPortraitVoffset", gblSettingsFile)
        gblBarometerVLocationPerc = fGetINISetting("Software\PzBarometerGauge", "barometerVLocationPerc", gblSettingsFile)
        gblBarometerHLocationPerc = fGetINISetting("Software\PzBarometerGauge", "barometerHLocationPerc", gblSettingsFile)
        gblBarometerFormHighDpiXPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormHighDpiXPos", gblSettingsFile)
        gblBarometerFormHighDpiYPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormHighDpiYPos", gblSettingsFile)
        gblBarometerFormLowDpiXPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormLowDpiXPos", gblSettingsFile)
        gblBarometerFormLowDpiYPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormLowDpiYPos", gblSettingsFile)
        gblPreventDraggingBarometer = fGetINISetting("Software\PzBarometerGauge", "preventDraggingBarometer", gblSettingsFile)
        
        gblPictorialGaugeSize = fGetINISetting(location, "pictorialGaugeSize", gblSettingsFile)
        gblPictorialLandscapeLocked = fGetINISetting(location, "pictorialLandscapeLocked", gblSettingsFile)
        gblPictorialPortraitLocked = fGetINISetting(location, "pictorialPortraitLocked", gblSettingsFile)
        gblPictorialLandscapeLockedHoffset = fGetINISetting(location, "pictorialLandscapeHoffset", gblSettingsFile)
        gblPictorialLandscapeLockedVoffset = fGetINISetting(location, "pictorialLandscapeYoffset", gblSettingsFile)
        gblPictorialPortraitLockedHoffset = fGetINISetting(location, "pictorialPortraitHoffset", gblSettingsFile)
        gblPictorialPortraitLockedVoffset = fGetINISetting(location, "pictorialPortraitVoffset", gblSettingsFile)
        gblPictorialVLocationPerc = fGetINISetting(location, "pictorialVLocationPerc", gblSettingsFile)
        gblPictorialHLocationPerc = fGetINISetting(location, "pictorialHLocationPerc", gblSettingsFile)
        gblPictorialFormHighDpiXPos = fGetINISetting("Software\PzPictorialGauge", "pictorialFormHighDpiXPos", gblSettingsFile)
        gblPictorialFormHighDpiYPos = fGetINISetting("Software\PzPictorialGauge", "pictorialFormHighDpiYPos", gblSettingsFile)
        gblPictorialFormLowDpiXPos = fGetINISetting("Software\PzPictorialGauge", "pictorialFormLowDpiXPos", gblSettingsFile)
        gblPictorialFormLowDpiYPos = fGetINISetting("Software\PzPictorialGauge", "pictorialFormLowDpiYPos", gblSettingsFile)
        gblPreventDraggingPictorial = fGetINISetting(location, "preventDraggingPictorial", gblSettingsFile)
             
        ' font
        gblTempFormFont = fGetINISetting(location, "tempFormFont", gblSettingsFile)
        gblPrefsFont = fGetINISetting(location, "prefsFont", gblSettingsFile)
        
        gblPrefsFontSizeHighDPI = fGetINISetting(location, "prefsFontSizeHighDPI", gblSettingsFile)
        gblPrefsFontSizeLowDPI = fGetINISetting(location, "prefsFontSizeLowDPI", gblSettingsFile)
        gblPrefsFontItalics = fGetINISetting(location, "prefsFontItalics", gblSettingsFile)
        gblPrefsFontColour = fGetINISetting(location, "prefsFontColour", gblSettingsFile)
        
        ' sound
        gblEnableSounds = fGetINISetting(location, "enableSounds", gblSettingsFile)
        
        ' development
        gblDebug = fGetINISetting(location, "debug", gblSettingsFile)
        gblDblClickCommand = fGetINISetting(location, "dblClickCommand", gblSettingsFile)
        gblOpenFile = fGetINISetting(location, "openFile", gblSettingsFile)
        gblDefaultVB6Editor = fGetINISetting(location, "defaultVB6Editor", gblSettingsFile)
        gblDefaultTBEditor = fGetINISetting(location, "defaultTBEditor", gblSettingsFile)
                
        ' other
        gblClipBFormHighDpiXPos = fGetINISetting("Software\PzClipB", "clipBFormHighDpiXPos", gblSettingsFile)
        gblClipBFormHighDpiYPos = fGetINISetting("Software\PzClipB", "clipBFormHighDpiYPos", gblSettingsFile)
        gblClipBFormLowDpiXPos = fGetINISetting("Software\PzClipB", "clipBFormLowDpiXPos", gblSettingsFile)
        gblClipBFormLowDpiYPos = fGetINISetting("Software\PzClipB", "clipBFormLowDpiYPos", gblSettingsFile)
         
        ' other
        gblSelectorFormHighDpiXPos = fGetINISetting("Software\PzSelector", "selectorFormHighDpiXPos", gblSettingsFile)
        gblSelectorFormHighDpiYPos = fGetINISetting("Software\PzSelector", "selectorFormHighDpiYPos", gblSettingsFile)
        gblSelectorFormLowDpiXPos = fGetINISetting("Software\PzSelector", "selectorFormLowDpiXPos", gblSettingsFile)
        gblSelectorFormLowDpiYPos = fGetINISetting("Software\PzSelector", "selectorFormLowDpiYPos", gblSettingsFile)
       
        gblLastSelectedTab = fGetINISetting(location, "lastSelectedTab", gblSettingsFile)
        gblSkinTheme = fGetINISetting(location, "skinTheme", gblSettingsFile)
        
        ' window
        gblWindowLevel = fGetINISetting(location, "windowLevel", gblSettingsFile)
        
        gblOpacity = fGetINISetting(location, "opacity", gblSettingsFile)
        
        gblLastUpdated = fGetINISetting(location, "lastUpdated", gblSettingsFile)
        gblMetarPref = fGetINISetting(location, "metarPref", gblSettingsFile)
        
        gblOldPressureStorage = fGetINISetting(location, "oldPressureStorage", gblSettingsFile)
        gblPressureStorageDate = fGetINISetting(location, "pressureStorageDate", gblSettingsFile)
        gblCurrentPressureValue = fGetINISetting(location, "currentPressureValue", gblSettingsFile)
    
        ' we do not want the widget to hide at startup
        'gblWidgetHidden = fGetINISetting(location, "widgetHidden", gblSettingsFile)
        gblWidgetHidden = "0"
        
        gblHidingTime = fGetINISetting(location, "hidingTime", gblSettingsFile)
        gblIgnoreMouse = fGetINISetting(location, "ignoreMouse", gblSettingsFile)
         
        gblFirstTimeRun = fGetINISetting(location, "firstTimeRun", gblSettingsFile)
        
    End If

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/06/2020
' Purpose   : validate the relevant entries from the settings.ini file in user appdata
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   On Error GoTo validateInputs_Error
            
        ' general
    If gblGaugeFunctions = vbNullString Then gblGaugeFunctions = "1" ' always turn on
'        If gblAnimationInterval = vbNullString Then gblAnimationInterval = "130"
    If gblStartup = vbNullString Then gblStartup = "1"
'    If gblPointerAnimate = vbNullString Then gblPointerAnimate = "0"
    If gblSamplingInterval = vbNullString Then gblSamplingInterval = "60"
    If gblSamplingInterval = "0" Then gblSamplingInterval = "60"
    If gblStormTestInterval = vbNullString Then gblStormTestInterval = "3600"
    If gblErrorInterval = vbNullString Then gblErrorInterval = "3"
    
    If gblAirportsURL = vbNullString Then gblAirportsURL = "https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports.dat"
    
    If gblTemperatureScale = vbNullString Then gblTemperatureScale = "0"
    If gblPressureScale = vbNullString Then gblPressureScale = "0" ' "Millibars"
    If gblWindSpeedScale = vbNullString Then gblWindSpeedScale = "0"
    If gblMetricImperial = vbNullString Then gblMetricImperial = "0"
    
    If gblIcao = vbNullString Then gblIcao = "EGSH"

    ' Configuration
    If gblEnableTooltips = vbNullString Then gblEnableTooltips = "0"
    If gblEnablePrefsTooltips = vbNullString Then gblEnablePrefsTooltips = "1"
    If gblEnableBalloonTooltips = vbNullString Then gblEnableBalloonTooltips = "1"
    If gblShowTaskbar = vbNullString Then gblShowTaskbar = "0"
    If gblDpiAwareness = vbNullString Then gblDpiAwareness = "0"
    
    If gblClipBSize = vbNullString Then gblClipBSize = "50"
    If gblSelectorSize = vbNullString Then gblSelectorSize = "100"
    
    If gblScrollWheelDirection = vbNullString Then gblScrollWheelDirection = "1"
           
    ' fonts
    If gblPrefsFont = vbNullString Then gblPrefsFont = "times new roman"
    If gblTempFormFont = vbNullString Then gblTempFormFont = gblPrefsFont
    
    If gblPrefsFontSizeHighDPI = vbNullString Then gblPrefsFontSizeHighDPI = "8"
    If gblPrefsFontSizeLowDPI = vbNullString Then gblPrefsFontSizeLowDPI = "8"
    If gblPrefsFontItalics = vbNullString Then gblPrefsFontItalics = "false"
    If gblPrefsFontColour = vbNullString Then gblPrefsFontColour = "0"

    ' sounds
    If gblEnableSounds = vbNullString Then gblEnableSounds = "1"

    ' position
    If gblAspectHidden = vbNullString Then gblAspectHidden = "0"
    If gblGaugeType = vbNullString Then gblGaugeType = "0"
    
    If gblWidgetPosition = vbNullString Then gblWidgetPosition = "0"
    
    If gblTemperatureGaugeSize = vbNullString Then gblTemperatureGaugeSize = "50"
    If gblTemperatureLandscapeLocked = vbNullString Then gblTemperatureLandscapeLocked = "0"
    If gblTemperaturePortraitLocked = vbNullString Then gblTemperaturePortraitLocked = "0"
    If gblTemperatureLandscapeLockedHoffset = vbNullString Then gblTemperatureLandscapeLockedHoffset = vbNullString
    If gblTemperatureLandscapeLockedVoffset = vbNullString Then gblTemperatureLandscapeLockedVoffset = vbNullString
    If gblTemperaturePortraitLockedHoffset = vbNullString Then gblTemperaturePortraitLockedHoffset = vbNullString
    If gblTemperaturePortraitLockedVoffset = vbNullString Then gblTemperaturePortraitLockedVoffset = vbNullString
    If gblTemperatureVLocationPerc = vbNullString Then gblTemperatureVLocationPerc = vbNullString
    If gblTemperatureHLocationPerc = vbNullString Then gblTemperatureHLocationPerc = vbNullString
    If gblPreventDraggingTemperature = vbNullString Then gblPreventDraggingTemperature = "0"
    
    If gblAnemometerGaugeSize = vbNullString Then gblAnemometerGaugeSize = "50"
    If gblAnemometerLandscapeLocked = vbNullString Then gblAnemometerLandscapeLocked = "0"
    If gblAnemometerPortraitLocked = vbNullString Then gblAnemometerPortraitLocked = "0"
    If gblAnemometerLandscapeLockedHoffset = vbNullString Then gblAnemometerLandscapeLockedHoffset = vbNullString
    If gblAnemometerLandscapeLockedVoffset = vbNullString Then gblAnemometerLandscapeLockedVoffset = vbNullString
    If gblAnemometerPortraitLockedHoffset = vbNullString Then gblAnemometerPortraitLockedHoffset = vbNullString
    If gblAnemometerPortraitLockedVoffset = vbNullString Then gblAnemometerPortraitLockedVoffset = vbNullString
    If gblAnemometerVLocationPerc = vbNullString Then gblAnemometerVLocationPerc = vbNullString
    If gblAnemometerHLocationPerc = vbNullString Then gblAnemometerHLocationPerc = vbNullString
    If gblPreventDraggingAnemometer = vbNullString Then gblPreventDraggingAnemometer = "0"
    
    If gblHumidityGaugeSize = vbNullString Then gblHumidityGaugeSize = "50"
    If gblHumidityLandscapeLocked = vbNullString Then gblHumidityLandscapeLocked = "0"
    If gblHumidityPortraitLocked = vbNullString Then gblHumidityPortraitLocked = "0"
    If gblHumidityLandscapeLockedHoffset = vbNullString Then gblHumidityLandscapeLockedHoffset = vbNullString
    If gblHumidityLandscapeLockedVoffset = vbNullString Then gblHumidityLandscapeLockedVoffset = vbNullString
    If gblHumidityPortraitLockedHoffset = vbNullString Then gblHumidityPortraitLockedHoffset = vbNullString
    If gblHumidityPortraitLockedVoffset = vbNullString Then gblHumidityPortraitLockedVoffset = vbNullString
    If gblHumidityVLocationPerc = vbNullString Then gblHumidityVLocationPerc = vbNullString
    If gblHumidityHLocationPerc = vbNullString Then gblHumidityHLocationPerc = vbNullString
    If gblPreventDraggingHumidity = vbNullString Then gblPreventDraggingHumidity = "0"
    
    If gblBarometerGaugeSize = vbNullString Then gblBarometerGaugeSize = "50"
    If gblBarometerLandscapeLocked = vbNullString Then gblBarometerLandscapeLocked = "0"
    If gblBarometerPortraitLocked = vbNullString Then gblBarometerPortraitLocked = "0"
    If gblBarometerLandscapeLockedHoffset = vbNullString Then gblBarometerLandscapeLockedHoffset = vbNullString
    If gblBarometerLandscapeLockedVoffset = vbNullString Then gblBarometerLandscapeLockedVoffset = vbNullString
    If gblBarometerPortraitLockedHoffset = vbNullString Then gblBarometerPortraitLockedHoffset = vbNullString
    If gblBarometerPortraitLockedVoffset = vbNullString Then gblBarometerPortraitLockedVoffset = vbNullString
    If gblBarometerVLocationPerc = vbNullString Then gblBarometerVLocationPerc = vbNullString
    If gblBarometerHLocationPerc = vbNullString Then gblBarometerHLocationPerc = vbNullString
    If gblPreventDraggingBarometer = vbNullString Then gblPreventDraggingBarometer = "0"
            
    If gblPictorialGaugeSize = vbNullString Then gblPictorialGaugeSize = "50"
    If gblPictorialLandscapeLocked = vbNullString Then gblPictorialLandscapeLocked = "0"
    If gblPictorialPortraitLocked = vbNullString Then gblPictorialPortraitLocked = "0"
    If gblPictorialLandscapeLockedHoffset = vbNullString Then gblPictorialLandscapeLockedHoffset = vbNullString
    If gblPictorialLandscapeLockedVoffset = vbNullString Then gblPictorialLandscapeLockedVoffset = vbNullString
    If gblPictorialPortraitLockedHoffset = vbNullString Then gblPictorialPortraitLockedHoffset = vbNullString
    If gblPictorialPortraitLockedVoffset = vbNullString Then gblPictorialPortraitLockedVoffset = vbNullString
    If gblPictorialVLocationPerc = vbNullString Then gblPictorialVLocationPerc = vbNullString
    If gblPictorialHLocationPerc = vbNullString Then gblPictorialHLocationPerc = vbNullString
    If gblPreventDraggingPictorial = vbNullString Then gblPreventDraggingPictorial = "0"
    
    ' development
    If gblDebug = vbNullString Then gblDebug = "0"
    If gblDblClickCommand = vbNullString Then gblDblClickCommand = vbNullString
    If gblOpenFile = vbNullString Then gblOpenFile = vbNullString
    If gblDefaultVB6Editor = vbNullString Then gblDefaultVB6Editor = vbNullString
    If gblDefaultTBEditor = vbNullString Then gblDefaultTBEditor = vbNullString
    
    ' window
    If gblWindowLevel = vbNullString Then gblWindowLevel = "1" 'WindowLevel", gblSettingsFile)
    If gblOpacity = vbNullString Then gblOpacity = "100"
    If gblWidgetHidden = vbNullString Then gblWidgetHidden = "0"
    If gblHidingTime = vbNullString Then gblHidingTime = "0"
    If gblIgnoreMouse = vbNullString Then gblIgnoreMouse = "0"
    
    ' other
    If gblFirstTimeRun = vbNullString Then gblFirstTimeRun = "true"
    If gblLastSelectedTab = vbNullString Then gblLastSelectedTab = "general"
    If gblSkinTheme = vbNullString Then gblSkinTheme = "dark"
    
    If gblLastUpdated = vbNullString Then gblLastUpdated = CStr(Now())
    If gblMetarPref = vbNullString Then gblMetarPref = "ICAO"
        
    If gblOldPressureStorage = vbNullString Then gblOldPressureStorage = "0"
    If gblPressureStorageDate = vbNullString Then gblPressureStorageDate = CStr(Now())
    If gblCurrentPressureValue = vbNullString Then gblCurrentPressureValue = "0"
    
 
   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of form modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getTrinketsFile
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/10/2019
' Purpose   : get this tool's entry in the trinkets settings file and assign the app.path
'---------------------------------------------------------------------------------------
'
Private Sub getTrinketsFile()
    On Error GoTo getTrinketsFile_Error
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gblTrinketsDir = fSpecialFolder(feUserAppData) & "\trinkets" ' just for this user alone
    gblTrinketsFile = gblTrinketsDir & "\" & widgetName1 & ".ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gblTrinketsDir) Then
        MkDir gblTrinketsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblTrinketsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblTrinketsFile For Output As #iFileNo
        Write #iFileNo, App.path & "\" & App.EXEName & ".exe"
        Write #iFileNo,
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getTrinketsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getTrinketsFile of Form modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    On Error GoTo getToolSettingsFile_Error
    ''If debugflg = 1  Then Debug.Print "%getToolSettingsFile"
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gblSettingsDir = fSpecialFolder(feUserAppData) & "\PzTemperatureGauge" ' just for this user alone
    gblSettingsFile = gblSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gblSettingsDir) Then
        MkDir gblSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gblSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gblSettingsFile For Output As #iFileNo
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form modMain"

End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : configureTimers
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 07/05/2023
' Purpose   : configure any global timers here
'---------------------------------------------------------------------------------------
'
Private Sub configureTimers()

    On Error GoTo configureTimers_Error
    
    oldgblSettingsModificationTime = FileDateTime(gblSettingsFile)

    frmTimer.rotationTimer.Enabled = True
    frmTimer.settingsTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

configureTimers_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configureTimers of Module modMain"
            Resume Next
          End If
    End With
 
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : setHidingTime
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 07/05/2023
' Purpose   : set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
'---------------------------------------------------------------------------------------
'
Private Sub setHidingTime()
    
    On Error GoTo setHidingTime_Error

    If gblHidingTime = "0" Then minutesToHide = 1
    If gblHidingTime = "1" Then minutesToHide = 5
    If gblHidingTime = "2" Then minutesToHide = 10
    If gblHidingTime = "3" Then minutesToHide = 20
    If gblHidingTime = "4" Then minutesToHide = 30
    If gblHidingTime = "5" Then minutesToHide = 60

    On Error GoTo 0
    Exit Sub

setHidingTime_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingTime of Module modMain"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createRCFormsOnCurrentDisplay
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createRCFormsOnCurrentDisplay()
    On Error GoTo createRCFormsOnCurrentDisplay_Error

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowAboutForm(widgetName1)
    End With
    
    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowHelpForm(widgetName1)
    End With

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndShowLicenceForm(widgetName1)
    End With
    
        On Error GoTo 0
    Exit Sub

createRCFormsOnCurrentDisplay_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createRCFormsOnCurrentDisplay of Module modMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : handleUnhideMode
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 13/05/2023
' Purpose   : when run in 'unhide' mode it writes the settings file then exits, the other
'             running but hidden process will unhide itself by timer.
'---------------------------------------------------------------------------------------
'
Private Sub handleUnhideMode(ByVal thisUnhideMode As String)
    
    On Error GoTo handleUnhideMode_Error

    If thisUnhideMode = "unhide" Then     'parse the command line
        gblUnhide = "true"
        sPutINISetting "Software\PzTemperatureGauge", "unhide", gblUnhide, gblSettingsFile
        Call thisForm_Unload
        End
    End If

    On Error GoTo 0
    Exit Sub

handleUnhideMode_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure handleUnhideMode of Module modMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadTemperatureExcludePathCollection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadTemperatureExcludePathCollection()

    'all of these will be rendered in cwOverlayTemp in the same order as below
    On Error GoTo loadTemperatureExcludePathCollection_Error

    With fTemperature.collTemperaturePSDNonUIElements ' the exclude list
        .Add Empty, "centigradeface"
        .Add Empty, "fahrenheitface"
        .Add Empty, "kelvinface"
        
        .Add Empty, "faceweathering"

        .Add Empty, "bigreflection"     'all reflections
        .Add Empty, "windowreflection"

        .Add Empty, "bluelamptrue"
        .Add Empty, "bluelampfalse"

        .Add Empty, "redlamptrue"
        .Add Empty, "redlampfalse"
        
        .Add Empty, "secondshadow" 'clock-hand-seconds-shadow
        .Add Empty, "secondhand"   'clock-hand-seconds

    End With

   On Error GoTo 0
   Exit Sub

loadTemperatureExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadTemperatureExcludePathCollection of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadSelectorExcludePathCollection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadSelectorExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadSelectorExcludePathCollection_Error

    With fSelector.collSelectorPSDNonUIElements ' the exclude list
        .Add Empty, "radioknobtwo"
        .Add Empty, "radioknobone"
    End With

   On Error GoTo 0
   Exit Sub

loadSelectorExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadSelectorExcludePathCollection of Module modMain"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : loadClipBExcludePathCollection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadClipBExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadClipBExcludePathCollection_Error

    With fClipB.collClipBPSDNonUIElements ' the exclude list

        '.Add Empty, "clipboard"
        .Add Empty, "clock"
        .Add Empty, "hourhand"
        .Add Empty, "minhand"
        .Add Empty, "text"
    End With

   On Error GoTo 0
   Exit Sub

loadClipBExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadClipBExcludePathCollection of Module modMain"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : loadAnemometerExcludePathCollection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadAnemometerExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadAnemometerExcludePathCollection_Error

    With fAnemometer.collAnemometerPSDNonUIElements ' the exclude list

        .Add Empty, "anemometerknotsface"
        .Add Empty, "anemometermetresface"
        
        .Add Empty, "bigreflection"     'all reflections
        .Add Empty, "windowreflection"

        .Add Empty, "amberlamp"
        .Add Empty, "purplelamp"
        .Add Empty, "redlamptrue"
        .Add Empty, "redlampfalse"
        
        .Add Empty, "directionpointer"
        .Add Empty, "directionshadow"

        .Add Empty, "pointershadow"
        .Add Empty, "pointer"
       
        
    End With

   On Error GoTo 0
   Exit Sub

loadAnemometerExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadAnemometerExcludePathCollection of Module modMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : loadHumidityExcludePathCollection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadHumidityExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadHumidityExcludePathCollection_Error

    With fHumidity.collHumidityPSDNonUIElements ' the exclude list

        .Add Empty, "Humidityface"
        
        .Add Empty, "bigreflection"     'all reflections
        .Add Empty, "windowreflection"
        
        .Add Empty, "pointerShadow"
        .Add Empty, "pointer"
       
        
    End With

   On Error GoTo 0
   Exit Sub

loadHumidityExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHumidityExcludePathCollection of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : loadBarometerExcludePathCollection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadBarometerExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadBarometerExcludePathCollection_Error

    With fBarometer.collBarometerPSDNonUIElements ' the exclude list

        .Add Empty, "barometerinhg23-31face"
        .Add Empty, "barometerinhg18-24face"
        .Add Empty, "barometerhpa800-1066face"
        .Add Empty, "barometerhpa600-800face"
        .Add Empty, "barometermb800-1066face"
        .Add Empty, "barometermb600-800face"
        .Add Empty, "barometermmhg450-600face"
        .Add Empty, "barometermmhg600-800face"
        
        .Add Empty, "greenlamp"
        .Add Empty, "redlamp"
        
        .Add Empty, "bigreflection"     'all reflections
        .Add Empty, "windowreflection"
        
        .Add Empty, "manualpointer"
       
        .Add Empty, "pointerShadow"
        .Add Empty, "pointer"
       
        
    End With

   On Error GoTo 0
   Exit Sub

loadBarometerExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadBarometerExcludePathCollection of Module modMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : loadPictorialExcludePathCollection
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 30/07/2023
' Purpose   : Do not create Widgets for those in the exclude list.
'             all non UI-interacting elements (no mouse events) must be inserted here
'---------------------------------------------------------------------------------------
'
Private Sub loadPictorialExcludePathCollection()

    'all of these will be rendered in cwOverlay in the same order as below
    On Error GoTo loadPictorialExcludePathCollection_Error

    With fPictorial.collPictorialPSDNonUIElements ' the exclude list
        
        .Add Empty, "greenlamp"
        .Add Empty, "redlamp"
        .Add Empty, "greenlampfalse"
        .Add Empty, "bigreflection"     'all reflections
        .Add Empty, "manualpointer"
        
    End With

   On Error GoTo 0
   Exit Sub

loadPictorialExcludePathCollection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPictorialExcludePathCollection of Module modMain"

End Sub

' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen, see Steamydock for a potential replacement?
'---------------------------------------------------------------------------------------
' Procedure : msgBoxA
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 20/05/2022
' Purpose   : ans = msgBoxA("main message", vbOKOnly, "title bar message", False)
'---------------------------------------------------------------------------------------
'
Public Function msgBoxA(ByVal msgBoxPrompt As String, Optional ByVal msgButton As VbMsgBoxResult, Optional ByVal msgTitle As String, Optional ByVal msgShowAgainChkBox As Boolean = False, Optional ByRef msgContext As String = "none") As Integer
     
    ' set the defined properties of a form
    On Error GoTo msgBoxA_Error

    frmMessage.propMessage = msgBoxPrompt
    frmMessage.propTitle = msgTitle
    frmMessage.propShowAgainChkBox = msgShowAgainChkBox
    frmMessage.propButtonVal = msgButton
    frmMessage.propMsgContext = msgContext
    Call frmMessage.Display ' run a subroutine in the form that displays the form

    msgBoxA = frmMessage.propReturnedValue

    On Error GoTo 0
    Exit Function

msgBoxA_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure msgBoxA of Module mdlMain"
            Resume Next
          End If
    End With

End Function


' ----------------------------------------------------------------
' Procedure Name: setTaskbarEntry
' Purpose: set taskbar entry
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 13/05/2024
' ----------------------------------------------------------------
Private Sub setTaskbarEntry()
    
    On Error GoTo setTaskbarEntry_Error
    
    If gblShowTaskbar = "0" Then
        fTemperature.temperatureGaugeForm.ShowInTaskbar = False
        fAnemometer.anemometerGaugeForm.ShowInTaskbar = False
        fHumidity.humidityGaugeForm.ShowInTaskbar = False
        fBarometer.barometerGaugeForm.ShowInTaskbar = False
        fPictorial.pictorialGaugeForm.ShowInTaskbar = False
        fClipB.clipBForm.ShowInTaskbar = False
    Else
        fTemperature.temperatureGaugeForm.ShowInTaskbar = True
        fAnemometer.anemometerGaugeForm.ShowInTaskbar = True
        fHumidity.humidityGaugeForm.ShowInTaskbar = True
        fBarometer.barometerGaugeForm.ShowInTaskbar = True
        fPictorial.pictorialGaugeForm.ShowInTaskbar = True
        fClipB.clipBForm.ShowInTaskbar = True
    End If

    On Error GoTo 0
    Exit Sub

setTaskbarEntry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setTaskbarEntry, line " & Erl & "."

End Sub



' ----------------------------------------------------------------
' Procedure Name: setMenuItems
' Purpose: set menu items
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 13/05/2024
' ----------------------------------------------------------------
Private Sub setMenuItems()
    Dim thisEditor As String: thisEditor = vbNullString
    
    On Error GoTo setMenuItems_Error
    
    If gblGaugeFunctions = "1" Then
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If gblDebug = "1" Then
        #If TWINBASIC Then
            If gblDefaultTBEditor <> vbNullString Then thisEditor = gblDefaultTBEditor
        #Else
            If gblDefaultVB6Editor <> vbNullString Then thisEditor = gblDefaultVB6Editor
        #End If
        
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & thisEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
    
    On Error GoTo 0
    Exit Sub

setMenuItems_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMenuItems, line " & Erl & "."

End Sub
    



'---------------------------------------------------------------------------------------
' Procedure : setAutomaticDPIState
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : check first usage via licence acceptance value and then set initial DPI awareness
'---------------------------------------------------------------------------------------
'
Private Sub setAutomaticDPIState(ByRef licenceState As Integer)
   On Error GoTo setAutomaticDPIState_Error

    licenceState = fLicenceState()
    If licenceState = 0 Then
        Call testDPIAndSetInitialAwareness ' determine High DPI awareness or not by default on first run
    Else
        Call setDPIaware ' determine the user settings for DPI awareness, for this program and all its forms.
    End If

   On Error GoTo 0
   Exit Sub

setAutomaticDPIState_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAutomaticDPIState of Module modMain"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : loadPreferenceForm
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : load the preferences form but don't yet show it, speeds up access to the prefs via the menu
'---------------------------------------------------------------------------------------
'
Private Sub loadPreferenceForm()
        
   On Error GoTo loadPreferenceForm_Error

    If panzerPrefs.IsLoaded = False Then
        Load panzerPrefs
        'gblPrefsFormResizedInCode = True
        Call panzerPrefs.PrefsForm_Resize_Event
    End If

   On Error GoTo 0
   Exit Sub

loadPreferenceForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPreferenceForm of Module modMain"
End Sub
