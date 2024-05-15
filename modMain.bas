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

Public sunriseSunset As cwSunriseSunset
Public widgetName1 As String
Public widgetName2 As String
Public widgetName3 As String
Public widgetName4 As String
Public widgetName5 As String
Public widgetName6 As String

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

    Dim licenceState As Integer: licenceState = 0

    On Error GoTo main_routine_Error
    
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
    
    prefsCurrentWidth = 9075
    prefsCurrentHeight = 16450
    
    gblOriginatingForm = "temperatureForm"
    
    firstPoll = True
    
    'startupFlg = True ' this is used to prevent some control initialisations from running code at startup

    extractCommand = Command$ ' capture any parameter passed, remove if a soft reload
    If restart = True Then extractCommand = vbNullString
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    ' create dictionary collection instead of an array to load dropdown list
    'Set collValidLocations = CreateObject("Scripting.Dictionary") ' tested with all three
    Set collValidLocations = New_c.Collection(False)
    
    'add Resources to the global ImageList
    Call addGeneralImagesToImageLists
    
    ' check the Windows version
    classicThemeCapable = fTestClassicThemeCapable
  
    ' get this tool's entry in the trinkets settings file and assign the app.path
    Call getTrinketsFile
  
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the dock settings from the new configuration file
    Call readSettingsFile("Software\PzTemperatureGauge", PzGSettingsFile)
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' check first usage via licence acceptance value and then set initial DPI awareness
    licenceState = fLicenceState()
    If licenceState = 0 Then
        Call testDPIAndSetInitialAwareness ' determine High DPI awareness or not by default on first run
    Else
        Call setDPIaware ' determine the user settings for DPI awareness, for this program and all its forms.
    End If

    'load the collection for storing the overlay surfaces with its relevant keys direct from each PSD
    If restart = False Then
        Call loadTemperatureExcludePathCollection ' no need to reload the collTemperaturePSDNonUIElements layer name keys on a reload
        Call loadSelectorExcludePathCollection
        Call loadClipBExcludePathCollection
        Call loadAnemometerExcludePathCollection
        Call loadHumidityExcludePathCollection
        Call loadBarometerExcludePathCollection
    End If
    
    ' start the load of the PSD files using the RC6 PSD-Parser.instance
    Call fTemperature.InitTemperatureFromPSD(temperaturePSDFullPath)
    Call fSelector.InitSelectorFromPSD(selectorPSDFullPath)
    Call fClipB.InitClipBFromPSD(clipBPSDFullPath)
    Call fAnemometer.InitAnemometerFromPSD(anemometerPSDFullPath)
    Call fHumidity.InitHumidityFromPSD(HumidityPSDFullPath)
    Call fBarometer.InitBarometerFromPSD(barometerPSDFullPath)
    
    ' resolve VB6 sizing width bug
    Call determineScreenDimensions
            
    ' initialise and create the three main RC forms on the current display
    Call createRCFormsOnCurrentDisplay
    
    ' check the selected monitor properties
    Call monitorProperties(fTemperature.temperatureGaugeForm)  ' might use RC6 for this?
    
    ' place the form at the saved location
    Call makeVisibleFormElements
    
    ' run the functions that are ALSO called at reload time elsewhere.
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' set menu items
    Call setMenuItems
    
    ' set taskbar entry
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
    Load panzerPrefs
    
    'load the message form but don't yet show it, speeds up access to the message form when needed.
    Load frmMessage
    
    ' display licence screen on first usage
    Call showLicence(fLicenceState)
    
    ' make the prefs appear on the first time running
    Call checkFirstTime
 
    ' configure any global timers here
    Call configureTimers
    
    'startupFlg = False
        
    ' RC message pump will auto-exit when Cairo Forms > 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload.
    If Cairo.WidgetForms.Count = 0 Then Cairo.WidgetForms.EnterMessageLoop
     
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

    If PzGFirstTimeRun = "true" Then
        'MsgBox "checkFirstTime"

        Call makeProgramPreferencesAvailable
        PzGFirstTimeRun = "false"
        sPutINISetting "Software\PzTemperatureGauge", "firstTimeRun", PzGFirstTimeRun, PzGSettingsFile
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
    PzGStartup = vbNullString
    PzGGaugeFunctions = vbNullString
'    PzGPointerAnimate = vbNullString
    PzGSamplingInterval = vbNullString
    PzGStormTestInterval = vbNullString
    PzGErrorInterval = vbNullString
    PzGAirportsURL = vbNullString
    
    PzGIcao = vbNullString

    ' config
    PzGEnableTooltips = vbNullString
    PzGEnablePrefsTooltips = vbNullString
    PzGEnableBalloonTooltips = vbNullString
    PzGShowTaskbar = vbNullString
    PzGDpiAwareness = vbNullString
    
    
    PzGClipBSize = vbNullString
    PzGSelectorSize = vbNullString
    
    PzGScrollWheelDirection = vbNullString
    
    ' position
    PzGAspectHidden = vbNullString
    PzGGaugeType = vbNullString
    PzGWidgetPosition = vbNullString
    
    PzGTemperatureLandscape = vbNullString
    PzGTemperaturePortrait = vbNullString
    PzGTemperatureGaugeSize = vbNullString
    PzGTemperatureLandscapeHoffset = vbNullString
    PzGTemperatureLandscapeVoffset = vbNullString
    PzGTemperaturePortraitHoffset = vbNullString
    PzGTemperaturePortraitVoffset = vbNullString
    PzGTemperatureVLocationPerc = vbNullString
    PzGTemperatureHLocationPerc = vbNullString
    PzGPreventDraggingTemperature = vbNullString
    PzGTemperatureFormHighDpiXPos = vbNullString
    PzGTemperatureFormHighDpiYPos = vbNullString
    PzGTemperatureFormLowDpiXPos = vbNullString
    PzGTemperatureFormLowDpiYPos = vbNullString
    
    PzGAnemometerGaugeSize = vbNullString
    PzGAnemometerLandscape = vbNullString
    PzGAnemometerPortrait = vbNullString
    PzGAnemometerFormHighDpiXPos = vbNullString
    PzGAnemometerFormHighDpiYPos = vbNullString
    PzGAnemometerFormLowDpiXPos = vbNullString
    PzGAnemometerFormLowDpiYPos = vbNullString
    PzGAnemometerLandscapeHoffset = vbNullString
    PzGAnemometerLandscapeVoffset = vbNullString
    PzGAnemometerPortraitHoffset = vbNullString
    PzGAnemometerPortraitVoffset = vbNullString
    PzGPreventDraggingAnemometer = vbNullString
    
    PzGBarometerGaugeSize = vbNullString
    PzGBarometerLandscape = vbNullString
    PzGBarometerPortrait = vbNullString
    PzGBarometerFormHighDpiXPos = vbNullString
    PzGBarometerFormHighDpiYPos = vbNullString
    PzGBarometerFormLowDpiXPos = vbNullString
    PzGBarometerFormLowDpiYPos = vbNullString
    PzGBarometerLandscapeHoffset = vbNullString
    PzGBarometerLandscapeVoffset = vbNullString
    PzGBarometerPortraitHoffset = vbNullString
    PzGBarometerPortraitVoffset = vbNullString
    PzGPreventDraggingBarometer = vbNullString
        
    ' sounds
    PzGEnableSounds = vbNullString
    
    ' development
    PzGDebug = vbNullString
    PzGDblClickCommand = vbNullString
    PzGOpenFile = vbNullString
    PzGDefaultEditor = vbNullString
         
    ' font
    PzGTempFormFont = vbNullString
    PzGPrefsFont = vbNullString
    PzGPrefsFontSizeHighDPI = vbNullString
    PzGPrefsFontSizeLowDPI = vbNullString
    PzGPrefsFontItalics = vbNullString
    PzGPrefsFontColour = vbNullString
    
    ' window
    PzGWindowLevel = vbNullString
    

    PzGOpacity = vbNullString

    
    PzGWidgetHidden = vbNullString
    PzGHidingTime = vbNullString
    PzGIgnoreMouse = vbNullString
    PzGFirstTimeRun = vbNullString
    
    ' general storage variables declared
    PzGSettingsDir = vbNullString
    PzGSettingsFile = vbNullString
    
    PzGTrinketsDir = vbNullString
    PzGTrinketsFile = vbNullString
    
    PzGClipBFormHighDpiXPos = vbNullString
    PzGClipBFormHighDpiYPos = vbNullString
    PzGClipBFormLowDpiXPos = vbNullString
    PzGClipBFormLowDpiYPos = vbNullString
    
    PzGSelectorFormHighDpiXPos = vbNullString
    PzGSelectorFormHighDpiYPos = vbNullString
    PzGSelectorFormLowDpiXPos = vbNullString
    PzGSelectorFormLowDpiYPos = vbNullString
    
    PzGLastSelectedTab = vbNullString
    PzGSkinTheme = vbNullString
    
    PzGLastUpdated = vbNullString
    PzGMetarPref = vbNullString
    PzGPressureScale = vbNullString
    
    PzGOldPressureStorage = vbNullString
    PzGPressureStorageDate = vbNullString
    PzGCurrentPressureValue = vbNullString
    
    
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
    oldPzGSettingsModificationTime = #1/1/2000 12:00:00 PM#
    
    gblJustAwoken = False

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
    Cairo.ImageList.AddImage "helpWeather", App.path & "\Resources\images\panzerWeather-help.png"
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
' Procedure : adjustSelectorMainControls
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the selector, individual controls
'---------------------------------------------------------------------------------------
'
Private Sub adjustSelectorMainControls()
    
    On Error GoTo adjustSelectorMainControls_Error
    
    fSelector.SelectorAdjustZoom Val(PzGSelectorSize) / 100
    
    With fSelector.SelectorForm.Widgets("optlocationgreen").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("optlocationred").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("opticaogreen").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("opticaored").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
    
    With fSelector.SelectorForm.Widgets("sbtnexit").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
    
    With fSelector.SelectorForm.Widgets("sbtnselect").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Enabled = False
    End With
        
    With fSelector.SelectorForm.Widgets("entericao").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
    
    With fSelector.SelectorForm.Widgets("enterlocation").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With

    With fSelector.SelectorForm.Widgets("sbtnsearch").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With

    
    With fSelector.SelectorForm.Widgets("radiobody").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(PzGOpacity) / 100
    End With
    
            
    With fSelector.SelectorForm.Widgets("glassblock").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
    
    
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
    
    fClipB.ClipBAdjustZoom Val(PzGClipBSize) / 100


'    With fClipB.ClipBForm.Widgets("hourhand").Widget
'        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
'        .MousePointer = IDC_HAND
'        .Alpha = val(PzGOpacity) / 100
'        .Tag = 0.25
'    End With
'
'    With fClipB.ClipBForm.Widgets("minhand").Widget
'        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
'        .MousePointer = IDC_HAND
'        .Alpha = val(PzGOpacity) / 100
'        .Tag = 0.25
'    End With
'
'    With fClipB.ClipBForm.Widgets("clock").Widget
'        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
'        .MousePointer = IDC_HAND
'        .Alpha = val(PzGOpacity) / 100
'        .Tag = 0.25
'    End With
'
    With fClipB.clipBForm.Widgets("clipboard").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With

    overlayClipbWidget.thisOpacity = Val(PzGOpacity)
    
    On Error GoTo 0
    Exit Sub

adjustClipBMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustClipBMainControls, line " & Erl & ". " & "Most likely a badly-named layer in the PSD file."

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
    
    fTemperature.tempAdjustZoom Val(PzGTemperatureGaugeSize) / 100
    
    If PzGGaugeFunctions = "1" Then
        overlayTemperatureWidget.Ticking = True
    Else
        overlayTemperatureWidget.Ticking = False
    End If
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fTemperature.temperatureGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
     
    With fTemperature.temperatureGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fTemperature.temperatureGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fTemperature.temperatureGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fTemperature.temperatureGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fTemperature.temperatureGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fTemperature.temperatureGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(PzGOpacity) / 100
    End With
    
'    If PzGPointerAnimate = "0" Then
'        overlayTemperatureWidget.pointerAnimate = False
'        fTemperature.temperatureGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(PzGOpacity) / 100
'    Else
'        overlayTemperatureWidget.pointerAnimate = True
'        fTemperature.temperatureGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
'    End If
        
    If PzGPreventDraggingTemperature = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayTemperatureWidget.Locked = False
        fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(PzGOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayTemperatureWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fTemperature.temperatureGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayTemperatureWidget.thisOpacity = Val(PzGOpacity)
    overlayTemperatureWidget.samplingInterval = Val(PzGSamplingInterval)
    overlayTemperatureWidget.thisFace = Val(PzGTemperatureScale)

    
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
    
    fAnemometer.anemoAdjustZoom Val(PzGAnemometerGaugeSize) / 100
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fAnemometer.anemometerGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
     
    With fAnemometer.anemometerGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fAnemometer.anemometerGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fAnemometer.anemometerGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fAnemometer.anemometerGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fAnemometer.anemometerGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fAnemometer.anemometerGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(PzGOpacity) / 100
    End With
    
'    If PzGPointerAnimate = "0" Then
'        overlayAnemoWidget.pointerAnimate = False
'        fAnemometer.anemometerGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(PzGOpacity) / 100
'    Else
'        overlayAnemoWidget.pointerAnimate = True
'        fAnemometer.anemometerGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
'    End If
        
    If PzGPreventDraggingAnemometer = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayAnemoWidget.Locked = False
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(PzGOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayAnemoWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fAnemometer.anemometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayAnemoWidget.thisOpacity = Val(PzGOpacity)
               
    
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
    
    fHumidity.humidAdjustZoom Val(PzGHumidityGaugeSize) / 100
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fHumidity.humidityGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
     
    With fHumidity.humidityGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fHumidity.humidityGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fHumidity.humidityGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fHumidity.humidityGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fHumidity.humidityGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fHumidity.humidityGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(PzGOpacity) / 100
    End With
    
'    If PzGPointerAnimate = "0" Then
'        overlayHumidWidget.pointerAnimate = False
'        fHumidity.humidityGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = Val(PzGOpacity) / 100
'    Else
'        overlayHumidWidget.pointerAnimate = True
'        fHumidity.humidityGaugeForm.Widgets("housing/tickbutton").Widget.Alpha = 0
'    End If
        
    If PzGPreventDraggingHumidity = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayHumidWidget.Locked = False
        fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(PzGOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayHumidWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fHumidity.humidityGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayHumidWidget.thisOpacity = Val(PzGOpacity)
               
    
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
    
    fBarometer.baromAdjustZoom Val(PzGBarometerGaugeSize) / 100
    
    ' set the characteristics of the interactive areas
    ' Note: set the Hover colour close to the original layer to avoid too much intrusion, 0 being grey
    With fBarometer.barometerGaugeForm.Widgets("housing/helpbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
     
    With fBarometer.barometerGaugeForm.Widgets("housing/startbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fBarometer.barometerGaugeForm.Widgets("housing/stopbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
        .Tag = 0.25
    End With
      
    With fBarometer.barometerGaugeForm.Widgets("housing/switchfacesbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
          
    With fBarometer.barometerGaugeForm.Widgets("housing/prefsbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Alpha = Val(PzGOpacity) / 100
    End With
          
    With fBarometer.barometerGaugeForm.Widgets("housing/tickbutton").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
    End With
    
    With fBarometer.barometerGaugeForm.Widgets("housing/surround").Widget
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_SIZEALL
        .Alpha = Val(PzGOpacity) / 100
    End With
        
    If PzGPreventDraggingBarometer = "0" Then
        menuForm.mnuLockTemperatureGauge.Checked = False
        overlayBaromWidget.Locked = False
        fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = Val(PzGOpacity) / 100
    Else
        menuForm.mnuLockTemperatureGauge.Checked = True
        overlayBaromWidget.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
        fBarometer.barometerGaugeForm.Widgets("housing/lockbutton").Widget.Alpha = 0
    End If

    overlayBaromWidget.thisOpacity = Val(PzGOpacity)
               
    
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

    If Val(PzGWindowLevel) = 0 Then
        Call SetWindowPos(fTemperature.temperatureGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fAnemometer.anemometerGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fHumidity.humidityGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fBarometer.barometerGaugeForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fClipB.clipBForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PzGWindowLevel) = 1 Then
        Call SetWindowPos(fTemperature.temperatureGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fAnemometer.anemometerGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fHumidity.humidityGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fBarometer.barometerGaugeForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fClipB.clipBForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(PzGWindowLevel) = 2 Then
        Call SetWindowPos(fTemperature.temperatureGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fAnemometer.anemometerGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fHumidity.humidityGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
        Call SetWindowPos(fBarometer.barometerGaugeForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
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
Public Sub readSettingsFile(ByVal location As String, ByVal PzGSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(PzGSettingsFile) Then
        
        ' general
        PzGStartup = fGetINISetting(location, "startup", PzGSettingsFile)
        PzGGaugeFunctions = fGetINISetting(location, "gaugeFunctions", PzGSettingsFile)
'        PzGPointerAnimate = fGetINISetting(location, "pointerAnimate", PzGSettingsFile)
        PzGSamplingInterval = fGetINISetting(location, "samplingInterval", PzGSettingsFile)
        PzGStormTestInterval = fGetINISetting(location, "stormTestInterval", PzGSettingsFile)
        PzGErrorInterval = fGetINISetting(location, "errorInterval", PzGSettingsFile)
        
        PzGAirportsURL = fGetINISetting(location, "airportsURL", PzGSettingsFile)
        
        PzGTemperatureScale = fGetINISetting(location, "temperatureScale", PzGSettingsFile)
        PzGPressureScale = fGetINISetting(location, "pressureScale", PzGSettingsFile)
        PzGWindSpeedScale = fGetINISetting(location, "windSpeedScale", PzGSettingsFile)
        PzGMetricImperial = fGetINISetting(location, "metricImperial", PzGSettingsFile)
        PzGIcao = fGetINISetting(location, "icao", PzGSettingsFile)

        ' configuration
        PzGEnableTooltips = fGetINISetting(location, "enableTooltips", PzGSettingsFile)
        PzGEnablePrefsTooltips = fGetINISetting(location, "enablePrefsTooltips", PzGSettingsFile)
        PzGEnableBalloonTooltips = fGetINISetting(location, "enableBalloonTooltips", PzGSettingsFile)
        PzGShowTaskbar = fGetINISetting(location, "showTaskbar", PzGSettingsFile)
        PzGDpiAwareness = fGetINISetting(location, "dpiAwareness", PzGSettingsFile)
        
        
        
        PzGClipBSize = fGetINISetting("Software\PzClipB", "clipBSize", PzGSettingsFile)
        PzGSelectorSize = fGetINISetting("Software\PzSelector", "selectorSize", PzGSettingsFile)
        
        PzGScrollWheelDirection = fGetINISetting(location, "scrollWheelDirection", PzGSettingsFile)
        
        ' position
        PzGAspectHidden = fGetINISetting(location, "aspectHidden", PzGSettingsFile)
        PzGGaugeType = fGetINISetting(location, "gaugeType", PzGSettingsFile)
        
        PzGWidgetPosition = fGetINISetting(location, "widgetPosition", PzGSettingsFile)
        
        PzGTemperatureGaugeSize = fGetINISetting(location, "temperatureGaugeSize", PzGSettingsFile)
        PzGTemperatureLandscape = fGetINISetting(location, "temperatureLandscape", PzGSettingsFile)
        PzGTemperaturePortrait = fGetINISetting(location, "temperaturePortrait", PzGSettingsFile)
        PzGTemperatureLandscapeHoffset = fGetINISetting(location, "temperatureLandscapeHoffset", PzGSettingsFile)
        PzGTemperatureLandscapeVoffset = fGetINISetting(location, "temperatureLandscapeYoffset", PzGSettingsFile)
        PzGTemperaturePortraitHoffset = fGetINISetting(location, "temperaturePortraitHoffset", PzGSettingsFile)
        PzGTemperaturePortraitVoffset = fGetINISetting(location, "temperaturePortraitVoffset", PzGSettingsFile)
        PzGTemperatureVLocationPerc = fGetINISetting(location, "temperatureVLocationPerc", PzGSettingsFile)
        PzGTemperatureHLocationPerc = fGetINISetting(location, "temperatureHLocationPerc", PzGSettingsFile)
        PzGTemperatureFormHighDpiXPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormHighDpiXPos", PzGSettingsFile)
        PzGTemperatureFormHighDpiYPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormHighDpiYPos", PzGSettingsFile)
        PzGTemperatureFormLowDpiXPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormLowDpiXPos", PzGSettingsFile)
        PzGTemperatureFormLowDpiYPos = fGetINISetting("Software\PzTemperatureGauge", "temperatureFormLowDpiYPos", PzGSettingsFile)
        PzGPreventDraggingTemperature = fGetINISetting(location, "preventDraggingTemperature", PzGSettingsFile)
        
        PzGAnemometerLandscape = fGetINISetting("Software\PzAnemometerGauge", "anemometerLandscape", PzGSettingsFile)
        PzGAnemometerPortrait = fGetINISetting("Software\PzAnemometerGauge", "anemometerPortrait", PzGSettingsFile)
        PzGAnemometerGaugeSize = fGetINISetting("Software\PzAnemometerGauge", "anemometerGaugeSize", PzGSettingsFile)
        PzGAnemometerLandscapeHoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerLandscapeHoffset", PzGSettingsFile)
        PzGAnemometerLandscapeVoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerLandscapeVoffset", PzGSettingsFile)
        PzGAnemometerPortraitHoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerPortraitHoffset", PzGSettingsFile)
        PzGAnemometerPortraitVoffset = fGetINISetting("Software\PzAnemometerGauge", "anemometerPortraitVoffset", PzGSettingsFile)
        PzGAnemometerVLocationPerc = fGetINISetting("Software\PzAnemometerGauge", "anemometerVLocationPerc", PzGSettingsFile)
        PzGAnemometerHLocationPerc = fGetINISetting("Software\PzAnemometerGauge", "anemometerHLocationPerc", PzGSettingsFile)
        PzGAnemometerFormHighDpiXPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormHighDpiXPos", PzGSettingsFile)
        PzGAnemometerFormHighDpiYPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormHighDpiYPos", PzGSettingsFile)
        PzGAnemometerFormLowDpiXPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormLowDpiXPos", PzGSettingsFile)
        PzGAnemometerFormLowDpiYPos = fGetINISetting("Software\PzAnemometerGauge", "anemometerFormLowDpiYPos", PzGSettingsFile)
        PzGPreventDraggingAnemometer = fGetINISetting("Software\PzAnemometerGauge", "preventDraggingAnemometer", PzGSettingsFile)
        
        PzGHumidityLandscape = fGetINISetting("Software\PzHumidityGauge", "humidityLandscape", PzGSettingsFile)
        PzGHumidityPortrait = fGetINISetting("Software\PzHumidityGauge", "humidityPortrait", PzGSettingsFile)
        PzGHumidityGaugeSize = fGetINISetting("Software\PzHumidityGauge", "humidityGaugeSize", PzGSettingsFile)
        PzGHumidityLandscapeHoffset = fGetINISetting("Software\PzHumidityGauge", "humidityLandscapeHoffset", PzGSettingsFile)
        PzGHumidityLandscapeVoffset = fGetINISetting("Software\PzHumidityGauge", "humidityLandscapeVoffset", PzGSettingsFile)
        PzGHumidityPortraitHoffset = fGetINISetting("Software\PzHumidityGauge", "humidityPortraitHoffset", PzGSettingsFile)
        PzGHumidityPortraitVoffset = fGetINISetting("Software\PzHumidityGauge", "humidityPortraitVoffset", PzGSettingsFile)
        PzGHumidityVLocationPerc = fGetINISetting("Software\PzHumidityGauge", "humidityVLocationPerc", PzGSettingsFile)
        PzGHumidityHLocationPerc = fGetINISetting("Software\PzHumidityGauge", "humidityHLocationPerc", PzGSettingsFile)
        PzGHumidityFormHighDpiXPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormHighDpiXPos", PzGSettingsFile)
        PzGHumidityFormHighDpiYPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormHighDpiYPos", PzGSettingsFile)
        PzGHumidityFormLowDpiXPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormLowDpiXPos", PzGSettingsFile)
        PzGHumidityFormLowDpiYPos = fGetINISetting("Software\PzHumidityGauge", "humidityFormLowDpiYPos", PzGSettingsFile)
        PzGPreventDraggingHumidity = fGetINISetting("Software\PzHumidityGauge", "preventDraggingHumidity", PzGSettingsFile)
         
        PzGBarometerLandscape = fGetINISetting("Software\PzBarometerGauge", "barometerLandscape", PzGSettingsFile)
        PzGBarometerPortrait = fGetINISetting("Software\PzBarometerGauge", "barometerPortrait", PzGSettingsFile)
        PzGBarometerGaugeSize = fGetINISetting("Software\PzBarometerGauge", "barometerGaugeSize", PzGSettingsFile)
        PzGBarometerLandscapeHoffset = fGetINISetting("Software\PzBarometerGauge", "barometerLandscapeHoffset", PzGSettingsFile)
        PzGBarometerLandscapeVoffset = fGetINISetting("Software\PzBarometerGauge", "barometerLandscapeVoffset", PzGSettingsFile)
        PzGBarometerPortraitHoffset = fGetINISetting("Software\PzBarometerGauge", "barometerPortraitHoffset", PzGSettingsFile)
        PzGBarometerPortraitVoffset = fGetINISetting("Software\PzBarometerGauge", "barometerPortraitVoffset", PzGSettingsFile)
        PzGBarometerVLocationPerc = fGetINISetting("Software\PzBarometerGauge", "barometerVLocationPerc", PzGSettingsFile)
        PzGBarometerHLocationPerc = fGetINISetting("Software\PzBarometerGauge", "barometerHLocationPerc", PzGSettingsFile)
        PzGBarometerFormHighDpiXPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormHighDpiXPos", PzGSettingsFile)
        PzGBarometerFormHighDpiYPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormHighDpiYPos", PzGSettingsFile)
        PzGBarometerFormLowDpiXPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormLowDpiXPos", PzGSettingsFile)
        PzGBarometerFormLowDpiYPos = fGetINISetting("Software\PzBarometerGauge", "barometerFormLowDpiYPos", PzGSettingsFile)
        PzGPreventDraggingBarometer = fGetINISetting("Software\PzBarometerGauge", "preventDraggingBarometer", PzGSettingsFile)
             
        ' font
        PzGTempFormFont = fGetINISetting(location, "tempFormFont", PzGSettingsFile)
        PzGPrefsFont = fGetINISetting(location, "prefsFont", PzGSettingsFile)
        
        PzGPrefsFontSizeHighDPI = fGetINISetting(location, "prefsFontSizeHighDPI", PzGSettingsFile)
        PzGPrefsFontSizeLowDPI = fGetINISetting(location, "prefsFontSizeLowDPI", PzGSettingsFile)
        PzGPrefsFontItalics = fGetINISetting(location, "prefsFontItalics", PzGSettingsFile)
        PzGPrefsFontColour = fGetINISetting(location, "prefsFontColour", PzGSettingsFile)
        
        ' sound
        PzGEnableSounds = fGetINISetting(location, "enableSounds", PzGSettingsFile)
        
        ' development
        PzGDebug = fGetINISetting(location, "debug", PzGSettingsFile)
        PzGDblClickCommand = fGetINISetting(location, "dblClickCommand", PzGSettingsFile)
        PzGOpenFile = fGetINISetting(location, "openFile", PzGSettingsFile)
        PzGDefaultEditor = fGetINISetting(location, "defaultEditor", PzGSettingsFile)
                
        ' other
        PzGClipBFormHighDpiXPos = fGetINISetting("Software\PzClipB", "clipBFormHighDpiXPos", PzGSettingsFile)
        PzGClipBFormHighDpiYPos = fGetINISetting("Software\PzClipB", "clipBFormHighDpiYPos", PzGSettingsFile)
        PzGClipBFormLowDpiXPos = fGetINISetting("Software\PzClipB", "clipBFormLowDpiXPos", PzGSettingsFile)
        PzGClipBFormLowDpiYPos = fGetINISetting("Software\PzClipB", "clipBFormLowDpiYPos", PzGSettingsFile)
         
        ' other
        PzGSelectorFormHighDpiXPos = fGetINISetting("Software\PzSelector", "selectorFormHighDpiXPos", PzGSettingsFile)
        PzGSelectorFormHighDpiYPos = fGetINISetting("Software\PzSelector", "selectorFormHighDpiYPos", PzGSettingsFile)
        PzGSelectorFormLowDpiXPos = fGetINISetting("Software\PzSelector", "selectorFormLowDpiXPos", PzGSettingsFile)
        PzGSelectorFormLowDpiYPos = fGetINISetting("Software\PzSelector", "selectorFormLowDpiYPos", PzGSettingsFile)
       
        PzGLastSelectedTab = fGetINISetting(location, "lastSelectedTab", PzGSettingsFile)
        PzGSkinTheme = fGetINISetting(location, "skinTheme", PzGSettingsFile)
        
        ' window
        PzGWindowLevel = fGetINISetting(location, "windowLevel", PzGSettingsFile)
        
        PzGOpacity = fGetINISetting(location, "opacity", PzGSettingsFile)
        
        PzGLastUpdated = fGetINISetting(location, "lastUpdated", PzGSettingsFile)
        PzGMetarPref = fGetINISetting(location, "metarPref", PzGSettingsFile)
        
        PzGOldPressureStorage = fGetINISetting(location, "oldPressureStorage", PzGSettingsFile)
        PzGPressureStorageDate = fGetINISetting(location, "pressureStorageDate", PzGSettingsFile)
        PzGCurrentPressureValue = fGetINISetting(location, "currentPressureValue", PzGSettingsFile)
    
        ' we do not want the widget to hide at startup
        'PzGWidgetHidden = fGetINISetting(location, "widgetHidden", PzGSettingsFile)
        PzGWidgetHidden = "0"
        
        PzGHidingTime = fGetINISetting(location, "hidingTime", PzGSettingsFile)
        PzGIgnoreMouse = fGetINISetting(location, "ignoreMouse", PzGSettingsFile)
         
        PzGFirstTimeRun = fGetINISetting(location, "firstTimeRun", PzGSettingsFile)
        
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
    If PzGGaugeFunctions = vbNullString Then PzGGaugeFunctions = "1" ' always turn on
'        If PzGAnimationInterval = vbNullString Then PzGAnimationInterval = "130"
    If PzGStartup = vbNullString Then PzGStartup = "1"
'    If PzGPointerAnimate = vbNullString Then PzGPointerAnimate = "0"
    If PzGSamplingInterval = vbNullString Then PzGSamplingInterval = "60"
    If PzGStormTestInterval = vbNullString Then PzGStormTestInterval = "3600"
    If PzGErrorInterval = vbNullString Then PzGErrorInterval = "3"
    
    If PzGAirportsURL = vbNullString Then PzGAirportsURL = "https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports.dat"
    
    If PzGTemperatureScale = vbNullString Then PzGTemperatureScale = "0"
    If PzGPressureScale = vbNullString Then PzGPressureScale = "0" ' "Millibars"
    If PzGWindSpeedScale = vbNullString Then PzGWindSpeedScale = "0"
    If PzGMetricImperial = vbNullString Then PzGMetricImperial = "0"
    
    If PzGIcao = vbNullString Then PzGIcao = "EGSH"

    ' Configuration
    If PzGEnableTooltips = vbNullString Then PzGEnableTooltips = "0"
    If PzGEnablePrefsTooltips = vbNullString Then PzGEnablePrefsTooltips = "1"
    If PzGEnableBalloonTooltips = vbNullString Then PzGEnableBalloonTooltips = "1"
    If PzGShowTaskbar = vbNullString Then PzGShowTaskbar = "0"
    If PzGDpiAwareness = vbNullString Then PzGDpiAwareness = "0"
    
    If PzGClipBSize = vbNullString Then PzGClipBSize = "50"
    If PzGSelectorSize = vbNullString Then PzGSelectorSize = "100"
    
    If PzGScrollWheelDirection = vbNullString Then PzGScrollWheelDirection = "1"
           
    ' fonts
    If PzGPrefsFont = vbNullString Then PzGPrefsFont = "times new roman"
    If PzGTempFormFont = vbNullString Then PzGTempFormFont = PzGPrefsFont
    
    If PzGPrefsFontSizeHighDPI = vbNullString Then PzGPrefsFontSizeHighDPI = "8"
    If PzGPrefsFontSizeLowDPI = vbNullString Then PzGPrefsFontSizeLowDPI = "8"
    If PzGPrefsFontItalics = vbNullString Then PzGPrefsFontItalics = "false"
    If PzGPrefsFontColour = vbNullString Then PzGPrefsFontColour = "0"

    ' sounds
    If PzGEnableSounds = vbNullString Then PzGEnableSounds = "1"

    ' position
    If PzGAspectHidden = vbNullString Then PzGAspectHidden = "0"
    If PzGGaugeType = vbNullString Then PzGGaugeType = "0"
    
    If PzGWidgetPosition = vbNullString Then PzGWidgetPosition = "0"
    
    If PzGTemperatureGaugeSize = vbNullString Then PzGTemperatureGaugeSize = "50"
    If PzGTemperatureLandscape = vbNullString Then PzGTemperatureLandscape = "0"
    If PzGTemperaturePortrait = vbNullString Then PzGTemperaturePortrait = "0"
    If PzGTemperatureLandscapeHoffset = vbNullString Then PzGTemperatureLandscapeHoffset = vbNullString
    If PzGTemperatureLandscapeVoffset = vbNullString Then PzGTemperatureLandscapeVoffset = vbNullString
    If PzGTemperaturePortraitHoffset = vbNullString Then PzGTemperaturePortraitHoffset = vbNullString
    If PzGTemperaturePortraitVoffset = vbNullString Then PzGTemperaturePortraitVoffset = vbNullString
    If PzGTemperatureVLocationPerc = vbNullString Then PzGTemperatureVLocationPerc = vbNullString
    If PzGTemperatureHLocationPerc = vbNullString Then PzGTemperatureHLocationPerc = vbNullString
    If PzGPreventDraggingTemperature = vbNullString Then PzGPreventDraggingTemperature = "0"
    
    If PzGAnemometerGaugeSize = vbNullString Then PzGAnemometerGaugeSize = "50"
    If PzGAnemometerLandscape = vbNullString Then PzGAnemometerLandscape = "0"
    If PzGAnemometerPortrait = vbNullString Then PzGAnemometerPortrait = "0"
    If PzGAnemometerLandscapeHoffset = vbNullString Then PzGAnemometerLandscapeHoffset = vbNullString
    If PzGAnemometerLandscapeVoffset = vbNullString Then PzGAnemometerLandscapeVoffset = vbNullString
    If PzGAnemometerPortraitHoffset = vbNullString Then PzGAnemometerPortraitHoffset = vbNullString
    If PzGAnemometerPortraitVoffset = vbNullString Then PzGAnemometerPortraitVoffset = vbNullString
    If PzGAnemometerVLocationPerc = vbNullString Then PzGAnemometerVLocationPerc = vbNullString
    If PzGAnemometerHLocationPerc = vbNullString Then PzGAnemometerHLocationPerc = vbNullString
    If PzGPreventDraggingAnemometer = vbNullString Then PzGPreventDraggingAnemometer = "0"
    
    If PzGHumidityGaugeSize = vbNullString Then PzGHumidityGaugeSize = "50"
    If PzGHumidityLandscape = vbNullString Then PzGHumidityLandscape = "0"
    If PzGHumidityPortrait = vbNullString Then PzGHumidityPortrait = "0"
    If PzGHumidityLandscapeHoffset = vbNullString Then PzGHumidityLandscapeHoffset = vbNullString
    If PzGHumidityLandscapeVoffset = vbNullString Then PzGHumidityLandscapeVoffset = vbNullString
    If PzGHumidityPortraitHoffset = vbNullString Then PzGHumidityPortraitHoffset = vbNullString
    If PzGHumidityPortraitVoffset = vbNullString Then PzGHumidityPortraitVoffset = vbNullString
    If PzGHumidityVLocationPerc = vbNullString Then PzGHumidityVLocationPerc = vbNullString
    If PzGHumidityHLocationPerc = vbNullString Then PzGHumidityHLocationPerc = vbNullString
    If PzGPreventDraggingHumidity = vbNullString Then PzGPreventDraggingHumidity = "0"
    
    If PzGBarometerGaugeSize = vbNullString Then PzGBarometerGaugeSize = "50"
    If PzGBarometerLandscape = vbNullString Then PzGBarometerLandscape = "0"
    If PzGBarometerPortrait = vbNullString Then PzGBarometerPortrait = "0"
    If PzGBarometerLandscapeHoffset = vbNullString Then PzGBarometerLandscapeHoffset = vbNullString
    If PzGBarometerLandscapeVoffset = vbNullString Then PzGBarometerLandscapeVoffset = vbNullString
    If PzGBarometerPortraitHoffset = vbNullString Then PzGBarometerPortraitHoffset = vbNullString
    If PzGBarometerPortraitVoffset = vbNullString Then PzGBarometerPortraitVoffset = vbNullString
    If PzGBarometerVLocationPerc = vbNullString Then PzGBarometerVLocationPerc = vbNullString
    If PzGBarometerHLocationPerc = vbNullString Then PzGBarometerHLocationPerc = vbNullString
    If PzGPreventDraggingBarometer = vbNullString Then PzGPreventDraggingBarometer = "0"
            
    ' development
    If PzGDebug = vbNullString Then PzGDebug = "0"
    If PzGDblClickCommand = vbNullString Then PzGDblClickCommand = "%systemroot%\system32\timedate.cpl"
    If PzGOpenFile = vbNullString Then PzGOpenFile = vbNullString
    If PzGDefaultEditor = vbNullString Then PzGDefaultEditor = vbNullString
    
    ' window
    If PzGWindowLevel = vbNullString Then PzGWindowLevel = "1" 'WindowLevel", PzGSettingsFile)
    If PzGOpacity = vbNullString Then PzGOpacity = "100"
    If PzGWidgetHidden = vbNullString Then PzGWidgetHidden = "0"
    If PzGHidingTime = vbNullString Then PzGHidingTime = "0"
    If PzGIgnoreMouse = vbNullString Then PzGIgnoreMouse = "0"
    
    ' other
    If PzGFirstTimeRun = vbNullString Then PzGFirstTimeRun = "true"
    If PzGLastSelectedTab = vbNullString Then PzGLastSelectedTab = "general"
    If PzGSkinTheme = vbNullString Then PzGSkinTheme = "dark"
    
    If PzGLastUpdated = vbNullString Then PzGLastUpdated = CStr(Now())
    If PzGMetarPref = vbNullString Then PzGMetarPref = "ICAO"
        
    If PzGOldPressureStorage = vbNullString Then PzGOldPressureStorage = "0"
    If PzGPressureStorageDate = vbNullString Then PzGPressureStorageDate = CStr(Now())
    If PzGCurrentPressureValue = vbNullString Then PzGCurrentPressureValue = "0"
    
 
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
    
    PzGTrinketsDir = fSpecialFolder(feUserAppData) & "\trinkets" ' just for this user alone
    PzGTrinketsFile = PzGTrinketsDir & "\" & widgetName1 & ".ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(PzGTrinketsDir) Then
        MkDir PzGTrinketsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(PzGTrinketsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open PzGTrinketsFile For Output As #iFileNo
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
    
    PzGSettingsDir = fSpecialFolder(feUserAppData) & "\PzTemperatureGauge" ' just for this user alone
    PzGSettingsFile = PzGSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(PzGSettingsDir) Then
        MkDir PzGSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(PzGSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open PzGSettingsFile For Output As #iFileNo
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
    
    oldPzGSettingsModificationTime = FileDateTime(PzGSettingsFile)

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

    If PzGHidingTime = "0" Then minutesToHide = 1
    If PzGHidingTime = "1" Then minutesToHide = 5
    If PzGHidingTime = "2" Then minutesToHide = 10
    If PzGHidingTime = "3" Then minutesToHide = 20
    If PzGHidingTime = "4" Then minutesToHide = 30
    If PzGHidingTime = "5" Then minutesToHide = 60

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
        PzGUnhide = "true"
        sPutINISetting "Software\PzTemperatureGauge", "unhide", PzGUnhide, PzGSettingsFile
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

        .Add Empty, "redlamptrue"
        .Add Empty, "redlampfalse"
        
        .Add Empty, "directionpointer"
        
        .Add Empty, "pointerShadow"
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

        .Add Empty, "barometermmhgface"
        .Add Empty, "barometerinhgface"
        .Add Empty, "barometerhpaface"
        .Add Empty, "barometermbface"
        
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
    
    If PzGShowTaskbar = "0" Then
        fTemperature.temperatureGaugeForm.ShowInTaskbar = False
        fAnemometer.anemometerGaugeForm.ShowInTaskbar = False
        fHumidity.humidityGaugeForm.ShowInTaskbar = False
        fBarometer.barometerGaugeForm.ShowInTaskbar = False
        fClipB.clipBForm.ShowInTaskbar = False
    Else
        fTemperature.temperatureGaugeForm.ShowInTaskbar = True
        fAnemometer.anemometerGaugeForm.ShowInTaskbar = True
        fHumidity.humidityGaugeForm.ShowInTaskbar = True
        fBarometer.barometerGaugeForm.ShowInTaskbar = True
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
    On Error GoTo setMenuItems_Error
    
    If PzGGaugeFunctions = "1" Then
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If PzGDefaultEditor <> vbNullString And PzGDebug = "1" Then
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & PzGDefaultEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
    
    On Error GoTo 0
    Exit Sub

setMenuItems_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setMenuItems, line " & Erl & "."

End Sub
    

