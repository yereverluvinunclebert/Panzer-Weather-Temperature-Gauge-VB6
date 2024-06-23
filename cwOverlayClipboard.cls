VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlayClipb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule AssignmentNotUsed, IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlayClipb
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfTemperature
'---------------------------------------------------------------------------------------

' Microsoft XML, v3.0
' c:/windows/SysWow64/msxml3.dll

Option Explicit

'these Public Properties are set in the Parent (here, cfTemperature)
Public cCollOverlayClipBImages As cCollection

Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1


'Private SDeg As Double


Private mLocked As Boolean ' property Locked
Private mHidden As Boolean ' property Hidden
Private mThisOpacity As String ' property ThisOpacity
Private mThisTime As Date ' property ThisTime


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Error
    
    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                        
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlayClipb"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlayClipb"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlayClipb"
End Property


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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayClipb"
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

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlayClipb"
End Property




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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayClipb"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlayClipb"
End Property

'---------------------------------------------------------------------------------------
' Property : thisTime
' Author    : Dean Beedell (yereverluvinunclebert)
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisTime() As Date
   On Error GoTo thisTimeGet_Error

    thisTime = mThisTime

   On Error GoTo 0
   Exit Property

thisTimeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisTime of Class Module cwoverlaytemp"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisTime
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisTime(ByVal newValue As Date)
    On Error GoTo thisTimeLet_Error

    If mThisTime <> newValue Then mThisTime = newValue Else Exit Property
        
    On Error GoTo 0
    Exit Property

thisTimeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisTime of Class Module cwoverlaytemp"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlayClipb"
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
    Dim hrs As Integer: hrs = 0
    Dim mins As Single: mins = 0
    Dim clipbHrsSDeg As Single: clipbHrsSDeg = 0
    Dim clipbMinsSDeg As Integer: clipbMinsSDeg = 0
     
    On Error GoTo W_Paint_Error
    
    'mThisTime = "12:15:00" ' debug

    clipbMinsSDeg = Minute(mThisTime) * 6
    hrs = Hour(mThisTime)
    If hrs >= 12 Then hrs = hrs - 12
    
    If Not mThisTime = "00:00:00" Then
        mins = Minute(mThisTime) / 60
        clipbHrsSDeg = (hrs * 30 + (mins * 30))
    End If
    
    For i = 0 To cCollOverlayClipBImages.Count - 1 ' this is the ExcludedPaths collection populated in cfTemperature by duplicating collTemperaturePSDNonUIElements.
      arr = cCollOverlayClipBImages.ItemByIndex(i)
      key = cCollOverlayClipBImages.KeyByIndex(i)
      
      If IsArray(arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case key
                 
                 Case "minhand"
                 
                         Set Pat.Matrix = Pat.Matrix.TranslateCoords(7, 69) 'shift to the "rotation-point" within the Pattern
                         CC.TranslateDrawings 105, 115  'shift to the rotation-point
                         CC.RotateDrawingsDeg clipbMinsSDeg

                 Case "hourhand"
                 ' 95 , 67
                         Set Pat.Matrix = Pat.Matrix.TranslateCoords(9, 50) 'shift to the "rotation-point" within the Pattern
                         CC.TranslateDrawings 105, 115  'shift to the rotation-point
                         CC.RotateDrawingsDeg clipbHrsSDeg

                 Case "clock"

                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))

                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-arr(0), -arr(1))
                   
             End Select
             CC.Paint arr(2), Pat 'render the Pattern
             CC.Restore
      End If
     Next
             
    W.Refresh


   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlayClipb"
End Sub




