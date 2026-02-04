VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XHRSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : XHRSink
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : used to capture the OnReadyStateChange event from MSXML2.XMLHTTP60 XMLHttpRequest2
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : OnReadyStateChange
' Author    : beededea
' Date      : 05/09/2025
' Purpose   : used to capture the OnReadyStateChange event from MSXML2.XMLHTTP60 XMLHttpRequest2
'---------------------------------------------------------------------------------------
'
Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0

    On Error GoTo OnReadyStateChange_Error

    Debug.Print XMLHttpRequest2.readyState
    If XMLHttpRequest2.readyState = 4 Then
        Debug.Print "sink code handling result"
        Debug.Print XMLHttpRequest2.responseText
        
        WeatherMeteo.processMetarData
    End If
    
    ' <The data necessary to complete this operation is not yet available.

    On Error GoTo 0
    Exit Sub

OnReadyStateChange_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OnReadyStateChange of Class Module XHRSink"
End Sub
