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
Option Explicit

Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0

    Debug.Print XMLHttpRequest2.readyState
    If XMLHttpRequest2.readyState = 4 Then
        Debug.Print "sink code handling result"
        Debug.Print XMLHttpRequest2.responseText
        
        WeatherMeteo.processMetarData
    End If
End Sub
