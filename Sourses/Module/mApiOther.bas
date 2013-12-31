Attribute VB_Name = "mApiOther"
Option Explicit

Public Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Type TRACKMOUSEEVENT_STRUCT
    cbSize                                  As Long
    dwFlags                                 As TRACKMOUSEEVENT_FLAGS
    hwndTrack                               As Long
    dwHoverTime                             As Long
End Type

Public Declare Function TrackMouseEvent Lib "user32.dll" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Public Declare Function TrackMouseEventComCtl _
               Lib "comctl32.dll" _
               Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Declare Function IsUserAnAdmin Lib "shell32.dll" () As Long

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

