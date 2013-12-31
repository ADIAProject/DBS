Attribute VB_Name = "mDEP"
Option Explicit

Public Declare Function SetProcessDEPPolicy Lib "kernel32.dll" (ByVal dwFlags As Long) As Boolean

Public Sub SetDEPDisable()

    Dim mboolCallback As Boolean

    DebugMode "Disable DEP: Try to Disable DEP for this Process"

    If APIFunctionPresent("SetProcessDEPPolicy", "kernel32.dll") Then
        mboolCallback = SetProcessDEPPolicy(0)
        DebugMode "Disable DEP: Result: " & mboolCallback & " - Err ¹" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
    Else
        DebugMode "Disable DEP: ApiFunction not Supported"
    End If
End Sub
