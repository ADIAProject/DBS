Attribute VB_Name = "mHideErrorWindow"
Option Explicit

Private Const SEM_NOGPFAULTERRORBOX As Long = &H2

Public mboolUnloadClean             As Boolean

Private Declare Function SetErrorMode Lib "kernel32.dll" (ByVal wMode As Long) As Long

Public Sub UnloadApp()

    If Forms.Count = 0 Then
        If Not mboolIsDesignMode Then

            ' ----------------------------------------------
            ' START: Added to allow testing
            If Not (mboolUnloadClean) Then
                Exit Sub
            End If

            MsgBox "UnloadApp Called"
            ' END: Added to allow testing
            ' ----------------------------------------------
            SetErrorMode SEM_NOGPFAULTERRORBOX
        End If
    End If
End Sub
