Attribute VB_Name = "mOtherFunction"
Option Explicit

' Функция расчета времени исходя из полученных значений в миллисекундах функции GetTickCount
Public Function CalculateTime(ByVal lngStartTime As Long, _
                              ByVal lngEndTime As Long, _
                              Optional ByVal mboolmSec As Boolean = False) As String

    Dim lngWorkTimeTemp         As Single
    Dim lngWorkTimeSecound      As Long
    Dim lngWorkTimeMinutes      As Long
    Dim lngWorkTimeHours        As Long
    Dim lngWorkTimeMilliSecound As Long
    Dim strWorkTimeSecound      As String
    Dim strWorkTimeMinutes      As String
    Dim strWorkTimeHours        As String
    Dim strWorkTimeMilliSecound As String

    If lngEndTime > lngStartTime Then
        lngWorkTimeTemp = (lngEndTime - lngStartTime) / 1000

        'время в секундах
        'Если надо то в миллисекундах
        If mboolmSec Then
            lngWorkTimeMilliSecound = (lngWorkTimeTemp - Fix(lngWorkTimeTemp)) * 100
        Else
            lngWorkTimeTemp = Fix(lngWorkTimeTemp)
        End If

        Select Case lngWorkTimeTemp

            Case 0 To 59
                lngWorkTimeSecound = lngWorkTimeTemp

            Case 60
                lngWorkTimeHours = 0
                lngWorkTimeMinutes = 1
                lngWorkTimeSecound = 0

            Case 61 To 3599
                lngWorkTimeHours = 0
                lngWorkTimeTemp = lngWorkTimeTemp - 60
                lngWorkTimeMinutes = 1

                Do While lngWorkTimeTemp > 60
                    lngWorkTimeTemp = lngWorkTimeTemp - 60
                    lngWorkTimeMinutes = lngWorkTimeMinutes + 1
                Loop
                lngWorkTimeSecound = lngWorkTimeTemp

            Case 3600
                lngWorkTimeHours = 1
                lngWorkTimeMinutes = 0
                lngWorkTimeSecound = 0

            Case Is > 3600
                lngWorkTimeTemp = lngWorkTimeTemp - 3600
                lngWorkTimeHours = 1

                Do While lngWorkTimeTemp > 3599
                    lngWorkTimeTemp = lngWorkTimeTemp - 3600
                    lngWorkTimeHours = lngWorkTimeHours + 1
                Loop
                lngWorkTimeSecound = lngWorkTimeTemp

                Do While lngWorkTimeTemp > 60
                    lngWorkTimeTemp = lngWorkTimeTemp - 60
                    lngWorkTimeMinutes = lngWorkTimeMinutes + 1
                Loop
                lngWorkTimeSecound = lngWorkTimeTemp

                Do While lngWorkTimeMinutes > 59
                    lngWorkTimeMinutes = lngWorkTimeMinutes - 60
                    lngWorkTimeHours = lngWorkTimeHours + 1
                Loop

            Case Else
                lngWorkTimeHours = 0
                lngWorkTimeMinutes = 0
                lngWorkTimeSecound = 0
        End Select
    End If

    ' Добавляем лидирующие нули при необходимости
    ' Часы
    If Len(CStr(lngWorkTimeHours)) = 1 Then
        strWorkTimeHours = "0" & CStr(lngWorkTimeHours)
    ElseIf Len(CStr(lngWorkTimeHours)) = 2 Then
        strWorkTimeHours = CStr(lngWorkTimeHours)
    Else
        strWorkTimeHours = "00"
    End If

    ' Минуты
    If Len(CStr(lngWorkTimeMinutes)) = 1 Then
        strWorkTimeMinutes = "0" & CStr(lngWorkTimeMinutes)
    ElseIf Len(CStr(lngWorkTimeMinutes)) = 2 Then
        strWorkTimeMinutes = CStr(lngWorkTimeMinutes)
    Else
        strWorkTimeMinutes = "00"
    End If

    ' Секунды
    If Len(CStr(lngWorkTimeSecound)) = 1 Then
        strWorkTimeSecound = "0" & CStr(lngWorkTimeSecound)
    ElseIf Len(CStr(lngWorkTimeSecound)) = 2 Then
        strWorkTimeSecound = CStr(lngWorkTimeSecound)
    Else
        strWorkTimeSecound = "00"
    End If

    ' МилиСекунды
    If mboolmSec Then
        If Len(CStr(lngWorkTimeMilliSecound)) = 1 Then
            strWorkTimeMilliSecound = "0" & CStr(lngWorkTimeMilliSecound)
        ElseIf Len(CStr(lngWorkTimeSecound)) = 2 Then
            strWorkTimeMilliSecound = CStr(lngWorkTimeMilliSecound)
        Else
            strWorkTimeMilliSecound = "00"
        End If

        ' Итоговое время
        CalculateTime = strWorkTimeHours & ":" & strWorkTimeMinutes & ":" & strWorkTimeSecound & "." & strWorkTimeMilliSecound & " (hh:mm:ss.ms)"
    Else
        ' Итоговое время
        CalculateTime = strWorkTimeHours & ":" & strWorkTimeMinutes & ":" & strWorkTimeSecound & " (hh:mm:ss)"
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  CollectCmdString
'!  Переменные  :
'!  Описание    :  Создание коммандной строки запуска программы DPInst
'! -----------------------------------------------------------
Public Function CollectCmdString() As String

    Dim strCmdStringDPInstTemp As String

    If mboolDpInstLegacyMode Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/LM "
    End If

    If mboolDpInstPromptIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/P "
    End If

    If mboolDpInstForceIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/F "
    End If

    If mboolDpInstSuppressAddRemovePrograms Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SA "
    End If

    If mboolDpInstSuppressWizard Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SW "
    End If

    If mboolDpInstQuietInstall Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/Q "
    End If

    If mboolDpInstScanHardware Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SH "
    End If

    ' Результирующая строка
    CollectCmdString = strCmdStringDPInstTemp
End Function

Public Function HPadding(ByVal Form As Form)

    Dim SaveMode As Integer

    With Form
        SaveMode = .ScaleMode
        .ScaleMode = vbTwips
        HPadding = .Width - .ScaleWidth
        .ScaleMode = SaveMode
    End With
End Function

Public Function VPadding(ByVal Form As Form)

    Dim SaveMode As Integer

    With Form
        SaveMode = .ScaleMode
        .ScaleMode = vbTwips
        VPadding = .Height - .ScaleHeight
        .ScaleMode = SaveMode
    End With
End Function
