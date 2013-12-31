Attribute VB_Name = "mDebug"
Option Explicit

'==========================================================================
'------------------ Параметры отладочного режима --------------------------'
'==========================================================================
Public mboolDebugEnable  As Boolean
Public strDebugLogPath   As String
Public mboolCleanHistory As Boolean     'Очистка истории отладочного режима
Public lngDetailMode     As Long

'! -----------------------------------------------------------
'!  Функция     :  DebugMode
'!  Переменные  :  Msg - Передаваемое сообщение
'!  Описание    :  Функция отладочных сообщений
'! -----------------------------------------------------------
Public Sub DebugMode(Msg As String, Optional lngDetailModeTemp As Long = 1)

    Dim tsLogFile As TextStream

    ' создается ли новый файл или открывается для дозаписи
    If mboolDebugEnable Then
        If Not mboolLogNotOnCDRoom Then
            If lngDetailModeTemp <= lngDetailMode Then
                Msg = CStr(Msg)

                If LenB(Msg) > 0 Then
                    If FSO.FileExists(strDebugLogPath) Then
                        Set tsLogFile = FSO.OpenTextFile(strDebugLogPath, ForAppending, False, TristateTrue)
                    Else
                        Set tsLogFile = FSO.OpenTextFile(strDebugLogPath, ForWriting, True, TristateTrue)
                    End If

                    tsLogFile.WriteLine CStr(Now()) & vbTab & Msg
                    tsLogFile.Close
                End If
            End If
        End If
    End If
End Sub

'! -----------------------------------------------------------
'!  Функция     :  LogNotOnCDRoom
'!  Переменные  :
'!  Описание    :  Проверка на хранение лог-файла на CD
'! -----------------------------------------------------------
Public Function LogNotOnCDRoom() As Boolean

    Dim strDriveName As String
    Dim xDrv         As Drive

    LogNotOnCDRoom = False
    strDriveName = Mid$(strDebugLogPath, 1, 2)

    ' Проверяем на запуск из сети
    If StrComp(strDriveName, "\\", vbTextCompare) <> 0 Then
        'получаем тип диска
        Set xDrv = FSO.GetDrive(strDriveName)

        If xDrv.DriveType = CDRom Then
            LogNotOnCDRoom = True
        End If
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  MakeCleanHistory
'!  Переменные  :
'!  Описание    :  Удаление истории отладочного режима
'! -----------------------------------------------------------
Public Sub MakeCleanHistory()

    Dim FileDel As File

    If mboolCleanHistory Then
        If FSO.FileExists(strDebugLogPath) Then
            If Not mboolLogNotOnCDRoom Then
                Set FileDel = FSO.GetFile(strDebugLogPath)
                FileDel.Delete
            End If
        End If
    End If
End Sub

' Запись в DebugLog сожержимого файла
Public Sub PrintFileInDebugLog(strFilePath As String)

    Dim objTxtFile    As TextStream
    Dim strTxtFileAll As String

    If PathFileExists(strFilePath) = 1 Then
        If Not IsPathAFolder(strFilePath) Then
            Set objTxtFile = FSO.OpenTextFile(strFilePath, ForReading, False, TristateUseDefault)
            strTxtFileAll = objTxtFile.ReadAll
            objTxtFile.Close
            DebugMode "***Content of file: " & strFilePath & vbNewLine & "*********************BEGIN FILE**************************" & vbNewLine & strTxtFileAll & vbNewLine & "**********************END FILE***************************"
        End If
    End If
End Sub
