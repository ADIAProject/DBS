Attribute VB_Name = "mOsVer"
Option Explicit

' Программные переменные
Public strOSArchitecture   As String
Public strOsCurrentVersion As String

'Получение расширенной информации о версии Windows
Public Type OSVERSIONINFO
    dwOSVersionInfoSize                            As Long
    dwMajorVersion                                 As Long
    dwMinorVersion                                 As Long
    dwBuildNumber                                  As Long
    dwPlatformId                                   As Long
    szCSDVersion                                   As String * 128
End Type

' Проверка процесса на 64 bit
Public Type SYSTEM_INFO
    wProcessorArchitecture                         As Integer
    wReserved                                      As Integer
    dwPageSize                                     As Long
    lpMinimumApplicationAddress                    As Long
    lpMaximumApplicationAddress                    As Long
    dwActiveProcessorMask                          As Long
    dwNumberOfProcessors                           As Long
    dwProcessorType                                As Long
    dwAllocationGranularity                        As Long
    wProcessorLevel                                As Integer
    wProcessorRevision                             As Integer
End Type

Public Const PROCESSOR_ARCHITECTURE_AMD64 As Long = &H9
Public Const PROCESSOR_ARCHITECTURE_IA64  As Long = &H6
Public Const PROCESSOR_ARCHITECTURE_INTEL As Long = 0
Public Const PROCESSOR_ARCHITECTURE_ALPHA = 2
Public Const PROCESSOR_ARCHITECTURE_ALPHA64 As Long = 7

'Windows NT - constants for unicode support
Public Const VER_PLATFORM_WIN32_NT          As Long = 2
Public Declare Function GetVersionEx _
               Lib "kernel32.dll" _
               Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)

Public Function IsWinXPAndGreater() As Boolean

    IsWinXPAndGreater = strOsCurrentVersion > "5.0"
End Function

'! -----------------------------------------------------------
'!  Функция     :  IsWow64
'!  Переменные  :
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Проверяет является ли запущенный процесс 64-битным
'! -----------------------------------------------------------
Public Function IsWow64() As Boolean

    Dim SI As SYSTEM_INFO

    strOSArchitecture = "x86"

    If APIFunctionPresent("GetNativeSystemInfo", "kernel32.dll") Then
        GetNativeSystemInfo SI

        Select Case SI.wProcessorArchitecture

            Case PROCESSOR_ARCHITECTURE_IA64
                IsWow64 = True
                strOSArchitecture = "ia64"

            Case PROCESSOR_ARCHITECTURE_AMD64
                IsWow64 = True
                strOSArchitecture = "amd64"

            Case Else
                IsWow64 = False
        End Select
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  OSInfo
'!  Переменные  :  Nfo As Integer
'!  Возвр. знач.:  As String
'!  Описание    :  Получение расширенной информации о версии Windows
'! -----------------------------------------------------------
Public Function OSInfo(ByVal Nfo As Long) As String

    Dim OSVerInfo As OSVERSIONINFO
    Dim OSN       As String

    On Error GoTo HandErr

    OSVerInfo.dwOSVersionInfoSize = Len(OSVerInfo)

    If GetVersionEx(OSVerInfo) <> 0 Then

        With OSVerInfo
            'Имя операционной системы
            OSN = "UnSupported\Unknown"

            If .dwMajorVersion = 5 Then
                If .dwMinorVersion = 0 Then
                    OSN = "2000"
                ElseIf .dwMinorVersion = 1 Then
                    OSN = "XP"
                ElseIf .dwMinorVersion = 2 Then
                    OSN = "Server 2003"
                End If
            End If

            If .dwMajorVersion = 6 Then
                If .dwMinorVersion = 0 Then
                    OSN = "Vista\Server 2008"
                ElseIf .dwMinorVersion = 1 Then
                    OSN = "7\Server 2008 R2"
                ElseIf .dwMinorVersion = 2 Then
                    OSN = "8"
                Else
                    OSN = "9 ?"
                End If
            End If

            If .dwMajorVersion > 6 Then
                OSN = "9 ?"
            End If

            Select Case Nfo

                Case 0
                    OSInfo = "Windows " & OSN

                Case 1
                    OSInfo = .dwBuildNumber

                Case 2
                    OSInfo = TrimNull(.szCSDVersion)

                Case 4
                    OSInfo = .dwMajorVersion & "." & .dwMinorVersion

                Case Else
                    OSInfo = "ERROR!"
            End Select
        End With
    End If

    Exit Function
HandErr:
    OSInfo = "GetWinVer Err.Number: " & err.Number & " Err.Description: " & err.Description
End Function
