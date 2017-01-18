Attribute VB_Name = "mCollectInfo"
Option Explicit

Public strControlSet As String

' From MSDN
'BOOL SetupGetInfDriverStoreLocation(
'  __in       PCTSTR FileName,
'  __in_opt   PSP_ALTPLATFORM_INFO AlternatePlatformInfo,
'  __in_opt   PCTSTR LocaleName,
'  __out      PTSTR ReturnBuffer,
'  __in       DWORD ReturnBufferSize,
'  __out_opt  PDWORD RequiredSize
');
Private Declare Function SetupGetInfDriverStoreLocationW Lib "setupapi.dll" (ByVal FileName As Long, AlternatePlatformInfo As PSP_ALTPLATFORM_INFO_V2, ByVal LocaleName As Long, ByVal ReturnBuffer As Long, ByVal ReturnBufferSize As Long, ByRef RequiredSize As Long) As Long

Private Type PSP_ALTPLATFORM_INFO_V2
    cbSize As Long
    Platform As Long
    MajorVersion As Long
    MinorVersion As Long
    ProcessorArchitecture As Integer
    Reserved As Integer
    flags As Integer
    FirstValidatedMajorVersion As Long
    FirstValidatedMinorVersion As Long
End Type

Private Const SP_ALTPLATFORM_FLAGS_VERSION_RANGE = &H1

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetInfDriverStorePath
'! Description (Описание)  :   [Получение пути расположения драйвера по inf-файлу]
'! Parameters  (Переменные):   sInfPath (String)
'!--------------------------------------------------------------------------------
Public Function GetInfDriverStorePath(ByVal sInfPath As String) As String

    Dim sBuffer     As String
    Dim PSPAI       As PSP_ALTPLATFORM_INFO_V2
    Dim lngRet      As Long
    Dim lngSizeBuff As Long
    Dim OSVI        As OSVERSIONINFOEX
    Dim SI          As SYSTEM_INFO

    If APIFunctionPresent("SetupGetInfDriverStoreLocationW", "setupapi.dll") Then
        ' получение расширенны данных о версии ОС и архитектуре процессора
        OSVI.dwOSVersionInfoSize = Len(OSVI)
        GetVersionEx OSVI

        ' Назначение полученных параметров для PSP_ALTPLATFORM_INFO_V2
        With PSPAI
            .FirstValidatedMajorVersion = 5
            .FirstValidatedMinorVersion = 1
            .MajorVersion = OSVI.dwMajorVersion
            .MinorVersion = OSVI.dwMinorVersion
            .Platform = OSVI.dwPlatformID

            If APIFunctionPresent("GetNativeSystemInfo", "kernel32.dll") Then
                GetNativeSystemInfo SI
                .ProcessorArchitecture = SI.wProcessorArchitecture
            Else
                .ProcessorArchitecture = PROCESSOR_ARCHITECTURE_INTEL
            End If

            .Reserved = 0
            .flags = SP_ALTPLATFORM_FLAGS_VERSION_RANGE
        End With
        If mbDebugStandart Then DebugMode "******GetInfDriverStorePath: " & sInfPath
        'lngRet = SetupGetInfDriverStoreLocationW(ByVal StrPtr(sInfPath & vbNullChar), PSPAI, StrPtr(vbNullString), ByVal 0&, 0&, lngSizeBuff)
        'sBuffer = String$(MAX_PATH_UNICODE, 0)
        sBuffer = FillNullChar(MAX_PATH_UNICODE)
        lngRet = SetupGetInfDriverStoreLocationW(ByVal StrPtr(sInfPath & vbNullChar), PSPAI, StrPtr(vbNullString), ByVal StrPtr(sBuffer), Len(sBuffer), 0&)

        GetInfDriverStorePath = TrimNull(sBuffer)
        
        If lngRet = 0 Then
            If mbDebugStandart Then DebugMode "******GetInfDriverStorePath: Err №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        Else
            If mbDebugStandart Then DebugMode "******GetInfDriverStorePath: ResultValue - " & GetInfDriverStorePath
        End If

    Else
        If mbDebugStandart Then DebugMode "******GetInfDriverStorePath: ApiFunction not Supported"
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReadDrivers
'! Description (Описание)  :   [Сбор информации о драйверах и занесение в массив arrHwidsLocal]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ReadDrivers()

    Dim arr_CH()            As String
    Dim arr_Z()             As String
    Dim arr_U()             As String
    Dim nn                  As Long
    Dim ii                  As Long
    Dim jj                  As Long
    Dim hh                  As Long
    Dim miMaxCountArr       As Long
    Dim miPBInterval        As Long
    Dim miPBNext            As Long
    Dim strClassID          As String
    Dim strProviderName     As String
    Dim strClassName        As String
    Dim strClass            As String
    Dim strDriverDesc       As String
    Dim strInfPath          As String
    Dim strDriverDate       As String
    Dim strDriverVersion    As String
    Dim strInfSection       As String
    Dim strInfSectionExt    As String
    Dim strMatchingDeviceId As String
    Dim regNameClass        As String
    Dim mbR                 As Boolean
    Dim strSS               As String
    Dim StringHash          As Scripting.Dictionary
    Dim lngUBoundCH         As Long
    Dim lngUBoundZ          As Long
    Dim lngUBoundU          As Long

    'Откуда берем данные CurrentControlSet
    strControlSet = "SYSTEM\CurrentControlSet"
    Set StringHash = CreateObject("Scripting.Dictionary")
    StringHash.CompareMode = 1
    If mbDebugDetail Then DebugMode "ReadDrivers-Start"

    '# sub to read drivers and populate grid #
    On Error Resume Next

    DoEvents
    miPBNext = 100
    ' Изменяем прогресс
    frmProgress.ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateInProgress
    frmProgress.ChangeProgressBarStatus miPBNext
    
    '# list all class of drivers installed
    If mbDebugStandart Then DebugMode "***ReadDrivers: ListKey - HKEY_LOCAL_MACHINE\" & strControlSet & "\Control\Class"
    arr_Z = ListKey(HKEY_LOCAL_MACHINE, strControlSet & "\Control\Class", False)
    If mbDebugStandart Then DebugMode "***ReadDrivers: ListKey Class - " & UBound(arr_Z)
    ' Изменяем прогресс
    frmProgress.ChangeProgressBarStatus miPBNext, 900
    
    ' максимальное кол-во элементов в массиве arr_CH
    miMaxCountArr = 500
    ReDim arr_CH(miMaxCountArr) As String

    ' Переменная для прогресса
    lngUBoundZ = UBound(arr_Z)
    If lngUBoundZ > 0 Then
        miPBInterval = Round(2000 / lngUBoundZ)
    Else
        miPBInterval = 1900
    End If

    For ii = 0 To lngUBoundZ
        arr_U = ListKey(HKEY_LOCAL_MACHINE, strControlSet & "\Control\Class\" & arr_Z(ii), False)

        lngUBoundU = UBound(arr_U)

        For jj = 0 To lngUBoundU

            ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
            If nn = miMaxCountArr Then
                miMaxCountArr = miMaxCountArr + miMaxCountArr
                ReDim Preserve arr_CH(miMaxCountArr)
            End If

            If LenB(arr_U(jj)) = 0 Then
                arr_CH(nn) = arr_Z(ii)
            Else
                arr_CH(nn) = arr_Z(ii) & vbBackslash & arr_U(jj)
            End If

            If mbDebugDetail Then DebugMode "******ReadDrivers: ListKey Result - " & arr_CH(nn)
            nn = nn + 1
        Next jj
        ' Изменяем прогресс
        frmProgress.ChangeProgressBarStatus miPBNext, miPBInterval
    Next ii

    If nn > 0 Then
        ReDim Preserve arr_CH(nn - 1)
    Else
        ReDim Preserve arr_CH(0)
    End If

    If mbDebugStandart Then DebugMode "***ReadDrivers-Start: ListKey Result- " & UBound(arr_CH)
    '# get all info of each instaled driver #
    hh = 0
    ' максимальное кол-во элементов в массиве arrHwidsLocal
    miMaxCountArr = 200
    ReDim arrHwidsLocal(miMaxCountArr)
    If mbDebugStandart Then DebugMode "***ReadDrivers: Collect Full Info"
    If mbDebugStandart Then DebugMode "*****************************************"

    ' Переменная для прогресса
    lngUBoundCH = UBound(arr_CH)
    If lngUBoundCH > 0 Then
        miPBInterval = Round(7000 / lngUBoundCH)
    Else
        miPBInterval = 6500
    End If

    For ii = 0 To lngUBoundCH

        ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
        If hh = miMaxCountArr Then
            miMaxCountArr = miMaxCountArr + miMaxCountArr
            ReDim Preserve arrHwidsLocal(miMaxCountArr)
        End If

        strClassID = arr_CH(ii)
        regNameClass = strControlSet & "\Control\Class\" & strClassID
        strDriverDesc = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDesc", True)

        ' Если устройство не имеет названия, значит это сокрее всего не устройство
        If LenB(strDriverDesc) > 0 Then
            ' собираем доп-инфо
            strDriverDate = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDate", True)

            ' если необходимо конвертировать дату в формат dd/mm/yyyy
            If LenB(strDriverDate) > 0 Then
                ConvertDate2Rus strDriverDate
            End If

            strDriverVersion = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverVersion", True)
            strProviderName = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "ProviderName", True)
            strClassName = GetKeyValue(HKEY_LOCAL_MACHINE, strControlSet & "\Control\Class\" & Mid$(strClassID, 1, Len(strClassID) - 5), vbNullString, True)
            strClass = GetKeyValue(HKEY_LOCAL_MACHINE, strControlSet & "\Control\Class\" & Mid$(strClassID, 1, Len(strClassID) - 5), "Class", True)
            strInfPath = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "InfPath", True)
            strInfSection = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "InfSection", True)
            strInfSectionExt = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "InfSectionExt", True)
            strMatchingDeviceId = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "MatchingDeviceId", True)
            ' Если нет повторов, то заносим данные в массив
            strSS = strDriverDesc & strInfPath & strInfSection & strInfSectionExt & strMatchingDeviceId
            mbR = StringHash.Exists(strSS)

            If Not mbR Then
                StringHash.item(strSS) = "+"
                'Заполняем массив даными
                arrHwidsLocal(hh).i0_DriverDesc = strDriverDesc
                arrHwidsLocal(hh).i1_DriverDate = strDriverDate
                arrHwidsLocal(hh).i2_DriverVersion = strDriverVersion
                arrHwidsLocal(hh).i3_ProviderName = strProviderName
                
                If LenB(strClassName) = 0 Then
                    arrHwidsLocal(hh).i4_ClassName = strClass
                Else
                    arrHwidsLocal(hh).i4_ClassName = strClassName
                End If

                arrHwidsLocal(hh).i5_Class = strClass
                arrHwidsLocal(hh).i6_InfPath = strInfPath
                arrHwidsLocal(hh).i7_InfSection = strInfSection & strInfSectionExt
                arrHwidsLocal(hh).i8_MatchingDeviceId = strMatchingDeviceId
                arrHwidsLocal(hh).i9_ClassID = strClassID

                'Вывод инфо в лог
                If mbDebugStandart Then DebugMode "RowNum: " & hh & " From: " & regNameClass
                If mbDebugStandart Then DebugMode "ClassID: " & strClassID
                If mbDebugStandart Then DebugMode "DriverDesc: " & strDriverDesc
                If mbDebugStandart Then DebugMode "ClassName: " & strClass & " : " & strClassName
                If mbDebugStandart Then DebugMode "ProviderName: " & strProviderName
                If mbDebugStandart Then DebugMode "InfPath: " & strInfPath & ", " & strInfSection & strInfSectionExt
                If mbDebugStandart Then DebugMode "DriverDate: " & strDriverDate
                If mbDebugStandart Then DebugMode "DriverVersion: " & strDriverVersion
                If mbDebugStandart Then DebugMode "MatchingDeviceId: " & strMatchingDeviceId
                If mbDebugStandart Then DebugMode "*****************************************"

                hh = hh + 1
            End If
        End If

        ' Изменяем прогресс
        frmProgress.ChangeProgressBarStatus miPBNext, miPBInterval
    Next
    
    ' Финишируем прогресс
    frmProgress.ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    frmProgress.ChangeProgressBarStatus 10000
    
    DoEvents
    
    ' Переобъявляем массив на реальное кол-во записей
    If hh > 0 Then
        ReDim Preserve arrHwidsLocal(hh - 1)
    Else
        ReDim Preserve arrHwidsLocal(0)
    End If

    If mbDebugStandart Then DebugMode "*****************************************"
    If mbDebugStandart Then DebugMode "ReadDrivers-Finish"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CollectCmdString
'! Description (Описание)  :   [Создание коммандной строки запуска программы DPInst]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function CollectCmdString() As String

    Dim strCmdStringDPInstTemp As String

    If mbDpInstLegacyMode Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/LM "
    End If

    If mbDpInstPromptIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/P "
    End If

    If mbDpInstForceIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/F "
    End If

    If mbDpInstSuppressAddRemovePrograms Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SA "
    End If

    If mbDpInstSuppressWizard Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SW "
    End If

    If mbDpInstQuietInstall Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/Q "
    End If

    If mbDpInstScanHardware Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SH "
    End If

    ' Результирующая строка
    CollectCmdString = strCmdStringDPInstTemp
End Function
