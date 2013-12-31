Attribute VB_Name = "mCollectInfo"
Option Explicit

Public strControlSet As String

Private Declare Function SetupGetInfDriverStoreLocationW _
                Lib "setupapi.dll" (ByVal FileName As Long, _
                                    AlternatePlatformInfo As PSP_ALTPLATFORM_INFO_V2, _
                                    ByVal LocaleName As Long, _
                                    ByVal ReturnBuffer As Long, _
                                    ByVal ReturnBufferSize As Long, _
                                    ByRef RequiredSize As Long) As Long

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

' From MSDN
'BOOL SetupGetInfDriverStoreLocation(
'  __in       PCTSTR FileName,
'  __in_opt   PSP_ALTPLATFORM_INFO AlternatePlatformInfo,
'  __in_opt   PCTSTR LocaleName,
'  __out      PTSTR ReturnBuffer,
'  __in       DWORD ReturnBufferSize,
'  __out_opt  PDWORD RequiredSize
');
Public Function GetInfDriverStorePath(sInfPath As String) As String

    Dim sBuffer     As String
    Dim as12        As PSP_ALTPLATFORM_INFO_V2
    Dim ret         As Long
    Dim lngSizeBuff As Long
    Dim OSVI        As OSVERSIONINFO
    Dim SI          As SYSTEM_INFO

    If APIFunctionPresent("SetupGetInfDriverStoreLocationW", "setupapi.dll") Then
        ' получение расширенны данных о версии ОС и архитектуре процессора
        OSVI.dwOSVersionInfoSize = Len(OSVI)
        GetVersionEx OSVI

        ' Назначение полученных параметров для PSP_ALTPLATFORM_INFO_V2
        With as12
            .FirstValidatedMajorVersion = 6
            .FirstValidatedMinorVersion = 0
            .MajorVersion = OSVI.dwMajorVersion
            .MinorVersion = OSVI.dwMinorVersion
            .Platform = OSVI.dwPlatformId

            If APIFunctionPresent("GetNativeSystemInfo", "kernel32.dll") Then
                GetNativeSystemInfo SI
                .ProcessorArchitecture = SI.wProcessorArchitecture
            Else
                .ProcessorArchitecture = PROCESSOR_ARCHITECTURE_INTEL
            End If

            .Reserved = 0
            .flags = SP_ALTPLATFORM_FLAGS_VERSION_RANGE
        End With

        DebugMode "******GetInfDriverStorePath: " & sInfPath
        ret = SetupGetInfDriverStoreLocationW(ByVal StrPtr(sInfPath), as12, StrPtr(vbNullString), ByVal 0&, 0&, lngSizeBuff)
        sBuffer = String$(lngSizeBuff, 0)
        ret = SetupGetInfDriverStoreLocationW(ByVal StrPtr(sInfPath), as12, StrPtr(vbNullString), ByVal StrPtr(sBuffer), Len(sBuffer), 0&)
        GetInfDriverStorePath = TrimNull(sBuffer)

        If ret = 0 Then
            DebugMode "******GetInfDriverStorePath: Err №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
        Else
            DebugMode "******GetInfDriverStorePath: ResultValue - " & GetInfDriverStorePath
        End If

    Else
        DebugMode "******GetInfDriverStorePath: ApiFunction not Supported"
    End If
End Function

Public Sub ReadDrivers()

    Dim CH()                As String
    Dim Z()                 As String
    Dim U()                 As String
    Dim n                   As Long
    Dim i                   As Long
    Dim J                   As Long
    Dim h                   As Long
    Dim miMaxCountArr       As Long
    Dim miPbInterval        As Long
    Dim miPbNext            As Long
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
    Dim r                   As Boolean
    Dim ss                  As String
    Dim StringHash          As Scripting.Dictionary

    'Откуда берем данные CurrentControlSet
    strControlSet = "SYSTEM\CurrentControlSet"
    'strControlSet = "SYSTEM\ControlSet001"
    Set StringHash = CreateObject("Scripting.Dictionary")
    StringHash.CompareMode = 1
    DebugMode "ReadDrivers-Start"

    '# sub to read drivers and populate grid #
    On Error Resume Next

    DoEvents
    miPbNext = 100
    ' Изменяем прогресс
    frmProgress.ChangeProgressBarStatus miPbNext, 100
    '# list all class of drivers installed
    DebugMode "***ReadDrivers: ListKey - HKEY_LOCAL_MACHINE\" & strControlSet & "\Control\Class"
    Z = ListKey(HKEY_LOCAL_MACHINE, strControlSet & "\Control\Class", False)
    DebugMode "***ReadDrivers: ListKey Class - " & UBound(Z)
    ' Изменяем прогресс
    frmProgress.ChangeProgressBarStatus miPbNext, 1000
    miMaxCountArr = 200
    ' максимальное кол-во элементов в массиве
    ReDim CH(miMaxCountArr) As String

    ' Переменная для прогресса
    If UBound(Z) > 0 Then
        miPbInterval = Round(2000 / UBound(Z))
    Else
        miPbInterval = 1900
    End If

    Dim lngUBoundZ As Long

    lngUBoundZ = UBound(Z)

    For i = 0 To lngUBoundZ
        U = ListKey(HKEY_LOCAL_MACHINE, strControlSet & "\Control\Class\" & Z(i), False)

        Dim lngUBoundU As Long

        lngUBoundU = UBound(U)

        For J = 0 To lngUBoundU

            ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
            If n = miMaxCountArr Then
                miMaxCountArr = miMaxCountArr + miMaxCountArr
                ReDim Preserve CH(miMaxCountArr)
            End If

            If U(J) = vbNullString Then
                CH(n) = Z(i)
            Else
                CH(n) = Z(i) & vbBackslash & U(J)
            End If

            DebugMode "******ReadDrivers: ListKey Result - " & CH(n), 2
            n = n + 1
        Next
        ' Изменяем прогресс
        frmProgress.ChangeProgressBarStatus miPbNext, miPbInterval
    Next

    If n > 0 Then
        ReDim Preserve CH(n - 1)
    Else
        ReDim Preserve CH(0)
    End If

    DebugMode "***ReadDrivers-Start: ListKey Result- " & UBound(CH)
    '# get all info of each instaled driver #
    h = 0
    miMaxCountArr = 200
    ' максимальное кол-во элементов в массиве
    ReDim arrHwidsLocal(10, miMaxCountArr) As String
    DebugMode "***ReadDrivers: Collect Full Info"
    DebugMode "*****************************************"

    ' Переменная для прогресса
    If UBound(CH) > 0 Then
        miPbInterval = Round(7000 / UBound(CH))
    Else
        miPbInterval = 6500
    End If

    Dim lngUBoundCH As Long

    lngUBoundCH = UBound(CH)

    For i = 0 To lngUBoundCH

        ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
        If h = miMaxCountArr Then
            miMaxCountArr = miMaxCountArr + miMaxCountArr
            ReDim Preserve arrHwidsLocal(10, miMaxCountArr)
        End If

        strClassID = CH(i)
        regNameClass = strControlSet & "\Control\Class\" & strClassID
        strDriverDesc = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDesc", True)

        ' Если устройство не имеет названия, значит это сокрее всего не устройство
        If LenB(strDriverDesc) > 0 Then
            ' собираем доп-инфо
            strDriverDate = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDate", True)

            ' если необходимо конвертировать дату в формат dd/mm/yyyy
            If LenB(strDriverDate) > 0 Then
                strDriverDate = ConvertDate2Rus(strDriverDate)
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
            ss = strDriverDesc & strInfPath & strInfSection & strInfSectionExt & strMatchingDeviceId
            r = StringHash.Exists(ss)

            If Not r Then
                StringHash.Item(ss) = "+"
                'Заполняем массив даными
                arrHwidsLocal(0, h) = strDriverDesc
                arrHwidsLocal(1, h) = strDriverDate
                arrHwidsLocal(2, h) = strDriverVersion
                arrHwidsLocal(3, h) = strProviderName
                
                If strClassName = vbNullString Then
                    arrHwidsLocal(4, h) = strClass
                Else
                    arrHwidsLocal(4, h) = strClassName
                End If

                arrHwidsLocal(5, h) = strClass
                arrHwidsLocal(6, h) = strInfPath
                arrHwidsLocal(7, h) = strInfSection & strInfSectionExt
                arrHwidsLocal(8, h) = strMatchingDeviceId
                arrHwidsLocal(9, h) = strClassID

                'Вывод инфо в лог
                If mboolDebugEnable Then
                    DebugMode "RowNum: " & h & " From: " & regNameClass
                    DebugMode "ClassID: " & strClassID
                    DebugMode "DriverDesc: " & strDriverDesc
                    DebugMode "ClassName: " & strClass & " : " & strClassName
                    DebugMode "ProviderName: " & strProviderName
                    DebugMode "InfPath: " & strInfPath & ", " & strInfSection & strInfSectionExt
                    DebugMode "DriverDate: " & strDriverDate
                    DebugMode "DriverVersion: " & strDriverVersion
                    DebugMode "MatchingDeviceId: " & strMatchingDeviceId
                    DebugMode "*****************************************"
                End If

                h = h + 1
            End If
        End If

        ' Изменяем прогресс
        frmProgress.ChangeProgressBarStatus miPbNext, miPbInterval
    Next
    ' Финишируем прогресс
    'miPbNext = 10000
    'ChangeProgressBarStatus frmProgress,frmProgress.ctlProgressBar1, miPbNext, 0
    frmProgress.ctlProgressBar1.Value = 10000
    DoEvents

    ' Переобъявляем массив на реальное кол-во записей
    If h > 0 Then
        ReDim Preserve arrHwidsLocal(10, h - 1)
    Else
        ReDim Preserve arrHwidsLocal(10, 0)
    End If

    DebugMode "*****************************************"
    DebugMode "ReadDrivers-Finish"
End Sub
