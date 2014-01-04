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
    Flags As Integer
    FirstValidatedMajorVersion As Long
    FirstValidatedMinorVersion As Long
End Type

Private Const SP_ALTPLATFORM_FLAGS_VERSION_RANGE = &H1

Public Function GetInfDriverStorePath(sInfPath As String) As String

    Dim sBuffer     As String
    Dim as12        As PSP_ALTPLATFORM_INFO_V2
    Dim ret         As Long
    Dim lngSizeBuff As Long
    Dim OSVI        As OSVERSIONINFOEX
    Dim SI          As SYSTEM_INFO

    If APIFunctionPresent("SetupGetInfDriverStoreLocationW", "setupapi.dll") Then
        ' ��������� ���������� ������ � ������ �� � ����������� ����������
        OSVI.dwOSVersionInfoSize = Len(OSVI)
        GetVersionEx OSVI

        ' ���������� ���������� ���������� ��� PSP_ALTPLATFORM_INFO_V2
        With as12
            .FirstValidatedMajorVersion = 6
            .FirstValidatedMinorVersion = 0
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
            .Flags = SP_ALTPLATFORM_FLAGS_VERSION_RANGE
        End With

        DebugMode "******GetInfDriverStorePath: " & sInfPath
        ret = SetupGetInfDriverStoreLocationW(ByVal StrPtr(sInfPath), as12, StrPtr(vbNullString), ByVal 0&, 0&, lngSizeBuff)
        sBuffer = String$(lngSizeBuff, 0)
        ret = SetupGetInfDriverStoreLocationW(ByVal StrPtr(sInfPath), as12, StrPtr(vbNullString), ByVal StrPtr(sBuffer), Len(sBuffer), 0&)
        GetInfDriverStorePath = TrimNull(sBuffer)

        If ret = 0 Then
            DebugMode "******GetInfDriverStorePath: Err �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
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
    Dim j                   As Long
    Dim H                   As Long
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
    Dim R                   As Boolean
    Dim ss                  As String
    Dim StringHash          As Scripting.Dictionary

    '������ ����� ������ CurrentControlSet
    strControlSet = "SYSTEM\CurrentControlSet"
    'strControlSet = "SYSTEM\ControlSet001"
    Set StringHash = CreateObject("Scripting.Dictionary")
    StringHash.CompareMode = 1
    DebugMode "ReadDrivers-Start"

    '# sub to read drivers and populate grid #
    On Error Resume Next

    DoEvents
    miPbNext = 100
    ' �������� ��������
    frmProgress.ProgressBar1.SetTaskBarProgressState PrbTaskBarStateInProgress
    frmProgress.ChangeProgressBarStatus miPbNext, 100
    
    '# list all class of drivers installed
    DebugMode "***ReadDrivers: ListKey - HKEY_LOCAL_MACHINE\" & strControlSet & "\Control\Class"
    Z = ListKey(HKEY_LOCAL_MACHINE, strControlSet & "\Control\Class", False)
    DebugMode "***ReadDrivers: ListKey Class - " & UBound(Z)
    ' �������� ��������
    frmProgress.ChangeProgressBarStatus miPbNext, 1000
    miMaxCountArr = 200
    ' ������������ ���-�� ��������� � �������
    ReDim CH(miMaxCountArr) As String

    ' ���������� ��� ���������
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

        For j = 0 To lngUBoundU

            ' ���� ������� � ������� ���������� ������ ��� ���������, �� ����������� ����������� �������
            If n = miMaxCountArr Then
                miMaxCountArr = miMaxCountArr + miMaxCountArr
                ReDim Preserve CH(miMaxCountArr)
            End If

            If LenB(U(j)) = 0 Then
                CH(n) = Z(i)
            Else
                CH(n) = Z(i) & vbBackslash & U(j)
            End If

            DebugMode "******ReadDrivers: ListKey Result - " & CH(n), 2
            n = n + 1
        Next
        ' �������� ��������
        frmProgress.ChangeProgressBarStatus miPbNext, miPbInterval
    Next

    If n > 0 Then
        ReDim Preserve CH(n - 1)
    Else
        ReDim Preserve CH(0)
    End If

    DebugMode "***ReadDrivers-Start: ListKey Result- " & UBound(CH)
    '# get all info of each instaled driver #
    H = 0
    miMaxCountArr = 200
    ' ������������ ���-�� ��������� � �������
    ReDim arrHwidsLocal(10, miMaxCountArr) As String
    DebugMode "***ReadDrivers: Collect Full Info"
    DebugMode "*****************************************"

    ' ���������� ��� ���������
    If UBound(CH) > 0 Then
        miPbInterval = Round(7000 / UBound(CH))
    Else
        miPbInterval = 6500
    End If

    Dim lngUBoundCH As Long

    lngUBoundCH = UBound(CH)

    For i = 0 To lngUBoundCH

        ' ���� ������� � ������� ���������� ������ ��� ���������, �� ����������� ����������� �������
        If H = miMaxCountArr Then
            miMaxCountArr = miMaxCountArr + miMaxCountArr
            ReDim Preserve arrHwidsLocal(10, miMaxCountArr)
        End If

        strClassID = CH(i)
        regNameClass = strControlSet & "\Control\Class\" & strClassID
        strDriverDesc = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDesc", True)

        ' ���� ���������� �� ����� ��������, ������ ��� ������ ����� �� ����������
        If LenB(strDriverDesc) > 0 Then
            ' �������� ���-����
            strDriverDate = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDate", True)

            ' ���� ���������� �������������� ���� � ������ dd/mm/yyyy
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
            ' ���� ��� ��������, �� ������� ������ � ������
            ss = strDriverDesc & strInfPath & strInfSection & strInfSectionExt & strMatchingDeviceId
            R = StringHash.Exists(ss)

            If Not R Then
                StringHash.Item(ss) = "+"
                '��������� ������ ������
                arrHwidsLocal(0, H) = strDriverDesc
                arrHwidsLocal(1, H) = strDriverDate
                arrHwidsLocal(2, H) = strDriverVersion
                arrHwidsLocal(3, H) = strProviderName
                
                If LenB(strClassName) = 0 Then
                    arrHwidsLocal(4, H) = strClass
                Else
                    arrHwidsLocal(4, H) = strClassName
                End If

                arrHwidsLocal(5, H) = strClass
                arrHwidsLocal(6, H) = strInfPath
                arrHwidsLocal(7, H) = strInfSection & strInfSectionExt
                arrHwidsLocal(8, H) = strMatchingDeviceId
                arrHwidsLocal(9, H) = strClassID

                '����� ���� � ���
                If mbDebugEnable Then
                    DebugMode "RowNum: " & H & " From: " & regNameClass
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

                H = H + 1
            End If
        End If

        ' �������� ��������
        frmProgress.ChangeProgressBarStatus miPbNext, miPbInterval
    Next
    
    ' ���������� ��������
    frmProgress.ChangeProgressBarStatus 10000, 0
    frmProgress.ProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    
    DoEvents
    
    ' ������������� ������ �� �������� ���-�� �������
    If H > 0 Then
        ReDim Preserve arrHwidsLocal(10, H - 1)
    Else
        ReDim Preserve arrHwidsLocal(10, 0)
    End If

    DebugMode "*****************************************"
    DebugMode "ReadDrivers-Finish"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CollectCmdString
'! Description (��������)  :   [�������� ���������� ������ ������� ��������� DPInst]
'! Parameters  (����������):
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

    ' �������������� ������
    CollectCmdString = strCmdStringDPInstTemp
End Function
