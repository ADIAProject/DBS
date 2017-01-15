Attribute VB_Name = "mMainIniSettings"
Option Explicit

Public mbPatnAbs                         As Boolean     ' ���� �� �������� �������� �����������, ������������ � ���������� frmOptions
Public mbAllFolderDRVNotExist            As Boolean     ' ��� �������� � �������� ���������, ��������� � ���������� �� ����������

' ��������� ��������� ����������� �� ini-�����
Public strSysIni                         As String      ' ������� ���� ��������
Public mbLoadIniTmpAfterRestart          As Boolean     ' ��������� ini �� ��������� �����
Public lngOSCount                        As Long        ' ���������� �� �������������� ����������
Public lngOSCountPerRow                  As Long        ' ���������� ��, ������������ �� ����� ������
Public lngUtilsCount                     As Long        ' ���������� ������, ����������� � ����������
Public strDevconCmdPath                  As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DevCon\devcon_c.cmd
Public strArh7zExePath                   As String      ' ���� �� ����������� ������ � ������ ������� ������ - ����������, � ����������� �� �����������, �� ���������� ����
Public strArh7zExePath86                 As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\Arc\7za.exe
Public strArh7zExePath64                 As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\Arc\7za64.exe
Public strArh7zParam1                    As String
Public strArh7zParam2                    As String
Public strArh7zSFXPATH                   As String
Public strArh7zSFXConfigPath             As String
Public strArh7zSFXConfigPathEn           As String
Public strDPInstExePath64                As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DPInst\DPInst64.exe
Public strDPInstExePath86                As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DPInst\DPInst.exe
Public strDPInstExePath                  As String      ' ���� �� ����������� ������ � ������ ������� ������ - ����������, � ����������� �� �����������, �� ���������� ����
Public mbDelTmpAfterClose                As Boolean
Public mbUpdateCheck                     As Boolean
Public mbUpdateCheckBeta                 As Boolean
Public mbUpdateToolTip                   As Boolean
Public miStartMode                       As Long
Public mbRecursion                       As Boolean
Public mbSaveSizeOnExit                  As Boolean
Public strExcludeHWID                    As String
Public lngStartModeTab2                  As Long        ' ��������� ������� ��� ����� �������
Public strThisBuildBy                    As String      ' ��������� � �������� � ������� ���� � �������� ���������
Public mbTabBlock                        As Boolean
Public mbTabHide                         As Boolean
Public mbButtonTextUpCase                As Boolean
Public mbLoadFinishFile                  As Boolean
Public mbReadClasses                     As Boolean
Public mbReadDPName                      As Boolean
Public mbConvertDPName                   As Boolean
Public strExcludeFileName                As String
Public strImageStatusButtonName          As String
Public strImageMainName                  As String
Public mbEULAAgree                       As Boolean
Public mbCompareDrvVerByDate             As Boolean     ' ��������� ������ ��������� �� ����
Public mbLoadUnSupportedOS               As Boolean     ' �������\��������� �������� ��� ������������� ��
Public mbAutoInfoAfterDelDRV             As Boolean     ' �������������� ������������ ��� �������� ��������
Public mbDateFormatRus                   As Boolean     ' �������������� ������������ ��� �������� ��������
'Public mbCreateRestorePoint              As Boolean     ' ���������� ��� ������ �������� ����� ��������������
'Public mbCreateRestorePointDone          As Boolean     ' ����, ������������ ��� ����� �������������� ��� ����������� �����
Public mbDisableDEP                      As Boolean     ' ���������� ��� ����������� ���������� DEP
Public mbHideOtherProcess                As Boolean     ' �������� ��������� �������� ��� �������
Public mbDP_Is_aFolder                   As Boolean     ' ������ ��������� � ���� ����� - �.� ������������� ����� ���������
Public mbStartMaximazed                  As Boolean     ' ��������� ��������� ����������� �� ���� �����
Public mbTempPath                        As Boolean     ' ������������ �������������� ������� Temp - �.� �������� �������
Public strAlternativeTempPath            As String      ' ���� ��� ��������������� �������� Temp
Public mbDpInstLegacyMode                As Boolean     ' ��������� DPinst
Public mbDpInstPromptIfDriverIsNotBetter As Boolean     ' ��������� DPinst
Public mbDpInstForceIfDriverIsNotBetter  As Boolean     ' ��������� DPinst
Public mbDpInstSuppressAddRemovePrograms As Boolean     ' ��������� DPinst
Public mbDpInstSuppressWizard            As Boolean     ' ��������� DPinst
Public mbDpInstQuietInstall              As Boolean     ' ��������� DPinst
Public mbDpInstScanHardware              As Boolean     ' ��������� DPinst
Public mbSearchOnStart                   As Boolean     ' ������ ����� ���������� ��� ������� ���������
Public lngPauseAfterSearch               As Long        ' ������ ����� ������ ����� ���������� ��� ������� ���������
Public mbCalcDriverScore                 As Boolean     ' ������������ ��� ������� ��������� ���� ���������� ��������, �� ��������� ��������� �������
Public mbCompatiblesHWID                 As Boolean     ' ������������ ��� ������ ���������� ��������� ������ CompatiblesHWID, ������� �� �������
Public mbSearchCompatibleDriverOtherOS   As Boolean     ' ������ ���������� �������� �� ���� ��������, � �� ������ �� �����������
Public lngSortMethodShell                As Long        ' �������� ����������� ��������� ����� ���������� �������
Public lngCompatiblesHWIDCount           As Long        ' ������� ������ ����������� HWID
Public mbMatchHWIDbyDPName               As Boolean     ' ������ ����� ����� ��� ���������� ������������� ��������
Public lngMainFormWidth                  As Long        ' ������ �������� �����
Public lngMainFormHeight                 As Long        ' ������ �������� �����
Public lngButtonWidth                    As Long        ' ������ ������
Public lngButtonHeight                   As Long        ' ������ ������
Public lngButtonLeft                     As Long        ' ������ ����� ��� ������
Public lngButtonTop                      As Long        ' ������ ������ ��� ������
Public lngBtn2BtnLeft                    As Long        ' �������� ����� �������� �� �����������
Public lngBtn2BtnTop                     As Long        ' �������� ����� �������� �� ���������
Public lngStatusBtnStyle                 As Long        ' ����� ������ ������ ���������
Public lngStatusBtnStyleColor            As Long        ' ���� ���������� ������ ������ ���������
Public lngStatusBtnBackColor             As Long        ' ���� ���������� ������ ������ ���������
Public lngFreeSpaceSysDrive              As Long        ' ��������� ����� �� ������� �����

'Public strImageMenuName                  As String
'Public mbExMenu                           As Boolean ' ����������� ����

'-------------------- ��������� �������� ���� � ������  ------------------'
Public Const lngMainFormWidthMin         As Long = 13000    ' ����������� �������� �������� �����
Public Const lngMainFormHeightMin        As Long = 6500     ' ����������� �������� �������� �����
'Public Const lngButtonWidthMin           As Long = 1500     ' ����������� �������� �������� ������ - ������
'Public Const lngButtonHeightMin          As Long = 350      ' ����������� �������� �������� ������ - ������
Private Const lngMainFormWidthDef        As Long = 13000    ' ��������� �������� �������� �����
Private Const lngMainFormHeightDef       As Long = 8400     ' ��������� �������� �������� �����
'Private Const lngButtonWidthDef          As Long = 2150     ' ��������� �������� �������� ������ - ������
'Private Const lngButtonHeightDef         As Long = 550      ' ��������� �������� �������� ������ - ������
'Private Const lngButtonLeftDef           As Long = 100      ' ��������� �������� �������� ������ - ������ ����� ��� ������
'Private Const lngButtonTopDef            As Long = 100      ' ��������� �������� �������� ������ - ������ ������ ��� ������
'Private Const lngBtn2BtnLeftDef          As Long = 100      ' ��������� �������� �������� ������ - �������� ����� �������� �� �����������
'Private Const lngBtn2BtnTopDef           As Long = 100      ' ��������� �������� �������� ������ - �������� ����� �������� �� ���������

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub CreateIni
'! Description (��������)  :   [���������� �������� � ��� ���� ���� ����� ���]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub CreateIni()

    If FileExists(strSysIni) = False Then
        If mbIsDriveCDRoom Then
            strSysIni = strWorkTempBackSL & strSettingIniFile
            MsgBox "File " & strSettingIniFile & " is not Exist!" & vbNewLine & _
                   "This program works from CD\DVD, so we create temporary " & strSettingIniFile & "-file" & vbNewLine & _
                   strSysIni, vbInformation + vbApplicationModal, strProductName
        End If

        '������ Main
        IniWriteStrPrivate "Main", "DelTmpAfterClose", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheck", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheckBeta", "0", strSysIni
        IniWriteStrPrivate "Main", "StartMode", "2", strSysIni
        IniWriteStrPrivate "Main", "EULAAgree", "0", strSysIni
        IniWriteStrPrivate "Main", "HideOtherProcess", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTemp", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTempPath", "%Temp%", strSysIni
        IniWriteStrPrivate "Main", "SilentDLL", "0", strSysIni
        IniWriteStrPrivate "Main", "IconMainSkin", "Standart", strSysIni
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", "0", strSysIni
        IniWriteStrPrivate "Main", "AutoLanguage", "1", strSysIni
        IniWriteStrPrivate "Main", "StartLanguageID", "0409", strSysIni
        IniWriteStrPrivate "Main", "DateFormatRus", "1", strSysIni
        IniWriteStrPrivate "Main", "CheckAllGroup", "1", strSysIni
        IniWriteStrPrivate "Main", "ListOnlyGroup", "1", strSysIni
        IniWriteStrPrivate "Main", "BlockListOnBackup", "1", strSysIni
        IniWriteStrPrivate "Main", "ArchMode", "0", strSysIni

        '������ Debug
        IniWriteStrPrivate "Debug", "DebugEnable", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogPath", "%WINDIR%\Logs\" & strProjectName & "Log\", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogName", strProjectName & "-LOG_%DATE%.txt", strSysIni
        IniWriteStrPrivate "Debug", "CleenHistory", "1", strSysIni
        IniWriteStrPrivate "Debug", "DetailMode", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLog2AppPath", "0", strSysIni
        IniWriteStrPrivate "Debug", "Time2File", "0", strSysIni
        '������ Arc
        IniWriteStrPrivate "Arc", "PathExe", "Tools\Arc\7za.exe", strSysIni
        IniWriteStrPrivate "Arc", "PathExe64", "Tools\Arc\7za64.exe", strSysIni
        IniWriteStrPrivate "Arc", "CompressParam1", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 *.ini -ir!*.inf", strSysIni
        IniWriteStrPrivate "Arc", "CompressParam2", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 -xr!*.inf -x!*.ini", strSysIni
        IniWriteStrPrivate "Arc", "PathSFX", "Tools\Arc\sfx\7zSD.sfx", strSysIni
        IniWriteStrPrivate "Arc", "PathSFXConfig", "Tools\Arc\sfx\config.txt", strSysIni
        IniWriteStrPrivate "Arc", "PathSFXConfigEn", "Tools\Arc\sfx\config_en.txt", strSysIni
        '[ARCName]
        IniWriteStrPrivate "ARCName", "StartMode", "1", strSysIni
        IniWriteStrPrivate "ARCName", "CustomName", "DP_%PCMODEL%_%OSVer%_%OSBit%_%DATE%", strSysIni
        'Folder=DP_%COMPUTERNAME%_%OS_Ver%_%OS_Bit%_%DATE%
        '7z=DP_%COMPUTERNAME%_%OS_Ver%_%OS_Bit%_%DATE%
        '7z-sfx=DriverAutoInstaller_%COMPUTERNAME%_%OS_Ver%_%OS_Bit%_%DATE%

        '������ DPInst
        IniWriteStrPrivate "DPInst", "PathExe", "Tools\DPInst\DPInst.exe", strSysIni
        IniWriteStrPrivate "DPInst", "PathExe64", "Tools\DPInst\DPInst64.exe", strSysIni
        'IniWriteStrPrivate "DPInst", "LegacyMode", 1, strSysIni
        'IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", 1, strSysIni
        'IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "SuppressWizard", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "QuietInstall", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "ScanHardware", 1, strSysIni

        '������ OS
        IniWriteStrPrivate "OS", "OSCount", "4", strSysIni
        '������ OS_1
        IniWriteStrPrivate "OS_1", "Ver", "5.0;5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_1", "drpFolder", "drivers\2k_xp_2003\x32\", strSysIni
        IniWriteStrPrivate "OS_1", "is64bit", "0", strSysIni
        '������ OS_2
        IniWriteStrPrivate "OS_2", "Ver", "5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_2", "drpFolder", "drivers\2k_xp_2003\x64\", strSysIni
        IniWriteStrPrivate "OS_2", "is64bit", "1", strSysIni

        '������ OS_3
        IniWriteStrPrivate "OS_3", "Ver", "6.0;6.1;6.2;6.3;6.4;10.0", strSysIni
        IniWriteStrPrivate "OS_3", "drpFolder", "drivers\vista_7_8_10\x32\", strSysIni
        IniWriteStrPrivate "OS_3", "is64bit", "0", strSysIni

        '������ OS_4
        IniWriteStrPrivate "OS_4", "Ver", "6.0;6.1;6.2;6.3;6.4;10.0", strSysIni
        IniWriteStrPrivate "OS_4", "drpFolder", "drivers\vista_7_8_10\x64\", strSysIni
        IniWriteStrPrivate "OS_4", "is64bit", "1", strSysIni
        '������ MainForm
        IniWriteStrPrivate "MainForm", "Width", CStr(lngMainFormWidthDef), strSysIni
        IniWriteStrPrivate "MainForm", "Height", CStr(lngMainFormHeightDef), strSysIni
        IniWriteStrPrivate "MainForm", "StartMaximazed", "0", strSysIni
        IniWriteStrPrivate "MainForm", "SaveSizeOnExit", "0", strSysIni
        IniWriteStrPrivate "MainForm", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "MainForm", "FontSize", "9", strSysIni
        IniWriteStrPrivate "MainForm", "HighlightColor", "32896", strSysIni
        
        '������ Buttons
        IniWriteStrPrivate "Button", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "Button", "FontSize", "9", strSysIni
        IniWriteStrPrivate "Button", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Button", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Button", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Button", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Button", "FontColor", "0", strSysIni
        IniWriteStrPrivate "Button", "Style", "8", strSysIni
        IniWriteStrPrivate "Button", "StyleColor", "2", strSysIni
        IniWriteStrPrivate "Button", "BackColor", "14933984", strSysIni

        ' �������� Ini ���� � ������������ ����
        NormIniFile strSysIni
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub GetMainIniParam
'! Description (��������)  :   [��������� �������� �� ��� �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function GetMainIniParam() As Boolean

    Dim ii         As Long
    Dim cntOsInIni As Integer


    GetMainIniParam = True
    
    '[Description]
    strThisBuildBy = GetIniValueString(strSysIni, "Description", "BuildBy", vbNullString)
    'strThisBuildBy = "www.SamLab.Ws"
    '[Debug]
    ' ���� �� ��� �����
    strDebugLogPathTemp = GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%WINDIR%\Logs\" & strProjectName & "Log\")
    strDebugLogPath = PathCollect(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%WINDIR%\Logs\" & strProjectName & "Log\"))
    ' ��� ���-�����
    strDebugLogNameTemp = GetIniValueString(strSysIni, "Debug", "DebugLogName", strProjectName & "-LOG_%DATE%.txt")
    strDebugLogName = ExpandFileNameByEnvironment(GetIniValueString(strSysIni, "Debug", "DebugLogName", strProjectName & "-LOG_%DATE%.txt"))
    ' ���������� ����� � ���-����
    mbDebugTime2File = GetIniValueBoolean(strSysIni, "Debug", "Time2File", 0)
    ' ��������� ���-���� � �������� "logs" ���������
    mbDebugLog2AppPath = GetIniValueBoolean(strSysIni, "Debug", "DebugLog2AppPath", 0)
    ' ��������� �������
    mbDebugStandart = GetIniValueBoolean(strSysIni, "Debug", "DebugEnable", 0)
    ' ������� �������
    mbCleanHistory = GetIniValueBoolean(strSysIni, "Debug", "CleenHistory", 1)

    If Not mbDebugLog2AppPath Then
        strDebugLogFullPath = PathCombine(strDebugLogPath, strDebugLogName)

        If mbDebugStandart Then
            If Not LogNotOnCDRoom Then
                If PathExists(strDebugLogPath) = False Then
                    CreateNewDirectory strDebugLogPath
                End If
            Else
                mbDebugStandart = False
            End If
        End If

    Else
        strDebugLogFullPath = strAppPathBackSL & "Logs\" & strDebugLogName

        If mbDebugStandart Then
            If Not LogNotOnCDRoom(strAppPathBackSL) Then
                If PathExists(strAppPathBackSL & "logs\") = False Then
                    CreateNewDirectory strAppPathBackSL & "logs\"
                End If
            Else
                If Not LogNotOnCDRoom Then
                    If PathExists(strDebugLogPath) = False Then
                        CreateNewDirectory strDebugLogPath
                    End If
                    strDebugLogFullPath = PathCombine(strDebugLogPath, strDebugLogName)
                Else
                    mbDebugStandart = False
                End If
            End If
        End If
    End If
    
    ' ����������� ������� - �� ���������=1
    lngDetailMode = GetIniValueLong(strSysIni, "Debug", "DetailMode", 1)
    If lngDetailMode < 1 Then
        lngDetailMode = 1
    End If
    If mbDebugStandart Then
        If lngDetailMode > 1 Then
            mbDebugDetail = True
        End If
    End If

    '[Main]
    ' �������� ��� ������
    mbDelTmpAfterClose = GetIniValueBoolean(strSysIni, "Main", "DelTmpAfterClose", 1)
    ' �������� ���������� ��� ������ (������ MAIN)
    mbUpdateCheck = GetIniValueBoolean(strSysIni, "Main", "UpdateCheck", 1)
    ' �������� ���������� ��� ������ (������ MAIN)
    mbUpdateCheckBeta = GetIniValueBoolean(strSysIni, "Main", "UpdateCheckBeta", 1)
    ' �������� EULA
    mbEULAAgree = GetIniValueBoolean(strSysIni, "Main", "EULAAgree", 0)
    ' ��������������� �����
    mbAutoLanguage = GetIniValueBoolean(strSysIni, "Main", "AutoLanguage", 1)

    If Not mbAutoLanguage Then
        strStartLanguageID = IniStringPrivate("Main", "StartLanguageID", strSysIni)
    End If

    ' ��������� ��������������� ���� Temp
    strAlternativeTempPath = IniStringPrivate("Main", "AlternativeTempPath", strSysIni)

    If StrComp(strAlternativeTempPath, "no_key") = 0 Then
        strAlternativeTempPath = strWinTemp
    End If

    ' ��� ������������� ���������� �������������� temp
    mbTempPath = GetIniValueBoolean(strSysIni, "Main", "AlternativeTemp", 0)

    If mbTempPath Then
        strAlternativeTempPath = PathCollect(strAlternativeTempPath)
        If mbDebugStandart Then DebugMode "AlternativeTempPath: " & strAlternativeTempPath

        If PathExists(strAlternativeTempPath) Then
            strWinTemp = strAlternativeTempPath
            strWorkTemp = strWinTemp & strProjectName

            ' ���� ���, �� ������� ��������� ������� �������
            If PathExists(strWorkTemp) = False Then
                CreateNewDirectory strWorkTemp
            End If

        Else
            If mbDebugStandart Then DebugMode "Alternative TempPath not Exist. Use Windows Temp"
        End If
    End If

    mbLoadIniTmpAfterRestart = GetIniValueBoolean(strSysIni, "Main", "LoadIniTmpAfterRestart", 0)
    mbDisableDEP = GetIniValueBoolean(strSysIni, "Main", "DisableDEP", 1)
    '--------------------- ��������� ����� �� ������ ---------------------
    '[DPInst]
    ' DPInst.exe
    strDPInstExePath86 = IniStringPrivate("DPInst", "PathExe", strSysIni)

    If InStr(strDPInstExePath86, strColon) Then
        mbPatnAbs = True
    End If

    strDPInstExePath86 = PathCollect(strDPInstExePath86)

    If FileExists(strDPInstExePath86) = False Then
        strDPInstExePath86 = strAppPathBackSL & "Tools\DPInst\DPInst.exe"

        If FileExists(strDPInstExePath86) = False Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath86, vbInformation, strProductName
        End If
    End If

    strDPInstExePath = strDPInstExePath86
    ' DPInst64.exe
    strDPInstExePath64 = IniStringPrivate("DPInst", "PathExe64", strSysIni)

    If InStr(strDPInstExePath64, strColon) Then
        mbPatnAbs = True
    End If

    strDPInstExePath64 = PathCollect(strDPInstExePath64)

    If FileExists(strDPInstExePath64) = False Then
        strDPInstExePath64 = strAppPathBackSL & "Tools\DPInst\DPInst64.exe"

        If FileExists(strDPInstExePath64) = False Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath64, vbInformation, strProductName
        End If
    End If

    ' ��������� DpInst
    mbDpInstLegacyMode = GetIniValueBoolean(strSysIni, "DPInst", "LegacyMode", 1)
    mbDpInstPromptIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "PromptIfDriverIsNotBetter", 1)
    mbDpInstForceIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "ForceIfDriverIsNotBetter", 0)
    mbDpInstSuppressAddRemovePrograms = GetIniValueBoolean(strSysIni, "DPInst", "SuppressAddRemovePrograms", 0)
    mbDpInstSuppressWizard = GetIniValueBoolean(strSysIni, "DPInst", "SuppressWizard", 0)
    mbDpInstQuietInstall = GetIniValueBoolean(strSysIni, "DPInst", "QuietInstall", 0)
    mbDpInstScanHardware = GetIniValueBoolean(strSysIni, "DPInst", "ScanHardware", 1)
    
    '[Arc]
    ' 7za.exe
    strArh7zExePath86 = IniStringPrivate("Arc", "PathExe", strSysIni)

    If InStr(strArh7zExePath86, strColon) Then
        mbPatnAbs = True
    End If

    strArh7zExePath86 = PathCollect(strArh7zExePath86)

    If FileExists(strArh7zExePath86) = False Then
        strArh7zExePath86 = strAppPathBackSL & "Tools\Arc\7za.exe"

        If FileExists(strArh7zExePath86) = False Then
            MsgBox strMessages(7) & vbNewLine & strArh7zExePath86, vbInformation, strProductName
        End If
    End If

    strArh7zExePath = strArh7zExePath86
    ' 7za.exe - x64
    strArh7zExePath64 = IniStringPrivate("Arc", "PathExe64", strSysIni)

    If InStr(strArh7zExePath64, strColon) Then
        mbPatnAbs = True
    End If
    
    strArh7zExePath64 = PathCollect(strArh7zExePath64)

    If FileExists(strArh7zExePath64) = False Then
        strArh7zExePath64 = strAppPathBackSL & "Tools\Arc\7za64.exe"

        If FileExists(strArh7zExePath64) = False Then
            MsgBox strMessages(7) & vbNewLine & strArh7zExePath64, vbInformation, strProductName
            If mbDebugStandart Then DebugMode "7zExePath64: " & " Not exist. Get from x86 - " & strArh7zExePath86
            strArh7zExePath64 = strArh7zExePath86
        End If
    End If

    strArh7zParam1 = GetIniValueString(strSysIni, "Arc", "CompressParam1", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 *.ini -ir!*.inf")
    strArh7zParam2 = GetIniValueString(strSysIni, "Arc", "CompressParam2", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 -xr!*.inf -x!*.ini")
    ' 7zSD.sfx
    strArh7zSFXPATH = IniStringPrivate("Arc", "PathSFX", strSysIni)
    strArh7zSFXPATH = PathCollect(strArh7zSFXPATH)

    If PathExists(strArh7zSFXPATH) = False Then
        strArh7zSFXPATH = strAppPath & "\Tools\Arc\sfx\7zSD.sfx"

        If PathExists(strArh7zSFXPATH) = False Then
            MsgBox strMessages(7) & vbNewLine & strArh7zSFXPATH, vbInformation, strProductName
        End If
    End If

    ' config.txt
    strArh7zSFXConfigPath = IniStringPrivate("Arc", "PathSFXConfig", strSysIni)
    strArh7zSFXConfigPath = PathCollect(strArh7zSFXConfigPath)

    If PathExists(strArh7zSFXConfigPath) = False Then
        strArh7zSFXConfigPath = strAppPath & "\Tools\Arc\sfx\config.txt"

        If PathExists(strArh7zSFXConfigPath) = False Then
            MsgBox strMessages(7) & vbNewLine & strArh7zSFXConfigPath, vbInformation, strProductName
        End If
    End If

    ' config_en.txt
    strArh7zSFXConfigPathEn = IniStringPrivate("Arc", "PathSFXConfigEn", strSysIni)
    strArh7zSFXConfigPathEn = PathCollect(strArh7zSFXConfigPathEn)

    If PathExists(strArh7zSFXConfigPathEn) = False Then
        strArh7zSFXConfigPathEn = strAppPath & "\Tools\Arc\sfx\config_en.txt"

        If PathExists(strArh7zSFXConfigPathEn) = False Then
            MsgBox strMessages(7) & vbNewLine & strArh7zSFXConfigPathEn, vbInformation, strProductName
        End If
    End If

    '[ARCName]
    lngArchNameMode = GetIniValueLong(strSysIni, "ARCName", "StartMode", 1)
    strArchNameCustom = GetIniValueString(strSysIni, "ARCName", "CustomName", "DP_%PCMODEL%_%OSVer%_%OSBit%_%DATE%")

    '[MainForm]
    ' ��������� ��������� ��� ������
    mbSaveSizeOnExit = GetIniValueBoolean(strSysIni, "MainForm", "SaveSizeOnExit", 0)
    '������ �������� �����
    lngMainFormWidth = GetIniValueLong(strSysIni, "MainForm", "Width", lngMainFormWidthDef)

    '���� ���������� �������� ������ ������������, �� ������������� �������� �� ���������
    If lngMainFormWidth < lngMainFormWidthMin Then
        lngMainFormWidth = lngMainFormWidthDef
    End If

    '������ �������� �����
    lngMainFormHeight = GetIniValueLong(strSysIni, "MainForm", "Height", lngMainFormHeightDef)

    '���� ���������� �������� ������ ������������, �� ������������� �������� �� ���������
    If lngMainFormHeight < lngMainFormHeightMin Then
        lngMainFormHeight = lngMainFormHeightDef
    End If

    ' ��������� ���� ������� (������ MainForm)
    mbStartMaximazed = GetIniValueBoolean(strSysIni, "MainForm", "StartMaximazed", 0)
    strFontMainForm_Name = GetIniValueString(strSysIni, "MainForm", "FontName", "Tahoma")
    lngFontMainForm_Size = GetIniValueLong(strSysIni, "MainForm", "FontSize", 8)
    ' ��������� ��������� ��������
    glHighlightColor = GetIniValueLong(strSysIni, "MainForm", "HighlightColor", 32896)
    ' ��������� ���� ������� (������ OtherForm)
    strFontOtherForm_Name = GetIniValueString(strSysIni, "OtherForm", "FontName", "Tahoma")
    lngFontOtherForm_Size = GetIniValueLong(strSysIni, "OtherForm", "FontSize", 8)


    '[OS]
    ' ��������� ���-�� ������ (������ OS) � ���������� ������� ��
    lngOSCount = IniLongPrivate("OS", "OSCount", strSysIni)

    If lngOSCount = 0 Or lngOSCount = 9999 Then
        If mbDebugStandart Then DebugMode "The List supported operating systems is empty. PreDefine BackUpfolder not accessible"
        mbBackFolderPredefine = False
    Else
        ReDim arrOSList(lngOSCount - 1)

        For ii = 0 To UBound(arrOSList)
            cntOsInIni = ii + 1
            arrOSList(ii).Ver = IniStringPrivate("OS_" & cntOsInIni, "Ver", strSysIni)
            arrOSList(ii).is64bit = IniLongPrivate("OS_" & cntOsInIni, "is64bit", strSysIni)

            If arrOSList(ii).is64bit = 9999 Then
                arrOSList(ii).is64bit = 0
            End If

            arrOSList(ii).drpFolder = IniStringPrivate("OS_" & cntOsInIni, "drpFolder", strSysIni)

            If arrOSList(ii).drpFolder <> "No Key" Then
                If PathExists(PathCollect(arrOSList(ii).drpFolder)) = False Then
                    If mbDebugStandart Then DebugMode "Not find folder for package driver backup" & vbNewLine & "aey IN: " & arrOSList(ii).Ver & " is64bit:" & arrOSList(ii).is64bit & vbNewLine & vbNewLine & "Folder is not Exist: " & vbNewLine & PathCollect(arrOSList(ii).drpFolder)
                    arrOSList(ii).DPFolderNotExist = True
                End If

            Else
                If mbDebugStandart Then DebugMode "Folder with package driver" & vbNewLine & "for OS: " & arrOSList(ii).Ver & " is64bit:" & arrOSList(ii).is64bit & vbNewLine & "Is Not present in options. Correct and start the program again."
            End If

        Next
        mbBackFolderPredefine = True
    End If

    '[Button]
    ' ����� ������
    strFontBtn_Name = GetIniValueString(strSysIni, "Button", "FontName", "Tahoma")
    miFontBtn_Size = GetIniValueLong(strSysIni, "Button", "FontSize", 8)
    mbFontBtn_Bold = GetIniValueBoolean(strSysIni, "Button", "FontBold", 0)
    mbFontBtn_Italic = GetIniValueBoolean(strSysIni, "Button", "FontItalic", 0)
    mbFontBtn_Underline = GetIniValueBoolean(strSysIni, "Button", "FontUnderline", 0)
    mbFontBtn_Strikethru = GetIniValueBoolean(strSysIni, "Button", "FontStrikethru", 0)
    lngFontBtn_Color = GetIniValueLong(strSysIni, "Button", "FontColor", 0)
    lngStatusBtnStyle = GetIniValueLong(strSysIni, "Button", "Style", "8")
    lngStatusBtnStyleColor = GetIniValueLong(strSysIni, "Button", "StyleColor", "2")
    lngStatusBtnBackColor = GetIniValueLong(strSysIni, "Button", "BackColor", "14933984")

    '[Main]
    ' ���������� ���� ������ � ������� dd/mm/yyyy
    mbDateFormatRus = GetIniValueBoolean(strSysIni, "Main", "DateFormatRus", 0)
    ' ����� �� ������
    strImageMainName = GetIniValueString(strSysIni, "Main", "IconMainSkin", "Standart")
    ' ����������� ����
    'mbExMenu = GetIniValueBoolean(strSysIni, "Main", "ExMenu", 1)
    'strImageMenuName = GetIniValueString(strSysIni, "Main", "IconMenuSkin", "Standart")
    ' �������� ������ ��������
    mbHideOtherProcess = GetIniValueBoolean(strSysIni, "Main", "HideOtherProcess", 1)
    ' ����� ����������� DLL
    mbSilentDLL = GetIniValueBoolean(strSysIni, "Main", "SilentDll", 0)
    ' ���������� ����������� �� ���������� (����������� ����)
    mbUpdateToolTip = GetIniValueBoolean(strSysIni, "Main", "UpdateToolTip", 1)

    ' ��������� �����
    miStartMode = GetIniValueLong(strSysIni, "Main", "StartMode", 2)
    ' �������� ��� ������
    mbCheckAllGroup = GetIniValueBoolean(strSysIni, "Main", "CheckAllGroup", 1)
    ' �������� ��� ������
    mbListOnlyGroup = GetIniValueBoolean(strSysIni, "Main", "ListOnlyGroup", 1)
    '������������ ���� listview ghb �������������
    mbBlockListOnBackup = GetIniValueBoolean(strSysIni, "Main", "BlockListOnBackup", 1)
    '����� ������������� �� ���������
    miArchMode = GetIniValueLong(strSysIni, "Main", "ArchMode", 0)

    Exit Function
    
ExitFunc:
    GetMainIniParam = False
End Function

