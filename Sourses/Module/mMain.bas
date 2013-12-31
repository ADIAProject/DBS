Attribute VB_Name = "mMain"
Option Explicit

' �������� ��������� ���������
Public Const strDateProgram                 As String = "21/09/2012"

' ������� ������
Public arrHwidsLocal()                      As String

' �������������� ������������ ��� �������� ��������
Public mboolDateFormatRus                   As Boolean

' ���������� �������� ���������
Public strProductName                       As String
Public strProductVersion                    As String

' ������� ���� ��������
Public strSysIni                            As String
Public mboolLoadIniTmpAfterRestart          As Boolean
Public strWorkTemp                          As String
Public strWinTemp                           As String
Public strWinDir                            As String
Public strSysDir                            As String
Public strSysDir64                          As String
Public strSysDir86                          As String
Public strSysDirCatRoot                     As String
Public strSysDirDrivers                     As String
Public strSysDirDrivers64                   As String
Public strSysDirDRVStore                    As String
Public strSysDrive                          As String
Public strWinDirHelp                        As String
Public strInfDir                            As String
Public mboolLogNotOnCDRoom                  As Boolean
Public mboolHideOtherProcess                As Boolean
Public mboolDelTmpAfterClose                As Boolean
Public mboolUpdateCheck                     As Boolean
Public mboolUpdateCheckBeta                 As Boolean
Public mboolUpdateToolTip                   As Boolean
Public mboolIsDesignMode                    As Boolean
Public mboolIsDriveCDRoom                   As Boolean
Public strArh7zExePATH                      As String
Public strArh7zParam1                       As String
Public strArh7zParam2                       As String
Public strArh7zSFXPATH                      As String
Public strArh7zSFXConfigPath                As String
Public strArh7zSFXConfigPathEn              As String

'����� ������ � ��������� listview - ���� �������� ���� ����������
Public mboolAddInList                       As Boolean

'����� ���������� �������� � ������ ��
Public LastIdOS                             As Long

'������ ����������� ���������
Public mboolRestartProgram                  As Boolean
Public mboolStartMaximazed                  As Boolean
Public strDPInstExePath                     As String
Public strDPInstExePath64                   As String
Public strDPInstExePath86                   As String

' ��������� DPinst
Public mboolDpInstLegacyMode                As Boolean
Public mboolDpInstPromptIfDriverIsNotBetter As Boolean
Public mboolDpInstForceIfDriverIsNotBetter  As Boolean
Public mboolDpInstSuppressAddRemovePrograms As Boolean
Public mboolDpInstSuppressWizard            As Boolean
Public mboolDpInstQuietInstall              As Boolean
Public mboolDpInstScanHardware              As Boolean
Public mboolCalculateHashMode               As Boolean
Public strImageMainName                     As String
Public mboolSilentDLL                       As Boolean

' ����������� ����
'Public mboolExMenu                              As Boolean
Public strImageMenuName                     As String

'������ ��������� ���������
Public mboolIsWin64                         As Boolean
Public mboolFirstStart                      As Boolean

' ����� �������� ����� � ������ ���������
Public strMainForm_FontName                 As String
Public lngMainForm_FontSize                 As Long

' ����� ������ ����
Public strOtherForm_FontName                As String
Public lngOtherForm_FontSize                As Long

' ������ � ���������� �������
Public mboolRunWithParam                    As Boolean

''Private strRunWithParam              As String
' ������������ �������������?
Private mboolIsUserAnAdmin                  As Boolean

Public Const strDONATE_MD5RTF               As String = "97f8178b2af5ba9377f76baf4ff71f78"
Public Const strDONATE_MD5RTF_Eng           As String = "59bbfbf6decbf91023da434cbe940d33"

' ����� �������� �����
Public strFrmMainCaptionTemp                As String
Public strFrmMainCaptionTempDate            As String

'-------------------- ���������� �������� ����� � ������ ------------------'
Public MainFormWidth                        As Long
Public MainFormHeight                       As Long

' ����������� �������� �������� �����
Public Const MainFormWidthMin               As Long = 12700
Public Const MainFormHeightMin              As Long = 6000

' ��������� �������� �������� �����
Private Const MainFormWidthDef              As Long = 12700
Private Const MainFormHeightDef             As Long = 8000

Public mboolSaveSizeOnExit                  As Boolean
Public mboolCheckAllGroup                   As Boolean
Public mboolListOnlyGroup                   As Boolean
Public miStartMode                          As Long
Public miArchMode                           As Long
Public arrOSList()                          As String
Public OSCount                              As Long
Public mboolBackFolderPredefine             As Boolean
Public mboolBlockListOnBackup               As Boolean

' ��������� �������� %Temp%
Public mboolTempPath          As Boolean
Public strAlternativeTempPath As String
Public mboolPatnAbs           As Boolean
Public strCompName            As String
Public strMB_Model            As String
Public strMB_Manufacturer     As String
Public strCompModel           As String
Public lngArchNameMode        As Long
Public strArchNameCustom      As String

' ���������� ��� ����������� ���������� DEP
Public mboolDisableDEP        As Boolean

'! -----------------------------------------------------------
'!  �������     :  ChangeStatusTextAndDebug
'!  ����������  :  Optional strSimpleText As String, Optional strDebugText As String
'!  ��������    :  ��������� ������ ���������� ������ � ���������� ����������
'! -----------------------------------------------------------
Public Sub ChangeStatusTextAndDebug(Optional strPanel2Text As String, _
                                    Optional strDebugText As String, _
                                    Optional ByVal mboolEqual As Boolean = False, _
                                    Optional ByVal mboolDoEvents As Boolean = True, _
                                    Optional strPanel1Text As String)

    If LenB(strPanel2Text) > 0 Then
        If mboolDoEvents Then
            DoEvents
        End If

        'frmMain.sbStatusBar.SimpleText = strSimpleText
        If frmMain.ctlUcStatusBar1.PanelCount >= 2 Then
            frmMain.ctlUcStatusBar1.PanelText(2) = strPanel2Text
        Else
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel2Text
        End If

        If LenB(strPanel1Text) > 0 Then
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel1Text
        End If
    End If

    If LenB(strDebugText) > 0 Then
        If mboolEqual Then
            If LenB(strPanel1Text) > 0 Then
                DebugMode strPanel1Text & ": " & strPanel2Text
            Else
                DebugMode strPanel2Text
            End If

        Else
            DebugMode strDebugText
        End If

    Else

        If mboolEqual Then
            If LenB(strPanel1Text) > 0 Then
                DebugMode strPanel1Text & ": " & strPanel2Text
            Else
                DebugMode strPanel2Text
            End If
        End If
    End If
End Sub

'! -----------------------------------------------------------
'!  �������     :  CreateIni
'!  ����������  :
'!  ��������    :  ���������� �������� � ��� ���� ���� ����� ���
'! -----------------------------------------------------------
Private Sub CreateIni()

    If PathFileExists(strSysIni) = 0 Then
        If mboolIsDriveCDRoom Then
            strSysIni = strWorkTemp & "\DriversBackuper.ini"
            MsgBox "File DriversBackuper.ini is not Exist!" & vbNewLine & "This program works from CD\DVD, so we create temporary DriversBackuper.ini-file" & vbNewLine & strSysIni, vbInformation + vbApplicationModal, strProductName
        End If

        '������ Main
        IniWriteStrPrivate "Main", "DelTmpAfterClose", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheck", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheckBeta", "0", strSysIni
        IniWriteStrPrivate "Main", "HideOtherProcess", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTemp", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTempPath", "%Temp%", strSysIni
        IniWriteStrPrivate "Main", "AutoLanguage", "1", strSysIni
        IniWriteStrPrivate "Main", "StartLanguageID", "0409", strSysIni
        IniWriteStrPrivate "Main", "IconMainSkin", "Standart", strSysIni
        IniWriteStrPrivate "Main", "SilentDLL", "0", strSysIni
        IniWriteStrPrivate "Main", "DateFormatRus", "1", strSysIni
        IniWriteStrPrivate "Main", "CheckAllGroup", "1", strSysIni
        IniWriteStrPrivate "Main", "ListOnlyGroup", "1", strSysIni
        IniWriteStrPrivate "Main", "StartMode", "2", strSysIni
        IniWriteStrPrivate "Main", "BlockListOnBackup", "1", strSysIni
        IniWriteStrPrivate "Main", "CalculateHashMode", "1", strSysIni
        IniWriteStrPrivate "Main", "ArchMode", "0", strSysIni
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", "0", strSysIni
        '������ Debug
        IniWriteStrPrivate "Debug", "DebugEnable", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogPath", "C:\debuglog_DBS.txt", strSysIni
        IniWriteStrPrivate "Debug", "CleenHistory", "1", strSysIni
        IniWriteStrPrivate "Debug", "DetailMode", "1", strSysIni
        '������ Arc
        IniWriteStrPrivate "Arc", "PathExe", "Tools\Arc\7za.exe", strSysIni
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
        '������ MainForm
        IniWriteStrPrivate "MainForm", "Width", CStr(MainFormWidthDef), strSysIni
        IniWriteStrPrivate "MainForm", "Height", CStr(MainFormHeightDef), strSysIni
        IniWriteStrPrivate "MainForm", "StartMaximazed", "0", strSysIni
        IniWriteStrPrivate "MainForm", "SaveSizeOnExit", "0", strSysIni
        IniWriteStrPrivate "MainForm", "FontName", "Lucida Console", strSysIni
        IniWriteStrPrivate "MainForm", "FontSize", "8", strSysIni
        '������ Buttons
        IniWriteStrPrivate "Button", "FontName", "Arial Unicode MS", strSysIni
        IniWriteStrPrivate "Button", "FontSize", "8", strSysIni
        IniWriteStrPrivate "Button", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Button", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Button", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Button", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Button", "FontColor", "0", strSysIni
        '������ OS
        IniWriteStrPrivate "OS", "OSCount", "4", strSysIni
        '������ OS_1
        IniWriteStrPrivate "OS_1", "Ver", "5.0;5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_1", "is64bit", "0", strSysIni
        IniWriteStrPrivate "OS_1", "drpFolder", "drivers\2k_xp_2003\x32\", strSysIni
        '������ OS_1
        IniWriteStrPrivate "OS_2", "Ver", "5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_2", "is64bit", "1", strSysIni
        IniWriteStrPrivate "OS_2", "drpFolder", "drivers\2k_xp_2003\x64\", strSysIni
        '������ OS_3
        IniWriteStrPrivate "OS_3", "Ver", "6.0;6.1;6.2", strSysIni
        IniWriteStrPrivate "OS_3", "is64bit", "0", strSysIni
        IniWriteStrPrivate "OS_3", "drpFolder", "drivers\vista_7_8\x32\", strSysIni
        '������ OS_4
        IniWriteStrPrivate "OS_4", "Ver", "6.0;6.1;6.2", strSysIni
        IniWriteStrPrivate "OS_4", "is64bit", "1", strSysIni
        IniWriteStrPrivate "OS_4", "drpFolder", "drivers\vista_7_8\x64\", strSysIni
        ' �������� Ini ���� � ������������ ����
        NormIniFile strSysIni
        ' ��������� ������� ����� �������� ini-�����
        mboolDebugEnable = True
        mboolCleanHistory = True
        strDebugLogPath = "C:\debuglog_DBS.txt"
    End If
End Sub

'! -----------------------------------------------------------
'!  �������     :  GetMainIniParam
'!  ����������  :
'!  ��������    :  ��������� �������� �� ��� �����
'! -----------------------------------------------------------
Private Sub GetMainIniParam()

    Dim i          As Long
    Dim cntOsInIni As Integer

    '[Debug]
    ' ��������� �������
    mboolDebugEnable = GetIniValueBoolean(strSysIni, "Debug", "DebugEnable", 1)
    ' ������� �������
    mboolCleanHistory = GetIniValueBoolean(strSysIni, "Debug", "CleenHistory", 1)
    ' ���� �� ��� �����
    strDebugLogPath = PathCollect(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "C:\debuglog_DBS.txt"))
    ' ����������� ������� - �� ���������=1
    lngDetailMode = GetIniValueLong(strSysIni, "Debug", "DetailMode", 1)

    If lngDetailMode < 1 Then
        lngDetailMode = 1
    ElseIf lngDetailMode > 2 Then
        lngDetailMode = 2
    End If

    '[Main]
    ' �������� ��� ������
    mboolDelTmpAfterClose = GetIniValueBoolean(strSysIni, "Main", "DelTmpAfterClose", 1)
    ' �������� ���������� ��� ������ (������ MAIN)
    mboolUpdateCheck = GetIniValueBoolean(strSysIni, "Main", "UpdateCheck", 1)
    ' �������� ���������� ��� ������ (������ MAIN)
    mboolUpdateCheckBeta = GetIniValueBoolean(strSysIni, "Main", "UpdateCheckBeta", 1)
    ' ��������������� �����
    mboolAutoLanguage = GetIniValueBoolean(strSysIni, "Main", "AutoLanguage", 1)

    If Not mboolAutoLanguage Then
        strStartLanguageID = IniStringPrivate("Main", "StartLanguageID", strSysIni)
    End If

    ' ���������� ���� ������ � ������� dd/mm/yyyy
    mboolDateFormatRus = GetIniValueBoolean(strSysIni, "Main", "DateFormatRus", 0)
    ' ����� �� ������
    strImageMainName = GetIniValueString(strSysIni, "Main", "IconMainSkin", "Standart")
    ' �������� ������ ��������
    mboolHideOtherProcess = GetIniValueBoolean(strSysIni, "Main", "HideOtherProcess", 1)
    ' ����� ����������� DLL
    mboolSilentDLL = GetIniValueBoolean(strSysIni, "Main", "SilentDll", 0)
    ' ���������� ����������� �� ���������� (����������� ����)
    mboolUpdateToolTip = GetIniValueBoolean(strSysIni, "Main", "UpdateToolTip", 1)
    ' �������� ��� ������
    mboolCheckAllGroup = GetIniValueBoolean(strSysIni, "Main", "CheckAllGroup", 1)
    ' �������� ��� ������
    mboolListOnlyGroup = GetIniValueBoolean(strSysIni, "Main", "ListOnlyGroup", 1)
    ' ��������� �����
    miStartMode = GetIniValueLong(strSysIni, "Main", "StartMode", 2)
    '������������ ���� listview ghb �������������
    mboolBlockListOnBackup = GetIniValueBoolean(strSysIni, "Main", "BlockListOnBackup", 1)
    ' ������������ ����� ������� ������� Hash-�����
    mboolCalculateHashMode = GetIniValueBoolean(strSysIni, "Main", "CalculateHashMode", 1)
    '����� ������������� �� ���������
    miArchMode = GetIniValueLong(strSysIni, "Main", "ArchMode", 0)
    ' ����������� ����
    'mboolExMenu = GetIniValueBoolean(strSysIni, "Main", "ExMenu", 1)
    'strImageMenuName = GetIniValueString(strSysIni, "Main", "IconMenuSkin", "Standart")
    mboolLoadIniTmpAfterRestart = GetIniValueBoolean(strSysIni, "Main", "LoadIniTmpAfterRestart", 0)
    mboolDisableDEP = GetIniValueBoolean(strSysIni, "Main", "DisableDEP", 1)
    '--------------------- ��������� ����� �� ������ ---------------------
    '[Arc]
    ' 7za.exe
    strArh7zExePATH = IniStringPrivate("Arc", "PathExe", strSysIni)

    If InStr(1, strArh7zExePATH, ":") > 0 Then
        mboolPatnAbs = True
    End If

    strArh7zExePATH = PathCollect(strArh7zExePATH)

    If PathFileExists(strArh7zExePATH) = 0 Then
        strArh7zExePATH = strAppPath & "\Tools\Arc\7za.exe"

        If PathFileExists(strArh7zExePATH) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strArh7zExePATH, vbInformation, strProductName
        End If
    End If

    strArh7zParam1 = GetIniValueString(strSysIni, "Arc", "CompressParam1", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 *.ini -ir!*.inf")
    strArh7zParam2 = GetIniValueString(strSysIni, "Arc", "CompressParam2", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 -xr!*.inf -x!*.ini")
    ' 7zSD.sfx
    strArh7zSFXPATH = IniStringPrivate("Arc", "PathSFX", strSysIni)
    strArh7zSFXPATH = PathCollect(strArh7zSFXPATH)

    If PathFileExists(strArh7zSFXPATH) = 0 Then
        strArh7zSFXPATH = strAppPath & "\Tools\Arc\sfx\7zSD.sfx"

        If PathFileExists(strArh7zSFXPATH) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strArh7zSFXPATH, vbInformation, strProductName
        End If
    End If

    ' config.txt
    strArh7zSFXConfigPath = IniStringPrivate("Arc", "PathSFXConfig", strSysIni)
    strArh7zSFXConfigPath = PathCollect(strArh7zSFXConfigPath)

    If PathFileExists(strArh7zSFXConfigPath) = 0 Then
        strArh7zSFXConfigPath = strAppPath & "\Tools\Arc\sfx\config.txt"

        If PathFileExists(strArh7zSFXConfigPath) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strArh7zSFXConfigPath, vbInformation, strProductName
        End If
    End If

    ' config_en.txt
    strArh7zSFXConfigPathEn = IniStringPrivate("Arc", "PathSFXConfigEn", strSysIni)
    strArh7zSFXConfigPathEn = PathCollect(strArh7zSFXConfigPathEn)

    If PathFileExists(strArh7zSFXConfigPathEn) = 0 Then
        strArh7zSFXConfigPathEn = strAppPath & "\Tools\Arc\sfx\config_en.txt"

        If PathFileExists(strArh7zSFXConfigPathEn) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strArh7zSFXConfigPathEn, vbInformation, strProductName
        End If
    End If

    '[DPInst]
    strDPInstExePath86 = IniStringPrivate("DPInst", "PathExe", strSysIni)

    If InStr(1, strDPInstExePath86, ":") > 0 Then
        mboolPatnAbs = True
    End If

    strDPInstExePath86 = PathCollect(strDPInstExePath86)

    If PathFileExists(strDPInstExePath86) = 0 Then
        strDPInstExePath86 = strAppPath & "\Tools\DPInst\DPInst.exe"

        If PathFileExists(strDPInstExePath86) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath86, vbInformation, strProductName
        End If
    End If

    strDPInstExePath = strDPInstExePath86
    ' DPInst64.exe
    strDPInstExePath64 = IniStringPrivate("DPInst", "PathExe64", strSysIni)

    If InStr(1, strDPInstExePath64, ":") > 0 Then
        mboolPatnAbs = True
    End If

    strDPInstExePath64 = PathCollect(strDPInstExePath64)

    If PathFileExists(strDPInstExePath64) = 0 Then
        strDPInstExePath64 = strAppPath & "\Tools\DPInst\DPInst64.exe"

        If PathFileExists(strDPInstExePath64) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath64, vbInformation, strProductName
        End If
    End If

    ' ��������� DpInst
    mboolDpInstLegacyMode = GetIniValueBoolean(strSysIni, "DPInst", "LegacyMode", 1)
    mboolDpInstPromptIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "PromptIfDriverIsNotBetter", 1)
    mboolDpInstForceIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "ForceIfDriverIsNotBetter", 0)
    mboolDpInstSuppressAddRemovePrograms = GetIniValueBoolean(strSysIni, "DPInst", "SuppressAddRemovePrograms", 0)
    mboolDpInstSuppressWizard = GetIniValueBoolean(strSysIni, "DPInst", "SuppressWizard", 0)
    mboolDpInstQuietInstall = GetIniValueBoolean(strSysIni, "DPInst", "QuietInstall", 0)
    mboolDpInstScanHardware = GetIniValueBoolean(strSysIni, "DPInst", "ScanHardware", 1)
    '[ARCName]
    lngArchNameMode = GetIniValueLong(strSysIni, "ARCName", "StartMode", 1)
    strArchNameCustom = GetIniValueString(strSysIni, "ARCName", "CustomName", "DP_%PCMODEL%_%OSVer%_%OSBit%_%DATE%")
    '[MainForm]
    ' ��������� ��������� ��� ������
    mboolSaveSizeOnExit = GetIniValueBoolean(strSysIni, "MainForm", "SaveSizeOnExit", 0)
    '������ �������� �����
    MainFormWidth = GetIniValueLong(strSysIni, "MainForm", "Width", MainFormWidthDef)

    '���� ���������� �������� ������ ������������, �� ������������� �������� �� ���������
    If MainFormWidth < MainFormWidthMin Then
        MainFormWidth = MainFormWidthDef
    End If

    '������ �������� �����
    MainFormHeight = GetIniValueLong(strSysIni, "MainForm", "Height", MainFormHeightDef)

    '���� ���������� �������� ������ ������������, �� ������������� �������� �� ���������
    If MainFormHeight < MainFormHeightMin Then
        MainFormHeight = MainFormHeightDef
    End If

    ' ��������� ���� ������� (������ MainForm)
    mboolStartMaximazed = GetIniValueBoolean(strSysIni, "MainForm", "StartMaximazed", 0)
    strMainForm_FontName = GetIniValueString(strSysIni, "MainForm", "FontName", "Arial Unicode MS")
    lngMainForm_FontSize = GetIniValueLong(strSysIni, "MainForm", "FontSize", 8)
    ' ��������� ��������� ��������
    glHighlightColor = GetIniValueLong(strSysIni, "MainForm", "HighlightColor", 32896)
    ' ��������� ���� ������� (������ OtherForm)
    strOtherForm_FontName = GetIniValueString(strSysIni, "OtherForm", "FontName", "Arial Unicode MS")
    lngOtherForm_FontSize = GetIniValueLong(strSysIni, "OtherForm", "FontSize", 8)
    ' ��������� ��������������� ���� Temp
    strAlternativeTempPath = IniStringPrivate("Main", "AlternativeTempPath", strSysIni)

    If strAlternativeTempPath = "No Key" Then
        strAlternativeTempPath = strWinTemp
    End If

    ' ��� ������������� ���������� �������������� temp
    mboolTempPath = GetIniValueBoolean(strSysIni, "Main", "AlternativeTemp", 0)

    If mboolTempPath Then
        'strAlternativeTempPath = IniStringPrivate("Main", "AlternativeTempPath", strSysIni)
        strAlternativeTempPath = PathCollect(strAlternativeTempPath)
        DebugMode "AlternativeTempPath: " & strAlternativeTempPath

        If PathFileExists(strAlternativeTempPath) = 1 Then
            strWinTemp = strAlternativeTempPath
            strWorkTemp = strWinTemp & "DriversInstaller"

            ' ������� ��������� ������� �������
            If PathFileExists(strWorkTemp) = 0 Then
                'MkDir (strWorkTemp)
                CreateNewDirectory strWorkTemp
            End If

        Else
            DebugMode "Alternative TempPath not Exist. Use Windows Temp"
        End If
    End If

    '[OS]
    ' ��������� ���-�� ������ (������ OS) � ���������� ������� ��
    OSCount = IniLongPrivate("OS", "OSCount", strSysIni)

    If OSCount = 0 Or OSCount = 9999 Then
        DebugMode "The List supported operating systems is empty. PreDefine BackUpfolder not accessible"
        mboolBackFolderPredefine = False
    Else
        ReDim arrOSList(OSCount - 1, 4)

        For i = 0 To UBound(arrOSList, 1)
            cntOsInIni = i + 1
            arrOSList(i, 0) = IniStringPrivate("OS_" & cntOsInIni, "Ver", strSysIni)
            arrOSList(i, 1) = IniLongPrivate("OS_" & cntOsInIni, "is64bit", strSysIni)

            If arrOSList(i, 1) = 9999 Then
                arrOSList(i, 1) = 0
            End If

            arrOSList(i, 2) = IniStringPrivate("OS_" & cntOsInIni, "drpFolder", strSysIni)

            If arrOSList(i, 2) <> "No Key" Then
                If PathFileExists(PathCollect(arrOSList(i, 2))) = 0 Then
                    DebugMode "Not find folder for package driver backup" & vbNewLine & "��� ��: " & arrOSList(i, 0) & " is64bit:" & arrOSList(i, 1) & vbNewLine & vbNewLine & "Folder is not Exist: " & vbNewLine & PathCollect(arrOSList(i, 2))
                    arrOSList(i, 3) = "DriverPack folder is not Exist"
                End If

            Else
                DebugMode "Folder with package driver" & vbNewLine & "for OS: " & arrOSList(i, 0) & " is64bit:" & arrOSList(i, 1) & vbNewLine & "Is Not present in options. Correct and start the program again."
            End If

        Next
        mboolBackFolderPredefine = True
    End If

    '[Button]
    ' ����� ������
    strDialog_FontName = GetIniValueString(strSysIni, "Button", "FontName", "Arial Unicode MS")
    miDialog_FontSize = GetIniValueLong(strSysIni, "Button", "FontSize", 8)
    mboolDialog_Bold = GetIniValueBoolean(strSysIni, "Button", "FontBold", 0)
    mboolDialog_Italic = GetIniValueBoolean(strSysIni, "Button", "FontItalic", 0)
    mboolDialog_Underline = GetIniValueBoolean(strSysIni, "Button", "FontUnderline", 0)
    mboolDialog_Strikethru = GetIniValueBoolean(strSysIni, "Button", "FontStrikethru", 0)
    lngDialog_Color = GetIniValueLong(strSysIni, "Button", "FontColor", 0)
End Sub

Public Function GetMB_Manufacturer() As String

    Dim objs As Object
    Dim obj  As Object
    Dim WMI  As Object
    Dim sAns As String

    Set WMI = CreateObject("WinMgmts:")
    Set objs = WMI.InstancesOf("Win32_ComputerSystem ")

    For Each obj In objs
        sAns = sAns & obj.Manufacturer

        If sAns < objs.Count Then
            sAns = sAns & "_"
        End If

    Next
    GetMB_Manufacturer = Trim$(sAns)
End Function

Public Function GetMB_Model() As String

    Dim objs As Object
    Dim obj  As Object
    Dim WMI  As Object
    Dim sAns As String

    Set WMI = CreateObject("WinMgmts:")
    Set objs = WMI.InstancesOf("Win32_ComputerSystem ")

    For Each obj In objs
        sAns = sAns & obj.Model

        If sAns < objs.Count Then
            sAns = sAns & "_"
        End If

    Next
    GetMB_Model = Trim$(sAns)
End Function

Public Sub Main()

    Dim LCID         As Long
    Dim strSysIniTMP As String

    ' ��������� ��� ������ ��������
    mboolUnloadClean = True
    strProductVersion = App.Major & "." & App.Minor & "." & App.Revision
    strProductName = App.ProductName & " " & strProductVersion & " @" & App.CompanyName
    ' ���� app.path ��� ����������
    GetCurAppPath

    On Error GoTo 0

    ' - �������������� ����� WindowsXP
    mboolInitXPStyle = InitXPStyle

    ' ��������� ��� ��������???
    If App.PrevInstance Then
        MsgBoxEx "Application is already running or quits ..." & vbNewLine & "This window will close automatically in 4 seconds. Please wait or click OK", 5, vbExclamation + vbSystemModal, strProductName
        ShowPrevInstance
    End If

    'MsgBox GetMD5("c:\WINDOWS\system32\DRVSTORE\igxp32_28D4AE6A4B66DD890D24C65EE34E5B62AB7E0BB9\igfxnt5.cat")
    ' ��������� �������� �� ��������� � ������ IDE
    mboolIsDesignMode = SetDebugMode
    kavichki = ChrW$(34)
    '�������� ��������� ������� windows � ������� windows
    strWinDir = BackslashAdd2Path(Environ$("WINDIR"))
    strWinTemp = BackslashAdd2Path(Environ$("TMP"))

    If InStr(1, strWinTemp, " ", vbTextCompare) > 0 Then
        strWinTemp = strWinDir & "TEMP"
    End If

    ' ���� ��������� ������� windows  (%windir%\temp)����������
    If PathFileExists(strWinTemp) = 0 Then
        MsgBox "Windows TempPath not Exist or Environ %TMP% undefined. Program is exit!!!", vbInformation, strProductName
        End
    End If

    ' ���� ������� tools ����������
    If PathFileExists(strAppPath & "\Tools\") = 0 Then
        MsgBox "Not found the main program subfolder '.\Tools'." & vbNewLine & "Program is exit!!!", vbInformation, strProductName
        End
    End If

    ' Disable DEP for current process
    If mboolDisableDEP Then
        SetDEPDisable
    End If

    'winDir = Getpath_WINDOWS
    strSysDir86 = Getpath_SYSTEM
    strSysDir = strSysDir86
    strOsCurrentVersion = CStr(OSInfo(4))

    If strOsCurrentVersion > "5.0" Then
        ' ����������� windows x64
        mboolIsWin64 = IsWow64

        If mboolIsWin64 Then
            Win64ReloadOptions
        End If

    ElseIf strOsCurrentVersion = "5.0" Then
        ' ��� win2k ���� ������ devcon
        'strDevConExePath = strDevConExePathW2k
    End If

    strSysDirCatRoot = strSysDir86 & "CatRoot\"
    strSysDirDrivers = strSysDir86 & "drivers\"
    strSysDirDRVStore = strSysDir86 & "DRVSTORE\"

    If strOsCurrentVersion >= "6.0" Then
        strSysDirDRVStore = strSysDir86 & "DriverStore\FileRepository\"
    End If

    strInfDir = strWinDir & "inf\"
    strWinDirHelp = strWinDir & "help\"
    strSysDrive = Environ$("SYSTEMDRIVE")
    ' ������� ��������� �������
    strWorkTemp = strWinTemp & "DriversBackuper"

    ' ������� ��������� ������� �������
    If PathFileExists(strAppPath & "\DriversBackuper.ini") = 0 Then
        strSysIni = CStr(strAppPath & "\Tools\DriversBackuper.ini")
    Else
        strSysIni = CStr(strAppPath & "\DriversBackuper.ini")
    End If

    ' �������� �� ��������� � CD
    mboolIsDriveCDRoom = IsDriveCDRoom
    ' ������� ���� �������� ��� �������������
    CreateIni
    ' ��������� ���� �����������
    LCID = GetSystemDefaultLCID()
    'language id
    strPCLangID = GetUserLocaleInfo(LCID, LOCALE_ILANGUAGE)
    'localized name of language
    strPCLangLocaliseName = GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE)
    'English name of language
    strPCLangEngName = GetUserLocaleInfo(LCID, LOCALE_SENGLANGUAGE)

    '��������� �������� �����
    If PathFileExists(strAppPath & "\Tools\LangDBS") = 1 Then
        mboolLanguageChange = LoadLanguage
    End If

    '��������� ����������� ���������
    LocaliseMessage strPCLangCurrentPath
    ' ��������� �������� �� ini-�����
    GetMainIniParam

    ' ���� ����� ��������� ��������� ��������� ���� �� ������� ini, �� ������������� ���� ����������
    If mboolLoadIniTmpAfterRestart Then
        If GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP", False) Then
            ' Reload Main ini
            strSysIniTMP = GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP_PATH", vbNullString)

            If LenB(strSysIniTMP) > 0 Then
                If PathFileExists(strSysIniTMP) = 1 Then
                    strSysIni = strSysIniTMP
                    ' ���������� ������������ ��������
                    GetMainIniParam
                End If
            End If
        End If
    End If

    If PathFileExists(strWorkTemp) = 0 Then
        CreateNewDirectory strWorkTemp
    End If

    '����������� �������� �����
    If PathFileExists(strAppPath & "\Tools\LangDBS") = 1 Then
        mboolLanguageChange = LoadLanguage
    End If

    '����������� ����������� ���������
    LocaliseMessage strPCLangCurrentPath
    strPathImageMain = strAppPath & "\Tools\GraphicsDBS\Main\"
    strPathImageMenu = strAppPath & "\Tools\GraphicsDBS\Menu\"
    LoadIconImagePath
    ' ��������� �� ��� �� CD
    mboolLogNotOnCDRoom = LogNotOnCDRoom
    ' ������� ���-�������
    MakeCleanHistory
    ' �������� ������� ������� ������� ���������
    GetWorkArea
    ' ���������� ��� ������������� ��� �������� ����� ������
    strMB_Manufacturer = GetMB_Manufacturer
    strMB_Model = GetMB_Model
    strCompName = SafeFileName(Environ$("COMPUTERNAME"))

    If LenB(strMB_Manufacturer) > 0 And LenB(strMB_Model) > 0 Then
        strCompModel = SafeFileName(strMB_Manufacturer & strMB_Model)
    ElseIf LenB(strMB_Manufacturer) = 0 And LenB(strMB_Model) > 0 Then
        strCompModel = SafeFileName(strMB_Model)
    ElseIf LenB(strMB_Manufacturer) > 0 And LenB(strMB_Model) = 0 Then
        strCompModel = SafeFileName(strMB_Manufacturer)
    Else
        strCompModel = "Unknown"
    End If

    DebugMode "Version: " & strProductName
    DebugMode "Build: " & strDateProgram
    DebugMode "ExeName: " & App.EXEName
    DebugMode "AppWork: " & strAppPath
    DebugMode "OsCurrentVersion: " & strOsCurrentVersion
    DebugMode "IsWow64: " & mboolIsWin64
    DebugMode "Architecture: " & strOSArchitecture

    ' ������������ �����?
    If APIFunctionPresent("IsUserAnAdmin", "shell32.dll") Then
        mboolIsUserAnAdmin = IsUserAnAdmin
    Else
        mboolIsUserAnAdmin = True
    End If

    DebugMode "is User an Admin?: " & mboolIsUserAnAdmin

    If Not mboolIsUserAnAdmin Then
        If Not mboolRunWithParam Then
            If MsgBox("Program needs Administrator privileges. You do not have such rights. You want to continue?", vbYesNo + vbQuestion, strProductName) = vbNo Then
                End
            End If
        End If
    End If

    DebugMode "SystemDrive: " & strSysDrive
    DebugMode "WinDir: " & strWinDir
    DebugMode "SysDir: " & strSysDir
    DebugMode "SysDir86: " & strSysDir86
    DebugMode "SysDir64: " & strSysDir64
    DebugMode "TmpDir: " & strWinTemp
    DebugMode "WorkTemp: " & strWorkTemp
    DebugMode "IsDriveCDRoom: " & mboolIsDriveCDRoom
    DebugMode "MotherBoard_Manufactured: " & strMB_Manufacturer
    DebugMode "MotherBoard_Model: " & strMB_Model
    'DebugMode Environ$("PROGRAMFILES")
    'DebugMode Environ$("SYSTEMROOT")
    'DebugMode Environ$("ALLUSERSPROFILE")
    'DebugMode Environ$("APPDATA")
    regParam = GetRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\Internet Explorer", "Version")
    DebugMode "IE Version: " & regParam
    DebugMode "OS Language: ID=" & strPCLangID & " Name=" & strPCLangEngName & "(" & strPCLangLocaliseName & ")"
    ' ����������� ������� ���������
    RegisterAddComponent
    DebugMode "InitXPStyle: " & mboolInitXPStyle

    If APIFunctionPresent("IsAppThemed", "uxtheme.dll") Then
        mboolAppThemed = IsAppThemed
        DebugMode "IsWindowsAppThemed: " & mboolAppThemed
    End If
    
    ' ������� ������ �������
    CmdLineParsing
    
    mboolFirstStart = True
    '# ���������� �������� �����
    frmMain.Show
End Sub

Private Sub CmdLineParsing()
    Dim argRetCMD As Collection
    Dim i   As Integer
    Dim intArgCount As Integer
    Dim strArg As String
    Dim strArg_x() As String

    With New cCMDArguments
        .CommandLine = "CMDLineParams " & Command$
        Set argRetCMD = .Arguments
        intArgCount = argRetCMD.Count
    End With
    
    For i = 2 To intArgCount
        strArg = argRetCMD(i)
        If InStr(1, strArg, "=", vbTextCompare) > 0 Then
            strArg_x = Split(strArg, "=")
            strArg = strArg_x(0)
        End If
        
        Debug.Print strArg
        Select Case LCase(strArg)
            Case "/?", "/h", "-help", "/help", "-h", "--h", "--help"
                ShowHelpMsg
                End
            Case "/extractdll", "-extractdll", "--extractdll"
                ExtractrResToFolder argRetCMD(i)
                End
            Case "/regdll", "-regdll", "--regdll"
                RegisterAddComponent
                End
            Case Else
                ShowHelpMsg
                End
        End Select
    Next i

End Sub

' ����� ���� � ����������� �������
Private Sub ShowHelpMsg()
    MsgBox strMessages(25), vbInformation & vbOKOnly, strProductName & " " & strProductVersion
End Sub

' ���������� �������� ��������� � �������
Private Sub ExtractrResToFolder(strArg As String)
Dim strArg_x() As String
Dim strPathToTemp As String
Dim strPathTo As String
    
    ' ��������� ���� �� ���������
    If InStr(1, strArg, "=", vbTextCompare) > 0 Then
        strArg_x = Split(strArg, "=")
        strPathToTemp = strArg_x(1)
    End If
    
    ' ��������� ������������ ��������
    If LenB(strPathToTemp) > 0 Then
        If Not IsPathAFolder(strPathToTemp) Then
            CreateNewDirectory strPathToTemp
        End If
        
        strPathTo = BackslashAdd2Path(strPathToTemp)
    Else
        strPathTo = strWorkTemp
    End If
    
    ' ������ ���������� ���� (dll-ocx) �������� ���������
    If ExtractResourceAll(strPathTo) Then
        If MsgBox(strMessages(21), vbYesNo + vbInformation, strProductName) = vbYes Then
            ShellEx strPathTo, essSW_SHOWNORMAL
        End If
    Else
        If MsgBox(strMessages(22), vbYesNo + vbInformation, strProductName) = vbYes Then
            ShellEx strPathTo, essSW_SHOWNORMAL
        End If
    End If
        
End Sub

'! -----------------------------------------------------------
'!  �������     :  SetDebugMode
'!  ����������  :
'!  ��������    :  �������� �� �� ������ �������� ��������� �� ������� ��� ������ exe
'! -----------------------------------------------------------
Private Function SetDebugMode() As Boolean

    On Error GoTo InIDE

    Debug.Print 1 / 0
    SetDebugMode = False
    Exit Function
InIDE:
    SetDebugMode = True
End Function

'! -----------------------------------------------------------
'!  �������     :  Win64ReloadOptions
'!  ����������  :
'!  ��������    :  �������������� ���������� ��� Win x64
'! -----------------------------------------------------------
Private Sub Win64ReloadOptions()

    DebugMode "Win64ReloadOptions"
    strSysDir86 = GetSpecialFolderPath(CSIDL_SYSTEM)
    strSysDir64 = GetSystemWow64Dir

    If strSysDir64 = vbNullString Then
        strSysDir64 = GetSpecialFolderPath(CSIDL_SYSTEMX86)
    End If

    strSysDir64 = BackslashAdd2Path(strSysDir64)
    strSysDir86 = BackslashAdd2Path(strSysDir86)
    DebugMode "CSIDL_SYSTEM: " & strSysDir86
    DebugMode "CSIDL_SYSTEMX86: " & strSysDir64

    ' ���� �������������� ���� ����������, �� ��������� ���, ���� ���, �� �����
    If PathFileExists(strSysDir64) And InStr(1, strSysDir64, "64", vbTextCompare) > 0 Then
        strSysDir = strSysDir64
    ElseIf PathFileExists(strWinDir & "SysWOW64") Then
        strSysDir = strWinDir & "SysWOW64"
    Else
        strSysDir = Getpath_SYSTEM
    End If

    strSysDir64 = strSysDir
End Sub
