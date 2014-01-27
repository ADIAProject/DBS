Attribute VB_Name = "mMain"
Option Explicit

' Основные параметры программы
Public Const strDateProgram         As String = "21/01/2013"

' Основные переменные проекта (название, версия и т.д)
Public strProductName               As String
Public strProductVersion            As String
Public Const strProjectName         As String = "DriversBackuper"
Public Const strUrl_MainWWWSite     As String = "http://adia-project.net/"                   ' Домашний сайт проекта
Public Const strUrl_MainWWWForum    As String = "http://adia-project.net/forum/index.php"    ' Домашний форум проекта
Public Const strUrlOsZoneNetThread  As String = "http://forum.oszone.net/thread-190814.html" ' Топик программы на сайте Oszone.net

'Константы путей основных каталогов и файла настроек (вынесены отдельно для универсальности кода под разные проекты)
Public Const strToolsLang_Path      As String = "Tools\LangDBS"         ' Каталог с языковыми файлами
Public Const strToolsDocs_Path      As String = "Tools\DocsDBS"         ' Каталог с документацией на программу
Public Const strToolsGraphics_Path  As String = "Tools\GraphicsDBS"     ' Каталог с графическими ресурсами программы
Public Const strSettingIniFile      As String = "DriversBackuper.ini"   ' INI-Файл настроек программы

' Версии лицензионного соглашения и файла Donate
Public Const strEULA_Version        As String = "02/02/2010"
Public Const strEULA_MD5RTF         As String = "68da44c8b1027547e4763472e0ecb727"
Public Const strEULA_MD5RTF_Eng     As String = "0cbd9d50eec41b26d24c5465c4be70bc"
Public Const strDONATE_MD5RTF       As String = "97f8178b2af5ba9377f76baf4ff71f78"
Public Const strDONATE_MD5RTF_Eng   As String = "59bbfbf6decbf91023da434cbe940d33"

' Массивы данных
Public arrHwidsLocal()                      As String

Public mbDateFormatRus                      As Boolean ' Автообновление конфигурации при удалении драйвера
Public strSysIni                            As String ' рабочий файл настроек
Public mbLoadIniTmpAfterRestart             As Boolean
Public mbEULAAgree                          As Boolean
Public strWorkTemp                          As String
Public strWorkTempBackSL                    As String
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
Public mbLogNotOnCDRoom                     As Boolean
Public mbHideOtherProcess                   As Boolean
Public mbDelTmpAfterClose                   As Boolean
Public mbUpdateCheck                        As Boolean
Public mbUpdateCheckBeta                    As Boolean
Public mbUpdateToolTip                      As Boolean
Public mbIsDesignMode                       As Boolean
Public mbIsDriveCDRoom                      As Boolean
Public strArh7zExePATH                      As String
Public strArh7zParam1                       As String
Public strArh7zParam2                       As String
Public strArh7zSFXPATH                      As String
Public strArh7zSFXConfigPath                As String
Public strArh7zSFXConfigPathEn              As String
Public mbAddInList                       As Boolean 'режим работы с элементом listview - либо изменние либо добавление
Public LastIdOS                          As Long 'номер последнего элемента в списке ОС
Public mbRestartProgram                  As Boolean 'Маркер перезапуска программы
Public mbStartMaximazed                  As Boolean
Public strDPInstExePath                  As String
Public strDPInstExePath64                As String
Public strDPInstExePath86                As String

' Параметры DPinst
Public mbDpInstLegacyMode                As Boolean
Public mbDpInstPromptIfDriverIsNotBetter As Boolean
Public mbDpInstForceIfDriverIsNotBetter  As Boolean
Public mbDpInstSuppressAddRemovePrograms As Boolean
Public mbDpInstSuppressWizard            As Boolean
Public mbDpInstQuietInstall              As Boolean
Public mbDpInstScanHardware              As Boolean

Public strImageMainName                  As String
Public mbSilentDLL                       As Boolean

' Расширенное меню
'Public mbExMenu                         As Boolean
Public strImageMenuName                  As String

'Прочие параметры программы
Public mbIsWin64                         As Boolean
Public mbFirstStart                      As Boolean

' Запуск с коммандной строкой
Public mbRunWithParam                    As Boolean
Private mbRunWithParamS                  As Boolean
Private strRunWithParam                  As String

Private mbIsUserAnAdmin                  As Boolean ' Пользователь администратор?

' кэпшн основной формы
Public strFrmMainCaptionTemp             As String
Public strFrmMainCaptionTempDate         As String

'-------------------- Переменные размеров Форм ------------------'
' Значения размеров формы
Public lngMainFormWidth                  As Long
Public lngMainFormHeight                 As Long
' Минимальные значения размеров формы
Public Const lngMainFormWidthMin         As Long = 12700
Public Const lngMainFormHeightMin        As Long = 6000
' Дефолтные значения размеров формы
Private Const lngMainFormWidthDef        As Long = 12700
Private Const lngMainFormHeightDef       As Long = 8000

Public mbSaveSizeOnExit                  As Boolean
Public mbCheckAllGroup                   As Boolean
Public mbListOnlyGroup                   As Boolean
Public miStartMode                       As Long
Public miArchMode                        As Long
Public arrOSList()                       As String
Public OSCount                           As Long
Public mbBackFolderPredefine             As Boolean
Public mbBlockListOnBackup               As Boolean

' Параметры каталога %Temp%
Public mbTempPath                        As Boolean
Public strAlternativeTempPath            As String
Public mbPatnAbs                         As Boolean
Public lngArchNameMode                   As Long
Public strArchNameCustom                 As String

Public mbDisableDEP                      As Boolean ' Переменная для определения выключения DEP

Private mbInitXPStyle                    As Boolean

' Переменные для определения модели компа
Public strCompName                       As String
Public strMB_Model                       As String
Public strMB_Manufacturer                As String
Public strCompModel                      As String
Public mbIsNotebok                       As Boolean ' Этот компьютер является ноутбуком

Public mbCheckUpdNotEnd                  As Boolean ' Маркер, показывающий что еще идет проверка обновления программы
Public mbChangeResolution                As Boolean ' Маркер, показывающий что проводилось изменение разрешения экрана
Public mbSilentRun                       As Boolean ' Работаем в тихом режиме
Public strThisBuildBy                    As String  ' Добавляем к описанию в главном окне в названии программы

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Main
'! Description (Описание)  :   [Основная функция запуска программы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Main()

    Dim mbShowFormLicence As Boolean
    Dim strSysIniTMP      As String
    Dim strLicenceDate    As String  ' дата лицензионного соглашения из реестра
    Dim mbShowLicence     As Boolean ' Показать лицензионное соглашение
    Dim mbIsUserAnAdmin   As Boolean ' Пользователь администратор?

    On Error Resume Next

    dtStartTimeProg = GetTickCount
    Set objFSO = New Scripting.FileSystemObject

    ' Запоминаем app.path и прочее в переменные
    GetCurAppPath
    strProductVersion = App.Major & "." & App.Minor & "." & App.Revision
    strProductName = App.ProductName & " v." & strProductVersion & " @" & App.CompanyName

    'считываем версию операционки
    If Not OsCurrVersionStruct.IsInitialize Then
        OsCurrVersionStruct = OSInfo
    End If

    strOsCurrentVersion = OsCurrVersionStruct.VerFull
    'Получаем временный каталог windows и каталог windows
    strWinDir = BackslashAdd2Path(Environ$("WINDIR"))
    strWinTemp = BackslashAdd2Path(Environ$("TMP"))

    If InStr(strWinTemp, " ") Then
        strWinTemp = BackslashAdd2Path(PathCombine(strWinDir, "TEMP"))
    End If

    ' Если временный каталог windows  (%windir%\temp)недоступен
    If PathExists(strWinTemp) = False Then
        MsgBox "Windows TempPath not Exist or Environ %TMP% undefined. Program is exit!!!", vbInformation, strProductName

        End

    End If
    '******************************************
    ' Проверяем работает ли программа в режиме IDE
    ' Программа уже запущена???
    If App.PrevInstance And Not InIDE() Then
        MsgBoxEx "Found a running application 'Drivers Installer Assistant'. If you restart the program from the settings menu, then save the settings, the program waits until the previous session..." & str2vbNewLine & _
                                    "This window will close automatically in 5 seconds. Please wait or click OK", 6, vbExclamation + vbSystemModal, strProductName
        ShowPrevInstance
    Else
        '******************************************
        ' - Инициализируем стиль WindowsXP
        Call ComCtlsInitIDEStopProtection
        Call InitVisualStyles
    End If

    ' Если каталог tools недоступен
    If PathExists(strAppPathBackSL & "Tools\") = False Then
        MsgBox "Not found the main program subfolder '.\Tools'." & vbNewLine & "Program is exit!!!", vbInformation, strProductName

        End

    End If

    ' Рабочий временный каталог
    strWorkTemp = strWinTemp & strProjectName
    strWorkTempBackSL = BackslashAdd2Path(strWorkTemp)

    ' Создаем временный рабочий каталог
    If PathExists(strAppPathBackSL & strSettingIniFile) = False Then
        strSysIni = strAppPathBackSL & "Tools\" & strSettingIniFile
    Else
        strSysIni = strAppPathBackSL & strSettingIniFile
    End If

    ' Запущена ли программа с CD
    mbIsDriveCDRoom = IsDriveCDRoom
    ' Создаем файл настроек при необходимости
    CreateIni
    ' Считаваем язык операционки
    LoadLanguageOS

    'загружаем языковые файлы
    If PathExists(strAppPathBackSL & strToolsLang_Path) Then
        mbMultiLanguage = LoadLanguageList
    End If

    'загружаем программные сообщения
    LocaliseMessage strPCLangCurrentPath
    ' Получение настроек из ini-файла
    GetMainIniParam

    ' Если стоит настройка проверять временный путь на наличие ini, то перезагружаем файл параметров
    If mbLoadIniTmpAfterRestart Then
        If GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP", False) Then
            ' Reload Main ini
            strSysIniTMP = GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP_PATH", vbNullString)

            If LenB(strSysIniTMP) > 0 Then
                If PathExists(strSysIniTMP) Then
                    strSysIni = strSysIniTMP
                    ' Собственно перезагрузка настроек
                    GetMainIniParam
                End If
            End If
        End If
    End If

    If PathExists(strWorkTemp) = False Then
        CreateNewDirectory strWorkTemp
    End If

    'Перегружаем языковые файлы
    If PathExists(strAppPathBackSL & strToolsLang_Path) Then
        mbMultiLanguage = LoadLanguageList
    End If

    'перегружаем программные сообщения
    LocaliseMessage strPCLangCurrentPath
    strPathImageStatusButton = strAppPathBackSL & strToolsGraphics_Path & "\StatusButton\"
    strPathImageMain = strAppPathBackSL & strToolsGraphics_Path & "\Main\"
    'strPathImageMenu = strAppPathBackSL & strToolsGraphics_Path & "\Menu\"
    LoadIconImagePath
    ' Находится ли лог на CD
    mbLogNotOnCDRoom = LogNotOnCDRoom
    ' Очищаем лог-историю
    MakeCleanHistory
    ' Получаем размеры рабочей области программы
    GetWorkArea
    ' Проверяем на запуск с параметрами
    strRunWithParam = CStr(Command)

    If LenB(strRunWithParam) > 0 Then
        ' Парсинг строки запуска
        CmdLineParsing
    End If

    If APIFunctionPresent("IsUserAnAdmin", "shell32.dll") Then
        mbIsUserAnAdmin = IsUserAnAdmin
    Else
        mbIsUserAnAdmin = True
    End If

    If Not mbDebugTime2File Then
        DebugMode "Current Date: " & Now()
    End If

    DebugMode "Version: " & strProductName & vbNewLine & _
              "Build: " & strDateProgram & vbNewLine & _
              "ExeName: " & App.EXEName & ".exe" & vbNewLine & _
              "AppWork: " & strAppPath & vbNewLine & _
              "is User an Admin?: " & mbIsUserAnAdmin

    If mbIsUserAnAdmin Then
        ' записываем в реестр мой сертификат, для ЭЦП на exe-файлы
        DebugMode "SaveSert2Reestr"
        SaveSert2Reestr
    Else

        If Not mbRunWithParam Then
            If MsgBox(strMessages(138), vbYesNo + vbQuestion, strProductName) = vbNo Then

                End

            End If
        End If
    End If

    DebugMode "WinDir: " & strWinDir & vbNewLine & _
              "TmpDir: " & strWinTemp & vbNewLine & _
              "WorkTemp: " & strWorkTemp & vbNewLine & _
              "IsDriveCDRoom: " & mbIsDriveCDRoom

    If strOsCurrentVersion > "5.0" Then
        ' Определение windows x64
        mbIsWin64 = IsWow64
        DebugMode "IsWow64: " & mbIsWin64

        If mbIsWin64 Then
            Win64ReloadOptions
        End If

    ElseIf strOsCurrentVersion = "5.0" Then
        ' Для win2k надо старый devcon
        'strDevConExePath = strDevConExePathW2k
    End If

    ' Disable DEP for current process
    If mbDisableDEP Then
        SetDEPDisable
    End If

    ' Регистрация внешних компонент
    RegisterAddComponent

    DebugMode "OsCurrentVersion: " & strOsCurrentVersion & vbNewLine & _
              "Architecture: " & strOSArchitecture & vbNewLine & _
              "OS Language: ID=" & strPCLangID & " Name=" & strPCLangEngName & "(" & strPCLangLocaliseName & ")"

    ' Определяем служебные каталоги Windows
    regParam = GetRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\Internet Explorer", "Version")
    strSysDir86 = Getpath_SYSTEM
    strSysDir = strSysDir86
    strSysDirCatRoot = strSysDir86 & "CatRoot\"
    strSysDirDrivers = strSysDir86 & "drivers\"
    strInfDir = strWinDir & "inf\"
    strWinDirHelp = strWinDir & "help\"
    strSysDrive = Environ$("SYSTEMDRIVE")

    strSysDirDRVStore = strSysDir86 & "DRVSTORE\"
    If strOsCurrentVersion >= "6.0" Then
        strSysDirDRVStore = strSysDir86 & "DriverStore\FileRepository\"
    End If

    DebugMode "InitXPStyle: " & mbInitXPStyle

    If APIFunctionPresent("IsAppThemed", "uxtheme.dll") Then
        mbAppThemed = IsAppThemed
        DebugMode "IsAppThemed: " & mbAppThemed
    End If

    mbAeroEnabled = IsAeroEnabled
    DebugMode "IsAeroEnabled : " & mbAeroEnabled
    ' изменяем разрешающую способность экрана монитора при необходимости
    SetVideoMode
    GetWorkArea
    ' Переменные для использовании при создании имени архива
    strCompModel = GetMBInfo()
    DebugMode "isNotebook: " & mbIsNotebok & vbNewLine & _
              "Notebook/Motherboard Model: " & strCompModel & vbNewLine & _
              "SystemDrive: " & strSysDrive & vbNewLine & _
              "SysDir: " & strSysDir & vbNewLine & _
              "SysDir86: " & strSysDir86 & vbNewLine & _
              "SysDir64: " & strSysDir64 & vbNewLine & _
              "IE Version: " & regParam
    
    mbFirstStart = True
    ' Показ лицензионного соглашения
    mbShowLicence = GetSetting(App.ProductName, "Licence", "Show at Startup", True)
    strLicenceDate = GetSetting(App.ProductName, "Licence", "EULA_DATE", strEULA_Version)

    If InStr(1, strLicenceDate, strEULA_Version, vbTextCompare) Then
        If mbShowLicence Then
            If Not mbRunWithParam Then
                mbShowFormLicence = True
            End If

            If mbEULAAgree Then
                mbShowFormLicence = False
            End If

        Else
            mbShowFormLicence = False
        End If

    Else

        If Not mbRunWithParam Then
            mbShowFormLicence = True
        End If

        If mbEULAAgree Then
            mbShowFormLicence = False
        End If
    End If

    If mbShowFormLicence Then
        'Открываем форму лицензионного соглашения
        frmLicence.Show
    Else
        'Открываем основную форму
        frmMain.Show vbModeless
    End If


End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ChangeStatusTextAndDebug
'! Description (Описание)  :   [Изменение текста статустной строки и отладочной информации]
'! Parameters  (Переменные):   strPanel2Text (String)
'                              strDebugText (String)
'                              mbEqual (Boolean = False)
'                              mbDoEvents (Boolean = True)
'                              strPanel1Text (String)
'!--------------------------------------------------------------------------------
Public Sub ChangeStatusTextAndDebug(Optional strPanel2Text As String, Optional strDebugText As String, Optional ByVal mbEqual As Boolean = False, Optional ByVal mbDoEvents As Boolean = True, Optional strPanel1Text As String)

    If LenB(strPanel2Text) > 0 Then
        If mbDoEvents Then
            DoEvents
        End If

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
        If mbEqual Then
            If LenB(strPanel1Text) > 0 Then
                DebugMode strPanel1Text & ": " & strPanel2Text
            Else
                DebugMode strPanel2Text
            End If

        Else
            DebugMode strDebugText
        End If

    Else

        If mbEqual Then
            If LenB(strPanel1Text) > 0 Then
                DebugMode strPanel1Text & ": " & strPanel2Text
            Else
                DebugMode strPanel2Text
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateIni
'! Description (Описание)  :   [Сохранение настроек в ини файл если файла нет]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CreateIni()


    If PathExists(strSysIni) = False Then
        If mbIsDriveCDRoom Then
            strSysIni = strWorkTempBackSL & strSettingIniFile
            MsgBox "File " & strSettingIniFile & " is not Exist!" & vbNewLine & "This program works from CD\DVD, so we create temporary " & strSettingIniFile & "-file" & vbNewLine & strSysIni, vbInformation + vbApplicationModal, strProductName
        End If

        'Секция Main
        IniWriteStrPrivate "Main", "DelTmpAfterClose", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheck", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheckBeta", "0", strSysIni
        IniWriteStrPrivate "Main", "StartMode", "2", strSysIni
        IniWriteStrPrivate "Main", "EULAAgree", "0", strSysIni
        IniWriteStrPrivate "Main", "HideOtherProcess", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTemp", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTempPath", "%Temp%", strSysIni
        IniWriteStrPrivate "Main", "AutoLanguage", "1", strSysIni
        IniWriteStrPrivate "Main", "StartLanguageID", "0409", strSysIni
        IniWriteStrPrivate "Main", "IconMainSkin", "Standart", strSysIni
        IniWriteStrPrivate "Main", "SilentDLL", "0", strSysIni
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", "0", strSysIni
        IniWriteStrPrivate "Main", "DateFormatRus", "1", strSysIni
        IniWriteStrPrivate "Main", "CheckAllGroup", "1", strSysIni
        IniWriteStrPrivate "Main", "ListOnlyGroup", "1", strSysIni
        IniWriteStrPrivate "Main", "BlockListOnBackup", "1", strSysIni
        IniWriteStrPrivate "Main", "CalculateHashMode", "1", strSysIni
        IniWriteStrPrivate "Main", "ArchMode", "0", strSysIni

        'Секция Debug
        IniWriteStrPrivate "Debug", "DebugEnable", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogPath", "%SYSTEMDRIVE%", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogName", "DBS-LOG_%DATE%.txt", strSysIni
        IniWriteStrPrivate "Debug", "CleenHistory", "1", strSysIni
        IniWriteStrPrivate "Debug", "DetailMode", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLog2AppPath", "0", strSysIni
        IniWriteStrPrivate "Debug", "Time2File", "0", strSysIni
        'Секция DPInst
        IniWriteStrPrivate "DPInst", "PathExe", "Tools\DPInst\DPInst.exe", strSysIni
        IniWriteStrPrivate "DPInst", "PathExe64", "Tools\DPInst\DPInst64.exe", strSysIni
        'IniWriteStrPrivate "DPInst", "LegacyMode", 1, strSysIni
        'IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", 1, strSysIni
        'IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "SuppressWizard", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "QuietInstall", 0, strSysIni
        'IniWriteStrPrivate "DPInst", "ScanHardware", 1, strSysIni
        'Секция Arc
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
        'Секция OS
        IniWriteStrPrivate "OS", "OSCount", "4", strSysIni
        'Секция OS_1
        IniWriteStrPrivate "OS_1", "Ver", "5.0;5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_1", "drpFolder", "drivers\2k_xp_2003\x32\", strSysIni
        IniWriteStrPrivate "OS_1", "is64bit", "0", strSysIni
        'Секция OS_2
        IniWriteStrPrivate "OS_2", "Ver", "5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_2", "drpFolder", "drivers\2k_xp_2003\x64\", strSysIni
        IniWriteStrPrivate "OS_2", "is64bit", "1", strSysIni









        'Секция OS_3
        IniWriteStrPrivate "OS_3", "Ver", "6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_3", "drpFolder", "drivers\vista_7_8\x32\", strSysIni
        IniWriteStrPrivate "OS_3", "is64bit", "0", strSysIni






        'Секция OS_4
        IniWriteStrPrivate "OS_4", "Ver", "6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_4", "drpFolder", "drivers\vista_7_8\x64\", strSysIni
        IniWriteStrPrivate "OS_4", "is64bit", "1", strSysIni
        'Секция MainForm
        IniWriteStrPrivate "MainForm", "Width", CStr(lngMainFormWidthDef), strSysIni
        IniWriteStrPrivate "MainForm", "Height", CStr(lngMainFormHeightDef), strSysIni
        IniWriteStrPrivate "MainForm", "StartMaximazed", "0", strSysIni
        IniWriteStrPrivate "MainForm", "SaveSizeOnExit", "0", strSysIni
        IniWriteStrPrivate "MainForm", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "MainForm", "FontSize", "8", strSysIni
        IniWriteStrPrivate "MainForm", "HighlightColor", "32896", strSysIni

        'Секция Buttons
        IniWriteStrPrivate "Button", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "Button", "FontSize", "8", strSysIni
        IniWriteStrPrivate "Button", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Button", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Button", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Button", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Button", "FontColor", "0", strSysIni

        ' Приводим Ini файл к читабельному виду
        NormIniFile strSysIni
        ' Активация отладки после создания ini-файла
        mbDebugEnable = True
        mbCleanHistory = True
        strDebugLogPathTemp = "%SYSTEMDRIVE%"
        strDebugLogNameTemp = "DBS-LOG_%DATE%.txt"
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetMainIniParam
'! Description (Описание)  :   [Получение настроек из ини файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub GetMainIniParam()

    Dim i          As Long
    Dim cntOsInIni As Integer
    Dim strDebugLogPathFolder       As String

    '[Description]
    strThisBuildBy = GetIniValueString(strSysIni, "Description", "BuildBy", vbNullString)
    'strThisBuildBy = "www.SamLab.Ws"
    '[Debug]
    ' Активация отладки
    mbDebugEnable = GetIniValueBoolean(strSysIni, "Debug", "DebugEnable", 1)
    ' Очистка истории
    mbCleanHistory = GetIniValueBoolean(strSysIni, "Debug", "CleenHistory", 1)
    ' Путь до лог файла
    strDebugLogPathTemp = PathNameFromPath(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%SYSTEMDRIVE%"))
    strDebugLogPath = PathCollect(PathNameFromPath(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%SYSTEMDRIVE%")))
    ' Имя лог-файла
    strDebugLogNameTemp = GetIniValueString(strSysIni, "Debug", "DebugLogName", "DBS-LOG_%DATE%.txt")
    strDebugLogName = ExpandFileNamebyEnvironment(GetIniValueString(strSysIni, "Debug", "DebugLogName", "DBS-LOG_%DATE%.txt"))
    ' Деталировка отладки - по умолчанию=1
    lngDetailMode = GetIniValueLong(strSysIni, "Debug", "DetailMode", 1)
    ' Записывать время в лог-файл
    mbDebugTime2File = GetIniValueBoolean(strSysIni, "Debug", "Time2File", 0)
    ' Создавать лог-файл в подпапке "logs" программы
    mbDebugLog2AppPath = GetIniValueBoolean(strSysIni, "Debug", "DebugLog2AppPath", 0)

    If Not mbDebugLog2AppPath Then
        strDebugLogFullPath = strDebugLogPath & strDebugLogName

        If mbDebugEnable Then
            strDebugLogPathFolder = strDebugLogPath

            If PathExists(strDebugLogPathFolder) = False Then
                CreateNewDirectory strDebugLogPathFolder
            End If
        End If

    Else
        strDebugLogPath2AppPath = strAppPathBackSL & "logs\" & strDebugLogName
        strDebugLogFullPath = strDebugLogPath2AppPath

        If Not LogNotOnCDRoom Then
            If mbDebugEnable Then
                If PathExists(strAppPathBackSL & "logs\") = False Then
                    CreateNewDirectory strAppPathBackSL & "logs\"
                End If
            End If

        Else
            strDebugLogFullPath = strDebugLogPath & strDebugLogName
        End If
    End If

    If lngDetailMode < 1 Then
        lngDetailMode = 1
    ElseIf lngDetailMode > 2 Then
        lngDetailMode = 2
    End If

    '[Main]
    ' удаление при выходе
    mbDelTmpAfterClose = GetIniValueBoolean(strSysIni, "Main", "DelTmpAfterClose", 1)
    ' проверка обновлений при старте (Секция MAIN)
    mbUpdateCheck = GetIniValueBoolean(strSysIni, "Main", "UpdateCheck", 1)
    ' проверка обновлений при старте (Секция MAIN)
    mbUpdateCheckBeta = GetIniValueBoolean(strSysIni, "Main", "UpdateCheckBeta", 1)
    ' погасить EULA
    mbEULAAgree = GetIniValueBoolean(strSysIni, "Main", "EULAAgree", 0)
    ' Автоопределение языка
    mbAutoLanguage = GetIniValueBoolean(strSysIni, "Main", "AutoLanguage", 1)

    If Not mbAutoLanguage Then
        strStartLanguageID = IniStringPrivate("Main", "StartLanguageID", strSysIni)
    End If
    ' Получение альтернативного пути Temp
    strAlternativeTempPath = IniStringPrivate("Main", "AlternativeTempPath", strSysIni)

    If strAlternativeTempPath = "no_key" Then
        strAlternativeTempPath = strWinTemp
    End If

    ' при необходимости используем альтернативный temp
    mbTempPath = GetIniValueBoolean(strSysIni, "Main", "AlternativeTemp", 0)

    If mbTempPath Then
        strAlternativeTempPath = PathCollect(strAlternativeTempPath)
        DebugMode "AlternativeTempPath: " & strAlternativeTempPath

        If PathExists(strAlternativeTempPath) Then
            strWinTemp = strAlternativeTempPath
            strWorkTemp = strWinTemp & strProjectName

            ' Если нет, то создаем временный рабочий каталог
            If PathExists(strWorkTemp) = False Then
                CreateNewDirectory strWorkTemp
            End If

        Else
            DebugMode "Alternative TempPath not Exist. Use Windows Temp"
        End If
    End If

    ' Отображать дату версии в формате dd/mm/yyyy
    mbDateFormatRus = GetIniValueBoolean(strSysIni, "Main", "DateFormatRus", 0)
    ' папка со скином
    strImageMainName = GetIniValueString(strSysIni, "Main", "IconMainSkin", "Standart")
    ' Скрывать прочие процессы
    mbHideOtherProcess = GetIniValueBoolean(strSysIni, "Main", "HideOtherProcess", 1)
    ' Тихая регистрация DLL
    mbSilentDLL = GetIniValueBoolean(strSysIni, "Main", "SilentDll", 0)
    ' Показывать напоминание об обновлении (всплывающее окно)
    mbUpdateToolTip = GetIniValueBoolean(strSysIni, "Main", "UpdateToolTip", 1)
    ' Выделять всю группу
    mbCheckAllGroup = GetIniValueBoolean(strSysIni, "Main", "CheckAllGroup", 1)
    ' Выделять всю группу
    mbListOnlyGroup = GetIniValueBoolean(strSysIni, "Main", "ListOnlyGroup", 1)
    ' Стартовый режим
    miStartMode = GetIniValueLong(strSysIni, "Main", "StartMode", 2)
    'Блокирование окна listview ghb бекапировании
    mbBlockListOnBackup = GetIniValueBoolean(strSysIni, "Main", "BlockListOnBackup", 1)
    'Режим архивирования по умолчанию
    miArchMode = GetIniValueLong(strSysIni, "Main", "ArchMode", 0)
    ' расширенное меню
    'mbExMenu = GetIniValueBoolean(strSysIni, "Main", "ExMenu", 1)
    'strImageMenuName = GetIniValueString(strSysIni, "Main", "IconMenuSkin", "Standart")
    mbLoadIniTmpAfterRestart = GetIniValueBoolean(strSysIni, "Main", "LoadIniTmpAfterRestart", 0)
    mbDisableDEP = GetIniValueBoolean(strSysIni, "Main", "DisableDEP", 1)
    '--------------------- Получение путей до файлов ---------------------
    '[DPInst]
    ' DPInst.exe
    strDPInstExePath86 = IniStringPrivate("DPInst", "PathExe", strSysIni)

    If InStr(strDPInstExePath86, ":") Then
        mbPatnAbs = True
    End If

    strDPInstExePath86 = PathCollect(strDPInstExePath86)

    If PathExists(strDPInstExePath86) = False Then
        strDPInstExePath86 = strAppPathBackSL & "Tools\DPInst\DPInst.exe"

        If PathExists(strDPInstExePath86) = False Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath86, vbInformation, strProductName
        End If
    End If

    strDPInstExePath = strDPInstExePath86
    ' DPInst64.exe
    strDPInstExePath64 = IniStringPrivate("DPInst", "PathExe64", strSysIni)

    If InStr(strDPInstExePath64, ":") Then
        mbPatnAbs = True
    End If

    strDPInstExePath64 = PathCollect(strDPInstExePath64)

    If PathExists(strDPInstExePath64) = False Then
        strDPInstExePath64 = strAppPathBackSL & "Tools\DPInst\DPInst64.exe"

        If PathExists(strDPInstExePath64) = False Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath64, vbInformation, strProductName
        End If
    End If

    ' Настройки DpInst
    mbDpInstLegacyMode = GetIniValueBoolean(strSysIni, "DPInst", "LegacyMode", 1)
    mbDpInstPromptIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "PromptIfDriverIsNotBetter", 1)
    mbDpInstForceIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "ForceIfDriverIsNotBetter", 0)
    mbDpInstSuppressAddRemovePrograms = GetIniValueBoolean(strSysIni, "DPInst", "SuppressAddRemovePrograms", 0)
    mbDpInstSuppressWizard = GetIniValueBoolean(strSysIni, "DPInst", "SuppressWizard", 0)
    mbDpInstQuietInstall = GetIniValueBoolean(strSysIni, "DPInst", "QuietInstall", 0)
    mbDpInstScanHardware = GetIniValueBoolean(strSysIni, "DPInst", "ScanHardware", 1)
    '[Arc]
    ' 7za.exe
    strArh7zExePATH = IniStringPrivate("Arc", "PathExe", strSysIni)

    If InStr(strArh7zExePATH, ":") Then
        mbPatnAbs = True
    End If

    strArh7zExePATH = PathCollect(strArh7zExePATH)

    If PathExists(strArh7zExePATH) = False Then
        strArh7zExePATH = strAppPathBackSL & "Tools\Arc\7za.exe"

        If PathExists(strArh7zExePATH) = False Then
            MsgBox strMessages(7) & vbNewLine & strArh7zExePATH, vbInformation, strProductName
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
    ' Сохранять настройки при выходе
    mbSaveSizeOnExit = GetIniValueBoolean(strSysIni, "MainForm", "SaveSizeOnExit", 0)
    'Ширина основной формы
    lngMainFormWidth = GetIniValueLong(strSysIni, "MainForm", "Width", lngMainFormWidthDef)

    'Если полученное значение меньше минимального, то устанавливаем значение по умолчанию
    If lngMainFormWidth < lngMainFormWidthMin Then
        lngMainFormWidth = lngMainFormWidthDef
    End If

    'Высота основной формы
    lngMainFormHeight = GetIniValueLong(strSysIni, "MainForm", "Height", lngMainFormHeightDef)

    'Если полученное значение меньше минимального, то устанавливаем значение по умолчанию
    If lngMainFormHeight < lngMainFormHeightMin Then
        lngMainFormHeight = lngMainFormHeightDef
    End If

    ' получение вида запуска (Секция MainForm)
    mbStartMaximazed = GetIniValueBoolean(strSysIni, "MainForm", "StartMaximazed", 0)
    strFontMainForm_Name = GetIniValueString(strSysIni, "MainForm", "FontName", "Tahoma")
    lngFontMainForm_Size = GetIniValueLong(strSysIni, "MainForm", "FontSize", 8)
    ' Подсветка активного элемента
    glHighlightColor = GetIniValueLong(strSysIni, "MainForm", "HighlightColor", 32896)
    ' получение вида запуска (Секция OtherForm)
    strFontOtherForm_Name = GetIniValueString(strSysIni, "OtherForm", "FontName", "Tahoma")
    lngFontOtherForm_Size = GetIniValueLong(strSysIni, "OtherForm", "FontSize", 8)


    '[OS]
    ' получение Кол-ва систем (Секция OS) и построение массива ОС
    OSCount = IniLongPrivate("OS", "OSCount", strSysIni)

    If OSCount = 0 Or OSCount = 9999 Then
        DebugMode "The List supported operating systems is empty. PreDefine BackUpfolder not accessible"
        mbBackFolderPredefine = False
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
                If PathExists(PathCollect(arrOSList(i, 2))) = False Then
                    DebugMode "Not find folder for package driver backup" & vbNewLine & "для ОС: " & arrOSList(i, 0) & " is64bit:" & arrOSList(i, 1) & vbNewLine & vbNewLine & "Folder is not Exist: " & vbNewLine & PathCollect(arrOSList(i, 2))
                    arrOSList(i, 3) = "DriverPack folder is not Exist"
                End If

            Else
                DebugMode "Folder with package driver" & vbNewLine & "for OS: " & arrOSList(i, 0) & " is64bit:" & arrOSList(i, 1) & vbNewLine & "Is Not present in options. Correct and start the program again."
            End If

        Next
        mbBackFolderPredefine = True
    End If

    '[Button]
    ' Шрифт Кнопок
    strFontBtn_Name = GetIniValueString(strSysIni, "Button", "FontName", "Tahoma")
    miFontBtn_Size = GetIniValueLong(strSysIni, "Button", "FontSize", 8)
    mbFontBtn_Bold = GetIniValueBoolean(strSysIni, "Button", "FontBold", 0)
    mbFontBtn_Italic = GetIniValueBoolean(strSysIni, "Button", "FontItalic", 0)
    mbFontBtn_Underline = GetIniValueBoolean(strSysIni, "Button", "FontUnderline", 0)
    mbFontBtn_Strikethru = GetIniValueBoolean(strSysIni, "Button", "FontStrikethru", 0)
    lngFontBtn_Color = GetIniValueLong(strSysIni, "Button", "FontColor", 0)
End Sub





'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CmdLineParsing
'! Description (Описание)  :   [Функция анализа коммандной строки и присвоение переменных на основании передеваемых комманд]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CmdLineParsing()

    Dim argRetCMD    As Collection
    Dim i            As Integer
    Dim intArgCount  As Integer
    Dim strArg       As String
    Dim strArg_x()   As String
    Dim iArgRavno    As Integer
    Dim iArgDvoetoch As Integer
    Dim strArgParam  As String

    With New cCMDArguments
        .CommandLine = "CMDLineParams " & Command$
        Set argRetCMD = .Arguments
        intArgCount = argRetCMD.Count
    End With

    For i = 2 To intArgCount
        strArg = argRetCMD(i)
        iArgRavno = InStr(strArg, "=")
        iArgDvoetoch = InStr(strArg, ":")

        If iArgRavno > 0 Then
            strArg_x = Split(strArg, "=")
            strArg = strArg_x(0)
            strArgParam = strArg_x(1)
        ElseIf iArgDvoetoch > 0 Then
            'strArg_x = Split(strArg, ":")
            strArg = Left$(argRetCMD(i), iArgDvoetoch - 1)
            strArgParam = Right$(argRetCMD(i), Len(argRetCMD(i)) - iArgDvoetoch)
        End If

        Select Case LCase$(strArg)

            Case "/?", "/h", "-help", "/help", "-h", "--h", "--help"
                ShowHelpMsg

                End

            Case "/extractdll", "-extractdll", "--extractdll"
                ExtractrResToFolder strArgParam

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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShowHelpMsg
'! Description (Описание)  :   [Показ окна с параметрами запуска]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ShowHelpMsg()
    MsgBox strMessages(137), vbInformation & vbOKOnly, strProductName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveSert2Reestr
'! Description (Описание)  :   [процедура прописывания сертификата для проверки валидности цифровой подписи моего exe]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveSert2Reestr()

    Dim strBuffer      As String
    Dim strBuffer_x()  As String
    Dim strByteArray() As Byte
    Dim i              As Long

    On Error Resume Next
    
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\A31D3E0A4D99335EBD9B6F18E0915490F13525CA
'"Blob"=hex:03,00,00,00,01,00,00,00,14,00,00,00,a3,1d,3e,0a,4d,99,33,5e,bd,9b,\
'  6f,18,e0,91,54,90,f1,35,25,ca,20,00,00,00,01,00,00,00,28,02,00,00,30,82,02,\
'  24,30,82,01,91,a0,03,02,01,02,02,10,82,58,85,44,28,61,9e,bc,48,c0,05,a4,40,\
'  6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,30,1f,31,1d,30,1b,06,03,55,04,03,\
'  13,14,77,77,77,2e,61,64,69,61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,1e,17,\
'  0d,31,33,30,33,31,31,30,39,35,37,34,30,5a,17,0d,33,39,31,32,33,31,32,33,35,\
'  39,35,39,5a,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69,61,\
'  2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,81,9f,30,0d,06,09,2a,86,48,86,f7,0d,\
'  01,01,01,05,00,03,81,8d,00,30,81,89,02,81,81,00,c4,4e,f8,78,d3,eb,fc,45,49,\
'  13,31,a0,fc,f6,50,1d,3c,b3,4b,9e,d5,73,45,4c,06,93,70,e7,ee,c8,6b,25,82,16,\
'  4b,58,ea,22,40,ab,82,d7,c7,c9,90,0c,31,45,aa,7f,79,27,e6,b5,47,fe,7d,48,ad,\
'  70,e6,9a,46,25,64,0b,50,74,ce,ea,f1,8c,92,6c,82,2e,08,4b,aa,a8,10,05,d1,e8,\
'  9b,9b,fb,ce,79,3e,42,a4,49,88,03,c8,22,6f,b6,21,a2,3f,68,f2,84,5d,ac,29,a5,\
'  02,71,87,6d,81,ec,e3,d0,17,be,cf,48,58,a3,ab,ed,f5,9d,5f,02,03,01,00,01,a3,\
'  69,30,67,30,13,06,03,55,1d,25,04,0c,30,0a,06,08,2b,06,01,05,05,07,03,03,30,\
'  50,06,03,55,1d,01,04,49,30,47,80,10,01,60,4c,5b,6f,d2,c8,c6,60,6b,50,24,03,\
'  4b,9b,a7,a1,21,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69,\
'  61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,82,10,82,58,85,44,28,61,9e,bc,48,c0,\
'  05,a4,40,6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,03,81,81,00,08,a6,57,6e,\
'  3c,a5,7c,ad,41,ab,61,f9,8f,41,0e,6e,e0,b2,6e,bd,35,16,cc,0c,05,d1,e2,d9,d4,\
'  b2,71,50,70,fd,28,a0,c7,7f,8f,23,63,4a,c4,e0,1b,0e,98,37,c1,24,1f,4f,ae,ae,\
'  db,8d,ce,b8,cb,9e,13,6e,b0,a8,b0,0f,90,1b,22,94,97,fa,47,b6,29,b1,eb,98,4a,\
'  26,28,23,a5,0a,ef,59,43,b1,be,25,49,2b,cf,8d,bc,82,37,20,cd,b7,db,90,0b,d7,\
'  3d,7b,e9,f5,87,7b,87,bb,ae,f2,53,de,5d,17,72,25,18,f9,61,bd,4e,cd,6c,c8
'

    strBuffer = "03,00,00,00,01,00,00,00,14,00,00,00,a3,1d,3e,0a,4d,99,33,5e,bd,9b," & "6f,18,e0,91,54,90,f1,35,25,ca,20,00,00,00,01,00,00,00,28,02,00,00,30,82,02," & "24,30,82,01,91,a0,03,02,01,02,02,10,82,58,85,44,28,61,9e,bc,48,c0,05,a4,40," & _
                                "6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,30,1f,31,1d,30,1b,06,03,55,04,03," & "13,14,77,77,77,2e,61,64,69,61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,1e,17," & _
                                "0d,31,33,30,33,31,31,30,39,35,37,34,30,5a,17,0d,33,39,31,32,33,31,32,33,35," & "39,35,39,5a,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69,61," & _
                                "2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,81,9f,30,0d,06,09,2a,86,48,86,f7,0d," & "01,01,01,05,00,03,81,8d,00,30,81,89,02,81,81,00,c4,4e,f8,78,d3,eb,fc,45,49," & _
                                "13,31,a0,fc,f6,50,1d,3c,b3,4b,9e,d5,73,45,4c,06,93,70,e7,ee,c8,6b,25,82,16," & "4b,58,ea,22,40,ab,82,d7,c7,c9,90,0c,31,45,aa,7f,79,27,e6,b5,47,fe,7d,48,ad," & _
                                "70,e6,9a,46,25,64,0b,50,74,ce,ea,f1,8c,92,6c,82,2e,08,4b,aa,a8,10,05,d1,e8," & "9b,9b,fb,ce,79,3e,42,a4,49,88,03,c8,22,6f,b6,21,a2,3f,68,f2,84,5d,ac,29,a5," & _
                                "02,71,87,6d,81,ec,e3,d0,17,be,cf,48,58,a3,ab,ed,f5,9d,5f,02,03,01,00,01,a3," & "69,30,67,30,13,06,03,55,1d,25,04,0c,30,0a,06,08,2b,06,01,05,05,07,03,03,30," & _
                                "50,06,03,55,1d,01,04,49,30,47,80,10,01,60,4c,5b,6f,d2,c8,c6,60,6b,50,24,03," & "4b,9b,a7,a1,21,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69," & _
                                "61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,82,10,82,58,85,44,28,61,9e,bc,48,c0," & "05,a4,40,6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,03,81,81,00,08,a6,57,6e," & _
                                "3c,a5,7c,ad,41,ab,61,f9,8f,41,0e,6e,e0,b2,6e,bd,35,16,cc,0c,05,d1,e2,d9,d4," & "b2,71,50,70,fd,28,a0,c7,7f,8f,23,63,4a,c4,e0,1b,0e,98,37,c1,24,1f,4f,ae,ae," & _
                                "db,8d,ce,b8,cb,9e,13,6e,b0,a8,b0,0f,90,1b,22,94,97,fa,47,b6,29,b1,eb,98,4a," & "26,28,23,a5,0a,ef,59,43,b1,be,25,49,2b,cf,8d,bc,82,37,20,cd,b7,db,90,0b,d7," & _
                                "3d,7b,e9,f5,87,7b,87,bb,ae,f2,53,de,5d,17,72,25,18,f9,61,bd,4e,cd,6c,c8"
    strBuffer_x = Split(strBuffer, ",")

    ReDim strByteArray(UBound(strBuffer_x))

    For i = LBound(strBuffer_x) To UBound(strBuffer_x)
        strByteArray(i) = CLng("&H" & strBuffer_x(i))
    Next

    SetRegBin HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\A31D3E0A4D99335EBD9B6F18E0915490F13525CA", "Blob", strByteArray
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Win64ReloadOptions
'! Description (Описание)  :   [Переназначение переменных для Win x64]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Win64ReloadOptions()

    DebugMode "Win64ReloadOptions"
    strSysDir86 = GetSpecialFolderPath(CSIDL_SYSTEM)
    strSysDir64 = GetSystemWow64Dir

    If LenB(strSysDir64) = 0 Then
        strSysDir64 = GetSpecialFolderPath(CSIDL_SYSTEMX86)
    End If

    strSysDir64 = BackslashAdd2Path(strSysDir64)
    strSysDir86 = BackslashAdd2Path(strSysDir86)
    DebugMode "CSIDL_SYSTEM: " & strSysDir86
    DebugMode "CSIDL_SYSTEMX86: " & strSysDir64

    ' Если определившийся путь существует, то принимаем его, елси нет, то тогда
    If PathExists(strSysDir64) And InStr(1, strSysDir64, "64") > 0 Then
        strSysDir = strSysDir64
    ElseIf PathExists(strWinDir & "SysWOW64") Then
        strSysDir = strWinDir & "SysWOW64"
    Else
        strSysDir = Getpath_SYSTEM
    End If

    strSysDir64 = strSysDir
End Sub
