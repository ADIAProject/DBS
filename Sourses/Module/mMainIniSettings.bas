Attribute VB_Name = "mMainIniSettings"
Option Explicit

Public mbPatnAbs                         As Boolean     ' Пути до программ являются абсолютными, используется в настройках frmOptions
Public mbAllFolderDRVNotExist            As Boolean     ' Все каталоги с пакетами драйверов, указанные в настройках не существуют

' Настройки программы считываемые из ini-файла
Public strSysIni                         As String      ' рабочий файл настроек
Public mbLoadIniTmpAfterRestart          As Boolean     ' Загружать ini из временной папки
Public lngOSCount                        As Long        ' Количество ОС поддерживаемых программой
Public lngOSCountPerRow                  As Long        ' Количество ОС, отображаемое на одной строке
Public lngUtilsCount                     As Long        ' Количество Утилит, прописанных в настройках
Public strDevconCmdPath                  As String      ' Пути до исполняемых файлов и других рабочих файлов - .\Tools\DevCon\devcon_c.cmd
Public strArh7zExePath                   As String      ' Пути до исполняемых файлов и других рабочих файлов - Выбирается, в зависимости от разрядности, из параметров выше
Public strArh7zExePath86                 As String      ' Пути до исполняемых файлов и других рабочих файлов - .\Tools\Arc\7za.exe
Public strArh7zExePath64                 As String      ' Пути до исполняемых файлов и других рабочих файлов - .\Tools\Arc\7za64.exe
Public strArh7zParam1                    As String
Public strArh7zParam2                    As String
Public strArh7zSFXPATH                   As String
Public strArh7zSFXConfigPath             As String
Public strArh7zSFXConfigPathEn           As String
Public strDPInstExePath64                As String      ' Пути до исполняемых файлов и других рабочих файлов - .\Tools\DPInst\DPInst64.exe
Public strDPInstExePath86                As String      ' Пути до исполняемых файлов и других рабочих файлов - .\Tools\DPInst\DPInst.exe
Public strDPInstExePath                  As String      ' Пути до исполняемых файлов и других рабочих файлов - Выбирается, в зависимости от разрядности, из параметров выше
Public mbDelTmpAfterClose                As Boolean
Public mbUpdateCheck                     As Boolean
Public mbUpdateCheckBeta                 As Boolean
Public mbUpdateToolTip                   As Boolean
Public miStartMode                       As Long
Public mbRecursion                       As Boolean
Public mbSaveSizeOnExit                  As Boolean
Public strExcludeHWID                    As String
Public lngStartModeTab2                  As Long        ' стартовая вкладка для типов пакетов
Public strThisBuildBy                    As String      ' Добавляем к описанию в главном окне в названии программы
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
Public mbCompareDrvVerByDate             As Boolean     ' Сравнение версий драйверов по дате
Public mbLoadUnSupportedOS               As Boolean     ' Грузить\негрузить драйвера для несовместимых ОС
Public mbAutoInfoAfterDelDRV             As Boolean     ' Автообновление конфигурации при удалении драйвера
Public mbDateFormatRus                   As Boolean     ' Автообновление конфигурации при удалении драйвера
'Public mbCreateRestorePoint              As Boolean     ' Переменная для режима создания точки восстановления
'Public mbCreateRestorePointDone          As Boolean     ' Флаг, показывающий что точка восстановления уже создавалась ранее
Public mbDisableDEP                      As Boolean     ' Переменная для определения выключения DEP
Public mbHideOtherProcess                As Boolean     ' Скрывать сторонние процессы при запуске
Public mbDP_Is_aFolder                   As Boolean     ' Пакеты драйверов в виде папок - т.е распакованные пакет драйверов
Public mbStartMaximazed                  As Boolean     ' Запускать программу развернутой на весь экран
Public mbTempPath                        As Boolean     ' Использовать альтернативный каталог Temp - т.е задается вручную
Public strAlternativeTempPath            As String      ' Путь для альтернативного каталога Temp
Public mbDpInstLegacyMode                As Boolean     ' Параметры DPinst
Public mbDpInstPromptIfDriverIsNotBetter As Boolean     ' Параметры DPinst
Public mbDpInstForceIfDriverIsNotBetter  As Boolean     ' Параметры DPinst
Public mbDpInstSuppressAddRemovePrograms As Boolean     ' Параметры DPinst
Public mbDpInstSuppressWizard            As Boolean     ' Параметры DPinst
Public mbDpInstQuietInstall              As Boolean     ' Параметры DPinst
Public mbDpInstScanHardware              As Boolean     ' Параметры DPinst
Public mbSearchOnStart                   As Boolean     ' Искать новые устройства при запуске программы
Public lngPauseAfterSearch               As Long        ' Паузка после поиска новых устройства при запуске программы
Public mbCalcDriverScore                 As Boolean     ' Использовать при анализе драйверов балл найденного драйвера, на основании различных условий
Public mbCompatiblesHWID                 As Boolean     ' Использовать для поиска подходящих драйверов секцию CompatiblesHWID, берется из реестра
Public mbSearchCompatibleDriverOtherOS   As Boolean     ' Искать подходящие драйвера на всех вкладках, а не только на подобранной
Public lngSortMethodShell                As Long        ' Параметр указывающий применять метод сортировки массива
Public lngCompatiblesHWIDCount           As Long        ' Глубина поиска совместимых HWID
Public mbMatchHWIDbyDPName               As Boolean     ' Анализ имени файла для определния совместимости драйвера
Public lngMainFormWidth                  As Long        ' Ширина основной формы
Public lngMainFormHeight                 As Long        ' Высота основной формы
Public lngButtonWidth                    As Long        ' Ширина кнопки
Public lngButtonHeight                   As Long        ' Высота кнопки
Public lngButtonLeft                     As Long        ' Отступ слева для кнопки
Public lngButtonTop                      As Long        ' Отступ сверху для кнопки
Public lngBtn2BtnLeft                    As Long        ' Интервал между кнопками по горизонтали
Public lngBtn2BtnTop                     As Long        ' Интервал между кнопками по вертикали
Public lngStatusBtnStyle                 As Long        ' Стиль кнопки пакета драйверов
Public lngStatusBtnStyleColor            As Long        ' Цвет оформления кнопки пакета драйверов
Public lngStatusBtnBackColor             As Long        ' Цвет оформления кнопки пакета драйверов
Public lngFreeSpaceSysDrive              As Long        ' Свободное место на жестком диске

'Public strImageMenuName                  As String
'Public mbExMenu                           As Boolean ' Расширенное меню

'-------------------- Константы размеров форм и кнопок  ------------------'
Public Const lngMainFormWidthMin         As Long = 13000    ' Минимальные значения размеров формы
Public Const lngMainFormHeightMin        As Long = 6500     ' Минимальные значения размеров формы
'Public Const lngButtonWidthMin           As Long = 1500     ' Минимальные значения размеров кнопки - Ширина
'Public Const lngButtonHeightMin          As Long = 350      ' Минимальные значения размеров кнопки - Высота
Private Const lngMainFormWidthDef        As Long = 13000    ' Дефолтные значения размеров формы
Private Const lngMainFormHeightDef       As Long = 8400     ' Дефолтные значения размеров формы
'Private Const lngButtonWidthDef          As Long = 2150     ' Дефолтные значения размеров кнопки - Ширина
'Private Const lngButtonHeightDef         As Long = 550      ' Дефолтные значения размеров кнопки - Высота
'Private Const lngButtonLeftDef           As Long = 100      ' Дефолтные значения размеров кнопки - Отступ слева для кнопки
'Private Const lngButtonTopDef            As Long = 100      ' Дефолтные значения размеров кнопки - Отступ сверху для кнопки
'Private Const lngBtn2BtnLeftDef          As Long = 100      ' Дефолтные значения размеров кнопки - Интервал между кнопками по горизонтали
'Private Const lngBtn2BtnTopDef           As Long = 100      ' Дефолтные значения размеров кнопки - Интервал между кнопками по вертикали

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateIni
'! Description (Описание)  :   [Сохранение настроек в ини файл если файла нет]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub CreateIni()

    If FileExists(strSysIni) = False Then
        If mbIsDriveCDRoom Then
            strSysIni = strWorkTempBackSL & strSettingIniFile
            MsgBox "File " & strSettingIniFile & " is not Exist!" & vbNewLine & _
                   "This program works from CD\DVD, so we create temporary " & strSettingIniFile & "-file" & vbNewLine & _
                   strSysIni, vbInformation + vbApplicationModal, strProductName
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

        'Секция Debug
        IniWriteStrPrivate "Debug", "DebugEnable", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogPath", "%WINDIR%\Logs\" & strProjectName & "Log\", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogName", strProjectName & "-LOG_%DATE%.txt", strSysIni
        IniWriteStrPrivate "Debug", "CleenHistory", "1", strSysIni
        IniWriteStrPrivate "Debug", "DetailMode", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLog2AppPath", "0", strSysIni
        IniWriteStrPrivate "Debug", "Time2File", "0", strSysIni
        'Секция Arc
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
        IniWriteStrPrivate "OS_3", "Ver", "6.0;6.1;6.2;6.3;6.4;10.0", strSysIni
        IniWriteStrPrivate "OS_3", "drpFolder", "drivers\vista_7_8_10\x32\", strSysIni
        IniWriteStrPrivate "OS_3", "is64bit", "0", strSysIni

        'Секция OS_4
        IniWriteStrPrivate "OS_4", "Ver", "6.0;6.1;6.2;6.3;6.4;10.0", strSysIni
        IniWriteStrPrivate "OS_4", "drpFolder", "drivers\vista_7_8_10\x64\", strSysIni
        IniWriteStrPrivate "OS_4", "is64bit", "1", strSysIni
        'Секция MainForm
        IniWriteStrPrivate "MainForm", "Width", CStr(lngMainFormWidthDef), strSysIni
        IniWriteStrPrivate "MainForm", "Height", CStr(lngMainFormHeightDef), strSysIni
        IniWriteStrPrivate "MainForm", "StartMaximazed", "0", strSysIni
        IniWriteStrPrivate "MainForm", "SaveSizeOnExit", "0", strSysIni
        IniWriteStrPrivate "MainForm", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "MainForm", "FontSize", "9", strSysIni
        IniWriteStrPrivate "MainForm", "HighlightColor", "32896", strSysIni
        
        'Секция Buttons
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

        ' Приводим Ini файл к читабельному виду
        NormIniFile strSysIni
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetMainIniParam
'! Description (Описание)  :   [Получение настроек из ини файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMainIniParam() As Boolean

    Dim ii         As Long
    Dim cntOsInIni As Integer


    GetMainIniParam = True
    
    '[Description]
    strThisBuildBy = GetIniValueString(strSysIni, "Description", "BuildBy", vbNullString)
    'strThisBuildBy = "www.SamLab.Ws"
    '[Debug]
    ' Путь до лог файла
    strDebugLogPathTemp = GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%WINDIR%\Logs\" & strProjectName & "Log\")
    strDebugLogPath = PathCollect(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%WINDIR%\Logs\" & strProjectName & "Log\"))
    ' Имя лог-файла
    strDebugLogNameTemp = GetIniValueString(strSysIni, "Debug", "DebugLogName", strProjectName & "-LOG_%DATE%.txt")
    strDebugLogName = ExpandFileNameByEnvironment(GetIniValueString(strSysIni, "Debug", "DebugLogName", strProjectName & "-LOG_%DATE%.txt"))
    ' Записывать время в лог-файл
    mbDebugTime2File = GetIniValueBoolean(strSysIni, "Debug", "Time2File", 0)
    ' Создавать лог-файл в подпапке "logs" программы
    mbDebugLog2AppPath = GetIniValueBoolean(strSysIni, "Debug", "DebugLog2AppPath", 0)
    ' Активация отладки
    mbDebugStandart = GetIniValueBoolean(strSysIni, "Debug", "DebugEnable", 0)
    ' Очистка истории
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
    
    ' Деталировка отладки - по умолчанию=1
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

    If StrComp(strAlternativeTempPath, "no_key") = 0 Then
        strAlternativeTempPath = strWinTemp
    End If

    ' при необходимости используем альтернативный temp
    mbTempPath = GetIniValueBoolean(strSysIni, "Main", "AlternativeTemp", 0)

    If mbTempPath Then
        strAlternativeTempPath = PathCollect(strAlternativeTempPath)
        If mbDebugStandart Then DebugMode "AlternativeTempPath: " & strAlternativeTempPath

        If PathExists(strAlternativeTempPath) Then
            strWinTemp = strAlternativeTempPath
            strWorkTemp = strWinTemp & strProjectName

            ' Если нет, то создаем временный рабочий каталог
            If PathExists(strWorkTemp) = False Then
                CreateNewDirectory strWorkTemp
            End If

        Else
            If mbDebugStandart Then DebugMode "Alternative TempPath not Exist. Use Windows Temp"
        End If
    End If

    mbLoadIniTmpAfterRestart = GetIniValueBoolean(strSysIni, "Main", "LoadIniTmpAfterRestart", 0)
    mbDisableDEP = GetIniValueBoolean(strSysIni, "Main", "DisableDEP", 1)
    '--------------------- Получение путей до файлов ---------------------
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
    ' Шрифт Кнопок
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
    ' Отображать дату версии в формате dd/mm/yyyy
    mbDateFormatRus = GetIniValueBoolean(strSysIni, "Main", "DateFormatRus", 0)
    ' папка со скином
    strImageMainName = GetIniValueString(strSysIni, "Main", "IconMainSkin", "Standart")
    ' расширенное меню
    'mbExMenu = GetIniValueBoolean(strSysIni, "Main", "ExMenu", 1)
    'strImageMenuName = GetIniValueString(strSysIni, "Main", "IconMenuSkin", "Standart")
    ' Скрывать прочие процессы
    mbHideOtherProcess = GetIniValueBoolean(strSysIni, "Main", "HideOtherProcess", 1)
    ' Тихая регистрация DLL
    mbSilentDLL = GetIniValueBoolean(strSysIni, "Main", "SilentDll", 0)
    ' Показывать напоминание об обновлении (всплывающее окно)
    mbUpdateToolTip = GetIniValueBoolean(strSysIni, "Main", "UpdateToolTip", 1)

    ' Стартовый режим
    miStartMode = GetIniValueLong(strSysIni, "Main", "StartMode", 2)
    ' Выделять всю группу
    mbCheckAllGroup = GetIniValueBoolean(strSysIni, "Main", "CheckAllGroup", 1)
    ' Выделять всю группу
    mbListOnlyGroup = GetIniValueBoolean(strSysIni, "Main", "ListOnlyGroup", 1)
    'Блокирование окна listview ghb бекапировании
    mbBlockListOnBackup = GetIniValueBoolean(strSysIni, "Main", "BlockListOnBackup", 1)
    'Режим архивирования по умолчанию
    miArchMode = GetIniValueLong(strSysIni, "Main", "ArchMode", 0)

    Exit Function
    
ExitFunc:
    GetMainIniParam = False
End Function

