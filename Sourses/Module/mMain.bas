Attribute VB_Name = "mMain"
Option Explicit

'�������� ��������� ���������
Public Const strDateProgram         As String = "09/01/2017"
Public Const strVerProgram          As String = "5.01.09"

'�������� ���������� ������� (��������, ������ � �.�)
Public strProductName               As String
Public strProductVersion            As String
'�������� ��������� ������� (��������, �����)
Public Const strProjectName         As String = "DBS"
Public Const strUrl_MainWWWSite     As String = "http://adia-project.net/"                   ' �������� ���� �������
Public Const strUrl_MainWWWForum    As String = "http://adia-project.net/forum/index.php"    ' �������� ����� �������
Public Const strUrlOsZoneNetThread  As String = "http://forum.oszone.net/thread-190814.html" ' ����� ��������� DBS �� ����� Oszone.net

'��������� ����� �������� ��������� � ����� �������� (�������� �������� ��� ��������������� ���� ��� ������ �������)
Public Const strToolsLang_Path      As String = "Tools\DBS\Lang"         ' ������� � ��������� �������
Public Const strToolsDocs_Path      As String = "Tools\DBS\Docs"         ' ������� � ������������� �� ���������
Public Const strToolsGraphics_Path  As String = "Tools\DBS\Graphics"     ' ������� � ������������ ��������� ���������
Public Const strSettingIniFile      As String = "DBS.ini"   ' INI-���� �������� ���������

' ������ ������������� ���������� � ����� Donate
Public Const strEULA_Version        As String = "02/02/2010"
Public Const strEULA_MD5RTF         As String = "68da44c8b1027547e4763472e0ecb727"
Public Const strEULA_MD5RTF_Eng     As String = "0cbd9d50eec41b26d24c5465c4be70bc"
Public Const strDONATE_MD5RTF       As String = "97f8178b2af5ba9377f76baf4ff71f78"
Public Const strDONATE_MD5RTF_Eng   As String = "59bbfbf6decbf91023da434cbe940d33"

'�������� ��������� ������� ���������� �� HWID (��� DBS)
Public Type arrHwidsStructDBS
    i0_DriverDesc               As String           ' �������� ���������
    i1_DriverDate               As String           ' ���� ��������
    i2_DriverVersion            As String           ' ������ ��������
    i3_ProviderName             As String           ' ������������� �������� ����������
    i4_ClassName                As String           ' ����� ���������
    i5_Class                    As String           ' ��� ������ ����������
    i6_InfPath                  As String           ' ������������� �������� ����������
    i7_InfSection               As String           ' ������ inf-����� � ������� ������ HWID
    i8_MatchingDeviceId         As String           ' ����������� ��������
    i9_ClassID                  As String           ' ID ������ ����������
End Type

'�������� ��������� ������� ��� �������������� �� (��� DBS)
Public Type arrOSStructDBS
    Ver                             As String           ' ������ ��
    is64bit                         As Long             ' 64-������ ��
    drpFolder                       As String           ' ������� � �������� ��������� (������������� ����)
    drpFolderFull                   As String           ' ������� � �������� ��������� (������ ����)
    DPFolderNotExist                As Boolean          ' ������� �� ���������
End Type

'������� ������
Public arrHwidsLocal()              As arrHwidsStructDBS   ' ������ ���������� � ��������� ���������
Public arrOSList()                  As arrOSStructDBS      ' ������ �������������� ��

'���� �� ��������� ��������� � ������ ������� ������
Public strWorkTemp                  As String           ' ������� ��������� �������
Public strWorkTempBackSL            As String           ' ������� ��������� �������   + \
Public strWinTemp                   As String           ' ��������� ��������� ������� + \
Public strWinDir                    As String           ' ��������� ������� Windows   + \
Public strSysDir                    As String           ' ��������� ������� System32  + \
Public strSysDir64                  As String           ' ��������� ������� Windows\System32  + \
Public strSysDir86                  As String           ' ��������� ������� Windows\Wow64  + \
Public strSysDirCatRoot             As String           ' c:\Windows\System32\catroot\
Public strSysDirDrivers             As String           ' ��������� ������� Windows\System32\drivers  + \
Public strSysDirDrivers64           As String           ' ��������� ������� Windows\Wow64\drivers  + \
Public strSysDirDRVStore            As String           ' ��������� ������� System32\DriverStore\
Public strSysDrive                  As String           ' ��������� ����
Public strWinDirHelp                As String           ' c:\Windows\Help\
Public strInfDir                    As String           ' c:\Windows\inf\

'���������� � ������� ������������ � ���� ���������
Public mbFirstStart                 As Boolean          ' ���� ����������� �������� ������� ���������
Public mbIsDriveCDRoom              As Boolean          ' ����, ����������� ��� ������� ���� �������� CDRoom
Public mbAddInList                  As Boolean '����� ������ � ��������� listview - ���� �������� ���� ����������
Public lngLastIdOS                     As Long '����� ���������� �������� � ������ ��
Public mbRestartProgram             As Boolean          ' ������ ����������� ���������
Public mbCheckAllGroup              As Boolean
Public mbListOnlyGroup              As Boolean
Public miArchMode                   As Long
Public mbBackFolderPredefine        As Boolean
Public mbBlockListOnBackup          As Boolean
Public strFrmMainCaptionTemp        As String           ' ����� �������� �����
Public strFrmMainCaptionTempDate    As String           ' ����� �������� ����� - ���� ������ ���������
Public lngArchNameMode              As Long
Public strArchNameCustom            As String

' ���������� ��� ����������� ������ �����
Public strCompModel                 As String
Public strMB_Model                  As String
Public strMB_Manufacturer           As String
Public strCompName                  As String
Public mbIsNotebok                  As Boolean ' ���� ��������� �������� ���������



'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ChangeStatusBarText
'! Description (��������)  :   [��������� ������ ���������� ������ � ���������� ����������]
'! Parameters  (����������):   strPanel2Text (String)
'                              strDebugText (String)
'                              mbEqual (Boolean = False)
'                              mbDoEvents (Boolean = True)
'                              strPanel1Text (String)
'!--------------------------------------------------------------------------------
Public Sub ChangeStatusBarText(ByVal strPanel2Text As String, Optional ByVal strPanel1Text As String = vbNullString, Optional ByVal mbDoEvents As Boolean = True)

    If LenB(strPanel2Text) Then

        If frmMain.ctlUcStatusBar1.PanelCount >= 2 Then
            frmMain.ctlUcStatusBar1.PanelText(2) = strPanel2Text
        Else
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel2Text
        End If

        If LenB(strPanel1Text) Then
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel1Text
        End If
        
        If mbDoEvents Then
            DoEvents
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Main
'! Description (��������)  :   [�������� ������� ������� ���������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Main()

    Dim mbShowFormLicence As Boolean
    Dim strSysIniTMP      As String
    Dim strLicenceDate    As String
    ' ���� ������������� ���������� �� �������
    Dim mbIsUserAnAdmin   As Boolean
    ' ������������ �������������?

    On Error Resume Next

    dtStartTimeProg = GetTimeStart

    ' ���������� app.path � ������ � ����������
    GetMyAppProperties

    '��������� ������ �����������
    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    '�������� ��������� ������� windows � ������� windows
    strWinDir = BackslashAdd2Path(Environ$("WINDIR"))
    strWinTemp = BackslashAdd2Path(Environ$("TMP"))
    strSysDrive = Environ$("SYSTEMDRIVE")
    strCompName = SafeFileName(Environ$("COMPUTERNAME"))
    
    lngFreeSpaceSysDrive = GetSystemDiskFreeSpace(strSysDrive)

    If InStr(strWinTemp, strSpace) Then
        strWinTemp = BackslashAdd2Path(PathCombine(strWinDir, "TEMP"))
    End If

    ' ���� ��������� ������� windows  (%windir%\temp)����������
    If PathExists(strWinTemp) = False Then
        MsgBox "Windows TempPath not Exist or Environ %TMP% undefined. Program is exit!!!", vbInformation, strProductName

        GoTo ExitSub

    End If
    '******************************************
    ' ��������� �������� �� ��������� � ������ IDE
    ' ��������� ��� ��������???
    If App.PrevInstance And Not InIDE() Then
        MsgBoxEx "Found a running application 'Drivers Installer Assistant'. If you restart the program from the settings menu, then save the settings, the program waits until the previous session..." & str2vbNewLine & _
                                    "This window will close automatically in 5 seconds. Please wait or click OK", vbExclamation + vbSystemModal, strProductName, 6
        ShowPrevInstance
    Else
        '******************************************
        ' - �������������� ����� WindowsXP
        Call ComCtlsInitIDEStopProtection
        Call InitVisualStyles
    End If

    ' ���� ������� tools ����������
    If PathExists(strAppPathBackSL & "Tools\") = False Then
        MsgBox "Not found the main program subfolder '.\Tools'." & vbNewLine & "Program is exit!!!", vbInformation, strProductName

        GoTo ExitSub

    End If
    
    ' ���� ������� tools ����������
    If PathExists(strAppPathBackSL & "Tools\" & strProjectName & "\") = False Then
        MsgBox "Not found the main program subfolder '.\Tools\" & strProjectName & "'." & vbNewLine & "Program is exit!!!", vbInformation, strProductName

        GoTo ExitSub

    End If

    ' ������� ��������� �������
    strWorkTemp = strWinTemp & strProjectName
    strWorkTempBackSL = BackslashAdd2Path(strWorkTemp)

    ' ���� �� ����� %strProjectName%.ini
    If FileExists(strAppPathBackSL & strSettingIniFile) = False Then
        strSysIni = strAppPathBackSL & "Tools\" & strSettingIniFile
    Else
        strSysIni = strAppPathBackSL & strSettingIniFile
    End If

    ' �������� �� ��������� � CD
    mbIsDriveCDRoom = IsDriveCDRoom
    ' ������� ���� �������� ��� �������������
    CreateIni
    ' ��������� ���� �����������
    LoadLanguageOS

    '��������� �������� �����
    If PathExists(strAppPathBackSL & strToolsLang_Path) Then
        mbMultiLanguage = LoadLanguageList
    End If

    '��������� ����������� ���������
    LocaliseMessage strPCLangCurrentPath
    ' ��������� �������� �� ini-�����
    If Not GetMainIniParam Then
        GoTo ExitSub
    End If

    ' ���� ����� ��������� ��������� ��������� ���� �� ������� ini, �� ������������� ���� ����������
    If mbLoadIniTmpAfterRestart Then
        If GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP", False) Then
            ' Reload Main ini
            strSysIniTMP = GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP_PATH", vbNullString)

            If LenB(strSysIniTMP) Then
                If PathExists(strSysIniTMP) Then
                    strSysIni = strSysIniTMP
                    ' ���������� ������������ ��������
                    GetMainIniParam
                End If
            End If
        End If
    End If

    ' ������� ��������� ������� �������
    If PathExists(strWorkTemp) = False Then
        CreateNewDirectory strWorkTemp
    End If

    '����������� �������� �����
    If PathExists(strAppPathBackSL & strToolsLang_Path) Then
        mbMultiLanguage = LoadLanguageList
    End If

    '����������� ����������� ���������
    LocaliseMessage strPCLangCurrentPath
    'strPathImageStatusButton = strAppPathBackSL & strToolsGraphics_Path & "\StatusButton\"
    strPathImageMain = strAppPathBackSL & strToolsGraphics_Path & "\Main\"
    'strPathImageMenu = strAppPathBackSL & strToolsGraphics_Path & "\Menu\"
    GetImageSkinPath
    ' ������� ���-�������
    MakeCleanHistory
    ' �������� ������� ������� ������� ���������
    GetWorkArea
    ' ��������� �� ������ � �����������
    If LenB(Command) Then
        ' ������ �������� ������ �������
        If CmdLineParsing Then
            ' ���� ������� CmdLineParsing=True, �� ��������� ����� �� ����������
            GoTo ExitSub
        End If
        
    End If

    If APIFunctionPresent("IsUserAnAdmin", "shell32.dll") Then
        mbIsUserAnAdmin = IsUserAnAdmin
        'mbIsUserAnAdmin = IsUserAnAdministrator
    Else
        If mbDebugStandart Then DebugMode vbTab & "APIFunctionPresent: " & IsUserAnAdmin & "=" & False
        mbIsUserAnAdmin = True
    End If

    If Not mbDebugTime2File Then
        If mbDebugStandart Then DebugMode "Current Date: " & Now()
    End If

    If mbDebugStandart Then DebugMode _
              "Version: " & strProductName & vbNewLine & _
              "Build: " & strDateProgram & vbNewLine & _
              "ExeName: " & strAppEXEName & ".exe" & vbNewLine & _
              "AppWork: " & strAppPath & vbNewLine & _
              "is User an Admin?: " & mbIsUserAnAdmin

    If mbIsUserAnAdmin Then
        ' ���������� � ������ ��� ����������, ��� ��� �� exe-�����
        If mbDebugStandart Then DebugMode "SaveSert2Reestr"
        SaveSert2Reestr
    Else

        If Not mbRunWithParam Then
            If MsgBox(strMessages(138), vbYesNo + vbQuestion, strProductName) = vbNo Then
                GoTo ExitSub
            End If
        End If
    End If

    If mbDebugStandart Then DebugMode _
              "WinDir: " & strWinDir & vbNewLine & _
              "TmpDir: " & strWinTemp & vbNewLine & _
              "WorkTemp: " & strWorkTemp & vbNewLine & _
              "FreeSpace: " & lngFreeSpaceSysDrive & " MB" & vbNewLine & _
              "IsDriveCDRoom: " & mbIsDriveCDRoom

    If IsWinXPOrLater Then

        ' ����������� windows x64
        mbIsWin64 = OS_Is_x64
        If mbDebugStandart Then DebugMode "OS-is-x64: " & mbIsWin64

        If mbIsWin64 Then
            Win64ReloadOptions
        End If
        
    End If

    ' Disable DEP for current process
    If mbDisableDEP Then
        SetDEPDisable
    End If

    ' ����������� ������� ���������
    If Not RegisterAddComponent Then
        GoTo ExitSub
    End If

    If mbDebugStandart Then DebugMode _
              "OsCurrentVersion: " & strOSCurrentVersion & vbNewLine & _
              "Architecture: " & strOSArchitecture & vbNewLine & _
              "OS Language: ID=" & strPCLangID & " Name=" & strPCLangEngName & "(" & strPCLangLocaliseName & ")"

    ' ���������� ��������� �������� Windows
    regParam = GetRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\Internet Explorer", "Version")
    strSysDir86 = Getpath_SYSTEM
    strSysDir = strSysDir86
    strSysDirCatRoot = strSysDir86 & "CatRoot\"
    strSysDirDrivers = strSysDir86 & "drivers\"
    strInfDir = strWinDir & "inf\"
    strWinDirHelp = strWinDir & "help\"


    strSysDirDRVStore = strSysDir86 & "DRVSTORE\"
    If IsWinVistaOrLater Then
        strSysDirDRVStore = strSysDir86 & "DriverStore\FileRepository\"
    End If

    If APIFunctionPresent("IsAppThemed", "uxtheme.dll") Then
        mbAppThemed = IsAppThemed
        If mbDebugStandart Then DebugMode "IsAppThemed: " & mbAppThemed
    Else
        If mbDebugStandart Then DebugMode vbTab & "APIFunctionPresent: " & IsAppThemed & "=" & False
    End If

    mbAeroEnabled = IsAeroEnabled
    If mbDebugStandart Then DebugMode "IsAeroEnabled : " & mbAeroEnabled
    ' �������� ����������� ����������� ������ �������� ��� �������������
    SetVideoMode
    GetWorkArea
    
    ' �������� ��� ������������� ����������� �����/��������
    strCompModel = GetMBInfo()
    If mbDebugStandart Then DebugMode _
              "PreDefined PC isNotebook: " & mbIsNotebook & vbNewLine & _
              "Notebook/Motherboard Model: " & strCompModel & vbNewLine & _
              "SystemDrive: " & strSysDrive & vbNewLine & _
              "SysDir: " & strSysDir & vbNewLine & _
              "SysDir86: " & strSysDir86 & vbNewLine & _
              "SysDir64: " & strSysDir64 & vbNewLine & _
              "IE Version: " & regParam

    ' ������ ����������� ��� ��� "������" ������ ���������, ����� ��� ������� ��������� ����� � ������ ��������
    mbFirstStart = True
    
    ' ���� ������ ��������� ��������� �� � �����������, ��....
    If Not mbRunWithParam Then
    ' ����� ������������� ����������
        strLicenceDate = GetSetting(App.ProductName, "Licence", "EULA_DATE", strEULA_Version)
        mbShowFormLicence = GetSetting(App.ProductName, "Licence", "Show at Startup", True)
        If mbShowFormLicence Then
            If Not mbEULAAgree Then
                mbShowFormLicence = StrComp(strLicenceDate, strEULA_Version) <> 0
            End If
        End If
    End If
    'Because Ambient.UserMode does not report IDE behavior properly, we use our own UserMode tracker.  Many thanks to
    ' Kroc of camendesign.com for suggesting this fix.
    g_UserModeFix = True

    If mbShowFormLicence Then
        '��������� ����� ������������� ����������
        frmLicence.Show
    Else
        '��������� �������� �����
        frmMain.Show vbModeless
    End If

ExitSub:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SaveSert2Reestr
'! Description (��������)  :   [��������� ������������ ����������� ��� �������� ���������� �������� ������� ����� exe]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub SaveSert2Reestr()

    Dim strBuffer      As String
    Dim strBuffer_x()  As String
    Dim strByteArray() As Byte
    Dim ii             As Long

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
    strBuffer_x = Split(strBuffer, strComma)

    ReDim strByteArray(UBound(strBuffer_x))

    For ii = 0 To UBound(strBuffer_x)
        strByteArray(ii) = CLng("&H" & strBuffer_x(ii))
    Next

    SetRegBin HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\A31D3E0A4D99335EBD9B6F18E0915490F13525CA", "Blob", strByteArray
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Win64ReloadOptions
'! Description (��������)  :   [�������������� ���������� ��� Win x64]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Win64ReloadOptions()

    If mbDebugStandart Then DebugMode "Win64ReloadOptions"

    strDPInstExePath = strDPInstExePath64
    strArh7zExePath = strArh7zExePath64

    strSysDir86 = GetSpecialFolderPath(CSIDL_SYSTEM)
    strSysDir64 = GetSystemWow64Dir

    If LenB(strSysDir64) = 0 Then
        strSysDir64 = GetSpecialFolderPath(CSIDL_SYSTEMX86)
    End If

    strSysDir64 = BackslashAdd2Path(strSysDir64)
    strSysDir86 = BackslashAdd2Path(strSysDir86)
    If mbDebugStandart Then DebugMode "CSIDL_SYSTEM: " & strSysDir86
    If mbDebugStandart Then DebugMode "CSIDL_SYSTEMX86: " & strSysDir64

    ' ���� �������������� ���� ����������, �� ��������� ���, ���� ���, �� �����
    If PathExists(strSysDir64) And InStr(1, strSysDir64, "64") > 0 Then
        strSysDir = strSysDir64
    ElseIf PathExists(strWinDir & "SysWOW64") Then
        strSysDir = strWinDir & "SysWOW64"
    Else
        strSysDir = Getpath_SYSTEM
    End If

    strSysDir64 = strSysDir
End Sub
