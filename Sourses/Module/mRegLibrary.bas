Attribute VB_Name = "mRegLibrary"
Option Explicit

' Переменные регистрации внешних компонентов
'TabCtl32.ocx
'Private Const strTabCtl32Reference              As String = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}"
'Private Const strTabCtl32Ver                    As String = "1.1"
'Private Const strTabCtl32VerDll                 As String = "6.1.98.16"
'MSFLXGRD.ocx
Private Const strMSFLXGRDReference As String = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}"
Private Const strMSFLXGRDVer       As String = "1.0"
Private Const strMSFLXGRDVerDll    As String = "6.1.98.14"

'RICHTX32.ocx
Private Const strRICHTX32Reference As String = "{3B7C8863-D78F-101B-B9B5-04021C009402}"
Private Const strRICHTX32Ver       As String = "1.2"
Private Const strRICHTX32VerDll    As String = "6.1.98.16"

'VBScript.dll
Private Const strVBScriptReference As String = "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}"
Private Const strVBScriptVer       As String = "5.5"
Private Const strVBScriptVerDll    As String = "5.8.6001.18700"

'Capicom.dll
Private Const strCAPICOMReference  As String = "{BD26B198-EE42-4725-9B23-AFA912434229}"
Private Const strCAPICOMVer        As String = "2.1"
Private Const strCAPICOMVerDll     As String = "2.1.0.2"

'Windows Script 5.7
'XP - http://www.microsoft.com/downloads/details.aspx?displaylang=en&FamilyID=47809025-d896-482e-a0d6-524e7e844d81
'2003 - http://www.microsoft.com/downloads/details.aspx?familyid=F00CB8C0-32E9-411D-A896-F2CD5EF21EB4&displaylang=en
'2000 - http://www.microsoft.com/downloads/details.aspx?familyid=C03D3E49-B40E-4CA1-A0C7-CC135EC4D2BE&displaylang=en
'http://support.microsoft.com/?kbid=951978
'! -----------------------------------------------------------
'!  Функция     :  APIFunctionPresent
'!  Переменные  :  ByVal FunctionName As String, ByVal DLLName As String
'!  Описание    :  Проверка на поддержку функции Api в текущей винде
'! -----------------------------------------------------------
Public Function APIFunctionPresent(ByVal FunctionName As String, ByVal DLLName As String) As Boolean

    Dim lHandle As Long
    Dim lAddr   As Long
    Dim FreeLib As Boolean

    lHandle = GetModuleHandle(DLLName)

    If lHandle = 0 Then
        lHandle = LoadLibraryEx(DLLName, 0&, DONT_RESOLVE_DLL_REFERENCES)
        FreeLib = True
    End If

    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)

        If FreeLib Then
            FreeLibrary lHandle
        End If
    End If

    APIFunctionPresent = (lAddr <> 0)
    DebugMode "***APIFunctionPresent: " & FunctionName & "=" & APIFunctionPresent, 2
End Function

'! -----------------------------------------------------------
'!  Функция     :  DLLOCX
'!  Переменные  :  ByVal hWnd As Long, ByVal Path As String, ByVal Register As Boolean
'!  Описание    :  Регистрация компонента ActiveX при необходимости (файл берется из ресурсов)
'! -----------------------------------------------------------
Private Function DLLOCX(ByVal lngHWnd As Long, ByVal Path As String, ByVal Register As Boolean) As Boolean

    Dim LB As Long
    Dim PA As Long

    On Error Resume Next

    LB = LoadLibrary(Path)

    If Not LB = 0 Then
        If Register Then
            PA = GetProcAddress(LB, "DllRegisterServer")
        Else
            PA = GetProcAddress(LB, "DllUnregisterServer")
        End If

        If PA = 0 Then
            GoTo FreeLib
        End If

        CallWindowProc PA, lngHWnd, ByVal 0&, ByVal 0&, ByVal 0&
        DLLOCX = True
FreeLib:
        FreeLibrary LB

        On Error GoTo FreeLib

    End If
End Function

' Регистрация внешних компонент
Public Sub RegisterAddComponent()

    DebugMode "RegisterAddComponent - Start"

    '    If RegOCX(strWorkTemp & "\TabCtl32.OCX", strTabCtl32Reference, strTabCtl32Ver, strTabCtl32VerDll) = False Then
    '        If MsgBox("System OCX or DLL: 'TabCtl32.OCX'" & vbNewLine & _
    '                   vbNewLine & _
    '                   strMessages(8), vbYesNo + vbQuestion, strProductName) = vbNo Then
    '            End
    '        End If
    '    End If
    '
    '    DebugMode "RegisterAddComponent - *****************Check Next File********************"
    If RegOCX(strWorkTemp & "\MSFLXGRD.OCX", strMSFLXGRDReference, strMSFLXGRDVer, strMSFLXGRDVerDll) = False Then
        If MsgBox("System OCX or DLL: 'MSFLXGRD.OCX'" & vbNewLine & vbNewLine & strMessages(8), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If
    End If

    DebugMode "RegisterAddComponent - *****************Check Next File********************"

    If RegOCX(strWorkTemp & "\RICHTX32.OCX", strRICHTX32Reference, strRICHTX32Ver, strRICHTX32VerDll) = False Then
        If MsgBox("System OCX or DLL: 'RICHTX32.OCX'" & vbNewLine & vbNewLine & strMessages(8), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If
    End If

    DebugMode "RegisterAddComponent - *****************Check Next File********************"

    If RegOCX(strWorkTemp & "\vbscript.dll", strVBScriptReference, strVBScriptVer, strVBScriptVerDll) = False Then
        If MsgBox("System OCX or DLL: 'vbscript.dll'" & vbNewLine & vbNewLine & strMessages(8), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If
    End If

    DebugMode "RegisterAddComponent - *****************Check Next File********************"

    If RegOCX(strWorkTemp & "\capicom.dll", strCAPICOMReference, strCAPICOMVer, strCAPICOMVerDll) = False Then
        If MsgBox("System DLL: 'capicom.dll'" & vbNewLine & vbNewLine & strMessages(8), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If
    End If

    DebugMode "RegisterAddComponent - Finish"
End Sub

'! -----------------------------------------------------------
'!  Функция     :  RegOCX
'!  Переменные  :  strPathOCX As String, strReference As String, strVerOcx
'!  Описание    :  Функция регистрация внешнего компонента
'! -----------------------------------------------------------
Public Function RegOCX(ByVal strPathOcx As String, _
                       ByVal strReference As String, _
                       ByVal strVerOcx As String, _
                       ByVal strVerFileOcx As String) As Boolean

    Dim strPathOcxSystem        As String
    Dim strPathOcxSystemSPFiles As String
    Dim strPathOcxSystemCache   As String
    Dim strOCXFileName          As String
    Dim strVersionFile          As String
    Dim strResultCompare        As String
    Dim mboolForceReg           As Boolean
    Dim MsgRet                  As Integer

    regParam = GetRegString(HKEY_CLASSES_ROOT, "TypeLib\" & strReference & vbBackslash & strVerOcx & "\0\win32", vbNullString)
    strOCXFileName = FileNameFromPath(strPathOcx)
    DebugMode "***" & strReference & ": Path to " & strOCXFileName & " from Reestr: " & regParam
    strPathOcxSystem = strWinDir & "system32\" & strOCXFileName
    strPathOcxSystemCache = strWinDir & "system32\dllcache\" & strOCXFileName
    strPathOcxSystemSPFiles = strWinDir & "ServicePackFiles\i386\" & strOCXFileName

    If InStr(1, regParam, "vbscript.dll\3", vbTextCompare) > 0 Then
        regParam = Replace$(regParam, "vbscript.dll\3", "vbscript.dll", , , vbTextCompare)
    End If

    If PathFileExists(regParam) = 0 Then
        If InStr(1, regParam, "vbscript.dll\3", vbTextCompare) = 0 Then
            If PathFileExists(strPathOcxSystem) = 0 Or LenB(regParam) = 0 Then
StartRegOCXForce:

                If GetBinaryFileFromResource(strPathOcx, "OCX_" & FileName_woExt(strOCXFileName), "CUSTOM") Then
                    DebugMode "***" & strReference & ": BinaryFileFromResourse: True"

                    'разрегистрация файла при необходимости
                    If mboolForceReg Then
                        RegOCX = DLLOCX(0, regParam, False)
                        DebugMode "***" & strReference & ": Unregistred old dll: " & RegOCX
                    End If

                    'Копируем новый файл с заменой в системный кеш
                    If StrComp(ExtFromFileName(strPathOcx), "dll", vbTextCompare) = 0 Then
                        If PathFileExists(strPathOcxSystemSPFiles) = 1 Then
                            If CopyFileTo(strPathOcx, strPathOcxSystemSPFiles) = False Then
                                RegOCX = False
                                DebugMode "***" & strReference & ": CopyOcxFileToServicePackFiles: False: " & strPathOcxSystemSPFiles
                                Exit Function
                            Else
                                DebugMode "***" & strReference & ": CopyOcxFileToServicePackFiles: True: " & strPathOcxSystemSPFiles
                            End If
                        End If

                        If PathFileExists(strPathOcxSystemCache) = 1 Then
                            If CopyFileTo(strPathOcx, strPathOcxSystemCache) = False Then
                                RegOCX = False
                                DebugMode "***" & strReference & ": CopyOcxFileToSystemCache: False: " & strPathOcxSystemCache
                                Exit Function
                            Else
                                DebugMode "***" & strReference & ": CopyOcxFileToSystemCache: True: " & strPathOcxSystemCache
                            End If
                        End If
                    End If

                    'Копируем новый файл с заменой
                    If CopyFileTo(strPathOcx, strPathOcxSystem) = False Then
                        RegOCX = False
                        DebugMode "***" & strReference & ": CopyOcxFileToSystem: False: " & strPathOcxSystem
                        Exit Function
                    Else
                        DebugMode "***" & strReference & ": CopyOcxFileToSystem: True: " & strPathOcxSystem
                    End If

                    strPathOcx = strPathOcxSystem
                Else
                    RegOCX = False
                    DebugMode "***" & strReference & ": BinaryFileFromResourse: False"
                    Exit Function
                End If
            End If

            DebugMode "***" & strReference & ": Registration in system"
            RegOCX = DLLOCX(0, strPathOcxSystem, True)
        Else
            RegOCX = True
            regParam = Replace$(regParam, "vbscript.dll\3", "vbscript.dll", , , vbTextCompare)

            If PathFileExists(regParam) = 1 Then
                ' Сравниваем версии библиотек
                strVersionFile = FSO.GetFileVersion(regParam)

                'Прерываем обновление библиотеки, если винда выше 2003
                If InStr(1, regParam, "vbscript.dll", vbTextCompare) > 0 Then
                    If strOsCurrentVersion >= "6.0" Or mboolIsWin64 Then
                        DebugMode "***" & strReference & ": Update file for your operating system does not provide"
                        Exit Function
                    End If
                End If

                strResultCompare = CompareByVersion(strVersionFile, strVerFileOcx)

                ' Анализ итога сравнения
                If strResultCompare = "<" Then
                    If Not mboolRunWithParam Then
                        If mboolSilentDLL Then
                            MsgRet = vbYes
                        Else
                            MsgRet = MsgBox("System OCX or DLL:" & strPathOcxSystem & vbNewLine & strVersionFile & " < " & strVerFileOcx & vbNewLine & vbNewLine & strMessages(106), vbYesNoCancel + vbQuestion, strProductName)
                        End If

                        If MsgRet = vbYes Then
                            mboolForceReg = True
                            GoTo StartRegOCXForce
                        ElseIf MsgRet = vbNo Then
                            Exit Function
                        Else
                            End
                        End If

                    Else
                        mboolForceReg = True
                        GoTo StartRegOCXForce
                    End If
                End If
            End If
        End If

    Else
        RegOCX = True

        If PathFileExists(regParam) = 1 And LenB(regParam) > 0 Then
            ' Сравниваем версии библиотек
            strVersionFile = FSO.GetFileVersion(regParam)

            'Прерываем обновление библиотеки, если винда выше 2003 или 64x
            If InStr(1, regParam, "vbscript.dll", vbTextCompare) > 0 Then
                If strOsCurrentVersion >= "6.0" Or mboolIsWin64 Then
                    DebugMode "***" & strReference & ": Update file for your operating system does not provide"
                    Exit Function
                End If
            End If

            strResultCompare = CompareByVersion(strVersionFile, strVerFileOcx)

            ' Анализ итога сравнения
            If strResultCompare = "<" Then
                If Not mboolRunWithParam Then
                    If mboolSilentDLL Then
                        MsgRet = vbYes
                    Else
                        MsgRet = MsgBox("System OCX or DLL:" & strPathOcxSystem & vbNewLine & strVersionFile & " < " & strVerFileOcx & vbNewLine & vbNewLine & strMessages(106), vbYesNoCancel + vbQuestion, strProductName)
                    End If

                    If MsgRet = vbYes Then
                        mboolForceReg = True
                        GoTo StartRegOCXForce
                    ElseIf MsgRet = vbNo Then
                        Exit Function
                    Else
                        End
                    End If

                Else
                    mboolForceReg = True
                    GoTo StartRegOCXForce
                End If
            End If
        End If

        DebugMode "***" & strReference & ": is already present in system - File version: " & strVersionFile
    End If

    DebugMode "***" & strReference & ": Result=" & RegOCX
End Function
