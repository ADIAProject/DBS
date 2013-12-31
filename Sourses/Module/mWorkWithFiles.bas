Attribute VB_Name = "mWorkWithFiles"
Option Explicit

' Переменные для работы с файловой системой
Public FSO                  As New Scripting.FileSystemObject

Private Root                As String
Private xFOL                As Folder
Private xFile               As File
Private strFileListInFolder As String

'Удаление слэша на конце
Public Function BacklashDelFromPath(ByVal strPath As String) As String

    strPath = strPath & strDoubleNull
    PathRemoveBackslash strPath
    BacklashDelFromPath = TrimNull(strPath)
End Function

'Добавление слэша на конце
Public Function BackslashAdd2Path(ByVal strPath As String) As String

    strPath = strPath & strDoubleNull
    PathAddBackslash strPath
    BackslashAdd2Path = TrimNull(strPath)
End Function

'! -----------------------------------------------------------
'!  Функция     :  cmdPathClick
'!  Переменные  :  vForm As Form, strTextBox As String, strDialog As String, mboolFile As Boolean
'!  Возвр. знач.:  As String
'!  Описание    :  Открыть диалоговое окно и выбрать файл или папку
'! -----------------------------------------------------------
Public Function cmdPathClick(vForm As Form, _
                             strStartDirectory As String, _
                             strDialog As String) As String

    Dim strStartDir As String
    Dim strPath As String

    strStartDir = PathCollect(strStartDirectory)

    If InStr(strStartDir, ".") > 0 Then
        strStartDir = PathNameFromPath(strStartDir)
    End If
    
    ' выбор каталога
    DebugMode "Show Open Dialog with Promt='" & strDialog & "' for InitDir=" & strStartDir, 1
    strPath = fBrowseForFolder(hWnd_Owner:=vForm.hwnd, sPrompt:=strDialog, WhatBr:=BIF_DEFAULT, InitDir:=strStartDir, CenterOnScreen:=True, TopMost:=True)

    If strPath <> vbNullString Then
        cmdPathClick = strPath
    End If
End Function

Public Function CompareFilesByHash(ByVal strFirstFile As String, ByVal strSecondFile As String) As Boolean

    Dim mobjSHAFirst     As New cSHA1
    Dim mobjSHASecond    As New cSHA1
    Dim strDataSHAFirst  As String
    Dim strDataSHASecond As String
    Dim lngResult        As Long
    Dim abytData()       As Byte
    Dim abytHashed()     As Byte

    If PathFileExists(strFirstFile) = 1 Then

        With mobjSHAFirst
            ' convert file location to byte array 
            abytData() = StrConv(strFirstFile, vbFromUnicode)
            ' hash data and return as Byte array
            abytHashed() = .HashFile(abytData())
            ' convert byte array to string data
            strDataSHAFirst = StrConv(CStr(abytHashed()), vbUnicode)
        End With

        strDataSHAFirst = CalcHashFile(strFirstFile, CAPICOM_HASH_ALGORITHM_SHA1)
    End If

    If PathFileExists(strSecondFile) = 1 Then

        With mobjSHASecond
            ' convert file location to byte array 
            abytData() = StrConv(strSecondFile, vbFromUnicode)
            ' hash data and return as Byte array
            abytHashed() = .HashFile(abytData())
            ' convert byte array to string data
            strDataSHASecond = StrConv(CStr(abytHashed()), vbUnicode)
        End With
    End If

    lngResult = StrComp(strDataSHAFirst, strDataSHASecond, vbTextCompare)

    If lngResult = 0 Then
        CompareFilesByHash = True
    Else
        CompareFilesByHash = False
    End If
End Function

Public Function CompareFilesByHashCAPICOM(ByVal strFirstFile As String, ByVal strSecondFile As String) As Boolean

    Dim strDataSHAFirst  As String
    Dim strDataSHASecond As String
    Dim lngResult        As Long

    If PathFileExists(strFirstFile) = 1 Then
        strDataSHAFirst = CalcHashFile(strFirstFile, CAPICOM_HASH_ALGORITHM_SHA1)
    End If

    If PathFileExists(strSecondFile) = 1 Then
        strDataSHASecond = CalcHashFile(strSecondFile, CAPICOM_HASH_ALGORITHM_SHA1)
    End If

    lngResult = StrComp(strDataSHAFirst, strDataSHASecond, vbTextCompare)

    If lngResult = 0 Then
        CompareFilesByHashCAPICOM = True
    Else
        CompareFilesByHashCAPICOM = False
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  CopyFileTo
'!  Переменные  :  PathFrom As String, PathTo As String
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Скопирует файл 'PathFrom' в директорию 'CopyFileTo', Если файл существует, то он будет перезаписан новым файлом.
'! -----------------------------------------------------------
Public Function CopyFileTo(ByVal PathFrom As String, ByVal PathTo As String) As Boolean

    Dim ret As Long

    If PathFileExists(PathFrom) Then
        ' Для всех файлов, сброс атрибута только для чтения, и системный если есть
        ResetReadOnly4File PathTo
        ' Собственно копирование
        'Если вы хотите, чтобы новый файл не записывался на место старого, замените 'False' на 'True'
        ret = CopyFile(PathFrom, PathTo, False)

        If ret <> 0 Then
            CopyFileTo = True
            ' Сброс атрибута только для чтения, если есть
            ResetReadOnly4File PathTo
        Else
            CopyFileTo = False
            MsgBox strMessages(42) & vbNewLine & "From: " & PathFrom & vbNewLine & "To:" & PathTo & vbNewLine & "Error: №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError), vbExclamation, strProductName
            DebugMode "***Copy file: False: " & PathFrom & " Error: №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
        End If

    Else
        CopyFileTo = False
        DebugMode "***Copy file: False : " & PathFrom & " Error: №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
    End If
End Function

Public Sub CreateNewDirectory(ByVal NewDirectory As String)

    Dim SecAttrib  As SECURITY_ATTRIBUTES
    Dim sPath      As String
    Dim iCounter   As Integer
    Dim sTempDir   As String
    Dim ret        As Long
    Dim retLasrErr As Long

    sPath = BackslashAdd2Path(NewDirectory)
    iCounter = 1

    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
        sTempDir = Left$(sPath, iCounter)
        iCounter = iCounter + 1

        'create directory
        With SecAttrib
            .lpSecurityDescriptor = &O0
            .bInheritHandle = False
            .nLength = Len(SecAttrib)
        End With

        ret = CreateDirectory(sTempDir, SecAttrib)

        If ret = 0 Then
            retLasrErr = err.LastDllError

            If PathFileExists(sTempDir) = 0 Then
                DebugMode "***CreateDirectory: False : " & sTempDir & " Error: №" & retLasrErr & " - " & ApiErrorText(retLasrErr)
            End If
        End If

    Loop
End Sub

Public Function DeleteFiles(ByVal PathFile As String) As Boolean

    Dim ret       As Long
    Dim retDllerr As Long

    ret = DeleteFile(PathFile)
    DeleteFiles = CBool(ret)

    If ret = 0 Then
        If PathFileExists(PathFile) = 1 Then

            On Error GoTo errhandler

            FSO.DeleteFile PathFile, True
        End If

        retDllerr = err.LastDllError

        If PathFileExists(PathFile) = 1 Then
            DebugMode "***DeleteFiles: False : " & PathFile & " Error: №" & retDllerr & " - " & ApiErrorText(retDllerr)
        End If
    End If

    Exit Function
errhandler:
    retDllerr = err.LastDllError
    DebugMode "***DeleteFiles: False : " & PathFile & " Error: №" & err.Number & ": " & err.Description
    DebugMode "***DeleteFiles: False : " & PathFile & " Error: №" & retDllerr & " - " & ApiErrorText(retDllerr)
    err.Clear

    Resume Next

End Function

'! -----------------------------------------------------------
'!  Функция     :  DelFolderBackUp
'!  Переменные  :
'!  Описание    :  Удаление временного каталога, если включена опция
'! -----------------------------------------------------------
Public Sub DelFolderBackUp(ByVal strFolderPath As String)

    Dim ret As Long

    On Error Resume Next

    DebugMode "DelFolder-Start: " & strFolderPath

    If PathFileExists(strFolderPath) = 1 Then
        DelRecursiveFolder strFolderPath
    End If

    If PathFileExists(strFolderPath) = 1 Then
        ret = RemoveDirectory(strFolderPath)

        If ret = 0 Then
            DebugMode "***RemoveDirectory: False : " & strFolderPath & " Error: №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
        End If
    End If

    On Error GoTo 0

    DebugMode "DelFolder-End"
End Sub

'! -----------------------------------------------------------
'!  Функция     :  DelRecursiveFolder
'!  Переменные  :  Folder As String
'!  Описание    :
'! -----------------------------------------------------------
Public Sub DelRecursiveFolder(ByVal Folder As String)

    Dim retDelete As Long
    Dim retStrMsg As String

    Root = BacklashDelFromPath(Folder)
    DebugMode "***DeleteFolder: " & Root

    If PathFileExists(Root) = 1 Then
        SearchFilesInRoot Root, ALL_FILES, True, False, True
        Set xFOL = FSO.GetFolder(Root)

        If xFOL.Files.Count > 0 Then

            For Each xFile In xFOL.Files
                DeleteFiles CStr(xFile.Path)
            Next
        End If

        ' Получение списка каталогов подлежащих удалению
        If PathFileExists(Root) = 1 Then
            GetAllFolderInRoot Root, True
        End If

        If PathFileExists(Root) = 1 Then
            GetAllFolderInRoot Root, True
            retDelete = DelTree(Root)

            If mboolDebugEnable Then

                Select Case retDelete

                    Case 0
                        retStrMsg = "Deleted"

                    Case -1
                        retStrMsg = "Invalid Directory"

                    Case Else
                        retStrMsg = "An Error was occured"
                End Select

                DebugMode "***DeleteFolder: " & " Result: " & retStrMsg
            End If
        End If
    End If
End Sub

'! -----------------------------------------------------------
'!  Функция     :  DelTemp
'!  Переменные  :
'!  Описание    :  Удаление временного каталога, если включена опция
'! -----------------------------------------------------------
Public Sub DelTemp()

    On Error Resume Next

    DebugMode "DelTemp-Start"

    If PathFileExists(strWorkTemp) = 1 Then
        DelRecursiveFolder strWorkTemp
    End If

    If PathFileExists(strWorkTemp) = 1 Then
        RemoveDirectory strWorkTemp
    End If

    On Error GoTo 0

    DebugMode "DelTemp-End"
End Sub

Private Function DelTree(ByVal strDir As String) As Long

    Dim X          As Long
    Dim intAttr    As Integer
    Dim strAllDirs As String
    Dim strFile    As String
    Dim ret        As Long
    Dim retLasrErr As Long

    DelTree = -1

    On Error Resume Next

    strDir = Trim$(strDir)

    If LenB(strDir) > 0 Then
        If Right$(strDir, 1) = vbBackslash Then
            strDir = Left$(strDir, Len(strDir) - 1)
        End If

        If InStr(strDir, "\") > 0 Then
            intAttr = GetAttr(strDir)

            If (intAttr And vbDirectory) Then
                strDir = BackslashAdd2Path(strDir)
                strFile = Dir$(strDir & ALL_FILES, vbSystem Or vbDirectory Or vbHidden)

                Do While Len(strFile)

                    If strFile <> "." Then
                        If strFile <> ".." Then
                            intAttr = GetAttr(strDir & strFile)

                            If (intAttr And vbDirectory) Then
                                strAllDirs = strAllDirs & strFile & vbNullChar
                            Else

                                If intAttr <> vbNormal Then
                                    SetAttr strDir & strFile, vbNormal

                                    If err Then
                                        DelTree = err.Number
                                    End If

                                    Exit Function
                                End If

                                DeleteFiles strDir & strFile

                                If err Then
                                    DelTree = err.Number
                                End If

                                Exit Function
                            End If
                        End If
                    End If

                    strFile = Dir
                Loop

                Do While Len(strAllDirs)
                    X = InStr(strAllDirs, vbNullChar)
                    strFile = Left$(strAllDirs, X - 1)
                    strAllDirs = Mid$(strAllDirs, X + 1)
                    X = DelTree(strDir & strFile)

                    If X Then
                        DelTree = X
                    End If

                Loop
                ret = RemoveDirectory(strDir)

                If ret = 0 Then
                    retLasrErr = err.LastDllError

                    If PathFileExists(strDir) = 0 Then
                        DebugMode "***RemoveDirectory: False : " & strDir & " Error: №" & retLasrErr & " - " & ApiErrorText(retLasrErr)
                    End If

                    DelTree = retLasrErr
                Else
                    DelTree = 0
                End If

                If err Then
                    DelTree = err.Number
                Else
                    DelTree = 0
                End If

                On Error GoTo 0

            End If
        End If
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  ExtFromFileName
'!  Переменные  :  FileName As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить расширение файла из пути или имени файла
'! -----------------------------------------------------------
Public Function ExtFromFileName(ByVal FileName As String) As String

    Dim intLastSeparator As Long

    intLastSeparator = InStrRev(FileName, ".")

    If intLastSeparator > 0 Then
        ExtFromFileName = Right$(FileName, Len(FileName) - intLastSeparator)
    Else
        ExtFromFileName = vbNullString
    End If
End Function

Public Function FileisReadOnly(ByVal PathFile As String) As Boolean

    FileisReadOnly = GetAttr(PathFile) And vbReadOnly
End Function

Public Function FileisSystemAttr(PathFile As String) As Boolean

    FileisSystemAttr = GetAttr(PathFile) And vbSystem
End Function

'! -----------------------------------------------------------
'!  Функция     :  FileName_woExt
'!  Переменные  :  FileName As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить имя файла без расширения, зная имя файла
'! -----------------------------------------------------------
Public Function FileName_woExt(ByVal FileName As String) As String

    Dim intLastSeparator As Long

    FileName_woExt = FileName

    If LenB(FileName) > 0 Then
        intLastSeparator = InStrRev(FileName, ".")

        If intLastSeparator > 0 Then
            FileName_woExt = Left$(FileName, intLastSeparator - 1)
        End If
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  FileNameFromPath
'!  Переменные  :  FilePath As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить имя файла из полного пути
'! -----------------------------------------------------------
Public Function FileNameFromPath(ByVal FilePath As String) As String

    Dim intLastSeparator As Long

    FileNameFromPath = FilePath

    If LenB(FilePath) > 0 Then
        intLastSeparator = InStrRev(FilePath, "\")

        If intLastSeparator >= 0 Then
            FileNameFromPath = Right$(FilePath, Len(FilePath) - intLastSeparator)
        End If
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  GetAllFileInFolder
'!  Переменные  :  xFolder As String, RealDelete As Boolean, Optional ExtFile As String
'!  Описание    :  Получение всех файлов в выбранном каталоге
'! -----------------------------------------------------------
Public Sub GetAllFileInFolder(ByVal xFolder As String, _
                              RealDelete As Boolean, _
                              Optional ExtFile As String, _
                              Optional ByVal mboolRecursFolder As Boolean = True)

    Dim strExtFile_x() As String
    Dim strExtFile     As String
    Dim strExtFileReal As String
    Dim strTemp        As String
    Dim strTempAll     As String
    Dim i              As Long

    DebugMode "******GetAllFileInFolder-Start: " & xFolder, 2

    If Not PathFileExists(xFolder) = 0 Then
        Set xFOL = FSO.GetFolder(xFolder)
        strExtFile_x = Split(ExtFile, ";")

        For Each xFile In xFOL.Files

            ' Если требуется удаление файла, то удалаем
            If RealDelete Then

                On Error GoTo errhandler

                xFile.Delete True
            Else
                ' Если расширение файл INF, то добавляем путь файла в массив
                strTemp = vbNullString

                Dim lngLBound As Long
                Dim lngUBound As Long

                lngLBound = LBound(strExtFile_x)
                lngUBound = UBound(strExtFile_x)

                For i = lngLBound To lngUBound
                    strExtFile = UCase$(strExtFile_x(i))
                    strExtFileReal = UCase$(ExtFromFileName(xFile.Path))

                    If strExtFile = strExtFileReal Then
                        strTemp = xFile.Path
                    End If

                Next

                If strTemp <> vbNullString Then
                    strTempAll = AppendStr(strTempAll, strTemp, ";")
                End If
            End If

        Next

        ' Если требуется удаление каталога, то удалаем
        If RealDelete Then

            With xFOL

                If .Files.Count = 0 Then
                    If .SubFolders.Count = 0 Then
                        .Delete True
                    End If
                End If
            End With

        Else

            If LenB(strTempAll) > 0 Then
                DebugMode "******ListFiles in Folder '" & CStr(xFOL.Name) & "': " & vbNewLine & "*****************************************" & vbNewLine & strTempAll & vbNewLine & "*****************************************"
                strFileListInFolder = AppendStr(strFileListInFolder, strTempAll, ";")
            End If
        End If

        If mboolRecursFolder Or RealDelete Then
            ' Проверяем есть ли подкаталоги в каталоге
            GetAllFolderInRoot xFolder, RealDelete, ExtFile
        End If
    End If

    DebugMode "******GetAllFileInFolder-End", 2
    Exit Sub
errhandler:
    DebugMode "***GetAllFileInFolder: False : " & xFolder & " Error: №" & err.Number & ": " & err.Description
    err.Clear

    Resume Next

End Sub

'! -----------------------------------------------------------
'!  Функция     :  GetAllFolderInFolder
'!  Переменные  :  rootFolder As String
'!  Описание    :  Получение всех подкаталогов в выбранном каталоге
'! -----------------------------------------------------------
Public Function GetAllFolderInFolder(ByVal rootFolder As String) As Variant

    Dim xFolder       As Folder
    Dim strListFolder As String

    DebugMode "******GetAllFolderInFolder-Start: "

    If PathFileExists(rootFolder) = 1 Then
        Set xFOL = FSO.GetFolder(rootFolder)

        If xFOL.SubFolders.Count > 0 Then

            For Each xFolder In xFOL.SubFolders
                strListFolder = AppendStr(strListFolder, CStr(xFolder.Name), ";")
            Next
        End If

        GetAllFolderInFolder = Split(strListFolder, ";", , vbTextCompare)
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  GetAllFolderInRoot
'!  Переменные  :  rootFolder As String, RealDelete As Boolean, Optional ExtFile As String
'!  Описание    :  Получение всех подкаталогов в выбранном каталоге
'! -----------------------------------------------------------
Private Sub GetAllFolderInRoot(ByVal rootFolder As String, ByVal RealDelete As Boolean, Optional ExtFile As String)

    Dim xFolder As Folder

    If PathFileExists(rootFolder) = 1 Then
        Set xFOL = FSO.GetFolder(rootFolder)

        If xFOL.SubFolders.Count > 0 Then

            For Each xFolder In xFOL.SubFolders
                DebugMode "******Analize Subfolder: " & CStr(xFolder.Path), 2
                GetAllFileInFolder xFolder.Path, RealDelete, ExtFile
            Next
        End If
    End If
End Sub

'! -----------------------------------------------------------
'!  Функция     :  GetEnviron
'!  Переменные  :  strEnv As String, Optional mboolCollectFull As Boolean = False
'!  Возвр. знач.:  As String
'!  Описание    :  Получение переменной системного окружения
'! -----------------------------------------------------------
Public Function GetEnviron(ByVal strEnv As String, Optional ByVal mboolCollectFull As Boolean = False) As String

    Dim strTemp        As String
    Dim strTempEnv     As String
    Dim strNumPosition As Long

    strNumPosition = InStr(1, strEnv, "%")

    If strNumPosition > 0 Then
        strTemp = Mid$(strEnv, strNumPosition + 1, Len(strEnv) - strNumPosition)
        strNumPosition = InStr(1, strTemp, "%")

        If strNumPosition > 0 Then
            strTemp = Mid$(strTemp, 1, strNumPosition - 1)
        End If
    End If

    strTempEnv = Environ$(strTemp)

    If mboolCollectFull Then
        GetEnviron = Replace$(strEnv, "%" & strTemp & "%", strTempEnv, , , vbTextCompare)
    Else
        GetEnviron = strTempEnv
    End If

    DebugMode "******GetEnviron: %" & strTemp & "%=" & strTempEnv
    DebugMode "******GetEnviron-End"
End Function

Public Function GetUniqueTempFile() As String

    Dim ll_Buffer       As Long
    Dim ls_TempFileName As String

    ll_Buffer = 255
    ls_TempFileName = Space$(255)
    ll_Buffer = GetTempFileName(strWinTemp, "xdia", 0, ls_TempFileName)

    'xxx is a three letter prefix - can be anything you want.
    '3rd parameter (0 above) is uUnique...If uUnique is nonzero, the function appends the hexadecimal string to lpPrefixString to form the temporary filename. In this case, the function does not create the specified file, and does not test whether the filename is unique.
    'If uUnique is zero, the function uses a hexadecimal string derived from the current system time. In this case, the function uses different values until it finds a unique filename, and then it creates the file in the lpPathName directory.
    If ll_Buffer = 0 Then
        MsgBox strMessages(44) & vbNewLine & strWinTemp, vbCritical, strProductName
    Else
        ls_TempFileName = Left$(ls_TempFileName, ll_Buffer)
        GetUniqueTempFile = ls_TempFileName
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  IsDriveCDRoom
'!  Переменные  :
'!  Описание    :  Проверка на запск программы с CD\DVD
'! -----------------------------------------------------------
Public Function IsDriveCDRoom() As Boolean

    Dim strDriveName As String
    Dim xDrv         As Drive

    IsDriveCDRoom = False
    strDriveName = Mid$(strAppPath, 1, 2)

    ' Проверяем на запуск из сети
    If StrComp(strDriveName, "\\", vbTextCompare) <> 0 Then
        'получаем тип диска
        Set xDrv = FSO.GetDrive(strDriveName)

        If xDrv.DriveType = CDRom Then
            IsDriveCDRoom = True
        End If
    End If
End Function

Public Function IsPathAFolder(ByVal sPath As String) As Boolean

    'Verifies that a path is a valid
    'directory, and returns True (1) if
    'the path is a valid directory,
    'or False otherwise. The path must
    'exist.
    'If the path is a directory on the
    'local machine, PathIsDirectory returns
    '16 (the file attribute for a folder).
    'If the path is a directory on a server
    'share, PathIsDirectory returns 1.
    'If it is neither PathIsDirectory returns 0.
    Dim Result As Long

    Result = PathIsDirectory(sPath)
    IsPathAFolder = (Result = vbDirectory) Or (Result = 1)
End Function

'! -----------------------------------------------------------
'!  Функция     :  MoveFileTo
'!  Переменные  :  PathFrom As String, PathTo As String
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Скопирует файл 'PathFrom' в директорию 'PathTo', Если файл существует, то он будет перезаписан новым файлом.
'! -----------------------------------------------------------
Public Function MoveFileTo(PathFrom As String, PathTo As String) As Boolean

    Dim ret As Long

    If StrComp(PathFrom, PathTo, vbTextCompare) <> 0 Then
        If PathFileExists(PathFrom) Then
            ' Для всех файлов, сброс атрибута только для чтения, и системный если есть
            ResetReadOnly4File PathTo
            ' Собственно копирование
            'Если вы хотите, чтобы новый файл не записывался на место старого, замените 'False' на 'True'
            ret = MoveFile(PathFrom, PathTo)

            If ret <> 0 Then
                MoveFileTo = True
                ' Сброс атрибута только для чтения, если есть
                ResetReadOnly4File PathTo
            Else
                MoveFileTo = False
                MsgBox strMessages(42) & vbNewLine & "From: " & PathFrom & vbNewLine & "To:" & PathTo & vbNewLine & "Error: №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError), vbExclamation, strProductName
                DebugMode "***Move file: False: " & PathFrom & " Error: №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
            End If

        Else
            MoveFileTo = False
            DebugMode "***Move file: False : " & PathFrom & " Error: №" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
        End If

    Else
        DebugMode "***Move file: Source and Destination are identicaly (" & PathFrom & " ; " & PathTo & ")"
    End If
End Function

Public Function ParserInf4Strings(ByVal strInfFilePath As String, ByVal strSearchString As String) As String

    Dim StringHash     As Scripting.Dictionary
    Dim objInfFile     As TextStream
    Dim RegExpStrSect  As RegExp
    Dim RegExpStrDefs  As RegExp
    Dim MatchesStrSect As MatchCollection
    Dim MatchesStrDefs As MatchCollection
    Dim objMatch       As Match
    Dim objMatch1      As Match
    Dim regex_strsect  As String
    Dim regex_strings  As String
    Dim r_beg          As String
    Dim r_identS       As String
    Dim r_str          As String
    Dim filecontent    As String
    Dim key            As String
    Dim Value          As String
    Dim r              As Boolean
    Dim i              As Long
    Dim Strings        As String
    Dim valval         As String
    Dim varname        As String
    Dim strFileDBSize  As String
    Dim pos            As Long

    r_beg = "^[ \t]*"
    r_identS = "([^; \t\r\n][^;\t\r\n]*[^; \t\r\n])"
    r_str = "(?:""([^\r\n""]*)""|([^\r\n;]*))"
    regex_strsect = r_beg & "\[strings\](?:([\s\S]*?)" & r_beg & "(?=\[)|([\s\S]*))"
    ' variable = "str"
    regex_strings = r_beg & r_identS & "[ \t]*=[ \t]*" & r_str
    ' Init regexps
    Set RegExpStrSect = CreateObject("VBScript.RegExp")

    With RegExpStrSect
        .Pattern = regex_strsect
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
        ' Note: "XP Alternative (by Greg)\D\3\M\A\12\prime.inf" has two [strings] sections
    End With

    Set RegExpStrDefs = CreateObject("VBScript.RegExp")

    With RegExpStrDefs
        .Pattern = regex_strings
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    ' Read INF file
    filecontent = vbNullString
    strFileDBSize = FileSizeApi(strInfFilePath)

    If InStr(1, strFileDBSize, "0 ", vbTextCompare) = 1 Then
        DebugMode "******DevParserByRegExp: File is zero = 0 bytes:" & strInfFilePath
    Else
        Set objInfFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(strInfFilePath, 1, False, -2)
        filecontent = objInfFile.ReadAll()
        objInfFile.Close
    End If

    ' Find [strings] section
    Strings = vbNullString
    Set StringHash = CreateObject("Scripting.Dictionary")
    StringHash.CompareMode = 1
    Set MatchesStrSect = RegExpStrSect.Execute(filecontent)

    If MatchesStrSect.Count >= 1 Then
        Set objMatch = MatchesStrSect.Item(0)
        Strings = objMatch.SubMatches(0) & objMatch.SubMatches(1)
        Set MatchesStrDefs = RegExpStrDefs.Execute(Strings)

        For i = 0 To MatchesStrDefs.Count - 1
            Set objMatch1 = MatchesStrDefs.Item(i)
            key = objMatch1.SubMatches(0)
            Value = objMatch1.SubMatches(1)

            If Value = vbNullString Then
                Value = objMatch1.SubMatches(2)
            End If

            r = StringHash.Exists(key)

            If Not r Then
                StringHash.Add key, Value
                StringHash.Add "%" & key & "%", Value
            End If

        Next
    End If

    ' Собственно ищем саму переменную
    pos = InStr(strSearchString, "%")

    If pos > 0 Then
        varname = Mid$(strSearchString, pos, InStrRev(strSearchString, "%"))
        valval = StringHash.Item(varname)

        If valval = vbNullString Then
            DebugMode "ParserInf4Strings: Error in inf: Cannot find '" & strSearchString & "'"
        Else
            ParserInf4Strings = Replace$(strSearchString, varname, valval)
        End If
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  PathNameFromPath
'!  Переменные  :  FilePath As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить путь к файлу из полного пути
'! -----------------------------------------------------------
Public Function PathNameFromPath(FilePath As String) As String

    Dim intLastSeparator As Long

    intLastSeparator = InStrRev(FilePath, "\")
    PathNameFromPath = Left$(FilePath, intLastSeparator)
End Function

Public Sub ResetReadOnly4File(ByVal strPathFile As String)

    If PathFileExists(strPathFile) Then
        If FileisReadOnly(strPathFile) Then
            SetAttr strPathFile, vbNormal
        End If

        If FileisSystemAttr(strPathFile) Then
            SetAttr strPathFile, vbNormal
        End If
    End If
End Sub

'# function to replace special chars to create dirs correctly #
Public Function SafeDir(ByVal str As String) As String

    Dim r As String

    r = str
    r = Replace$(r, "\", "_")
    r = Replace$(r, "/", "-")
    r = Replace$(r, "*", "_")
    r = Replace$(r, ":", "_")
    r = Replace$(r, ";", "_")
    r = Replace$(r, "?", "_")
    r = Replace$(r, ">", "_")
    r = Replace$(r, "<", "_")
    r = Replace$(r, "|", "_")
    r = Replace$(r, "@", "_")
    r = Replace$(r, "'", "")
    r = Replace$(r, " ", "_")
    r = Replace$(r, "_-_", "_")
    r = Replace$(r, "(R)", "_")
    r = Replace$(r, "___", "_")
    r = Replace$(r, "__", "_")
    r = Trim$(r)
    SafeDir = r
End Function

'# function to replace special chars to create dirs correctly #
Public Function SafeFileName(ByVal strString) As String

    ' Заменяем VbTab
    strString = Replace$(strString, vbTab, vbNullString, , , vbTextCompare)
    strString = TrimNull(strString)

    ' Отбрасываем все после ","
    If InStr(1, strString, ",", vbTextCompare) > 0 Then
        strString = Left$(strString, InStr(strString, ",") - 1)
    End If

    ' Отбрасываем все после ";"
    If InStr(1, strString, ";", vbTextCompare) > 0 Then
        strString = Left$(strString, InStr(strString, ";") - 1)
    End If

    strString = Trim$(TrimNull(strString))
    SafeFileName = strString
End Function

'# function to discover dirs with inf code #
Public Function WhereIsDir(ByVal str As String, ByVal strInfFilePath As String) As String

    Dim cDir                As String
    Dim str_x()             As String
    Dim mboolAdditionalPath As Boolean

    If InStr(1, str, ";", vbTextCompare) > 0 Then
        str_x = Split(str, ";", , vbTextCompare)
        str = Trim$(str_x(0))
    End If

    If InStr(1, str, ",", vbTextCompare) > 0 Then
        str_x = Split(str, ",", , vbTextCompare)
        mboolAdditionalPath = True
        str = str_x(0)
    End If

    If InStr(1, str, vbNullChar, vbTextCompare) > 0 Then
        str = TrimNull(str)
    End If

    If InStr(1, str, vbTab, vbTextCompare) > 0 Then
        str = Replace$(str, vbTab, vbNullString, , , vbTextCompare)
    End If

    'http://msdn.microsoft.com/en-us/library/ff553598.aspx
    Select Case str

        Case "01"
            cDir = strSysDrive

        Case "10"
            cDir = strWinDir

            'system32 независимо от винды
        Case "11"
            cDir = strSysDir86

        Case "12"
            cDir = strSysDir86 & "Drivers"

        Case "17"
            cDir = strInfDir

        Case "18"
            cDir = strWinDir & "Help"

        Case "20"
            cDir = GetSpecialFolderPath(CSIDL_FONTS)

        Case "21"
            cDir = vbNullString

            'viewer dir
        Case "23"
            cDir = strSysDir86 & "spool\drivers\color"

        Case "24"
            cDir = strSysDrive

        Case "25"
            cDir = vbNullString

            'shared dir
        Case "30"
            cDir = strSysDrive

        Case "50"
            cDir = strWinDir & "system"

        Case "51"
            cDir = strSysDir86 & "Spool"

        Case "52"
            cDir = strSysDir86 & "Spool\Drivers"

        Case "53"
            cDir = vbNullString

            'user profile dir
        Case "54"
            cDir = vbNullString

            ' ntldr.exe dir
        Case "55"
            cDir = strSysDir86 & "spool\prtprocs"

        Case "16384"
            cDir = GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY)

        Case "16386"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAMS)

        Case "16389"
            cDir = GetSpecialFolderPath(CSIDL_MYDOCUMENTS)

        Case "16391"
            cDir = GetSpecialFolderPath(CSIDL_STARTUP)

        Case "16392"
            cDir = GetSpecialFolderPath(CSIDL_RECENT)

        Case "16393"
            cDir = GetSpecialFolderPath(CSIDL_SENDTO)

        Case "16395"
            cDir = GetSpecialFolderPath(CSIDL_STARTMENU)

        Case "16397"
            cDir = GetSpecialFolderPath(CSIDL_MYMUSIC)

        Case "16397"
            cDir = GetSpecialFolderPath(CSIDL_MYVIDEO)

        Case "16400"
            cDir = GetSpecialFolderPath(CSIDL_DESKTOP)

        Case "16403"
            cDir = GetSpecialFolderPath(CSIDL_NETHOOD)

        Case "16404"
            cDir = GetSpecialFolderPath(CSIDL_FONTS)

        Case "16405"
            cDir = GetSpecialFolderPath(CSIDL_TEMPLATES)

        Case "16406"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_STARTMENU)

        Case "16407"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_PROGRAMS)

        Case "16408"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_STARTUP)

        Case "16409"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_DESKTOPDIRECTORY)

        Case "16410"
            cDir = GetSpecialFolderPath(CSIDL_APPDATA)

        Case "16411"
            cDir = GetSpecialFolderPath(CSIDL_PRINTHOOD)

        Case "16412"
            cDir = GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)

        Case "16415"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_FAVORITES)

        Case "16416"
            cDir = GetSpecialFolderPath(CSIDL_INTERNET_CACHE)

        Case "16417"
            cDir = GetSpecialFolderPath(CSIDL_COOKIES)

        Case "16418"
            cDir = GetSpecialFolderPath(CSIDL_HISTORY)

        Case "16419"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_APPDATA)

        Case "16420"
            cDir = strWinDir

        Case "16421"
            cDir = strSysDir86

        Case "16422"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES)

        Case "16423"
            cDir = GetSpecialFolderPath(CSIDL_MYPICTURES)

        Case "16424"
            cDir = GetSpecialFolderPath(CSIDL_PROFILE)

        Case "16425"
            cDir = strSysDir64

        Case "16426"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86)

        Case "16427"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMON)

        Case "16428"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMONX86)

        Case "16429"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_TEMPLATES)

        Case "16430"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_DOCUMENTS)

        Case "16432"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86)

        Case "16437"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_MUSIC)

        Case "16438"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_PICTURES)

        Case "16439"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_VIDEO)

        Case "16440"
            cDir = strWinDir & "resources"

        Case "16441"
            cDir = strWinDir & "resources\0409"

        Case "-1"
            cDir = vbNullString

            ' absolute path
            'http://msdn.microsoft.com/en-us/library/ff560821.aspx
        Case "66000"
            cDir = Getpath_PrinterDriverDirectory

            If LenB(cDir) = 0 Then
                cDir = strSysDir86 & "spool\Drivers\w32x86"
            End If

        Case "66001"
            cDir = Getpath_PrintProcessorDirectory

            If LenB(cDir) = 0 Then
                cDir = strSysDir86 & "spool\prtprocs\w32x86"
            End If

        Case "66002"
            cDir = strSysDir86

        Case "66003"
            cDir = Getpath_PrinterColorDirectory

            If LenB(cDir) = 0 Then
                cDir = strSysDir86 & "spool\drivers\color"
            End If

        Case "66004"
            cDir = strSysDir86 & "spool\Drivers\w32x86"

        Case Else
            cDir = vbNullString
    End Select

    If InStr(1, cDir, vbNullChar, vbTextCompare) > 0 Then
        cDir = TrimNull(cDir)
    End If

    If mboolAdditionalPath Then
        cDir = BackslashAdd2Path(cDir) & Trim$(str_x(1))

        If InStr(cDir, "%") > 0 Then
            cDir = ParserInf4Strings(strInfFilePath, cDir)
        End If
    End If

    cDir = Replace$(cDir, vbTab, vbNullString, , , vbTextCompare)
    cDir = Replace$(cDir, kavichki, vbNullString, , , vbTextCompare)
    cDir = BackslashAdd2Path(cDir)
    WhereIsDir = TrimNull(cDir)
End Function

' процедура получения глобальных переменных путей программы
Public Sub GetCurAppPath()

    strAppPath = App.Path
    strAppPathBackSL = BackslashAdd2Path(strAppPath)
End Sub
