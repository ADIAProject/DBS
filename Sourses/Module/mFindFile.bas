Attribute VB_Name = "mFindFile"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private fp                     As FILE_PARAMS    'holds search parameters
Private fp2                    As FOLDER_PARAMS    'holds search parameters
Private sResultFileList()      As String
Private sResultFileListCount   As Long
Private sResultFolderList()    As String
Private sResultFolderListCount As Long

Public Function FileSizeApi(sSource As String) As String

    Dim WFD   As WIN32_FIND_DATA
    Dim hFile As Long
    Dim sSize As String

    hFile = FindFirstFile(sSource, WFD)

    If hFile <> INVALID_HANDLE_VALUE Then
        sSize = Space$(30)
        StrFormatByteSizeW WFD.nFileSizeLow, WFD.nFileSizeHigh, ByVal StrPtr(sSize), 30
        FileSizeApi = TrimNull(sSize)
    End If

    FindClose hFile
End Function

' Проверка на соответствие условиям поиска
Public Function MatchSpec(sFile As String, sSpec As String) As Boolean

    If sSpec <> vbNullString Then
        MatchSpec = PathMatchSpec(StrPtr(sFile), StrPtr(sSpec))
    End If
End Function

Public Function rgbCopyFiles(ByVal sSourcePath As String, ByVal sDestination As String, ByVal sFiles As String) As Long

    Dim WFD                   As WIN32_FIND_DATA
    Dim SA                    As SECURITY_ATTRIBUTES
    Dim hFile                 As Long
    Dim bNext                 As Long
    Dim copied                As Long
    Dim currFile              As String
    Dim currSourcePath        As String
    Dim lngNumFilesFromFolder As Long

    sSourcePath = BackslashAdd2Path(sSourcePath)
    sDestination = BackslashAdd2Path(sDestination)

    'Create the target directory if it doesn't exist
    If PathFileExists(sDestination) = 0 Then
        Call CreateDirectory(sDestination, SA)
    End If

    'Start searching for files in the Target directory.
    hFile = FindFirstFile(sSourcePath & sFiles, WFD)

    If (hFile = INVALID_HANDLE_VALUE) Then
        'nothing to do, so bail out
        DebugMode "******CopyAllFilesFromFolder: " & sSourcePath & " No " & sFiles & " files found."
        Exit Function
    End If

    'Copy each file to the new directory
    If hFile Then
        Do
            'trim trailing nulls, leaving one to terminate the string
            'currFile = Left$(WFD.cFileName, InStr(WFD.cFileName, Chr$(0)))
            currFile = TrimNull(WFD.cFileName)

            If Asc(WFD.cFileName) <> vbDot Then
                currSourcePath = sSourcePath & currFile

                If Not IsPathAFolder(currSourcePath) Then
                    If MatchSpec(currFile, sFiles) Then
                        'copy the file to the destination directory & increment the count
                        Call CopyFileTo(currSourcePath, sDestination & currFile)
                        copied = copied + 1
                    End If

                Else
                    ' Копируем содержимое архива
                    DebugMode "******CopyFiles from SubFolder: " & currFile
                    lngNumFilesFromFolder = rgbCopyFiles(currSourcePath, sDestination & currFile, ALL_FILES)
                    DebugMode "******CopyFiles SubFolder - count files: " & lngNumFilesFromFolder
                    copied = copied + lngNumFilesFromFolder
                End If
            End If

            'just to check what's happening
            'List1.AddItem sSourcePath & currFile
            'find the next file matching the initial file spec
            bNext = FindNextFile(hFile, WFD)
        Loop Until bNext = 0

    End If

    'Close the search handle
    Call FindClose(hFile)
    'and return the number of files copied
    rgbCopyFiles = copied
End Function

Public Function SearchFilesInRoot(strRootDir As String, _
                                  ByVal strSearchMask As String, _
                                  ByVal mboolSearchRecursion As Boolean, _
                                  ByVal mboolOnlyFirstFile As Boolean, _
                                  Optional mboolDelete As Boolean = False)

    With fp
        .sFileRoot = BackslashAdd2Path(strRootDir)
        .sFileNameExt = strSearchMask
        .bRecurse = mboolSearchRecursion
    End With

    'FP
    SearchForFiles fp.sFileRoot, True, 100, mboolDelete

    If Not mboolDelete Then
        If mboolOnlyFirstFile Then
            SearchFilesInRoot = sResultFileList(0, 0)
        Else
            SearchFilesInRoot = sResultFileList
        End If
    End If
End Function

Public Function SearchFoldersInRoot(strRootDir As String, _
                                    ByVal strSearchMask As String, _
                                    ByVal mboolSearchRecursion As Boolean, _
                                    ByVal mboolOnlyFirstFile As Boolean)

    With fp2
        .sFileRoot = BackslashAdd2Path(strRootDir)
        .sFileNameExt = strSearchMask
        .bRecurse = mboolSearchRecursion
    End With

    'FP
    SearchForFolders fp2.sFileRoot, True, 100

    If mboolOnlyFirstFile Then
        SearchFoldersInRoot = sResultFolderList(0, 0)
    Else
        SearchFoldersInRoot = sResultFolderList
    End If
End Function

Private Sub SearchForFiles(sRoot As String, _
                           ByVal mboolInitial As Boolean, _
                           miMaxCountArr As Long, _
                           Optional mboolDelete As Boolean = False)

    Dim WFD   As WIN32_FIND_DATA
    Dim hFile As Long
    Dim sSize As String

    hFile = FindFirstFile(sRoot & ALL_FILES, WFD)

    If Not mboolDelete Then
        If mboolInitial Then
            sResultFileListCount = 0
            ReDim sResultFileList(1, miMaxCountArr)
        Else
            ReDim Preserve sResultFileList(1, miMaxCountArr)
        End If
    End If

    If hFile <> INVALID_HANDLE_VALUE Then
        Do

            'if a folder, and recurse specified, call
            'method again
            If (WFD.dwFileAttributes And vbDirectory) Then
                If Asc(WFD.cFileName) <> vbDot Then
                    If fp.bRecurse Then
                        SearchForFiles sRoot & TrimNull(WFD.cFileName) & vbBackslash, False, miMaxCountArr, mboolDelete
                    End If
                End If

            Else

                'must be a file..
                If MatchSpec(WFD.cFileName, fp.sFileNameExt) Then
                    If mboolDelete Then
                        DeleteFiles sRoot & TrimNull(WFD.cFileName)
                    Else

                        ' Переопределение массива если превышаем заданную размерность
                        If sResultFileListCount = miMaxCountArr Then
                            miMaxCountArr = 2 * miMaxCountArr
                            ReDim Preserve sResultFileList(1, miMaxCountArr)
                        End If

                        ' Полный путь файла
                        sResultFileList(0, sResultFileListCount) = sRoot & TrimNull(WFD.cFileName)
                        ' размер файла
                        sSize = Space$(30)
                        StrFormatByteSizeW WFD.nFileSizeLow, WFD.nFileSizeHigh, ByVal StrPtr(sSize), 30
                        sResultFileList(1, sResultFileListCount) = TrimNull(sSize)
                        sResultFileListCount = sResultFileListCount + 1
                    End If
                End If
            End If

        Loop While FindNextFile(hFile, WFD)

0   End If

    FindClose hFile

    ' Переопределение массива на реальное кол-во записей
    If Not mboolDelete Then
        If mboolInitial Then
            If sResultFileListCount > 0 Then
                ReDim Preserve sResultFileList(1, sResultFileListCount - 1)
            Else
                ReDim Preserve sResultFileList(1, sResultFileListCount)
            End If
        End If
    End If
End Sub

Private Sub SearchForFolders(sRoot As String, ByVal mboolInitial As Boolean, miMaxCountArr As Long)

    Dim WFD   As WIN32_FIND_DATA
    Dim hFile As Long

    hFile = FindFirstFile(sRoot & ALL_FILES, WFD)

    If mboolInitial Then
        sResultFolderListCount = 0
        ReDim sResultFolderList(1, miMaxCountArr)
    Else
        ReDim Preserve sResultFolderList(1, miMaxCountArr)
    End If

    If hFile <> INVALID_HANDLE_VALUE Then
        Do

            'if a folder, and recurse specified, call
            'method again
            If (WFD.dwFileAttributes And vbDirectory) Then
                If Asc(WFD.cFileName) <> vbDot Then
                    If MatchSpec(WFD.cFileName, fp2.sFileNameExt) Then

                        ' Переопределение массива если превышаем заданную размерность
                        If sResultFolderListCount = miMaxCountArr Then
                            miMaxCountArr = 2 * miMaxCountArr
                            ReDim Preserve sResultFolderList(1, miMaxCountArr)
                        End If

                        ' Полный путь файла
                        sResultFolderList(0, sResultFolderListCount) = sRoot & TrimNull(WFD.cFileName)
                        sResultFolderList(1, sResultFolderListCount) = Mid$(TrimNull(WFD.cFileName), 1, InStrRev(TrimNull(WFD.cFileName), "_", , vbTextCompare) - 1)
                        sResultFolderListCount = sResultFolderListCount + 1
                    End If

                    If fp2.bRecurse Then
                        SearchForFolders sRoot & TrimNull(WFD.cFileName) & vbBackslash, False, miMaxCountArr
                    End If
                End If
            End If

        Loop While FindNextFile(hFile, WFD)

    End If

    FindClose hFile

    ' Переопределение массива на реальное кол-во записей
    If mboolInitial Then
        If sResultFolderListCount > 0 Then
            ReDim Preserve sResultFolderList(1, sResultFolderListCount - 1)
        Else
            ReDim Preserve sResultFolderList(1, sResultFolderListCount)
        End If
    End If
End Sub
