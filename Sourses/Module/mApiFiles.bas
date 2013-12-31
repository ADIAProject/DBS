Attribute VB_Name = "mApiFiles"
Option Explicit

Public Const vbDot                As Integer = 46
Public Const MAX_PATH             As Long = 260
Public Const MAX_PATH_UNICODE = 2 * MAX_PATH - 1
Public Const INVALID_HANDLE_VALUE As Integer = -1
Public Const vbBackslash          As String = "\"
Public Const ALL_FILES            As String = "*.*"
Public Const ForWriting           As Long = 2
Public Const ForAppending         As Long = 8    '‘‡ÈÎ‡ ÓÚÍ˚Ú ‰Îˇ ƒŒ¡¿¬À≈Õ»ﬂ
Public Const ForReading           As Long = 1    '‘‡ÈÎ‡ ÓÚÍ˚Ú ‰Îˇ ◊“≈Õ»ﬂ

Public Type SECURITY_ATTRIBUTES
    nLength                        As Long
    lpSecurityDescriptor           As Long
    bInheritHandle                 As Long
End Type

Public Type FILETIME
    dwLowDateTime                          As Long
    dwHighDateTime                         As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes                       As Long
    ftCreationTime                         As FILETIME
    ftLastAccessTime                       As FILETIME
    ftLastWriteTime                        As FILETIME
    nFileSizeHigh                          As Long
    nFileSizeLow                           As Long
    dwReserved0                            As Long
    dwReserved1                            As Long
    cFileName                              As String * MAX_PATH
    cAlternate                             As String * 14
End Type

Public Type FILE_PARAMS
    bRecurse                               As Boolean
    sFileNameExt                           As String
    sFileRoot                              As String
End Type

Public Type FOLDER_PARAMS
    bRecurse                               As Boolean
    sFileNameExt                           As String
    sFileRoot                              As String
End Type

Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function CopyFile _
               Lib "kernel32.dll" _
               Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                  ByVal lpNewFileName As String, _
                                  ByVal bFailIfExists As Long) As Long

Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function CreateDirectory _
               Lib "kernel32.dll" _
               Alias "CreateDirectoryA" (ByVal lpPathName As String, _
                                         lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal Path As String) As Long
Public Declare Function PathRemoveBackslash _
               Lib "shlwapi.dll" _
               Alias "PathRemoveBackslashA" (ByVal Path As String) As Long

Public Declare Function MoveFile _
               Lib "kernel32.dll" _
               Alias "MoveFileA" (ByVal lpExistingFileName As String, _
                                  ByVal lpNewFileName As String) As Long

Public Declare Function CreateFile _
               Lib "kernel32.dll" _
               Alias "CreateFileA" (ByVal lpFileName As String, _
                                    ByVal dwDesiredAccess As Long, _
                                    ByVal dwShareMode As Long, _
                                    ByVal lpSecurityAttributes As Any, _
                                    ByVal dwCreationDisposition As Long, _
                                    ByVal dwFlagsAndAttributes As Long, _
                                    ByVal hTemplateFile As Long) As Long

Public Declare Function ReadFile _
               Lib "kernel32.dll" (ByVal hFile As Long, _
                                   lpBuffer As Any, _
                                   ByVal nNumberOfBytesToRead As Long, _
                                   lpNumberOfBytesRead As Long, _
                                   ByVal lpOverlapped As Any) As Long

Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function StrFormatByteSizeW _
               Lib "shlwapi" (ByVal qdwLow As Long, _
                              ByVal qdwHigh As Long, _
                              pwszBuf As Any, _
                              ByVal cchBuf As Long) As Long

Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFile _
               Lib "kernel32.dll" _
               Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                       lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function FindNextFile _
               Lib "kernel32.dll" _
               Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                      lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function PathMatchSpec _
               Lib "shlwapi" _
               Alias "PathMatchSpecW" (ByVal pszFileParam As Long, _
                                       ByVal pszSpec As Long) As Long

Public Declare Function GetTempFileName _
                Lib "kernel32.dll" _
                Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                          ByVal lpPrefixString As String, _
                                          ByVal wUnique As Long, _
                                          ByVal lpTempFileName As String) As Long

