Attribute VB_Name = "mApiProcess"
Option Explicit

' необходимо для регистрации компонента
Public Const DONT_RESOLVE_DLL_REFERENCES As Long = &H1
Public Const GMEM_FIXED                  As Long = 0 'Fixed memory GlobalAlloc flag
Public Const PATCH_04                    As Long = 88                                   'Table B (before) address patch offset
Public Const PATCH_05                    As Long = 93                                   'Table B (before) entry count patch offset
Public Const PATCH_08                    As Long = 132                                  'Table A (after) address patch offset
Public Const PATCH_09                    As Long = 137                                  'Table A (after) entry count patch offset
Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub CopyMemory _
               Lib "Kernel32" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal Length As Long)
Public Declare Sub CopyMemoryLong _
               Lib "Kernel32" _
               Alias "RtlMoveMemory" (ByVal Destination As Long, _
                                      ByVal Source As Long, _
                                      ByVal Length As Long)
                                      
Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GetModuleHandle _
               Lib "kernel32.dll" _
               Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Public Declare Function GetModuleHandleA Lib "kernel32.dll" (ByVal lpModuleName As String) As Long
Public Declare Function LoadLibraryEx _
               Lib "kernel32.dll" _
               Alias "LoadLibraryExA" (ByVal lpLibFileName As String, _
                                       ByVal hFile As Long, _
                                       ByVal dwFlags As Long) As Long

Public Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CallWindowProc _
               Lib "user32.dll" _
               Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                        ByVal hwnd As Long, _
                                        ByVal Msg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long
Public Declare Function OpenProcess _
               Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
                                   ByVal bInheritHandle As Long, _
                                   ByVal dwProcessId As Long) As Long

Public Declare Function WriteProcessMemory _
               Lib "kernel32.dll" (ByVal hProcess As Long, _
                                   lpBaseAddress As Any, _
                                   lpBuffer As Any, _
                                   ByVal nSize As Long, _
                                   Optional lpNumberOfBytesWritten As Long) As Long

