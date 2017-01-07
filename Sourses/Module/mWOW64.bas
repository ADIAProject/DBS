Attribute VB_Name = "mWOW64"
Option Explicit

Public Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (ByRef OldValue As Long) As Long
Public Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByRef OldValue As Long) As Long

Private Declare Function GetSystemWow64Directory Lib "kernel32.dll" Alias "GetSystemWow64DirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Long

Public Function GetSystemWow64Dir() As String

    Dim evar As String
    Dim elen As Long

    If APIFunctionPresent("GetSystemWow64Directory", "kernel32.dll") Then
        evar = String$(256, " ")
        elen = GetSystemWow64Directory(evar, Len(evar))

        If elen = 0 Then
            'ERROR_CALL_NOT_IMPLEMENTED= 120 (0x78) ???
            GetSystemWow64Dir = vbNullString
        Else
            evar = TrimNull(evar)
            GetSystemWow64Dir = evar
        End If
    End If
End Function

'Dim OldValue As Long
'Dim bRet As Long
'bRet = Wow64DisableWow64FsRedirection(OldValue)
'Call Wow64RevertWow64FsRedirection(OldValue)
'If bRet Then
'End If
' Restore the previous WOW64 file system redirection value.
'Call Wow64RevertWow64FsRedirection(OldValue)
'End If
