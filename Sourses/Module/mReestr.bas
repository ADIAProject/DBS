Attribute VB_Name = "mReestr"
Option Explicit

' ������ � ��������
Public regParam                      As String

Private Const REG_SZ                 As Long = 1
Private Const REG_EXPAND_SZ          As Long = 2
Private Const REG_MULTI_SZ           As Long = 7
Private Const REG_DWORD              As Long = 4
Private Const ERROR_SUCCESS          As Long = 0

Public Const HKEY_CLASSES_ROOT       As Long = &H80000000
Public Const HKEY_CURRENT_USER       As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE      As Long = &H80000002

Private Const KEY_ALL_ACCESS         As Long = &H3F
Private Const ERROR_NONE             As Long = 0
Private Const STANDARD_RIGHTS_READ   As Long = &H20000
Private Const KEY_QUERY_VALUE        As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY             As Long = &H10
Private Const SYNCHRONIZE            As Long = &H100000
Private Const KEY_READ               As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

' for wow64 access
Private Const KEY_WOW64_32KEY        As Long = &H100
Private Const KEY_WOW64_64KEY        As Long = &H200

Private Declare Function RegQueryValueExString _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          ByVal lpData As String, _
                                          lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          lpData As Long, _
                                          lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          ByVal lpData As Long, _
                                          lpcbData As Long) As Long

Private Declare Function RegEnumKey _
                Lib "advapi32.dll" _
                Alias "RegEnumKeyA" (ByVal hKey As Long, _
                                     ByVal dwIndex As Long, _
                                     ByVal lpName As String, _
                                     ByVal cbName As Long) As Long

Private Declare Function RegOpenKey _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                     ByVal lpSubKey As String, _
                                     phkResult As Long) As Long

Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       ByRef phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          ByRef lpType As Long, _
                                          ByVal lpData As String, _
                                          ByRef lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function GetKeyValue(ByVal KeyRoot As Long, _
                            ByVal KeyName As String, _
                            ByVal SubKeyRef As String, _
                            Optional ByVal mboolReadKeyRights As Boolean = False) As String

    Dim i                    As Long
    Dim rc                   As Long
    Dim hKey                 As Long
    Dim sKeyVal              As String
    Dim sKeyVal_x()          As String
    Dim lKeyValType          As Long
    Dim tmpVal               As String
    Dim KeyValSize           As Long
    Dim GetKeyValueMultiSZ() As String
    Dim intTemp              As Long
    Dim intTempSmallBuff     As Long
    Dim Index                As Long
    Dim strstr               As String
    Dim lngKeyRights         As Long

    ' Loop Counter
    ' Return Code
    ' Handle To An Open Registry Key
    ' Data Type Of A Registry Key
    ' Tempory Storage For A Registry Key Value
    ' Size Of Registry Key Variable
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    If mboolReadKeyRights Then
        lngKeyRights = KEY_READ
    Else
        lngKeyRights = KEY_READ Or KEY_ALL_ACCESS
    End If

    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, lngKeyRights, hKey)

    ' Open Registry Key
    If rc <> ERROR_SUCCESS Then
        GoTo GetKeyError
    End If

    ' Handle Error...
    tmpVal = String$(1024, 0)
    ' Allocate Variable Space
    KeyValSize = 1024
    ' Mark Variable Size
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, lKeyValType, tmpVal, KeyValSize)

    ' Get/Create Key Value
    If rc <> ERROR_SUCCESS Then
        GoTo GetKeyError
    End If

    ' Handle Errors
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType

            ' Search Data Types...
        Case REG_SZ, REG_EXPAND_SZ
            tmpVal = Left$(tmpVal, InStr(tmpVal, vbNullChar) - 1)
            ' String Registry Key Data Type
            sKeyVal = tmpVal

            ' Copy String Value
        Case REG_DWORD
            tmpVal = Left$(tmpVal, InStr(tmpVal, vbNullChar) - 1)

            ' Double Word Registry Key Data Type
            If LenB(tmpVal) Then

                For i = Len(tmpVal) To 1 Step -1
                    ' Convert Each Bit
                    ' Build Value Char. By Char.
                    sKeyVal = sKeyVal + Hex$(Asc(Mid$(tmpVal, i, 1)))
                Next
                sKeyVal = Format$("&h" & sKeyVal)
                ' Convert Double Word To String
            Else
                sKeyVal = 0
            End If

        Case REG_MULTI_SZ
            intTemp = InStr(1, tmpVal, strDoubleNull)
            tmpVal = Mid$(tmpVal, 1, intTemp)
            sKeyVal_x = Split(tmpVal, vbNullChar)
            ReDim GetKeyValueMultiSZ(UBound(sKeyVal_x)) As String

            Do Until LenB(tmpVal) = 0
                intTempSmallBuff = InStrRev(tmpVal, vbNullChar)

                If intTempSmallBuff > 0 Then
                    intTemp = InStr(1, tmpVal, vbNullChar)
                    strstr = Mid$(tmpVal, 1, intTemp)
                    GetKeyValueMultiSZ(Index) = TrimNull(strstr)
                    Index = Index + 1
                    tmpVal = Mid$(tmpVal, intTemp + 1, Len(tmpVal))
                Else
                    sKeyVal = vbNullString
                End If

            Loop
            ReDim Preserve GetKeyValueMultiSZ(Index) As String

            Dim lngLBound As Long
            Dim lngUBound As Long

            lngLBound = LBound(GetKeyValueMultiSZ)
            lngUBound = UBound(GetKeyValueMultiSZ) - 1

            For i = lngLBound To lngUBound

                If LenB(sKeyVal) > 0 Then
                    If GetKeyValueMultiSZ(i) <> vbNullString Then
                        sKeyVal = sKeyVal & " | " & GetKeyValueMultiSZ(i)
                    End If

                Else
                    sKeyVal = GetKeyValueMultiSZ(i)
                End If

            Next

        Case Else
            tmpVal = Left$(tmpVal, InStr(tmpVal, vbNullChar) - 1)
    End Select

    GetKeyValue = Trim$(sKeyVal)
    ' Return Value
    rc = RegCloseKey(hKey)
    ' Close Registry Key
    Exit Function
    ' Exit
GetKeyError:
    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString
    DebugMode "Error read RegistryParam: Key:" & KeyName & " Param:" & SubKeyRef & " Error: �" & rc & " - " & ApiErrorText(rc), 2
    ' Set Return Val To Empty String
    rc = RegCloseKey(hKey)
    ' Close Registry Key
End Function

'! -----------------------------------------------------------
'!  �������     :  GetRegString
'!  ����������  :  hKey As Long, strSubKey As String, strValueName As String
'!  �����. ����.:  As String
'!  ��������    :  ��������� �������� ����� �� �������
'! -----------------------------------------------------------
Public Function GetRegString(hKey As Long, strSubKey As String, strValueName As String) As String

    Dim strSetting As String
    Dim lngDataLen As Long
    Dim lngRes     As Long

    If RegOpenKey(hKey, strSubKey, lngRes) = ERROR_SUCCESS Then
        strSetting = Space$(255)
        lngDataLen = LenB(strSetting)

        If RegQueryValueEx(lngRes, strValueName, ByVal 0, REG_EXPAND_SZ, ByVal strSetting, lngDataLen) = ERROR_SUCCESS Then
            If lngDataLen > 1 Then
                GetRegString = Left$(strSetting, lngDataLen - 1)
            End If
        End If

        If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
            MsgBox "RegCloseKey Failed: " & strSubKey, vbCritical, strProductName
        End If
    End If
End Function

'! -----------------------------------------------------------
'!  �������     :  IsAppPresent
'!  ����������  :  strSubKey$, strValueName$
'!  �����. ����.:  As Boolean
'!  ��������    :  ������������ �� ������ ���������� � �������
'! -----------------------------------------------------------
Public Function IsAppPresent(strSubKey As String, strValueName As String) As Boolean

    regParam = GetRegString(HKEY_CLASSES_ROOT, strSubKey, strValueName)
    IsAppPresent = CBool(Len(regParam))
End Function

Public Function ListKey(ByVal hKey, ByVal key, Optional ByVal mboolReadKeyRights As Boolean = True) As String()

    Dim strValue      As String
    Dim lDataLen      As Long
    Dim lResult       As Long
    Dim lValueLen     As Long
    Dim lCurIdx       As Long
    Dim lRetVal       As Long
    Dim hKeyResult    As Long
    Dim K()           As String
    Dim lngKeyRights  As Long
    Dim miMaxCountArr As Long

    'Variable to hold current enumerated key
    'Length of data
    'Result of RegEnumKey
    'Current Index which gets incremented with each pass through the loop
    'Result of RegOpenKeyEx
    If mboolReadKeyRights Then
        lngKeyRights = KEY_READ
    Else
        lngKeyRights = KEY_READ Or KEY_ALL_ACCESS
    End If

    ' ������������ ���-�� ��������� � �������
    miMaxCountArr = 500
    ReDim K(miMaxCountArr) As String
    lRetVal = RegOpenKeyEx(hKey, key, 0, lngKeyRights, hKeyResult)

    'Open key with Full Access Rights
    If lRetVal = ERROR_SUCCESS Then
        'data Length
        lDataLen = 64
        lValueLen = 64
        Do
            strValue = String$(lValueLen, 0)
            'get current key's value
            lResult = RegEnumKey(hKeyResult, lCurIdx, strValue, lDataLen)

            'Enumerate keys
            'if successful, add current enumerated key to the txtEnumKeys textbox
            If lResult = ERROR_SUCCESS Then

                ' ���� ������� � ������� ���������� ������ ��� ���������, �� ����������� ����������� �������
                If lCurIdx = miMaxCountArr Then
                    miMaxCountArr = miMaxCountArr + miMaxCountArr
                    ReDim Preserve K(miMaxCountArr)
                End If

                K(lCurIdx) = Replace$(strValue, vbNullChar, vbNullString)
            End If

            'Increment counter for next enumeration
            lCurIdx = lCurIdx + 1
            'continue while successful
        Loop While lResult = ERROR_SUCCESS

        'Close key
        RegCloseKey hKeyResult
        'If lRetVal is unsuccessful
    Else
        MsgBox "Cannot Open Key"
    End If

    ' �������� �������������� ������� �� �������� ���-�� �������
    If lCurIdx > 0 Then
        ReDim Preserve K(lCurIdx - 1) As String
    Else
        ReDim Preserve K(0)
    End If

    ListKey = K
End Function