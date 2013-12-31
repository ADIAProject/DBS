Attribute VB_Name = "mStrings"
' mStrings.bas
Option Explicit

Public Enum SplitCompareMethod
    [Split BinaryCompare] = VbCompareMethod.vbBinaryCompare         ' InStrB
    '[Split TextCompare] = VbCompareMethod.vbTextCompare            ' InStr(TextCompare)
    [Split CharacterCompare] = VbCompareMethod.vbDatabaseCompare    ' InStr(BinaryCompare)
End Enum

Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60.dll" (ByVal ptr As Long, Value As Long)
Private Declare Function InitStringArray _
                Lib "oleaut32.dll" _
                Alias "SafeArrayCreate" (Optional ByVal VarType As VbVarType = vbString, _
                                         Optional ByVal Dims As Integer = 1, _
                                         Optional saBound As Currency) As Long

Private Declare Sub PutMem4 Lib "msvbvm60.dll" (ByVal ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal ptr As Long, ByVal Length As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32.dll" (ByVal ptr As Long, ByVal Length As Long) As Long

Private Function InIDE(Optional IDE) As Boolean

    If IsMissing(IDE) Then
        Debug.Assert Not InIDE(InIDE)
    Else
        IDE = True
    End If
End Function

Private Property Get Procedure(ByVal AddressOfDest As Long) As Long

    Procedure = AddressOfDest
End Property

Private Property Let Procedure(ByVal AddressOfDest As Long, ByVal AddressOfSrc As Long)

    Dim JMP As Currency, PID As Long

    ' get process handle
    PID = OpenProcess(&H1F0FFF, 0&, GetCurrentProcessId)

    If PID Then
        If InIDE Then
            ' get correct pointers to procedures in IDE
            GetMem4 AddressOfDest + &H16&, AddressOfDest
            GetMem4 AddressOfSrc + &H16&, AddressOfSrc
        End If

        Debug.Assert App.hInstance
        ' ASM JMP (0xE9) followed by bytes to jump in memory
        JMP = (&HE9& * 0.0001@) + (AddressOfSrc - AddressOfDest - 5@) * 0.0256@
        ' write the JMP over the destination procedure
        WriteProcessMemory PID, ByVal AddressOfDest, JMP, 5
        ' close process handle
        CloseHandle PID
    End If
End Property

Public Function Split(Expression As String, _
                      Optional Delimiter As String = " ", _
                      Optional ByVal Limit As Long = -1, _
                      Optional ByVal Compare As SplitCompareMethod) As String()

    Procedure(AddressOf mStrings.Split) = Procedure(AddressOf mStrings.z_Split)
    Split = mStrings.Split(Expression, Delimiter, Limit, Compare)
End Function

Public Function z_Split(Expression As String, _
                        Optional Delimiter As String = " ", _
                        Optional ByVal Limit As Long = -1, _
                        Optional ByVal Compare As SplitCompareMethod) As Long

    ' general variables that we need
    Dim p() As Long
    Dim r() As Long
    Dim C   As Long
    Dim i   As Long
    Dim J   As Long
    Dim K   As Long
    Dim LD  As Long
    Dim LE  As Long
    Dim PL  As Long
    Dim PS  As Long

    ' get pointer
    PS = StrPtr(Expression)
    ' length information
    LE = LenB(Expression)
    LD = LenB(Delimiter)

    ' unlimited or limited?
    If Limit = -1 Then
        If LD Then
            Limit = LE \ LD + 1
        End If
    End If

    ' validate lengths and limit
    If LE > 0 And LD > 0 And Limit >= 0 Then

        ' find the first item
        If Limit > 1 Then
            If Compare = [Split BinaryCompare] Then
                Do
                    i = InStrB(i + 1, Expression, Delimiter)
                Loop Until (i And 1) = 1 Or (i = 0)

            Else
                i = InStr(Expression, Delimiter)
            End If
        End If

        ' did we find an item?
        If i Then
            ' space for knowing the positions
            PL = Limit \ 80
            ReDim p(0 To PL)

            ' InStrB?
            If Compare = [Split BinaryCompare] Then
                Do
                    ' remember position
                    p(C) = i - 1
                    ' find next
                    i = i + LD - 1
                    Do
                        i = InStrB(i + 1, Expression, Delimiter)
                    Loop Until (i And 1) = 1 Or (i = 0)

                    ' increase counter
                    C = C + 1

                    If C > PL Then
                        PL = PL + C
                        ReDim Preserve p(PL)
                    End If

                Loop While i > 0 And C <= Limit

            Else
                ' InStr
                Do
                    ' remember position
                    p(C) = (i - 1) * 2
                    ' find next
                    i = InStr(i + LD \ 2, Expression, Delimiter)
                    ' increase counter
                    C = C + 1

                    If C > PL Then
                        PL = PL + C
                        ReDim Preserve p(PL)
                    End If

                Loop While i > 0 And C <= Limit

            End If

            p(C) = LE
            ' make space for the new items
            z_Split = InitStringArray(, , (C + 1) * 0.0001@)
            PutMem4 ArrPtr(r), z_Split
            ' keep it simple, stupid!
            i = 0

            For C = 0 To C
                K = p(C)
                J = K - i

                If J Then
                    r(C) = SysAllocStringByteLen(PS + i, J)
                End If

                i = K + LD
            Next
        Else
            ' one item
            z_Split = InitStringArray(, , 0.0001@)
            PutMem4 ArrPtr(r), z_Split
            r(0) = SysAllocStringByteLen(PS, LE)
        End If

        ' clean up z_Split reference
        PutMem4 ArrPtr(r), 0
    Else
        z_Split = InitStringArray
    End If
End Function
