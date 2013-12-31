Attribute VB_Name = "mStringFunction"
Option Explicit

Public kavichki As String

Public Const strDoubleNull = vbNullChar & vbNullChar

Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long

' Добавляет строку к строке с нужным разделителем
Public Function AppendStr(ByVal strHead As String, ByVal strAdd As String, Optional ByVal sep As String = " ") As String

    If LenB(strHead) > 0 Then
        AppendStr = strHead & sep & strAdd
    Else
        AppendStr = strAdd
    End If
End Function

Public Function CompareByVersion(ByVal strVersionBD As String, ByVal strVersionLocal As String) As String

    Dim strDevVer_x()       As String
    Dim strDevVerLocal_x()  As String
    Dim strDevVer_xx        As String
    Dim strDevVerLocal_xx   As String
    Dim miDimension         As Integer
    Dim miDimensionLocal    As Integer
    Dim strVersionBD_x()    As String
    Dim strVersionLocal_x() As String
    Dim i                   As Integer
    Dim ResultTemp          As String

    DebugMode "*********CompareByVersion-Start"
    DebugMode "*********CompareByVersion-Start: " & strVersionBD & " compare with " & strVersionLocal
    ResultTemp = "?"
    strDevVer_x = Split(Trim$(strVersionBD), ",", , vbTextCompare)
    miDimension = UBound(strDevVer_x)

    If strVersionBD <> "Unknown" Then
        If strVersionLocal <> "Unknown" Then
            If miDimension > 0 Then
                strDevVer_xx = Trim$(strDevVer_x(1))
            Else
                ResultTemp = "<"
                strDevVer_xx = strVersionBD
            End If

            strDevVerLocal_x = Split(Trim$(strVersionLocal), ",", , vbTextCompare)
            miDimensionLocal = UBound(strDevVerLocal_x)

            If miDimensionLocal > 0 Then
                strDevVerLocal_xx = Trim$(strDevVerLocal_x(1))
            Else
                ResultTemp = ">"
                strDevVerLocal_xx = strVersionLocal
            End If

            If Right$(strDevVer_xx, 1) = "." Then
                strDevVer_xx = Mid$(strDevVer_xx, 1, Len(strDevVer_xx) - 1)
            End If

            If Right$(strDevVerLocal_xx, 1) = "." Then
                strDevVer_xx = Mid$(strDevVerLocal_xx, 1, Len(strDevVerLocal_xx) - 1)
            End If

            strVersionBD_x = Split(strDevVer_xx, ".")
            strVersionLocal_x = Split(strDevVerLocal_xx, ".")

            If LenB(Trim$(strDevVerLocal_xx)) > 0 Then
                If UBound(strVersionBD_x) > UBound(strVersionLocal_x) Then

                    For i = LBound(strVersionLocal_x) To UBound(strVersionLocal_x)

                        If IsNumeric(strVersionBD_x(i)) Then
                            If IsNumeric(strVersionLocal_x(i)) Then
                                If CLng(strVersionBD_x(i)) < CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = "<"
                                    Exit For
                                ElseIf CLng(strVersionBD_x(i)) > CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = ">"
                                    Exit For
                                Else

                                    If i = UBound(strVersionBD_x) Then
                                        ResultTemp = "="
                                    End If
                                End If
                            End If

                        Else
                            ResultTemp = "?"
                        End If

                    Next
                Else

                    For i = LBound(strVersionBD_x) To UBound(strVersionBD_x)

                        If IsNumeric(strVersionBD_x(i)) Then
                            If IsNumeric(strVersionLocal_x(i)) Then
                                If CLng(strVersionBD_x(i)) < CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = "<"
                                    Exit For
                                ElseIf CLng(strVersionBD_x(i)) > CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = ">"
                                    Exit For
                                Else

                                    If i = UBound(strVersionBD_x) Then
                                        ResultTemp = "="
                                    End If
                                End If
                            End If

                        Else
                            ResultTemp = "?"
                        End If

                    Next
                End If

            Else
                ResultTemp = "?"
            End If

        Else
            ResultTemp = ">"
        End If

    Else
        ResultTemp = "<"
    End If

CompareFinish:
    CompareByVersion = ResultTemp
    DebugMode "*********CompareByVersion-Result: " & strVersionBD & " " & ResultTemp & " " & strVersionLocal
    DebugMode "*********CompareByVersion-End"
End Function

Public Function ConvertDate2Rus(ByVal dtDate As String) As String

    Dim DD         As String
    Dim MM         As String
    Dim YYYY       As String
    Dim dtDateTemp As String
    Dim objRegExp  As Object
    Dim objMatch   As Match
    Dim objMatches As MatchCollection

    dtDateTemp = dtDate

    If LenB(dtDate) > 0 Then
        If StrComp(dtDate, "Unknown", vbTextCompare) <> 0 Then
            Set objRegExp = CreateObject("VBScript.RegExp")

            With objRegExp
                .Pattern = "(\d+).(\d+).(\d+)"
                .IgnoreCase = True
                .Global = True
            End With

            'получаем date1
            Set objMatches = objRegExp.Execute(dtDate)

            With objMatches

                If .Count > 0 Then
                    Set objMatch = .Item(0)
                    MM = Format$(objMatch.SubMatches(0), "00")
                    DD = Format$(objMatch.SubMatches(1), "00")
                    YYYY = DateTime.Year(dtDate)
                End If
            End With

            'OBJMATCHES
            ' если необходимо конвертировать дату в формат dd/mm/yyyy
            If mboolDateFormatRus Then
                dtDateTemp = DD & "/" & MM & "/" & YYYY
            Else
                dtDateTemp = MM & "/" & DD & "/" & YYYY
            End If
        End If
    End If

    ConvertDate2Rus = dtDateTemp
End Function

' Заменяем в строке некоторые символы RegExp на константы VB
Public Function ConvertString(strStringText As String) As String

    If InStr(1, strStringText, "\t") > 0 Then
        strStringText = Replace$(strStringText, "\t", vbTab, , , vbTextCompare)
    End If

    If InStr(1, strStringText, "\r\n") > 0 Then
        strStringText = Replace$(strStringText, "\r\n", vbNewLine, , , vbTextCompare)
    End If

    If InStr(1, strStringText, "\r") > 0 Then
        strStringText = Replace$(strStringText, "\r", vbCr, , , vbTextCompare)
    End If

    If InStr(1, strStringText, "\n") > 0 Then
        strStringText = Replace$(strStringText, "\n", vbLf, , , vbTextCompare)
    End If

    ConvertString = strStringText
End Function

Public Function FilterArray(ByVal Source As String, _
                            ByVal Search As String, _
                            Optional ByVal Keep As Boolean = True) As String

    Dim i                   As Long
    Dim SearchArray()       As String
    Dim IntermediateArray() As String
    Dim iSearchLower        As Long
    Dim iSearchUpper        As Long

    If LenB(Source) <> 0 And LenB(Search) <> 0 Then
        SearchArray = Split(Search, " ")
    Else
        FilterArray = Source
        Exit Function
    End If

    iSearchLower = LBound(SearchArray)
    iSearchUpper = UBound(SearchArray)
    IntermediateArray = Split(Source, " ")

    For i = iSearchLower To iSearchUpper
        DoEvents
        IntermediateArray = Filter(IntermediateArray, SearchArray(i), Keep, vbTextCompare)
    Next
    FilterArray = Join(IntermediateArray, " ")
End Function

'! -----------------------------------------------------------
'!  Функция     :  PathCollect
'!  Переменные  :  Path As String
'!  Возвр. знач.:  As String
'!  Описание    :
'! -----------------------------------------------------------
Public Function PathCollect(Path As String) As String

    If InStr(1, Path, ":") = 2 Then
        PathCollect = Path
    Else

        If Left$(Path, 2) = ".\" Then
            PathCollect = strAppPath & Mid$(Path, 2, Len(Path) - 1)
        Else

            If InStr(1, Path, "\") = 1 Then
                PathCollect = strAppPath & Path
            Else

                If Left$(Path, 3) = "..\" Then
                    PathCollect = PathNameFromPath(strAppPath) & Mid$(Path, 4, Len(Path) - 1)
                Else

                    If InStr(1, Path, "%", vbTextCompare) > 0 Then
                        PathCollect = GetEnviron(Path, True)
                    Else

                        If ExtFromFileName(Path) <> vbNullString Then
                            If FileNameFromPath(Path) = Path Then
                                PathCollect = Path
                            Else
                                PathCollect = strAppPathBackSL & Path
                            End If

                        Else
                            PathCollect = strAppPathBackSL & Path
                        End If
                    End If
                End If
            End If
        End If
    End If

    If InStr(PathCollect, "\\") > 0 Then
        PathCollect = Replace$(PathCollect, "\\", "\", , , vbTextCompare)

        If Left$(strAppPath, 2) = "\\" Then
            If InStr(1, PathCollect, "\") = 1 Then
                PathCollect = vbBackslash & PathCollect
            Else
                PathCollect = "\\" & PathCollect
            End If
        End If
    End If

    If IsPathAFolder(PathCollect) Then
        PathCollect = BackslashAdd2Path(PathCollect)
    End If
End Function

'! -----------------------------------------------------------
'!  Функция     :  PathCollect4Dest
'!  Переменные  :  Path As String
'!  Возвр. знач.:  As String
'!  Описание    :
'! -----------------------------------------------------------
Public Function PathCollect4Dest(ByVal Path As String, ByVal strDest As String) As String

    If InStr(1, Path, ":") = 2 Then
        PathCollect4Dest = Path
    Else

        If Left$(Path, 2) = ".\" Then
            PathCollect4Dest = strDest & Mid$(Path, 2, Len(Path) - 1)
        Else

            If InStr(1, Path, "\") = 1 Then
                PathCollect4Dest = strDest & Path
            Else

                If Left$(Path, 3) = "..\" Then
                    PathCollect4Dest = PathNameFromPath(strDest) & Mid$(Path, 4, Len(Path) - 1)
                Else

                    If InStr(1, Path, "%", vbTextCompare) > 0 Then
                        PathCollect4Dest = GetEnviron(Path, True)
                    Else

                        If ExtFromFileName(Path) <> vbNullString Then
                            If FileNameFromPath(Path) = Path Then
                                PathCollect4Dest = Path
                            Else
                                PathCollect4Dest = BackslashAdd2Path(strDest) & Path
                            End If

                        Else
                            PathCollect4Dest = BackslashAdd2Path(strDest) & Path
                        End If
                    End If
                End If
            End If
        End If
    End If

    If InStr(PathCollect4Dest, "\\") > 0 Then
        PathCollect4Dest = Replace$(PathCollect4Dest, "\\", "\", , , vbTextCompare)

        If Left$(strDest, 2) = "\\" Then
            If InStr(1, PathCollect4Dest, "\") = 1 Then
                PathCollect4Dest = vbBackslash & PathCollect4Dest
            Else
                PathCollect4Dest = "\\" & PathCollect4Dest
            End If
        End If
    End If

    PathCollect4Dest = BackslashAdd2Path(PathCollect4Dest)
End Function

' получаем значение из буфера данных
Public Function TrimNull(ByVal startstr As String) As String

    TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
End Function
