Attribute VB_Name = "mUpdate"
Option Explicit

Public strLink()           As String
Public strLinkFull()       As String
Public strLinkHistory      As String
Public strLinkHistory_en   As String
Public strVersion          As String
Public strDateProg         As String
Public strDescription      As String
Public strDescription_en   As String
Public strRelease          As String
Public strUpdVersions()    As String
Public strUpdDescription() As String

' Проверка существования файла на сервере
Function CheckConnection2Server(ByVal URL As String) As String

    ' Функция скачивает файл по ссылке URL$
    ' и сохраняет его под именем LocalPath$
    Dim XMLHTTP
    Dim strResultText As String
    Dim strResultCode As String

    On Error Resume Next

    Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")

    With XMLHTTP
        .Open "GET", Replace$(URL, "\", "/"), "False"
        .sEnd
        strResultText = .statusText
        strResultCode = .Status
    End With

    If StrComp(strResultText, "OK", vbTextCompare) = 0 Then
        CheckConnection2Server = "OK"
    Else
        CheckConnection2Server = strResultCode & "-" & strResultText
    End If

    Set XMLHTTP = Nothing
End Function

'! -----------------------------------------------------------
'!  Функция     :  CheckUpd
'!  Переменные  :  Optional start As Boolean
'!  Описание    :  Проверка новых версий программы с использованием MSXML
'! -----------------------------------------------------------
Public Sub CheckUpd(Optional ByVal start As Boolean = True)

    Dim xmlDoc           As DOMDocument
    Dim nodeList         As IXMLDOMNodeList
    Dim xmlNode          As IXMLDOMNode
    Dim propertyNode     As IXMLDOMElement
    Dim Url_Request      As String
    Dim Url_Test         As String
    Dim Url_Test_Result  As String
    Dim TextNodeName     As String
    Dim NodeIndex        As Integer
    Dim strVerTemp       As String
    Dim strResultCompare As String

    DebugMode "CheckUpd-Start"
    DebugMode "***CheckUpd-Options: " & start

    On Error Resume Next

    'Узнаем версию программы (установленной)
    strVerTemp = strProductVersion
    Set xmlDoc = New DOMDocument
    xmlDoc.async = False
    Url_Request = "http://www.adia-project.net/ProjectDBS/dbs_update2.xml"
    Url_Test = "http://www.adia-project.net/test.txt"
    'Url_Request = strAppPath & "\dia_update2.xml"
    ' проверка наличия доступа до сервера
    Url_Test_Result = CheckConnection2Server(Url_Test)

    If StrComp(Url_Test_Result, "OK", vbTextCompare) = 0 Then

        ' загружаем файл
        If Not xmlDoc.Load(Url_Request) Then
            If Not start Then
                MsgBox strMessages(126), vbInformation, strMessages(54)
            End If

        Else
            Set nodeList = xmlDoc.documentElement.selectNodes("//driversbackuper")
            Set xmlNode = nodeList(0)
            NodeIndex = 0

            For Each propertyNode In xmlNode.childNodes
                TextNodeName = vbNullString
                TextNodeName = xmlNode.childNodes(NodeIndex).nodeName

                Select Case TextNodeName

                        ' Данные из файла dbs_update2.xml
                        ' Версия проги
                    Case "version"
                        strVersion = xmlNode.childNodes(NodeIndex).Text

                        ' Дата проги
                    Case "date"
                        strDateProg = xmlNode.childNodes(NodeIndex).Text

                    Case "release"
                        strRelease = xmlNode.childNodes(NodeIndex).Text

                        ' Ссылка на Полную историю изменений
                    Case "linkHistory"
                        strLinkHistory = xmlNode.childNodes(NodeIndex).Text

                    Case "linkHistory_en"
                        strLinkHistory_en = xmlNode.childNodes(NodeIndex).Text
                End Select

                NodeIndex = NodeIndex + 1
            Next
            '**** Сравнение версий программ
            strResultCompare = CompareByVersion(strVersion, strVerTemp)

            ' Анализ итога сравнения и показ окна
            Select Case strResultCompare

                Case ">"

                    If StrComp(strRelease, "beta", vbTextCompare) = 0 Then
                        If Not mboolUpdateCheckBeta Then
                            DebugMode "***The version on the site is Beta. In options check for beta are disable. Break function!!!"

                            If Not start Then
                                If MsgBox(strMessages(56), vbQuestion + vbYesNo, strProductName) = vbYes Then
                                    frmCheckUpdate.Show vbModal, frmMain
                                Else
                                    Exit Sub
                                End If
                            End If

                        Else
                            frmCheckUpdate.Show vbModal, frmMain
                        End If

                    Else
                        frmCheckUpdate.Show vbModal, frmMain
                    End If

                Case "="

                    If Not start Then
                        If MsgBox(strMessages(56), vbQuestion + vbYesNo, strProductName) = vbYes Then
                            frmCheckUpdate.Show vbModal, frmMain
                        End If
                    End If

                Case "<"

                    If Not start Then
                        If MsgBox(strMessages(55), vbQuestion + vbYesNo, strProductName) = vbYes Then
                            frmCheckUpdate.Show vbModal, frmMain
                        End If
                    End If

                Case Else

                    If Not start Then
                        MsgBox strMessages(102), vbInformation, strProductName
                    End If
            End Select

            Set xmlNode = Nothing
            Set nodeList = Nothing
        End If

    Else
        DebugMode "***CheckUPD: " & strMessages(53) & vbNewLine & "Error: " & Url_Test_Result

        If Not start Then
            MsgBox strMessages(53) & vbNewLine & "Error: " & Url_Test_Result, vbInformation, strMessages(54)
        End If
    End If

    Set xmlDoc = Nothing

    On Error GoTo 0

    DebugMode "CheckUpd-End"
End Sub

Private Function DeltaDay() As Integer

    Dim CurrentDate As Date
    Dim BuildDate   As Date
    Dim DeltaTemp   As Integer

    CurrentDate = Date
    BuildDate = CDate(strDateProgram)
    DeltaTemp = CInt(CurrentDate - BuildDate)
    DeltaDay = DeltaTemp
End Function

Private Function DeltaDayNew(ByVal dtFirstDate As Date, ByVal dtSecondDate As Date) As Integer

    Dim DeltaTemp As Integer

    DeltaTemp = CInt(dtFirstDate - dtSecondDate)
    DeltaDayNew = DeltaTemp
End Function

'! -----------------------------------------------------------
'!  Функция     :  CheckUpd
'!  Переменные  :  Optional start As Boolean
'!  Описание    :  Проверка новых версий программы с использованием MSXML
'! -----------------------------------------------------------
Public Sub LoadUpdateData()

    Dim xmlDoc          As DOMDocument
    Dim nodeList        As IXMLDOMNodeList
    Dim xmlNode         As IXMLDOMNode
    Dim propertyNode    As IXMLDOMElement
    Dim Url_Request     As String
    Dim TextNodeName    As String
    Dim NodeIndex       As Integer
    Dim strVersionsTemp As String
    Dim i               As Long

    On Error Resume Next

    Set xmlDoc = New DOMDocument
    xmlDoc.async = False
    Url_Request = "http://www.adia-project.net/ProjectDBS/dbs_update2.xml"

    'Url_Request = strAppPath & "\dbs_update2.xml"
    If Not xmlDoc.Load(Url_Request) Then
        MsgBox strMessages(53), vbInformation, strMessages(54)
    Else
        Set nodeList = xmlDoc.documentElement.selectNodes("//driversbackuper")
        Set xmlNode = nodeList(0)
        NodeIndex = 0

        For Each propertyNode In xmlNode.childNodes
            TextNodeName = vbNullString
            TextNodeName = xmlNode.childNodes(NodeIndex).nodeName

            Select Case TextNodeName

                    ' Данные из файла dbs_update2.xml
                    ' массив версий
                Case "versions"
                    strVersionsTemp = xmlNode.childNodes(NodeIndex).Text
                    strUpdVersions = Split(strVersionsTemp, ";", , vbTextCompare)
                    ReDim strUpdDescription(UBound(strUpdVersions), 2) As String
                    ReDim strLink(UBound(strUpdVersions), 6) As String
                    ReDim strLinkFull(UBound(strUpdVersions), 6) As String

                    ' Данные из файла %ver%.xml
                    'Загрузка описаний изменений
                    For i = LBound(strUpdVersions) To UBound(strUpdVersions)
                        LoadUpdDescription strUpdVersions(i), i
                    Next
            End Select

            NodeIndex = NodeIndex + 1
        Next
        Set xmlNode = Nothing
        Set nodeList = Nothing
    End If

    Set xmlDoc = Nothing

    On Error GoTo 0

End Sub

Public Function LoadUpdDescription(ByVal strVer As String, ByVal lngIndexVer As Long) As String

    Dim xmlDocVers       As DOMDocument
    Dim nodeListVers     As IXMLDOMNodeList
    Dim xmlNodeVers      As IXMLDOMNode
    Dim propertyNodeVers As IXMLDOMElement
    Dim Url_Request      As String
    Dim TextNodeName     As String
    Dim NodeIndex        As Integer

    Set xmlDocVers = New DOMDocument
    xmlDocVers.async = False
    Url_Request = "http://www.adia-project.net/ProjectDBS/" & strVer & ".xml"

    'Url_Request = strAppPath & vbBackslash & strVer & ".xml"
    If Not xmlDocVers.Load(Url_Request) Then
        MsgBox strMessages(53), vbInformation, strMessages(54)
    Else
        Set nodeListVers = xmlDocVers.documentElement.selectNodes("//driversbackuper")
        Set xmlNodeVers = nodeListVers(0)
        NodeIndex = 0

        For Each propertyNodeVers In xmlNodeVers.childNodes
            TextNodeName = vbNullString
            TextNodeName = xmlNodeVers.childNodes(NodeIndex).nodeName

            Select Case TextNodeName

                    ' Описание изменений
                Case "description"
                    strUpdDescription(lngIndexVer, 0) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "description_en"
                    strUpdDescription(lngIndexVer, 1) = xmlNodeVers.childNodes(NodeIndex).Text

                    ' Ссылка на обновление
                Case "link"
                    strLink(lngIndexVer, 0) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_header"
                    strLink(lngIndexVer, 1) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_Mirror1"
                    strLink(lngIndexVer, 2) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_header1"
                    strLink(lngIndexVer, 3) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_Mirror2"
                    strLink(lngIndexVer, 4) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_header2"
                    strLink(lngIndexVer, 5) = xmlNodeVers.childNodes(NodeIndex).Text

                    ' Ссылка на дистрибутив
                Case "linkFull"
                    strLinkFull(lngIndexVer, 0) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_header"
                    strLinkFull(lngIndexVer, 1) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_Mirror1"
                    strLinkFull(lngIndexVer, 2) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_header1"
                    strLinkFull(lngIndexVer, 3) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_Mirror2"
                    strLinkFull(lngIndexVer, 4) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_header2"
                    strLinkFull(lngIndexVer, 5) = xmlNodeVers.childNodes(NodeIndex).Text
            End Select

            NodeIndex = NodeIndex + 1
        Next
    End If
End Function

' Показ напоминания об обновлении
Public Sub ShowUpdateToolTip()

    Dim mboolShowToolTip As Boolean
    Dim intDeltaDay      As Integer
    Dim dtToolTipDate    As Date
    Dim strToolTipDate   As String

    If DeltaDay > 45 Then
        If mboolUpdateToolTip Then
            strToolTipDate = GetSetting(App.ProductName, "UpdateToolTip", "Show at Date", vbNullString)

            If strToolTipDate = vbNullString Then
                mboolShowToolTip = True
            Else
                dtToolTipDate = CDate(strToolTipDate)
                intDeltaDay = DeltaDayNew(Date, dtToolTipDate)

                If intDeltaDay >= 5 Then
                    mboolShowToolTip = True
                End If
            End If

        Else
            mboolShowToolTip = False
        End If

    Else
        mboolShowToolTip = False
    End If

    ' Если все условия выполнены, то показываем сообщение
    ' "Возможно, используемая вами, версия программы 'Помощник установки драйверов' уже устарела! "
    If mboolShowToolTip Then
        ShowNotifyMessage strMessages(107)
    End If
End Sub
