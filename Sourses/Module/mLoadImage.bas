Attribute VB_Name = "mLoadImage"
Option Explicit

'Private strPathImageStatusButton         As String
Public strPathImageMain     As String
Public strPathImageMenu     As String

'Private strPathImageStatusButtonWork     As String
Public strPathImageMainWork As String
Public strPathImageMenuWork As String

Private Const lngIMG_SIZE   As Long = &H20

Public Sub LoadIconImage2Btn(ByVal ObjectName As ctlXpButton, _
                             ByVal strPictureName As String, _
                             ByVal strPathImageDir As String)

    LoadIconImageFromFileBtn ObjectName, SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
End Sub

Public Sub LoadIconImage2BtnJC(ByVal ObjectName As ctlJCbutton, _
                               ByVal strPictureName As String, _
                               ByVal strPathImageDir As String)

    LoadIconImageFromFileBtnJC ObjectName, SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
End Sub

Public Sub LoadIconImage2FrameJC(ByVal ObjectName As ctlJCFrames, _
                                 ByVal strPictureName As String, _
                                 ByVal strPathImageDir As String)

    LoadIconImageFromFileJC ObjectName, SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
End Sub

Private Sub LoadIconImageFromFileBtn(ByVal imgName As ctlXpButton, ByVal PicturePath As String)

    DebugMode "***LoadIconImageFromFileBtn-Start", 2
    DebugMode "******LoadIconImageFromFileBtn: PicturePath=" & PicturePath, 2

    If PathFileExists(PicturePath) = 1 Then

        With imgName

            If Not (.Picture Is Nothing) Then
                If .Picture <> stdole.LoadPicture(PicturePath) Then
                    Set .Picture = stdole.LoadPicture(PicturePath)
                    DebugMode "******LoadIconImageFromFileBtn: Picture is Installed", 2
                Else
                    DebugMode "******LoadIconImageFromFileBtn: Picture is already set", 2
                End If

            Else
                Set .Picture = stdole.LoadPicture(PicturePath)
                DebugMode "******LoadIconImageFromFileBtn: Picture is Installed", 2
            End If
        End With

        'imgName
    Else
        DebugMode "******LoadIconImageFromFileBtn: Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2
    End If

    DebugMode "***LoadIconImageFromFileBtn-End", 2
End Sub

Private Sub LoadIconImageFromFileBtnJC(ByVal btnName As ctlJCbutton, ByVal PicturePath As String)

    DebugMode "***LoadIconImageFromFileBtnJC-Start", 2
    DebugMode "******LoadIconImageFromFileBtnJC: PicturePath=" & PicturePath, 2

    If PathFileExists(PicturePath) = 1 Then

        With btnName

            If Not (.PictureNormal Is Nothing) Then
                If .PictureNormal <> stdole.LoadPicture(PicturePath) Then
                    Set .PictureNormal = stdole.LoadPicture(PicturePath)
                    DebugMode "******LoadIconImageFromFileBtnJC: Picture is Installed", 2
                Else
                    DebugMode "******LoadIconImageFromFileBtnJC: Picture is already set", 2
                End If

            Else
                Set .PictureNormal = stdole.LoadPicture(PicturePath)
                DebugMode "******LoadIconImageFromFileBtnJC: Picture is Installed", 2
            End If
        End With

        'imgName
    Else
        DebugMode "******Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2
    End If

    DebugMode "***LoadIconImageFromFileBtnJC-End", 2
End Sub

Private Sub LoadIconImageFromFileJC(ByVal imgName As ctlJCFrames, ByVal PicturePath As String)

    DebugMode "***LoadIconImageFromFileJC-Start", 2
    DebugMode "******LoadIconImageFromFileJC: PicturePath=" & PicturePath, 2

    If PathFileExists(PicturePath) = 1 Then
        'imgName.Picture = stdole.LoadPicture(PicturePath, lngIMG_SIZE, lngIMG_SIZE, Color)
        Set imgName.Picture = stdole.LoadPicture(PicturePath)
    Else
        DebugMode "******Path to picture: " & PicturePath & " not Exist. Standard picture Will is used", 2
    End If

    DebugMode "***LoadIconImageFromFileJC-End", 2
End Sub

Public Function LoadIconImageFromPath(strPictureName As String, strPathImageDir As String) As IPictureDisp

    Dim strPicturePath As String

    DebugMode "LoadIconImageFromPath-Start", 2
    strPicturePath = SearchFilesInRoot(strPathImageDir, strPictureName & ".*", False, True)
    DebugMode "***LoadIconImageFromPath: PicturePath=" & strPicturePath, 2

    If PathFileExists(strPicturePath) = 1 Then
        'Set LoadIconImageFromPath = LoadPicture(strPicturePath)
        Set LoadIconImageFromPath = stdole.LoadPicture(strPicturePath)
    Else
        DebugMode "***LoadIconImageFromPath: Path to picture: " & strPicturePath & " not Exist. Standard picture Will is used", 2
    End If

    DebugMode "LoadIconImageFromPath-End", 2
End Function

Public Sub LoadIconImagePath()

    DebugMode "***LoadIconImagePath-Start", 2
    strPathImageMainWork = strPathImageMain & strImageMainName

    If PathFileExists(strPathImageMainWork) = 0 Then
        MsgBox strMessages(15), vbCritical, strProductName
        strPathImageMainWork = strPathImageMain & "Standart"
    End If

    'strPathImageMenuWork = strPathImageMenu & strImageMenuName
    'If PathFileExists(strPathImageMenuWork) = 0 Then
    'MsgBox strMessages(15), vbCritical, strProductName
    'strPathImageMenuWork = strPathImageMenu & "Standart"
    'End If
    DebugMode "***LoadIconImagePath-End", 2
End Sub
