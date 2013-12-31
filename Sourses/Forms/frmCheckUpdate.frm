VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form frmCheckUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Обновление: Обнаружена новая версия программы"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCheckUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ComboBoxW cmbVersions 
      Height          =   345
      Left            =   4680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   450
      Width           =   1335
   End
   Begin prjDIADBS.ctlXpButton cmdExit 
      Height          =   735
      Left            =   9345
      TabIndex        =   0
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      Caption         =   "Закрыть"
      ButtonStyle     =   3
      PictureWidth    =   0
      PictureHeight   =   0
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjDIADBS.ctlXpButton cmdHistory 
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Caption         =   "История изменений"
      ButtonStyle     =   3
      PictureWidth    =   48
      PictureHeight   =   48
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjDIADBS.ctlXpButton cmdUpdate 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Caption         =   "Скачать обновление"
      ButtonStyle     =   3
      PictureWidth    =   48
      PictureHeight   =   48
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      MaskColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjDIADBS.ctlXpButton cmdDonate 
      Height          =   735
      Left            =   6990
      TabIndex        =   4
      Top             =   5160
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   1296
      Caption         =   "Поддержать проект"
      ButtonStyle     =   3
      PictureWidth    =   51
      PictureHeight   =   28
      ShowFocusRect   =   0   'False
      XPColor_Pressed =   15116940
      XPColor_Hover   =   4692449
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   0
      MenuCaption0    =   "#"
   End
   Begin prjDIADBS.RichTextBox rtfDescription 
      Height          =   4275
      Left            =   120
      TabIndex        =   5
      Top             =   800
      Visible         =   0   'False
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   7541
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmCheckUpdate.frx":000C
   End
   Begin prjDIADBS.LabelW lblWait 
      Height          =   375
      Left            =   100
      Top             =   2640
      Width           =   11160
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Идет загрузка данных с официального сайта. Пожалуйста, подождите...."
      ShadowStyle     =   0
      Alignment       =   2
   End
   Begin prjDIADBS.LabelW lblVersionList 
      Height          =   375
      Left            =   120
      Top             =   450
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Список изменений для версии:"
      ShadowStyle     =   0
   End
   Begin prjDIADBS.LabelW lblWWW 
      Height          =   315
      Left            =   6120
      Top             =   450
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "www.adia-project.net"
      ForeColor       =   12582912
      ShadowStyle     =   0
      Alignment       =   1
   End
   Begin prjDIADBS.LabelW lblVersion 
      Height          =   315
      Left            =   120
      Top             =   45
      Width           =   11085
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Последняя версия программы:"
      ShadowStyle     =   0
      Alignment       =   2
   End
End
Attribute VB_Name = "frmCheckUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbFirstStartUpdate As Boolean

Private Sub cmbVersions_Click()

    With cmbVersions

        If .ListIndex > -1 Then
            strDescription = strUpdDescription(.ListIndex, 0)
            strDescription_en = strUpdDescription(.ListIndex, 1)
        Else
            strDescription = vbNullString
            strDescription_en = vbNullString
        End If
    End With

    LoadDescriptionAndLinks
End Sub

Private Sub cmdDonate_Click()

    frmDonate.Show vbModal, Me
End Sub

Private Sub cmdExit_Click()

    Unload Me
End Sub

Private Sub cmdHistory_Click()

    Dim nRetShellEx As Boolean
    Dim cmdString   As String

    Select Case strPCLangCurrentID

        Case "0419"
            cmdString = kavichki & strLinkHistory & kavichki

        Case Else
            cmdString = kavichki & strLinkHistory_en & kavichki
    End Select

    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdUpdate_Click()

    Dim nRetShellEx As Boolean
    Dim cmdString   As String

    cmdString = kavichki & strLink(cmbVersions.ListIndex, 0) & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdUpdate_ClickMenu(mnuIndex As Integer)

    Dim nRetShellEx As Boolean
    Dim cmdString   As String

    Select Case mnuIndex

        Case 0
            cmdString = kavichki & strLink(cmbVersions.ListIndex, 0) & kavichki

        Case 2
            cmdString = kavichki & strLink(cmbVersions.ListIndex, 2) & kavichki

        Case 4
            cmdString = kavichki & strLink(cmbVersions.ListIndex, 4) & kavichki
    End Select

    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub FontCharsetChange()

    ' Выставляем шрифт
    Me.Font.Name = strOtherForm_FontName
    Me.Font.Size = lngOtherForm_FontSize
    Me.Font.Charset = lngDialog_Charset
    SetButtonProperties cmdUpdate, , False
    SetButtonProperties cmdHistory, , False
    SetButtonProperties cmdDonate, , False
    SetButtonProperties cmdExit, , False
End Sub

Private Sub Form_Activate()

    Dim i As Long

    If mbFirstStartUpdate Then
        ' Загрузка данных с сайта
        LoadUpdateData
        ' установка параметров для кнопок
        LoadDescriptionAndLinks
        ' Показываем список изменений
        lblWait.Visible = False
        rtfDescription.Visible = True
        cmbVersions.Left = lblVersionList.Left + lblVersionList.Width + 50

        For i = LBound(strUpdVersions) To UBound(strUpdVersions)
            cmbVersions.AddItem strUpdVersions(i), i
        Next
        cmbVersions.ListIndex = 0
    End If

    mbFirstStartUpdate = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    'SetSmallIcon Me.hWnd
    Call SetIcon(Me.hwnd, "FRMUPDATE", False)
    
    mbFirstStartUpdate = True
    lblWait.Left = 100
    lblWait.Width = Me.Width - 200
    Me.Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - Me.Width / 2
    Me.Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - Me.Height / 2
    LoadIconImage2Btn cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Btn cmdUpdate, "BTN_UPDATE", strPathImageMainWork
    LoadIconImage2Btn cmdHistory, "BTN_HISTORY", strPathImageMainWork
    LoadIconImage2Btn cmdDonate, "BTN_DONATE", strPathImageMainWork

    ' Локализациz приложения
    If mbLanguageChange Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If
End Sub

Private Sub Form_Terminate()

    If Forms.Count = 0 Then
        UnloadApp
    End If
End Sub

Private Sub lblWWW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = kavichki & "http://www.adia-project.net" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub LoadButtonLink(ButtonName As ctlXpButton, strMassivLink() As String)

    Dim strMirrorText As String

    If cmbVersions.ListIndex > -1 Then

        ' Отличия работы если русский или английский
        Select Case strPCLangCurrentID

            Case "0419"
                strMirrorText = "Зеркало"

            Case Else
                strMirrorText = "Mirror"
        End Select

        With ButtonName

            If InStr(1, strMassivLink(cmbVersions.ListIndex, 0), "http", vbTextCompare) > 0 Then
                .MenuExist = True
            ElseIf InStr(1, strMassivLink(cmbVersions.ListIndex, 2), "http", vbTextCompare) > 0 Then
                .MenuExist = True
            Else
                .MenuExist = False
            End If

            If .MenuExist Then
                If .MenuCount = 0 Then
                    .AddMenu strMirrorText & " 1"
                    .AddMenu "-"
                    .AddMenu strMirrorText & " 2"
                    .AddMenu "-"
                    .AddMenu strMirrorText & " 3"
                End If

                If InStr(1, strMassivLink(cmbVersions.ListIndex, 2), "http", vbTextCompare) = 0 Then
                    .MenuEnabled(2) = False
                End If

                If InStr(1, strMassivLink(cmbVersions.ListIndex, 4), "http", vbTextCompare) = 0 Then
                    .MenuEnabled(4) = False
                End If

                If strMassivLink(cmbVersions.ListIndex, 1) = vbNullString Then
                    .MenuVisible(0) = False
                    .MenuVisible(1) = False
                Else
                    .MenuCaption(0) = strMassivLink(cmbVersions.ListIndex, 1)
                End If

                If strMassivLink(cmbVersions.ListIndex, 3) = vbNullString Then
                    .MenuVisible(1) = False
                    .MenuVisible(2) = False
                Else
                    .MenuCaption(2) = strMassivLink(cmbVersions.ListIndex, 3)
                End If

                If strMassivLink(cmbVersions.ListIndex, 5) = vbNullString Then
                    .MenuVisible(3) = False
                    .MenuVisible(4) = False
                Else
                    .MenuCaption(4) = strMassivLink(cmbVersions.ListIndex, 5)
                End If
            End If
        End With
    End If
End Sub

Private Sub LoadDescriptionAndLinks()

    Dim strDescriptionTemp As String

    ' Отличия работы если русский или английский
    Select Case strPCLangCurrentID

        Case "0419"
            strDescriptionTemp = Replace$(strDescription, vbLf, vbNewLine, , , vbTextCompare)

        Case Else
            strDescriptionTemp = Replace$(strDescription_en, vbLf, vbNewLine, , , vbTextCompare)
    End Select

    ' Кнопка Скачать обновление
    LoadButtonLink cmdUpdate, strLink

    ' Описание изменений
    If strDescriptionTemp <> vbNullString Then
        rtfDescription.TextRTF = strDescriptionTemp
    Else
        rtfDescription.TextRTF = "Error on load ChangeLog. Please inform the developer"
    End If
End Sub

Private Sub Localise(strPathFile As String)

    Dim strFormName As String

    strFormName = CStr(Me.Name)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    'Кнопки
    cmdUpdate.Caption = LocaliseString(strPathFile, strFormName, "cmdUpdate", cmdUpdate.Caption)
    cmdHistory.Caption = LocaliseString(strPathFile, strFormName, "cmdHistory", cmdHistory.Caption)
    cmdDonate.Caption = LocaliseString(strPathFile, strFormName, "cmdDonate", cmdDonate.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    ' Лейблы
    lblVersion.Caption = LocaliseString(strPathFile, strFormName, "lblVersion", lblVersion.Caption) & " " & strVersion & " (" & strDateProg & ")"

    If StrComp(strRelease, "beta", vbTextCompare) = 0 Then
        lblVersion.Caption = lblVersion.Caption & " This version may be Unstable!!!"
        lblVersion.ForeColor = vbRed
    End If

    lblVersionList.Caption = LocaliseString(strPathFile, strFormName, "lblVersionList", lblVersionList.Caption)
    lblWait.Caption = LocaliseString(strPathFile, strFormName, "lblWait", lblWait.Caption)
End Sub
