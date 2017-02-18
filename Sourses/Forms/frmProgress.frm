VERSION 5.00
Begin VB.Form frmProgress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Сбор информации о драйверах. Пожалуйста подождите..."
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8235
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin prjDIADBS.ProgressBar ctlProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      Top             =   120
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   873
      Max             =   10000
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbRunProgress       As Boolean
Private mbFullExit          As Boolean
Private strFormName         As String
Private strFormCaptionTemp  As String

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Const MF_BYPOSITION     As Long = &H400&
Private Const WS_EX_APPWINDOW   As Long = &H40000
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const SW_HIDE           As Long = 0
Private Const SW_SHOW           As Long = 5

Private m_bActivated As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Get CaptionW
'! Description (Описание)  :   [Получение Caption-формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Let CaptionW
'! Description (Описание)  :   [Изменение Caption-формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ChangeFrmMainCaption
'! Description (Описание)  :   [Изменение Caption Формы]
'! Parameters  (Переменные):   lngPercentage (Long)
'!--------------------------------------------------------------------------------
Private Sub ChangeFrmMainCaption(Optional ByVal lngPercentage As Long)

    Dim strProgressValue    As String
    
    If lngPercentage Mod 9999 Then
        If ctlProgressBar1.Visible Then
            strProgressValue = (lngPercentage \ 100) & "% - "
        End If
    End If

    If LenB(strThisBuildBy) = 0 Then
        Me.CaptionW = strProgressValue & strFormCaptionTemp
    Else
        Me.CaptionW = strProgressValue & strFormCaptionTemp
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ChangeProgressBarStatus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ChangeProgressBarStatus(ByRef lngProgressValue As Long, Optional ByVal lngProgressValuePlus As Long = 0)
Attribute ChangeProgressBarStatus.VB_UserMemId = 1610809348

    lngProgressValue = lngProgressValue + lngProgressValuePlus

    If lngProgressValue > 10000 Then
        lngProgressValue = 10000
        Sleep 50
    End If

    With ctlProgressBar1
        .Value = lngProgressValue
        .SetTaskBarProgressValue lngProgressValue, .Max
        ChangeFrmMainCaption lngProgressValue
    End With
    
    DoEvents
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FontCharsetChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' Выставляем шрифт
    With Me.Font
        .Name = strFontOtherForm_Name
        .Size = lngFontOtherForm_Size
        .Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()

    ' Отображаем форму на таскбаре, по умолчанию модальные формы не отображаются
    If Not m_bActivated Then
        m_bActivated = True
        Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW)
        Call ShowWindow(hWnd, SW_HIDE)
        Call ShowWindow(hWnd, SW_SHOW)
    End If
    
    '# call function to read drivers #
    If mbRunProgress Then
        MousePointer = 11
        '# display hourglass cursor while read #
        ReadDrivers
        'ReadDriversByEmun
        mbRunProgress = False
        MousePointer = 0
        '# display default cursor #
    End If

    ' Фиктивная пауза
    'Sleep 300
    Me.Hide
    Unload Me
    Set frmProgress = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, "frmMain", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        ctlProgressBar1.Height = .Height - VPadding(Me)
    End With
    
    '// remove the close menu (which then disables the close button)
    RemoveMenus
    
    mbRunProgress = True

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [Localise message and controls]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(strPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    strFormCaptionTemp = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub RemoveMenus
'! Description (Описание)  :   [Remove "Close" menu from form]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub RemoveMenus()
    Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(Me.hWnd, False)
    DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub

