VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Drivers BackUp Solution"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   12765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   12765
   Begin prjDIADBS.ctlUcStatusBar ctlUcStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6525
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   1244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjDIADBS.ProgressBar ctlProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      Top             =   6030
      Visible         =   0   'False
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   873
      Max             =   10000
   End
   Begin prjDIADBS.ctlJCFrames frPanel 
      Height          =   6120
      Left            =   0
      Top             =   0
      Width           =   12800
      _ExtentX        =   22569
      _ExtentY        =   10795
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14215660
      FillColor       =   14215660
      Style           =   8
      RoundedCorner   =   0   'False
      Caption         =   ""
      IconSize        =   48
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlJCFrames frGroup 
         Height          =   2100
         Left            =   100
         Top             =   75
         Width           =   6100
         _ExtentX        =   10769
         _ExtentY        =   3704
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15783104
         FillColor       =   15783104
         Style           =   4
         RoundedCorner   =   0   'False
         Caption         =   "Выделение группы драйверов"
         TextBoxHeight   =   21
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin prjDIADBS.CheckBoxW chkHideOther 
            Height          =   405
            Left            =   1920
            TabIndex        =   12
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmMain.frx":000C
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.OptionButtonW optGrp1 
            Height          =   400
            Left            =   120
            TabIndex        =   7
            Top             =   500
            Width           =   1700
            _ExtentX        =   2990
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0   'False
            Caption         =   "frmMain.frx":0072
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.OptionButtonW optGrp2 
            Height          =   400
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1700
            _ExtentX        =   2990
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0   'False
            Caption         =   "frmMain.frx":00A4
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.OptionButtonW optGrp3 
            Height          =   400
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   1700
            _ExtentX        =   2990
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0   'False
            Caption         =   "frmMain.frx":00CA
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.OptionButtonW optGrp4 
            Height          =   400
            Left            =   120
            TabIndex        =   10
            Top             =   1560
            Width           =   1700
            _ExtentX        =   2990
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   0   'False
            Caption         =   "frmMain.frx":00F0
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.ctlJCbutton cmdCheckAll 
            Height          =   630
            Left            =   1920
            TabIndex        =   13
            Top             =   1320
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   1111
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonStyle     =   8
            BackColor       =   12244692
            Caption         =   "Выделить всё"
            CaptionEffects  =   0
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            ColorScheme     =   1
         End
         Begin prjDIADBS.ctlJCbutton cmdUnCheckAll 
            Height          =   630
            Left            =   4020
            TabIndex        =   14
            Top             =   1320
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   1111
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonStyle     =   8
            BackColor       =   12244692
            Caption         =   "Снять выделение"
            CaptionEffects  =   0
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            ColorScheme     =   1
         End
         Begin prjDIADBS.CheckBoxW chkCheckAll 
            Height          =   405
            Left            =   1920
            TabIndex        =   11
            Top             =   480
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "frmMain.frx":0122
            Transparent     =   -1  'True
         End
      End
      Begin prjDIADBS.ctlJCFrames frBackUp 
         Height          =   2100
         Left            =   6300
         Top             =   75
         Width           =   6400
         _ExtentX        =   11298
         _ExtentY        =   3704
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   15783104
         FillColor       =   15783104
         Style           =   4
         RoundedCorner   =   0   'False
         Caption         =   "Создание резервной копии выбранных драйверов"
         TextBoxHeight   =   21
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin prjDIADBS.ComboBoxW cmbTypeBackUp 
            Height          =   315
            Left            =   1785
            TabIndex        =   2
            Top             =   495
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
         End
         Begin prjDIADBS.ctlJCbutton cmdStartBackUp 
            Height          =   540
            Left            =   3900
            TabIndex        =   0
            Top             =   930
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   953
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonStyle     =   8
            BackColor       =   12244692
            Caption         =   "Start Backup"
            CaptionEffects  =   0
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            ColorScheme     =   1
         End
         Begin prjDIADBS.ctlJCbutton cmdBreak 
            Height          =   540
            Left            =   3900
            TabIndex        =   1
            Top             =   1500
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   953
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonStyle     =   8
            Enabled         =   0   'False
            BackColor       =   12244692
            Caption         =   "Break"
            CaptionEffects  =   0
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            ColorScheme     =   1
         End
         Begin prjDIADBS.ctlJCFrames frArchName 
            Height          =   1170
            Left            =   0
            Top             =   930
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2064
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14215660
            FillColor       =   14215660
            TextBoxColor    =   12244692
            Style           =   5
            RoundedCorner   =   0   'False
            Caption         =   "Имя Архива"
            Alignment       =   0
            Begin prjDIADBS.TextBoxW txtArchName 
               Height          =   350
               Left            =   120
               TabIndex        =   6
               Top             =   725
               Width           =   3615
               _ExtentX        =   0
               _ExtentY        =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin prjDIADBS.OptionButtonW optArchModelPC 
               Height          =   255
               Left            =   1800
               TabIndex        =   4
               Top             =   360
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   0   'False
               Caption         =   "frmMain.frx":017E
               Transparent     =   -1  'True
            End
            Begin prjDIADBS.OptionButtonW optArchNamePC 
               Height          =   255
               Left            =   1800
               TabIndex        =   3
               Top             =   50
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   450
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   0   'False
               Caption         =   "frmMain.frx":01C0
               Transparent     =   -1  'True
            End
            Begin prjDIADBS.OptionButtonW optArchCustom 
               Height          =   325
               Left            =   120
               TabIndex        =   5
               Top             =   325
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   0   'False
               Caption         =   "frmMain.frx":01FC
               Transparent     =   -1  'True
            End
         End
         Begin prjDIADBS.LabelW lblTypeBackUp 
            Height          =   405
            Left            =   75
            TabIndex        =   17
            Top             =   495
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
            Caption         =   "Type of backup:"
         End
      End
      Begin prjDIADBS.ctlJCFrames frPanelLV 
         Height          =   3690
         Left            =   100
         Top             =   2280
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   6509
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14016736
         FillColor       =   14016736
         TextBoxColor    =   11595760
         TxtBoxShadow    =   1
         Style           =   3
         RoundedCorner   =   0   'False
         Caption         =   "Список найденных драйверов устройств"
         TextBoxHeight   =   21
         ThemeColor      =   3
         GradientHeaderStyle=   1
         Begin prjDIADBS.ListView lvDevices 
            Height          =   3255
            Left            =   60
            TabIndex        =   15
            Top             =   360
            Width           =   12450
            _ExtentX        =   21960
            _ExtentY        =   5741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   1
            OLEDragDropScroll=   0   'False
            Redraw          =   0   'False
            View            =   3
            Arrange         =   1
            AllowColumnReorder=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            LabelEdit       =   2
            LabelWrap       =   0   'False
            Sorted          =   -1  'True
            Checkboxes      =   -1  'True
            HideSelection   =   0   'False
            HoverSelection  =   -1  'True
            HotTracking     =   -1  'True
            HighlightHot    =   -1  'True
            TextBackground  =   1
         End
         Begin prjDIADBS.LabelW lblWait 
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Visible         =   0   'False
            Width           =   11640
            _ExtentX        =   17383
            _ExtentY        =   688
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            BackStyle       =   0
            Caption         =   "Идет обновление конфигурации оборудования. Пожалуйста, подождите...."
         End
      End
   End
   Begin VB.Menu mnuReCollectHWID 
      Caption         =   "Обновить информацию"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Параметры"
   End
   Begin VB.Menu mnuMainAbout 
      Caption         =   "Справка"
      Begin VB.Menu mnuLinks 
         Caption         =   "Ссылки"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "История изменения"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Справка по работе"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHomePage 
         Caption         =   "Домашная страница программы"
      End
      Begin VB.Menu mnuHomePageForum 
         Caption         =   "Форум программы"
      End
      Begin VB.Menu mnuOsZoneNet 
         Caption         =   "Обсуждение программы на OsZone.net"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUpd 
         Caption         =   "Проверить обновление программы"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModulesVersion 
         Caption         =   "Модули..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDonate 
         Caption         =   "Поблагодарить автора..."
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "Лицензионное соглашение"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе..."
      End
   End
   Begin VB.Menu mnuMainLang 
      Caption         =   "Язык"
      Begin VB.Menu mnuLangStart 
         Caption         =   "Использовать выбранный язык при запуске (отмена автовыбора)"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLang 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbBreakUpdateDBAll        As Boolean
Private cmbListTypeBackupElement1 As String
Private cmbListTypeBackupElement2 As String
Private cmbListTypeBackupElement3 As String
Private strTableHwidHeader1       As String
Private strTableHwidHeader2       As String
Private strTableHwidHeader3       As String
Private strTableHwidHeader4       As String
Private strTableHwidHeader5       As String
Private strTableHwidHeader6       As String
Private strTableHwidHeader7       As String
Private strTableHwidHeader8       As String
Private strTableHwidHeader9       As String
Private strTableHwidHeader10      As String
Private strTableHwidHeader11      As String
Private arrSourceDisksFiles_x()   As String
Private arrSourceDisksNames_x()   As String
Private lngFrameTime              As Long
Private lngFrameCount             As Long
Private lngBorderWidthX           As Long
Private lngBorderWidthY           As Long
Private strFormName               As String
Private mbMassCheckInLV           As Boolean

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
'! Procedure   (Функция)   :   Sub BlockControl
'! Description (Описание)  :   [Блокировка(Разблокировка) некоторых элементов формы при работе сложных функций]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub BlockControl(ByVal mbEnable As Boolean)

    cmdCheckAll.Enabled = Not mbEnable
    cmdUnCheckAll.Enabled = Not mbEnable
    optGrp1.Enabled = Not mbEnable
    optGrp2.Enabled = Not mbEnable
    optGrp3.Enabled = Not mbEnable
    optGrp4.Enabled = Not mbEnable
    chkHideOther.Enabled = Not mbEnable
    cmdStartBackUp.Enabled = Not mbEnable
    cmdBreak.Enabled = mbEnable
    cmbTypeBackUp.Enabled = Not mbEnable
    frPanelLV.Enabled = Not mbEnable
    chkCheckAll.Enabled = Not mbEnable
    mnuReCollectHWID.Enabled = Not mbEnable
    mnuOptions.Enabled = Not mbEnable
    mnuMainAbout.Enabled = Not mbEnable
    mnuMainLang.Enabled = Not mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ChangeFrmMainCaption
'! Description (Описание)  :   [Изменение Caption Формы]
'! Parameters  (Переменные):   lngPercentage (Long)
'!--------------------------------------------------------------------------------
Private Sub ChangeFrmMainCaption(Optional ByVal lngPercentage As Long)

    Dim strProgressValue As String
    Dim strStatusBarText As String
    Dim lngRet           As Long

    Select Case strPCLangCurrentID

        Case "0419"
            strFrmMainCaptionTemp = strProjectNameFull
            strFrmMainCaptionTempDate = " (Дата релиза: " & strDateProgram & ")"

        Case Else
            strFrmMainCaptionTemp = strProjectNameFull
            strFrmMainCaptionTempDate = " (Date Build: " & strDateProgram & ")"
    End Select

    If lngPercentage Mod 9999 Then
        If ctlProgressBar1.Visible Then
            strStatusBarText = ctlUcStatusBar1.PanelText(1)
            lngRet = InStr(strStatusBarText, ":")
            If lngRet Then
                strStatusBarText = Left$(strStatusBarText, lngRet - 1)
            End If
            lngRet = InStr(strStatusBarText, ".")
            If lngRet Then
                strStatusBarText = Left$(strStatusBarText, lngRet - 1)
            End If
            strProgressValue = (lngPercentage \ 100) & "% (" & strStatusBarText & ") - "
        End If
    End If

    If LenB(strThisBuildBy) = 0 Then
        Me.CaptionW = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & strFrmMainCaptionTempDate & " @" & App.CompanyName & " - " & strPCLangCurrentLangName
    Else
        Me.CaptionW = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & strFrmMainCaptionTempDate & " " & strThisBuildBy & " - " & strPCLangCurrentLangName
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkHideOther_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkHideOther_Click()
    chkCheckAll.Enabled = CBool(chkHideOther.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdBreak_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdBreak_Click()
    mbBreakUpdateDBAll = True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdCheckAll_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdCheckAll_Click()

    Dim ii As Long

    mbMassCheckInLV = True
    
    With lvDevices.ListItems

        For ii = 1 To .count

            If Not .item(ii).Checked Then
                .item(ii).Checked = True
            End If

        Next

    End With

    mbMassCheckInLV = False
    FindCheckCountList
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdStartBackUp_Click
'! Description (Описание)  :   [do backup of selected drivers]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdStartBackUp_Click()
    StartBackUp
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdUnCheckAll_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdUnCheckAll_Click()

    Dim ii As Long

    mbMassCheckInLV = True
    With lvDevices.ListItems

        For ii = 1 To .count

            If .item(ii).Checked Then
                .item(ii).Checked = False
            End If

        Next

    End With

    mbMassCheckInLV = False
    FindCheckCountList
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CollectDestPathFiles
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CollectDestPathFiles(ByVal strPathInfFile As String)

    Dim arr1_SDN()               As String
    Dim arr2_SDF()               As String
    Dim strSecNameSDN            As String
    Dim strSecNameSDF            As String
    Dim lngArrCount              As Long
    Dim lngArrCount2             As Long
    Dim strDestPathTemp          As String
    Dim strDestPathTransform     As String
    Dim strDestPathTransform_x() As String
    Dim strDestTempPart1         As String

    If mbDebugDetail Then DebugMode "***CollectDestPathFiles-Start"
    'SourceDisksNames
    strSecNameSDN = "SourceDisksNames" & "." & strOSArchitecture

    If Not CheckIniSectionExists(strSecNameSDN, strPathInfFile) Then
        strSecNameSDN = "SourceDisksNames"
    End If

    arr1_SDN = GetSectionMass(strSecNameSDN, strPathInfFile, False)

    For lngArrCount = 1 To UBound(arr1_SDN, 1)
        strDestPathTemp = vbNullString
        strDestPathTransform = vbNullString
        strDestPathTemp = arr1_SDN(lngArrCount, 2)
        strDestPathTransform_x() = Split(strDestPathTemp, ",")

        If UBound(strDestPathTransform_x) = 3 Then
            strDestPathTransform = strDestPathTransform_x(3)
            strDestPathTransform = StringCleaner(strDestPathTransform)
        Else
            strDestPathTransform = vbNullString
        End If

        arr1_SDN(lngArrCount, 2) = strDestPathTransform
    Next
    
    'SourceDisksFiles
    strSecNameSDF = "SourceDisksFiles." & strOSArchitecture

    If Not CheckIniSectionExists(strSecNameSDF, strPathInfFile) Then
        strSecNameSDF = "SourceDisksFiles"
    End If

    arr2_SDF = GetSectionMass(strSecNameSDF, strPathInfFile, False)

    For lngArrCount = 1 To UBound(arr2_SDF, 1)
        strDestPathTemp = vbNullString
        strDestPathTransform = vbNullString
        strDestPathTemp = arr2_SDF(lngArrCount, 2)
        strDestPathTransform_x() = Split(strDestPathTemp, ",")

        If UBound(strDestPathTransform_x) >= 1 Then
            strDestPathTransform = strDestPathTransform_x(1)
            strDestPathTransform = StringCleaner(strDestPathTransform)
        End If

        For lngArrCount2 = 1 To UBound(arr1_SDN, 1)

            If StrComp(arr1_SDN(lngArrCount2, 1), strDestPathTransform_x(0), vbTextCompare) = 0 Then
                strDestTempPart1 = arr1_SDN(lngArrCount2, 2)
                strDestPathTransform = PathCollect4Dest(strDestPathTransform, strDestTempPart1)
                Exit For
            Else
                strDestPathTransform = vbNullString
            End If

        Next
        arr2_SDF(lngArrCount, 2) = strDestPathTransform
    Next
    
    arrSourceDisksFiles_x = arr2_SDF
    arrSourceDisksNames_x = arr1_SDN
    If mbDebugDetail Then DebugMode "***CollectDestPathFiles-Finish"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CollectDPName
'! Description (Описание)  :   [Собрать имя архива]
'! Parameters  (Переменные):   strPCName (String)
'!--------------------------------------------------------------------------------
Private Function CollectDPName(ByVal strPCName As String) As String

    Dim strDPName       As String
    Dim strDPName_Part1 As String
    Dim strDPName_Part2 As String
    Dim strDPName_Part3 As String

    strDPName_Part1 = "_wnt" & OSCurrVersionStruct.VerMajor

    If mbIsWin64 Then
        strDPName_Part2 = "_x64_"
    Else
        strDPName_Part2 = "_x32_"
    End If

    strDPName_Part3 = Replace$(CStr(Date), ".", "-")
    strDPName_Part3 = SafeDir(strDPName_Part3)
    strDPName = "DP_" & strPCName & strDPName_Part1 & strDPName_Part2 & strDPName_Part3
    strDPName = SafeDir(strDPName)
    CollectDPName = Replace$(strDPName, " ", "_")
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CopyFile2Dest
'! Description (Описание)  :   [Копирование файлов из секции inf в каталог назначения]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub CopyFile2Dest(ByRef arrZ() As String, _
                          ByVal strDestination As String, _
                          ByVal strDestFolderSection As String, _
                          ByVal strInfFile As String, _
                          Optional ByVal mbSectCopyFiles As Boolean = False)

    Dim strFileName         As String
    Dim strFileName_x()     As String
    Dim strFileNameFrom     As String
    Dim strFileNameFrom_x() As FindListStruct
    Dim strFileNameTo       As String
    Dim strDestPath4File    As String
    Dim strDestinationTemp  As String
    Dim ii                  As Long
    Dim strExt              As String
    Dim strSpecDir          As String
    Dim strCustomDir        As String
    Dim lngOldValue         As Long
    Dim lngArrCount         As Long
    Dim lngUbound           As Long
    Dim lngUBoundZ          As Long
    Dim lngUBoundFileName   As Long

    lngUBoundZ = UBound(arrZ)

    For ii = 0 To lngUBoundZ
        strFileName = arrZ(ii)

        ' если пустое значение, то пропускаем
        If LenB(strFileName) Then
            If mbSectCopyFiles Then
                If InStr(1, strFileName, ",") Then
                    strFileName_x = Split(strFileName, ",")
                    lngUBoundFileName = UBound(strFileName_x)

                    If lngUBoundFileName >= 1 Then
                        strFileName = strFileName_x(0)
                        strFileNameTo = SafeFileName(strFileName_x(1))
                    End If
                End If
            End If

            ' Убираем все лишнее из имени файла
            strFileName = SafeFileName(strFileName)

            ' если пустое значение, то пропускаем
            If LenB(strFileName) Then

                ' Если строка содержит ".", значит это скорее все имя файла
                If InStr(1, strFileName, ".") Then

                    ' Куда будет скопирован файл
                    lngUbound = UBound(arrSourceDisksFiles_x, 1)

                    For lngArrCount = 1 To lngUbound

                        If StrComp(arrSourceDisksFiles_x(lngArrCount, 1), strFileName, vbTextCompare) = 0 Then
                            strDestinationTemp = arrSourceDisksFiles_x(lngArrCount, 2)
                            strDestinationTemp = PathCollect4Dest(strDestinationTemp, strDestination)
                            Exit For
                        Else
                            strDestinationTemp = strDestination
                        End If

                    Next

                    ' создаем каталог назначения, если его нет
                    If PathExists(strDestinationTemp) = False Then
                        CreateNewDirectory strDestinationTemp
                    End If

                    ' собственно полный путь копируемого файла
                    If LenB(strFileNameTo) Then
                        If mbSectCopyFiles Then
                            strDestPath4File = PathCombine(strDestinationTemp, strFileNameTo)
                        Else
                            strDestPath4File = PathCombine(strDestinationTemp, strFileName)
                        End If

                    Else
                        strDestPath4File = PathCombine(strDestinationTemp, strFileName)
                    End If

                    ' определяем каталог, где должен лежать файл по числовому коду
                    strCustomDir = ReadFromINI("DestinationDirs", strDestFolderSection, strInfFile, vbNullString)

                    'Если каталог не определен, то используем каталог по дефолту
                    If LenB(strCustomDir) = 0 Then
                        strCustomDir = ReadFromINI("DestinationDirs", "DefaultDestDir", strInfFile, vbNullString)
                    End If

                    'если все равно не определен, то пропускаем
                    If LenB(strCustomDir) Then
                        '# if it is #
                        strSpecDir = WhereIsDir(strCustomDir, strInfFile)

                        ' если x64, то устанавливаем отключение перенаправления для папки system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64DisableWow64FsRedirection", "kernel32.dll") Then
                                Wow64DisableWow64FsRedirection lngOldValue
                            End If
                        End If

                        ' Копирование файла
                        strFileNameFrom = PathCombine(strSpecDir, strFileName)

                        If FileExists(strFileNameFrom) Then
                            If FileExists(strDestPath4File) = False Then
                                CopyFileTo strSpecDir & strFileName, strDestPath4File
                                If mbDebugStandart Then DebugMode "******Backup File: FROM=" & strFileNameFrom & " TO=" & strDestPath4File
                            End If
                        End If

                        ' Если это драйвера принтера, то ищем по всей папке
                        If InStr(1, strSpecDir, strSysDir86 & "spool\Drivers\w32x86", vbTextCompare) > 0 Then

                            '# search for correctly driver if has more tha one printer #
                            ' ищем файл по всей папке strSysDir & "\spool\Drivers\w32x86"
                            If FileExists(strDestPath4File) = False Then
                                strFileNameFrom_x = SearchFilesInRoot(strSpecDir, strFileName, True, True)

                                If LenB(strFileNameFrom_x(0).FullPath) Then
                                    CopyFileTo strFileNameFrom_x(0).FullPath, strDestPath4File
                                End If
                            End If
                        End If

                        ' если x64, то включаем обратно перенаправления для папки system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64RevertWow64FsRedirection", "kernel32.dll") Then
                                Wow64RevertWow64FsRedirection lngOldValue
                            End If
                        End If
                    End If

                    ' Дополнительный поиск файлов по расширению, если файл все еще не найден
                    If FileExists(strDestPath4File) = False Then
                        'Расширение файла
                        strExt = GetFileNameExtension(strFileName)

                        ' если x64, то устанавливаем отключение перенаправления для папки system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64DisableWow64FsRedirection", "kernel32.dll") Then
                                Wow64DisableWow64FsRedirection lngOldValue
                            End If
                        End If

                        If strExt = "hlp" Then
                            strFileNameFrom = PathCombine(strWinDirHelp, strFileName)
                            If FileExists(strFileNameFrom) Then
                                CopyFileTo strFileNameFrom, strDestPath4File
                            End If

                        ElseIf strExt = "sys" Then

                            strFileNameFrom = PathCombine(strSysDirDrivers, strFileName)
                            If FileExists(strFileNameFrom) Then
                                CopyFileTo strFileNameFrom, strDestPath4File
                            End If
                            
                            strFileNameFrom = PathCombine(strSysDirDrivers64, strFileName)
                            If FileExists(strFileNameFrom) Then
                                CopyFileTo strFileNameFrom, strDestPath4File
                            End If
                            
                        Else

                            strFileNameFrom = PathCombine(strSysDir86, strFileName)
                            If FileExists(strFileNameFrom) Then
                                CopyFileTo strFileNameFrom, strDestPath4File
                            End If
                            
                            strFileNameFrom = PathCombine(strSysDir64, strFileName)
                            If FileExists(strFileNameFrom) Then
                                CopyFileTo strFileNameFrom, strDestPath4File
                            End If
                            
                        End If

                        ' если x64, то включаем обратно перенаправления для папки system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64RevertWow64FsRedirection", "kernel32.dll") Then
                                Wow64RevertWow64FsRedirection lngOldValue
                            End If
                        End If
                    End If
                End If
            End If
        End If

    Next
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateMenuLng
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strMenuCaption (String)
'!--------------------------------------------------------------------------------
Private Sub CreateMenuLng()
    Dim ii  As Long
    Dim iii As Long

    If Not mnuLang(0).Visible Then
        'если меню еще не создано
        mnuLang(0).Visible = True
    End If
    
    ' Создаем динамическое меню
    iii = 0
    For ii = UBound(arrLanguage, 2) To 0 Step -1
        If iii > 0 Then Load mnuLang(iii)
        mnuLang(iii).Visible = True
        mnuLang(iii).Caption = "Lang " & iii
        iii = iii + 1
    Next ii
    
    ' Присваиваем свойство Caption для меню
    For ii = 0 To UBound(arrLanguage, 2)
        '3  mnuMainLang - "Язык"
        ' 2    mnuLang - "" - Index0 - Visible'False
        SetUniMenu 3, 2 + ii, -1, mnuLang(ii), arrLanguage(1, ii)
    Next ii

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DefineFolderBackUp
'! Description (Описание)  :   [Определение каталога назначения для резервных копий]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function DefineFolderBackUp() As String

    Dim ii                As Long
    Dim strDestFolder     As String
    Dim strDestFolderTemp As String
    Dim str_x64           As String

    If mbBackFolderPredefine Then
        
        ' Просматриваем в цикле настройки
        For ii = 0 To UBound(arrOSList)
            str_x64 = arrOSList(ii).is64bit
            strDestFolderTemp = arrOSList(ii).drpFolder
        
            If InStr(1, arrOSList(ii).Ver, strOSCurrentVersion) Then
                If CBool(str_x64) = mbIsWin64 Then
                    strDestFolder = PathCollect(strDestFolderTemp)

                    If PathExists(strDestFolder) = False Then
                        strDestFolder = vbNullString
                    End If

                    Exit For
                End If
            End If

        Next
    End If

    If LenB(strDestFolder) Then
        DefineFolderBackUp = strDestFolder
    Else
        DefineFolderBackUp = strAppPathBackSL & "drivers\"
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DoZip
'! Description (Описание)  :   [Создание архива с драйверами]
'! Parameters  (Переменные):   strPackFolder As String, ByVal strDPName As String
'!--------------------------------------------------------------------------------
Private Function DoZip(ByVal strPackFolder As String, ByVal strDPName As String) As Boolean

    Dim cmdString             As String
    Dim strDpName7z           As String
    Dim strDpNameExt          As String
    Dim strDpNamewoExt        As String
    Dim mbCreateSFX           As Boolean
    Dim strDPInstPath         As String
    Dim lngNumFilesFromFolder As Long

    ' получаем расширение файла архива (exe или 7Z)
    strDpNameExt = GetFileNameExtension(strDPName)
    strDpNamewoExt = GetFileName_woExt(strDPName)

    If StrComp(strDpNameExt, "exe", vbTextCompare) = 0 Then
        strDpName7z = strDpNamewoExt & ".7z"
        mbCreateSFX = True
    Else
        strDpName7z = strDPName
    End If

    ' Удаляем старые архивы если есть
    If FileExists(strDpName7z) Then
        If mbDebugStandart Then DebugMode "***DoZip: Clean previous drivers archive "
        DeleteFiles strDpName7z
    End If

    If mbCreateSFX Then
        If FileExists(strDPName) Then
            If mbDebugStandart Then DebugMode "***DoZip: Clean previous drivers archive "
            DeleteFiles strDPName
        End If

        ' Копируем файлы DPInst для автозапуска
        strDPInstPath = GetPathNameFromPath(strDPInstExePath)
        If mbDebugStandart Then DebugMode "******CopyFiles DPINST : " & strDPInstPath
        ChangeStatusBarText "Copying files from DPInst folder: " & strDPInstPath
        'Изменяем прогресс на caption
        With ctlProgressBar1
            .Value = 9100
            .SetTaskBarProgressValue .Value, .Max
            ChangeFrmMainCaption .Value
        End With
        lngNumFilesFromFolder = rgbCopyFiles(strDPInstPath, strPackFolder, ALL_FILES)
        If mbDebugStandart Then DebugMode "******CopyFiles - count files: " & lngNumFilesFromFolder
    End If

    ' Первая стадия упаковки
    '..\7za.exe a ..\out\%1 -mmt=off -m0=BCJ2 -m1=LZMA2:d%dict%m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 *.ini -ir!*.inf
    cmdString = strQuotes & strArh7zExePath & strQuotes & " a " & strQuotes & strDpName7z & strQuotes & " " & strArh7zParam1
    ChangeStatusBarText strMessages(97) & " " & strDpName7z, "Compressing...: " & cmdString
    'Изменяем прогресс на caption
    With ctlProgressBar1
        .Value = 9200
        .SetTaskBarProgressValue .Value, .Max
        ChangeFrmMainCaption .Value
    End With
    
    
    If RunAndWait(cmdString, strPackFolder, vbHide) = False Then
        MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
        DoZip = False
        ChangeStatusBarText strMessages(13) & " " & strDpName7z, "Error on run : " & cmdString
    Else

        ' Архиватор отработал на все 100%? Если нет то сообщаем
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusBarText strMessages(13) & strDpName7z
            MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
            DoZip = False
        End If

        DoZip = True
        ChangeStatusBarText "7z-archive (STEP 1) successfully done!!!"
    End If

    ' Вторая стадия упаковки
    '..\7za.exe a ..\out\%1 -mmt=off -m0=BCJ2 -m1=LZMA2:d%dict%m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 -xr!*.inf -x!*.ini
    cmdString = strQuotes & strArh7zExePath & strQuotes & " a " & strQuotes & strDpName7z & strQuotes & " " & strArh7zParam2
    ChangeStatusBarText strMessages(97) & " " & strDpName7z, "Compressing...: " & cmdString
    'Изменяем прогресс на caption
    With ctlProgressBar1
        .Value = 9300
        .SetTaskBarProgressValue .Value, .Max
        ChangeFrmMainCaption .Value
    End With

    If RunAndWait(cmdString, strPackFolder, vbHide) = False Then
        MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
        DoZip = False
        ChangeStatusBarText strMessages(13) & " " & strDpName7z, "Error on run : " & cmdString
    Else

        ' Архиватор отработал на все 100%? Если нет то сообщаем
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusBarText strMessages(13) & strDpName7z
            MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
            DoZip = False
        End If

        DoZip = True
        ChangeStatusBarText "7z-archive (STEP 2) successfully done!!!"
    End If

    If mbCreateSFX Then

        ' Третья стадия упаковки SFX
        'copy /b "d:\aWork\myProg\DriversBackuper\Tools\Arc\sfx\7zSD.sfx" + "d:\aWork\myProg\DriversBackuper\Tools\Arc\sfx\config.txt" + "D:\aWork\myProg\DriversBackuper\drivers\2k_xp_2003\x64\DP_0300-B01951_wnt5_x32_03-03-2011.7z" "D:\aWork\myProg\DriversBackuper\drivers\2k_xp_2003\x64\DP_0300-B01951_wnt5_x32_03-03-2011.exe"
        Select Case strPCLangCurrentID

            Case "0419"
                cmdString = "cmd.exe /C copy /b " & strQuotes & strArh7zSFXPATH & strQuotes & " + " & strQuotes & strArh7zSFXConfigPath & strQuotes & " + " & strQuotes & strDpName7z & strQuotes & " " & strQuotes & strDPName & strQuotes

            Case Else
                cmdString = "cmd.exe /C copy /b " & strQuotes & strArh7zSFXPATH & strQuotes & " + " & strQuotes & strArh7zSFXConfigPathEn & strQuotes & " + " & strQuotes & strDpName7z & strQuotes & " " & strQuotes & strDPName & strQuotes
        End Select

        ChangeStatusBarText strMessages(97) & " " & strDPName, "Creating SFX...: " & cmdString
        'Изменяем прогресс на caption
        With ctlProgressBar1
            .Value = 9900
            .SetTaskBarProgressValue .Value, .Max
            ChangeFrmMainCaption .Value
        End With

        If RunAndWait(cmdString, strWorkTemp, vbHide) = False Then
            MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
            DoZip = False
            ChangeStatusBarText strMessages(13) & " " & strDPName, "Error on run : " & cmdString
        Else

            If FileExists(strDPName) Then
                If FileExists(strDpName7z) Then
                    If mbDebugStandart Then DebugMode "***DoZip: Clean temp drivers archive "
                    DeleteFiles strDpName7z
                End If

                DoZip = True
                ChangeStatusBarText "7z-archive (STEP 3) successfully done!!! SFX-archive created"
            Else
                MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
                DoZip = False
                ChangeStatusBarText strMessages(13) & " " & strDPName, "Error on run : " & cmdString
            End If
        End If
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ExpandArchNamebyEnvironment
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strArchName As String
'!--------------------------------------------------------------------------------
Private Function ExpandArchNamebyEnvironment(ByVal strArchName As String) As String

    Dim strRet          As String
    Dim strDPName_OSVer As String
    Dim strDPName_OSBit As String
    Dim strDPName_DATE  As String

    If InStr(1, strArchName, "%") Then
        ' Макроподстановка версия ОС %OSVer%
        strDPName_OSVer = "wnt" & OSCurrVersionStruct.VerMajor

        ' Макроподстановка битность ОС %OSBit%
        If mbIsWin64 Then
            strDPName_OSBit = "x64"
        Else
            strDPName_OSBit = "x32"
        End If

        ' Макроподстановка ДАТА %DATE%
        strDPName_DATE = Replace$(CStr(Date), ".", "-")
        strDPName_DATE = SafeDir(strDPName_DATE)
        ' Замена макросов значениями
        strRet = strArchName
        strRet = Replace$(strRet, "%PCNAME%", strCompName)
        strRet = Replace$(strRet, "%PCMODEL%", Replace$(strCompModel, " ", "_"))
        strRet = Replace$(strRet, "%OSVer%", strDPName_OSVer)
        strRet = Replace$(strRet, "%OSBit%", strDPName_OSBit)
        strRet = Replace$(strRet, "%DATE%", strDPName_DATE)
        strRet = Trim$(strRet)
        ExpandArchNamebyEnvironment = strRet
    Else
        ExpandArchNamebyEnvironment = strArchName
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindCheckCountList
'! Description (Описание)  :   [Поиск выделенных строк]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function FindCheckCountList(Optional ByVal mbOnlyCollect As Boolean = False) As Long

    Dim miCount As Integer
    Dim ii      As Integer

    For ii = 1 To lvDevices.ListItems.count

        If lvDevices.ListItems.item(ii).Checked Then
            miCount = miCount + 1
        End If

    Next
    
    If Not mbOnlyCollect Then
        cmdStartBackUp.Caption = LocaliseString(strPCLangCurrentPath, Me.Name, "cmdStartBackUp", "Start Backup")
    
        If miCount Then
    
            With cmdStartBackUp
    
                If Not .Enabled Then
                    .Enabled = True
                End If
    
                .Caption = .Caption & " (" & miCount & ")"
            End With
    
        Else
    
            With cmdStartBackUp
    
                If .Enabled Then
                    .Enabled = False
                End If
            End With
        End If
        cmdStartBackUp.ToolTipText = cmdStartBackUp.Caption
    End If
    
    FindCheckCountList = miCount
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FindCopyCatFile
'! Description (Описание)  :   [Поиск и копирование файла cat]
'! Parameters  (Переменные):   ByVal strInfFilePath As String, ByVal strDestination As String
'!--------------------------------------------------------------------------------
Private Function FindCopyCatFile(ByVal strInfFilePath As String, ByVal strDestination As String) As String

    Dim strCatFile         As String
    Dim strCatFilePathTemp As String
    Dim strCatFile_ntx86   As String
    Dim strCatFile_ntamd64 As String
    Dim strCatFile_nt      As String
    Dim strCatFilePath_x() As FindListStruct
    Dim strCatFileFromInf  As String
    Dim mbExitGoto         As Boolean

    '# Ищем в файле inf - catalog file (Каталог безопасности)
    strCatFile = ReadFromINI("Version", "CatalogFile", strInfFilePath, vbNullString)
    strCatFile_nt = ReadFromINI("Version", "CatalogFile.nt", strInfFilePath, vbNullString)
    strCatFile_ntx86 = ReadFromINI("Version", "CatalogFile.ntx86", strInfFilePath, vbNullString)
    strCatFile_ntamd64 = ReadFromINI("Version", "CatalogFile.ntamd64", strInfFilePath, vbNullString)
    strCatFile = SafeFileName(strCatFile)

    If LenB(strCatFile) = 0 Then
        If LenB(strCatFile_ntx86) Then
            strCatFile = strCatFile_ntx86
        ElseIf LenB(strCatFile_ntamd64) Then
            strCatFile = strCatFile_ntamd64
        ElseIf LenB(strCatFile_nt) Then
            strCatFile = strCatFile_nt
        Else
            strCatFile = vbNullString
        End If
    End If

    'strCatFileFromInf = GetFileName_woExt(GetFileNameFromPath(strInfFilePath)) & ".cat"
    strCatFileFromInf = GetFileNameOnly_woExt(strInfFilePath) & ".cat"
    
CopyCatAgain:

    '# if has catalog file #
    If LenB(strCatFile) Then
        
        strCatFilePathTemp = PathCombine(strDestination, strCatFile)
        
        ' ищем файл cat его по всей папке strSysDirCatRoot c именем из полученным из файла inf
        If FileExists(strCatFilePathTemp) = False Then
            strCatFilePath_x = SearchFilesInRoot(strSysDirCatRoot, strCatFile, True, True)

            If LenB(strCatFilePath_x(0).FullPath) Then
                CopyFileTo strCatFilePath_x(0).FullPath, BackslashAdd2Path(strDestination) & strCatFile
                If mbDebugStandart Then DebugMode "***CatalogFile find in: " & strCatFilePath_x(0).FullPath
            End If
        End If
        DoEvents

        ' ищем файл cat его по всей папке strSysDirCatRoot c именем аналогичным файлу inf
        If FileExists(strCatFilePathTemp) = False Then
            strCatFilePath_x = SearchFilesInRoot(strSysDirCatRoot, strCatFileFromInf, True, True)

            If LenB(strCatFilePath_x(0).FullPath) Then
                CopyFileTo strCatFilePath_x(0).FullPath, BackslashAdd2Path(strDestination) & strCatFile
                If mbDebugStandart Then DebugMode "***CatalogFile find in: " & strCatFilePath_x(0).FullPath
            End If
        End If
        DoEvents

        ' ищем файл cat его по всей папке strSysDirDRVStore
        If FileExists(strCatFilePathTemp) = False Then
            strCatFilePath_x = SearchFilesInRoot(strSysDirDRVStore, strCatFile, True, True)

            If LenB(strCatFilePath_x(0).FullPath) Then
                CopyFileTo strCatFilePath_x(0).FullPath, BackslashAdd2Path(strDestination) & strCatFile
                If mbDebugStandart Then DebugMode "***CatalogFile find in: " & strCatFilePath_x(0).FullPath
            End If
        End If
        DoEvents

        ' Если файл cat все еще не найден, то ищем его по всей папке windows
        If FileExists(strCatFilePathTemp) = False Then
            strCatFilePath_x = SearchFilesInRoot(strWinDir, strCatFile, True, True)

            If LenB(strCatFilePath_x(0).FullPath) Then
                CopyFileTo strCatFilePath_x(0).FullPath, BackslashAdd2Path(strDestination) & strCatFile
                If mbDebugStandart Then DebugMode "***CatalogFile find in: " & strCatFilePath_x(0).FullPath
            End If
        End If
        DoEvents

        ' Если файл найден, то имя файла передаем обратно функции для дальнейшего использования
        If FileExists(strCatFilePathTemp) Then
            FindCopyCatFile = strCatFile
        Else

            'если не найден файл? то пытаемся найти его используя ключи  strCatFile_ntx86 и strCatFile_ntamd64
            If LenB(strCatFile_ntx86) Then
                If LenB(strCatFile_ntamd64) Then
                    If Not mbExitGoto Then
                        mbExitGoto = True
                        strCatFile = strCatFile_ntamd64
                        DoEvents
                        GoTo CopyCatAgain
                    End If
                End If
            End If
        End If
    End If

    If FileExists(BackslashAdd2Path(strDestination) & strCatFile) = False Then
        If mbDebugStandart Then DebugMode "***CatalogFile not find: " & strCatFile
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FontCharsetChange
'! Description (Описание)  :   [Изменение шрифта для элементов]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' Выставляем шрифт формы
    With Me.Font
        .Name = strFontMainForm_Name
        .Size = lngFontMainForm_Size
        .Charset = lngFont_Charset
    End With
    
    ' Выставляем шрифт кнопок
    SetBtnFontProperties cmdCheckAll
    SetBtnFontProperties cmdUnCheckAll
    SetBtnFontProperties cmdStartBackUp
    SetBtnFontProperties cmdBreak
        
    ' Выставляем шрифт других элементов
    frGroup.Font.Charset = lngFont_Charset
    frBackUp.Font.Charset = lngFont_Charset
    frArchName.Font.Charset = lngFont_Charset
    frPanelLV.Font.Charset = lngFont_Charset
    ctlUcStatusBar1.Font.Charset = lngFont_Charset
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [Активация формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()

    If mbFirstStart Then
        If mbStartMaximazed Then
            Me.WindowState = vbMaximized
            DoEvents
        ElseIf mbChangeResolution Then
            Me.WindowState = vbMaximized
            DoEvents
        End If

        DoEvents

        ' Выгрузка формы Показ Лицензионного соглашения, если есть
        If IsFormLoaded("frmLicence") Then
            Unload frmLicence
            Set frmLicence = Nothing
        End If
            
        ' Проверка обновлений при старте
        If mbUpdateCheck Then
            ChangeStatusBarText strMessages(58)
            CheckUpd
            mbFirstStart = False
        Else
            ShowUpdateToolTip
        End If

        ChangeStatusBarText strMessages(1)
        mbFirstStart = False
    End If

    mbFirstStart = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [обработка нажатий клавиш клавиатуры]
'! Parameters  (Переменные):   KeyCode As Integer, Shift As Integer
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Выход из программы по "Escape"
    If KeyCode = vbKeyEscape Then
        If MsgBox(strMessages(34), vbQuestion + vbYesNo, strProductName) = vbYes Then
            Unload Me
        End If
    End If

    If KeyCode = vbKeyF5 Then
        ' Сбор инормации о компе
        mnuReCollectHWID_Click
    End If

    ' Нажата кнопка "Ctrl"
    If Shift = 2 Then

        Select Case KeyCode

            Case 65
                ' Ctrl+A (Выделение рекомендуемых пакетов для установки)
                cmdCheckAll_Click

            Case 90
                ' Ctrl+Z (Выделение рекомендуемых пакетов для установки)
                cmdUnCheckAll_Click

            Case 19
                ' CTRL+Break (Прерывание групповой обработки)
                If cmdBreak.Visible Then
                    cmdBreak_Click
                End If

            Case 49
                ' CTRL+1 (Переключение между группами)
                optGrp1.Value = True
                
            Case 50
                ' CTRL+2 (Переключение между группами)
                optGrp2.Value = True
                
            Case 51
                ' CTRL+3 (Переключение между группами)
                optGrp3.Value = True
                
            Case 52
                ' CTRL+4 (Переключение между группами)
                optGrp4.Value = True
                
        End Select
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [Событие загрузки формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()

    Dim ii  As Long
    Dim iii As Long

    If mbDebugStandart Then DebugMode "MainForm Show"
    SetupVisualStyles Me

    With Me
        ' изменяем иконки формы и приложения
        ' Icon for Exe-file
        SetIcon .hWnd, "APPICO", True
        SetIcon .hWnd, "FRMMAIN", False
        ' Смена заголовка формы
        strFormName = .Name
        ChangeFrmMainCaption
        ' Разворачиваем форму на весь экран
        .Width = lngMainFormWidth
        .Height = lngMainFormHeight
        ' Центрируем форму на экране
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    LoadIconImage
    ' Подчеркавание меню (аля 3D)
    Me.Line (0, 15)-(ScaleWidth, 15), vbWhite
    Me.Line (0, 0)-(ScaleWidth, 0), GetSysColor(COLOR_BTNSHADOW)

    lngBorderWidthY = VPadding(Me)
    lngBorderWidthX = HPadding(Me)

    ' Создаем StatusBar
    ctlUcStatusBar1.AddPanel strProductName
    ctlUcStatusBar1.PanelAutoSize(1) = False
    PrintFileInDebugLog strSysIni
    
    ' Загрузка меню языков
    mnuMainLang.Visible = mbMultiLanguage

    ' Загрузка меню языков и локализация приложения
    If mbMultiLanguage Then
        If mbDebugStandart Then DebugMode "CreateLangList: " & UBound(arrLanguage)

        ' Создаем меню поддержки языков
        CreateMenuLng
        
        ' Локализация приложения
        Localise strPCLangCurrentPath
        
        ' Устанавливаем галочку на активном языке
        For iii = mnuLang.LBound To mnuLang.UBound
            mnuLang(iii).Checked = arrLanguage(0, iii) = strPCLangCurrentPath
        Next
        
        ' Устанавливаем галочку на автовыборе языка
        mnuLangStart.Checked = Not mbAutoLanguage
    End If
    
    ' Выставляем внешний вид кнопки
    SetBtnStyle cmdCheckAll
    SetBtnStyle cmdUnCheckAll
    SetBtnStyle cmdStartBackUp
    SetBtnStyle cmdBreak
    
    ' Выставляем шрифт
    FontCharsetChange

    ChangeStatusBarText strMessages(3), , True

    'заполнение списка типами создания резервных копий
    LoadComboList
    ' Загружаем список драйверов из реестра - прогресс на отдельной форме
    frmProgress.Show vbModal, Me
    ' Параметры выделения при старте
    chkCheckAll.Value = Abs(mbCheckAllGroup)
    chkHideOther.Value = Abs(mbListOnlyGroup)
    ' Режим при старте (Построение ListView из данных полученных выше)
    SelectStartMode
    ' Имя архива при старте
    SelectStartArchName
    ' Подсчет кол-ва выделенных
    FindCheckCountList

'    If lngFrameTime < 0 Then lngFrameTime = 1
'    If lngFrameCount < 1 Then lngFrameCount = 20
'    If Me.WindowState <> vbMinimized Then
'        AnimateForm Me, aLoad, eFoldOut, lngFrameTime, lngFrameCount
'    End If

    If mbDebugStandart Then DebugMode "FrmMainLoad-Finish" & vbNewLine & _
              "======================================================================="
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [Корректная выгрузка формы]
'! Parameters  (Переменные):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' Проверяем закончена ли проверка обновления, если нет то прерываем выход из программы, иначе программа вылетит
    If mbCheckUpdNotEnd Then
        Cancel = UnloadMode = vbFormControlMenu Or vbFormCode
        Exit Sub
    End If
    
    ' Удаление временных файлов если есть и если опция включена
    If mbDelTmpAfterClose Then
        ChangeStatusBarText strMessages(81), strMessages(130)

        'Чистим если только не перезапуск программы
        If Not mbRestartProgram Then
            DelTemp
        End If
    End If
    
    ' сохранение параметров при выходе
    If mbSaveSizeOnExit Then
        FRMStateSave
    End If

    ' Сохраняем язык при старте
    If Not mbIsDriveCDRoom Then
        If mnuLangStart.Checked Then
            IniWriteStrPrivate "Main", "StartLanguageID", strPCLangCurrentID, strSysIni
        End If

        IniWriteStrPrivate "Main", "AutoLanguage", CStr(Abs(Not mnuLangStart.Checked)), strSysIni
    End If

    SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP", False

    If mbLoadIniTmpAfterRestart Then
        SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP_PATH", "-"

        If StrComp(GetFileNameFromPath(strSysIni), "Settings_" & strProjectName & "_TMP.ini", vbTextCompare) = 0 Then
            DeleteFiles strSysIni
        End If
    End If

'    If lngFrameTime < 0 Then lngFrameTime = 2
'    If lngFrameCount < 1 Then lngFrameCount = 40
'    If Me.WindowState <> vbMinimized Then
'        AnimateForm Me, aUnload, eZoomOut, lngFrameTime, lngFrameCount
'    End If

    ' Выгружаем из памяти форму и другие компоненты
    Set frmMain = Nothing

    ' Выгружаем из памяти формы
    UnloadAllForms strFormName
    
    Unload Me
    Set frmMain = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Resize
'! Description (Описание)  :   [Изменение размеров контролов при изменении размеров формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    With Me
        If .WindowState <> vbMinimized Then

            ' Применение минимального размера по ширине
            If .Width < lngMainFormWidthMin Then
                .Width = lngMainFormWidthMin
                .Enabled = False
                .Enabled = True
    
                Exit Sub
    
            End If
    
            ' Применение минимального размера по высоте
            If .Height < lngMainFormHeightMin Then
                .Height = lngMainFormHeightMin
                .Enabled = False
                .Enabled = True
    
                Exit Sub
    
            End If

            If Not (frPanel Is Nothing) Then
                frPanel.Top = 0
                frPanel.Left = -20
                frPanel.Height = .Height - ctlUcStatusBar1.Height - lngBorderWidthY
                frPanel.Width = .Width
            End If

            If Not (ctlUcStatusBar1 Is Nothing) Then
                If ctlUcStatusBar1.PanelCount = 1 Then
                    ctlUcStatusBar1.PanelWidth(1) = (.Width \ Screen.TwipsPerPixelX)
                    ctlUcStatusBar1.Refresh
                End If
            End If

            ' Изменение размеров lvDevices
            ListViewResize
            
            ' Удаление иконки в трее если есть
            SetTrayIcon NIM_DELETE, Me.hWnd, 0&, vbNullString
        Else
            ' Добавляеи иконку в трей
            SetTrayIcon NIM_ADD, Me.hWnd, Me.Icon, strProjectNameFull
        End If
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub FRMStateSave
'! Description (Описание)  :   [Запись положения форм в ini-шку]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub FRMStateSave()

    Dim miHeight      As Long
    Dim miWidth       As Long
    Dim miWindowState As Long

    ' Если настройка активна, то выполняем сохранение
    miHeight = CLng(Me.Height)
    miWidth = CLng(Me.Width)

    If Me.WindowState = vbMaximized Then
        miWindowState = 1
    Else
        miWindowState = 0
    End If

    IniWriteStrPrivate "MainForm", "Height", CStr(miHeight), strSysIni
    IniWriteStrPrivate "MainForm", "Width", CStr(miWidth), strSysIni
    IniWriteStrPrivate "MainForm", "StartMaximazed", CStr(miWindowState), strSysIni
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ListViewResize
'! Description (Описание)  :   [Изменение размера панели с ListView и самого ListView]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ListViewResize()

    Dim lngLVPanelHeight    As Long
    Dim lngLVPanelWidht     As Long
    Dim lngLVPanelWidhtTemp As Long
    Dim lngLVPanelTop       As Long
    Dim lngLVPanelLeft      As Long
    Dim lngLVHeight         As Long
    Dim lngLVWidht          As Long
    Dim lngLVTop            As Long

    With Me
        ' Расчет размеров frPanelLV
        lngLVPanelTop = frGroup.Top + frGroup.Height + 80
        lngLVPanelLeft = frGroup.Left
        lngLVPanelHeight = frPanel.Height - lngLVPanelTop - 120
        lngLVPanelWidhtTemp = frBackUp.Left + frBackUp.Width - frGroup.Left
        lngLVPanelWidht = .Width - lngBorderWidthX - lngLVPanelLeft * 2

        If lngLVPanelWidht < lngLVPanelWidhtTemp Then
            lngLVPanelWidht = lngLVPanelWidhtTemp
        End If
    End With
    
    If Not (frPanelLV Is Nothing) Then
        With frPanelLV
            .Top = lngLVPanelTop
            .Left = lngLVPanelLeft
            .Height = lngLVPanelHeight
            .Width = lngLVPanelWidht
            ' Расчет размеров lvDevices
            lngLVTop = .TextBoxHeight * Screen.TwipsPerPixelY + 45
            lngLVHeight = .Height - lngLVTop - 60
            lngLVWidht = .Width - 120
            lblWait.Left = 100
            lblWait.Width = .Width - 200
        End With
    End If
    
    ' Изменение размеров lvDevices
    If Not (lvDevices Is Nothing) Then
        lvDevices.Move 60, lngLVTop, lngLVWidht, lngLVHeight
    End If

        
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadComboList
'! Description (Описание)  :   [заполнение списка типами создания резервных копий]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadComboList()

    ' Режимы выделения
    cmbListTypeBackupElement1 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbListTypeBackupElement1", "Структурированная папка с драйверами")
    cmbListTypeBackupElement2 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbListTypeBackupElement2", "7z-архив с драйверами")
    cmbListTypeBackupElement3 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbListTypeBackupElement3", "7zSFX-архив c автоустановкой через DPInst")

    With cmbTypeBackUp
        .Clear
        .AddItem cmbListTypeBackupElement1, 0
        .AddItem cmbListTypeBackupElement2, 1
        .AddItem cmbListTypeBackupElement3, 2

        ' Режим архивирования при запуске
        If miArchMode < 0 Or miArchMode > .ListCount - 1 Then
            .ListIndex = 0
        Else
            .ListIndex = miArchMode
        End If
    End With
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadIconImage
'! Description (Описание)  :   []
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadIconImage()

    If mbDebugDetail Then DebugMode "LoadIconImage-Start"
    '--------------------- Иконки кнопок
    LoadIconImage2Object cmdStartBackUp, "BTN_STARTBACKUP", strPathImageMainWork
    LoadIconImage2Object cmdBreak, "BTN_BREAK", strPathImageMainWork
    LoadIconImage2Object cmdCheckAll, "BTN_CHECKMARK", strPathImageMainWork
    LoadIconImage2Object cmdUnCheckAll, "BTN_UNCHECKMARK", strPathImageMainWork
    '--------------------- Остальные Иконки
    LoadIconImage2Object frBackUp, "FRAME_BACKUP", strPathImageMainWork
    LoadIconImage2Object frGroup, "FRAME_GROUP", strPathImageMainWork
    LoadIconImage2Object frPanelLV, "FRAME_LIS", strPathImageMainWork
    If mbDebugDetail Then DebugMode "LoadIconImage-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadListbyMode
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadListbyMode()

    Dim lngModeList As Long
    Dim mbOpt1      As Boolean
    Dim mbOpt2      As Boolean
    Dim mbOpt3      As Boolean
    Dim mbOpt4      As Boolean

    mbOpt1 = optGrp1.Value
    mbOpt2 = optGrp2.Value
    mbOpt3 = optGrp3.Value
    mbOpt4 = optGrp4.Value

    ' Microsoft
    If mbOpt1 Then
        lngModeList = 1

    ' OEM
    ElseIf mbOpt2 Then
        lngModeList = 2

    ' Все
    ElseIf mbOpt3 Then
        lngModeList = 3
        
    ' Ничего
    ElseIf mbOpt4 Then
        lngModeList = 9999
    End If

    If lngModeList <> 9999 Then
        LoadList_Device lngModeList
    Else
        If Not (lvDevices Is Nothing) Then
            lvDevices.ListItems.Clear
        End If
        With lvDevices.ColumnHeaders
            If .count Then
                .item(1).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(2).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(3).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(4).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(5).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(6).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(7).AutoSize LvwColumnHeaderAutoSizeToHeader
                .item(8).AutoSize LvwColumnHeaderAutoSizeToHeader
            End If
        End With
    End If

    'LoadFormCaption
    FindCheckCountList
End Sub

'strTableHwidHeader1 = "*Наименование устройства*")
'strTableHwidHeader2 = "*Дата драйвера*")
'strTableHwidHeader3 = "*Версия драйвера*")
'strTableHwidHeader4 = "*Производитель*")
'strTableHwidHeader5 = "*Класс драйвера*")
'strTableHwidHeader6 = "*Код класса*")
'strTableHwidHeader7 = "*Inf-файл*")
'strTableHwidHeader8 = "*Секция Inf-файла*")
'strTableHwidHeader9 = "*HWID*")
'strTableHwidHeader10 ="-ID Класса-")
'strTableHwidHeader11 ="-ID Экземпляра устройства-")

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_Device
'! Description (Описание)  :   [Построение полного спиcка устройств]
'! Parameters  (Переменные):   lngMode (Long = 0)
'!--------------------------------------------------------------------------------
Private Sub LoadList_Device(Optional ByVal lngMode As Long = 0)

    Dim strDevHwid        As String
    Dim strDevDriverLocal As String
    Dim strDevStatus      As String
    Dim strDevName        As String
    Dim strProvider       As String
    Dim strCompatID       As String
    Dim strStrDescription As String
    Dim strOrigHwid       As String
    Dim ii                As Long
    Dim strInDPacks       As String
    Dim lngNumRow         As Long

    If mbDebugDetail Then DebugMode "LoadList_Device-Start"
    If mbDebugStandart Then DebugMode "***LoadList_Device: Mode=" & lngMode
    
    With lvDevices
        .Redraw = False
        .ListItems.Clear
        .ColumnHeaders.Clear

        If .ColumnHeaders.count = 0 Then
            .ColumnHeaders.Add 1, , strTableHwidHeader1
            .ColumnHeaders.Add 2, , strTableHwidHeader2
            .ColumnHeaders.Add 3, , strTableHwidHeader3
            .ColumnHeaders.Add 4, , strTableHwidHeader4
            .ColumnHeaders.Add 5, , strTableHwidHeader5
            .ColumnHeaders.Add 6, , strTableHwidHeader6
            .ColumnHeaders.Add 7, , strTableHwidHeader7
            .ColumnHeaders.Add 8, , strTableHwidHeader8
            .ColumnHeaders.Add 9, , strTableHwidHeader9
            .ColumnHeaders.Add 10, , strTableHwidHeader10
        End If

        For ii = 0 To UBound(arrHwidsLocal)
    
            strProvider = arrHwidsLocal(ii).i3_ProviderName
            
            Select Case lngMode
    
                ' All - ALL
                Case 0, 3
    
                    With .ListItems.Add(, , arrHwidsLocal(ii).i0_DriverDesc)
                        .SubItems(1) = arrHwidsLocal(ii).i1_DriverDate
                        .SubItems(2) = arrHwidsLocal(ii).i2_DriverVersion
                        .SubItems(3) = strProvider
                        .SubItems(4) = arrHwidsLocal(ii).i4_ClassName
                        .SubItems(5) = arrHwidsLocal(ii).i5_Class
                        .SubItems(6) = arrHwidsLocal(ii).i6_InfPath
                        .SubItems(7) = arrHwidsLocal(ii).i7_InfSection
                        .SubItems(8) = arrHwidsLocal(ii).i8_MatchingDeviceId
                        .SubItems(9) = arrHwidsLocal(ii).i9_ClassID
                        If Not .Checked Then
                            If chkCheckAll.Value Or chkCheckAll.Enabled = False Then
                                .Checked = True
                            End If
                        End If
                    '.ListItems.Add
                    End With
                    
                ' Microsoft - All
                Case 1
                    If InStr(1, strProvider, "microsoft", vbTextCompare) Or InStr(1, strProvider, "майкрософт", vbTextCompare) Or InStr(1, strProvider, "standard", vbTextCompare) Then
    
                        With .ListItems.Add(, , arrHwidsLocal(ii).i0_DriverDesc)
                            .SubItems(1) = arrHwidsLocal(ii).i1_DriverDate
                            .SubItems(2) = arrHwidsLocal(ii).i2_DriverVersion
                            .SubItems(3) = strProvider
                            .SubItems(4) = arrHwidsLocal(ii).i4_ClassName
                            .SubItems(5) = arrHwidsLocal(ii).i5_Class
                            .SubItems(6) = arrHwidsLocal(ii).i6_InfPath
                            .SubItems(7) = arrHwidsLocal(ii).i7_InfSection
                            .SubItems(8) = arrHwidsLocal(ii).i8_MatchingDeviceId
                            .SubItems(9) = arrHwidsLocal(ii).i9_ClassID
                            If Not .Checked Then
                                If chkCheckAll.Value Or chkCheckAll.Enabled = False Then
                                    .Checked = True
                                End If
                            End If
                        '.ListItems.Add
                        End With
    
                        lngNumRow = lngNumRow + 1
                    Else
                        If chkHideOther.Value = 0 Then
                            With .ListItems.Add(, , arrHwidsLocal(ii).i0_DriverDesc)
                                .SubItems(1) = arrHwidsLocal(ii).i1_DriverDate
                                .SubItems(2) = arrHwidsLocal(ii).i2_DriverVersion
                                .SubItems(3) = strProvider
                                .SubItems(4) = arrHwidsLocal(ii).i4_ClassName
                                .SubItems(5) = arrHwidsLocal(ii).i5_Class
                                .SubItems(6) = arrHwidsLocal(ii).i6_InfPath
                                .SubItems(7) = arrHwidsLocal(ii).i7_InfSection
                                .SubItems(8) = arrHwidsLocal(ii).i8_MatchingDeviceId
                                .SubItems(9) = arrHwidsLocal(ii).i9_ClassID
                                .Checked = False
                            '.ListItems.Add
                            End With
                            
                            lngNumRow = lngNumRow + 1
                        End If
                    
                    End If
    
                ' OEM - All
                Case 2
    
                    If InStr(1, strProvider, "microsoft", vbTextCompare) = 0 Then
                        If InStr(1, strProvider, "майкрософт", vbTextCompare) = 0 Then
                            If InStr(1, strProvider, "standard", vbTextCompare) = 0 Then
    
                                With .ListItems.Add(, , arrHwidsLocal(ii).i0_DriverDesc)
                                    .SubItems(1) = arrHwidsLocal(ii).i1_DriverDate
                                    .SubItems(2) = arrHwidsLocal(ii).i2_DriverVersion
                                    .SubItems(3) = strProvider
                                    .SubItems(4) = arrHwidsLocal(ii).i4_ClassName
                                    .SubItems(5) = arrHwidsLocal(ii).i5_Class
                                    .SubItems(6) = arrHwidsLocal(ii).i6_InfPath
                                    .SubItems(7) = arrHwidsLocal(ii).i7_InfSection
                                    .SubItems(8) = arrHwidsLocal(ii).i8_MatchingDeviceId
                                    .SubItems(9) = arrHwidsLocal(ii).i9_ClassID
                                    If Not .Checked Then
                                        If chkCheckAll.Value Or chkCheckAll.Enabled = False Then
                                            .Checked = True
                                        End If
                                    End If
                                '.ListItems.Add
                                End With
                                
                                lngNumRow = lngNumRow + 1
                            End If
                        End If
                    Else
                        If chkHideOther.Value = 0 Then
                            With .ListItems.Add(, , arrHwidsLocal(ii).i0_DriverDesc)
                                .SubItems(1) = arrHwidsLocal(ii).i1_DriverDate
                                .SubItems(2) = arrHwidsLocal(ii).i2_DriverVersion
                                .SubItems(3) = strProvider
                                .SubItems(4) = arrHwidsLocal(ii).i4_ClassName
                                .SubItems(5) = arrHwidsLocal(ii).i5_Class
                                .SubItems(6) = arrHwidsLocal(ii).i6_InfPath
                                .SubItems(7) = arrHwidsLocal(ii).i7_InfSection
                                .SubItems(8) = arrHwidsLocal(ii).i8_MatchingDeviceId
                                .SubItems(9) = arrHwidsLocal(ii).i9_ClassID
                                .Checked = False
                            '.ListItems.Add
                            End With
                            lngNumRow = lngNumRow + 1
                        End If
                    End If
                    
            End Select
    
        Next
    
        With .ColumnHeaders
            If .count Then
                If lvDevices.ListItems.count Then
                    .item(1).AutoSize LvwColumnHeaderAutoSizeToItems
                    .item(2).AutoSize LvwColumnHeaderAutoSizeToItems
                    If .item(2).Width < lvDevices.ListItems.item(1).Width Then
                        .item(2).AutoSize LvwColumnHeaderAutoSizeToHeader
                    End If
                    .item(3).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(4).AutoSize LvwColumnHeaderAutoSizeToItems
                    .item(5).AutoSize LvwColumnHeaderAutoSizeToItems
                    .item(6).AutoSize LvwColumnHeaderAutoSizeToItems
                    .item(7).AutoSize LvwColumnHeaderAutoSizeToItems
                    .item(8).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(9).AutoSize LvwColumnHeaderAutoSizeToHeader
                Else
                    .item(1).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(2).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(3).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(4).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(5).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(6).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(7).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(8).AutoSize LvwColumnHeaderAutoSizeToHeader
                    .item(9).AutoSize LvwColumnHeaderAutoSizeToHeader
                End If
            End If
            
        '.ColumnHeaders
        End With
        
        .Redraw = True
        .Sorted = True
    'lvDevices
    End With

    If mbDebugStandart Then DebugMode "LoadList_Device-Finish"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    ChangeFrmMainCaption
    ' frGroup
    frGroup.Caption = LocaliseString(strPathFile, strFormName, "frGroup", frGroup.Caption)
    optGrp1.Caption = LocaliseString(strPathFile, strFormName, "optGrp1", optGrp1.Caption)
    optGrp2.Caption = LocaliseString(strPathFile, strFormName, "optGrp2", optGrp2.Caption)
    optGrp3.Caption = LocaliseString(strPathFile, strFormName, "optGrp3", optGrp3.Caption)
    optGrp4.Caption = LocaliseString(strPathFile, strFormName, "optGrp4", optGrp4.Caption)
    lblWait.Caption = LocaliseString(strPathFile, strFormName, "lblWait", lblWait.Caption)
    chkHideOther.Caption = LocaliseString(strPathFile, strFormName, "chkHideOther", chkHideOther.Caption)
    cmdCheckAll.Caption = LocaliseString(strPathFile, strFormName, "cmdCheckAll", cmdCheckAll.Caption)
    cmdUnCheckAll.Caption = LocaliseString(strPathFile, strFormName, "cmdUnCheckAll", cmdUnCheckAll.Caption)
    chkCheckAll.Caption = LocaliseString(strPathFile, strFormName, "chkCheckAll", chkCheckAll.Caption)
    cmdBreak.Caption = LocaliseString(strPathFile, strFormName, "cmdBreak", cmdBreak.Caption)
    ' frBackUp
    frBackUp.Caption = LocaliseString(strPathFile, strFormName, "frBackUp", frBackUp.Caption)
    cmdStartBackUp.Caption = LocaliseString(strPathFile, strFormName, "cmdStartBackUp", cmdStartBackUp.Caption)
    lblTypeBackUp.Caption = LocaliseString(strPathFile, strFormName, "lblTypeBackUp", lblTypeBackUp.Caption)
    ' frPanelLV
    frPanelLV.Caption = LocaliseString(strPathFile, strFormName, "frPanelLV", frPanelLV.Caption)
    strTableHwidHeader1 = LocaliseString(strPathFile, strFormName, "TableHeader1", "*Наименование устройства*")
    strTableHwidHeader2 = LocaliseString(strPathFile, strFormName, "TableHeader2", "*Дата драйвера*")
    strTableHwidHeader3 = LocaliseString(strPathFile, strFormName, "TableHeader3", "*Версия драйвера*")
    strTableHwidHeader4 = LocaliseString(strPathFile, strFormName, "TableHeader4", "*Производитель*")
    strTableHwidHeader5 = LocaliseString(strPathFile, strFormName, "TableHeader5", "*Класс драйвера*")
    strTableHwidHeader6 = LocaliseString(strPathFile, strFormName, "TableHeader6", "*Код класса*")
    strTableHwidHeader7 = LocaliseString(strPathFile, strFormName, "TableHeader7", "*Inf-файл*")
    strTableHwidHeader8 = LocaliseString(strPathFile, strFormName, "TableHeader8", "*Секция Inf-файла*")
    strTableHwidHeader9 = LocaliseString(strPathFile, strFormName, "TableHeader9", "*HWID*")
    strTableHwidHeader10 = LocaliseString(strPathFile, strFormName, "TableHeader10", "-ID Класса-")
    strTableHwidHeader11 = LocaliseString(strPathFile, strFormName, "TableHeader11", "-ID Экземпляра устройства-")
    ' frArchName
    frArchName.Caption = LocaliseString(strPathFile, strFormName, "frArchName", frArchName.Caption)
    optArchNamePC.Caption = LocaliseString(strPathFile, strFormName, "optArchNamePC", optArchNamePC.Caption)
    optArchModelPC.Caption = LocaliseString(strPathFile, strFormName, "optArchModelPC", optArchModelPC.Caption)
    optArchCustom.Caption = LocaliseString(strPathFile, strFormName, "optArchCustom", optArchCustom.Caption)
    ' Меню - Вызов основной функции для вывода Caption меню с поддержкой Unicode
    Call LocaliseMenu(strPathFile)
    ' Типы архивов
    LoadComboList
    'загружаем программные сообщения
    LocaliseMessage strPCLangCurrentPath
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LocaliseMenu
'! Description (Описание)  :   [Загрузка текста меню с поддеркой Unicode]
'! Parameters  (Переменные):   strPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub LocaliseMenu(ByVal strPathFile As String)
    
'0  mnuReCollectHWID - "Обновить информацию"
    SetUniMenu -1, 0, -1, mnuReCollectHWID, LocaliseString(strPathFile, strFormName, "mnuReCollectHWID", mnuReCollectHWID.Caption)

'1  mnuOptions - "Параметры"
    SetUniMenu -1, 1, -1, mnuOptions, LocaliseString(strPathFile, strFormName, "mnuOptions", mnuOptions.Caption), , "Ctrl+O"
       
'2  mnuMainAbout - "Справка"
' 0    mnuLinks - "Ссылки"
' 1    mnuHistory - "История изменения"
' 2    mnuHelp - "Справка по работе" - Shortcut{F1}
' 3    mnuSep11 - "-"
' 4    mnuHomePage1 - "Домашная страница программы"
' 5    mnuHomePage - "Обсуждение программы на OsZone.net"
' 6    mnuOsZoneNet
' 7    mnuSep12 - "-"
' 8    mnuCheckUpd - "Проверить обновление программы"
' 9   mnuSep13 - "-"
' 10   mnuModulesVersion - "Модули..."
' 11   mnuSep14 - "-"
' 12   mnuDonate - "Поблагодарить автора..."
' 13   mnuLicence - "Лицензионное соглашение..."
' 14   mnuAbout - "О программе..."
    SetUniMenu -1, 2, -1, mnuMainAbout, LocaliseString(strPathFile, strFormName, "mnuMainAbout", mnuMainAbout.Caption)
    SetUniMenu 2, 0, -1, mnuLinks, LocaliseString(strPathFile, strFormName, "mnuLinks", mnuLinks.Caption)
    SetUniMenu 2, 1, -1, mnuHistory, LocaliseString(strPathFile, strFormName, "mnuHistory", mnuHistory.Caption)
    SetUniMenu 2, 2, -1, mnuHelp, LocaliseString(strPathFile, strFormName, "mnuHelp", mnuHelp.Caption), , "F1"
    SetUniMenu 2, 4, -1, mnuHomePage, LocaliseString(strPathFile, strFormName, "mnuHomePage", mnuHomePage.Caption)
    SetUniMenu 2, 5, -1, mnuHomePageForum, LocaliseString(strPathFile, strFormName, "mnuHomePageForum", mnuHomePageForum.Caption)
    SetUniMenu 2, 6, -1, mnuOsZoneNet, LocaliseString(strPathFile, strFormName, "mnuOsZoneNet", mnuOsZoneNet.Caption)
    SetUniMenu 2, 8, -1, mnuCheckUpd, LocaliseString(strPathFile, strFormName, "mnuCheckUpd", mnuCheckUpd.Caption)
    SetUniMenu 2, 10, -1, mnuModulesVersion, LocaliseString(strPathFile, strFormName, "mnuModulesVersion", mnuModulesVersion.Caption)
    SetUniMenu 2, 12, -1, mnuDonate, LocaliseString(strPathFile, strFormName, "mnuDonate", mnuDonate.Caption)
    SetUniMenu 2, 13, -1, mnuLicence, LocaliseString(strPathFile, strFormName, "mnuLicence", mnuLicence.Caption)
    SetUniMenu 2, 14, -1, mnuAbout, LocaliseString(strPathFile, strFormName, "mnuAbout", mnuAbout.Caption)
    
'3  mnuMainLang - "Язык"
' 0    mnuLangStart - "Использовать выбранный язык при запуске (отмена автовыбора)"
' 1    mnuSep15 - "-"
' 2    mnuLang - "" - Index0 - Visible'False
    SetUniMenu -1, 3, -1, mnuMainLang, LocaliseString(strPathFile, strFormName, "mnuMainLang", mnuMainLang.Caption)
    SetUniMenu 3, 0, -1, mnuLangStart, LocaliseString(strPathFile, strFormName, "mnuLangStart", mnuLangStart.Caption)
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvDevices_ColumnClick
'! Description (Описание)  :   [Сортировка списка при щелчке по колонке]
'! Parameters  (Переменные):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub lvDevices_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)

    Dim ii As Long

    With lvDevices
        .Sorted = False
        .SortKey = ColumnHeader.Index - 1

        If ComCtlsSupportLevel() >= 1 Then

            For ii = 1 To .ColumnHeaders.count

                If ii <> ColumnHeader.Index Then
                    .ColumnHeaders(ii).SortArrow = LvwColumnHeaderSortArrowNone
                Else

                    If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowNone Then
                        ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
                    Else

                        If ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown Then
                            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp
                        ElseIf ColumnHeader.SortArrow = LvwColumnHeaderSortArrowUp Then
                            ColumnHeader.SortArrow = LvwColumnHeaderSortArrowDown
                        End If
                    End If
                End If

            Next ii

            Select Case ColumnHeader.SortArrow

                Case LvwColumnHeaderSortArrowDown, LvwColumnHeaderSortArrowNone
                    .SortOrder = LvwSortOrderAscending

                Case LvwColumnHeaderSortArrowUp
                    .SortOrder = LvwSortOrderDescending
            End Select

            .SelectedColumn = ColumnHeader
        Else

            For ii = 1 To .ColumnHeaders.count

                If ii <> ColumnHeader.Index Then
                    .ColumnHeaders(ii).Icon = 0
                Else

                    If ColumnHeader.Icon = 0 Then
                        ColumnHeader.Icon = 1
                    Else

                        If ColumnHeader.Icon = 2 Then
                            ColumnHeader.Icon = 1
                        ElseIf ColumnHeader.Icon = 1 Then
                            ColumnHeader.Icon = 2
                        End If
                    End If
                End If

            Next ii

            Select Case ColumnHeader.Icon

                Case 1, 0
                    .SortOrder = LvwSortOrderAscending

                Case 2
                    .SortOrder = LvwSortOrderDescending
            End Select

        End If

        .Sorted = True

        If Not .SelectedItem Is Nothing Then .SelectedItem.EnsureVisible
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvDevices_ItemCheck
'! Description (Описание)  :   [Выбор элемента списка]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Checked (Boolean)
'!--------------------------------------------------------------------------------
Private Sub lvDevices_ItemCheck(ByVal item As LvwListItem, ByVal Checked As Boolean)
    If Not mbFirstStart Then
        If Not mbMassCheckInLV Then
            FindCheckCountList
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvDevices_ItemDblClick
'! Description (Описание)  :   [Двойной клик по записи - вызов свойства устройства]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvDevices_ItemDblClick(ByVal item As LvwListItem, ByVal Button As Integer)

    Dim strOrigHwid As String

    If Button = vbLeftButton Then
        strOrigHwid = item.SubItems(8)
        OpenDeviceProp strOrigHwid
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuAbout_Click
'! Description (Описание)  :   [Меню - О программе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuCheckUpd_Click
'! Description (Описание)  :   [Меню - Проверить обновление]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuCheckUpd_Click()
    CheckUpd False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuDonate_Click
'! Description (Описание)  :   [Меню - Помощь проекту]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuDonate_Click()
    frmDonate.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuHistory_Click
'! Description (Описание)  :   [Меню - История изменений]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuHistory_Click()

    Dim strFilePathTemp As String

    strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\" & strPCLangCurrentID & "\history.txt"

    If FileExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\0409\history.txt"
    End If

    RunUtilsShell strFilePathTemp, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuHomePageForum_Click
'! Description (Описание)  :   [Меню - Домашняя страница форум]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuHomePageForum_Click()
    RunUtilsShell strUrl_MainWWWForum, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuHomePage_Click
'! Description (Описание)  :   [Меню - Домашняя страница]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuHomePage_Click()
    RunUtilsShell strUrl_MainWWWSite, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLangStart_Click
'! Description (Описание)  :   [Меню - Язык]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuLangStart_Click()
    mnuLangStart.Checked = Not mnuLangStart.Checked
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLang_Click
'! Description (Описание)  :   [Меню - выбор языка]
'! Parameters  (Переменные):   Index (Integer)
'!--------------------------------------------------------------------------------
Private Sub mnuLang_Click(Index As Integer)

    Dim ii                     As Long
    Dim strPathLng             As String
    Dim strPCLangCurrentIDTemp As String
    Dim strPCLangCurrentID_x() As String

    For ii = mnuLang.LBound To mnuLang.UBound
        mnuLang(ii).Checked = ii = Index
    Next

    strPathLng = arrLanguage(0, Index)
    strPCLangCurrentPath = strPathLng
    strPCLangCurrentIDTemp = arrLanguage(2, Index)
    strPCLangCurrentLangName = arrLanguage(1, Index)
    lngFont_Charset = GetCharsetFromLng(CLng(arrLanguage(5, Index)))

    If InStr(strPCLangCurrentIDTemp, strSemiColon) Then
        strPCLangCurrentID_x = Split(strPCLangCurrentIDTemp, strSemiColon)
        strPCLangCurrentID = strPCLangCurrentID_x(0)
    Else
        strPCLangCurrentID = strPCLangCurrentIDTemp
    End If
    
    ' Собственно локализация
    Localise strPCLangCurrentPath

    ' ПереВыставляем шрифт основной формы
    With Me.Font
        .Name = strFontMainForm_Name
        .Size = lngFontMainForm_Size
        .Charset = lngFont_Charset
    End With
    
    ChangeFrmMainCaption
    mnuReCollectHWID_Click
    
    ChangeStatusBarText strMessages(142) & strSpace & arrLanguage(1, Index), , False

End Sub

'! Procedure   (Функция)   :   Sub mnuLicence_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuLicence_Click()
    frmLicence.Show vbModal, Me
End Sub


'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuLinks_Click
'! Description (Описание)  :   [Меню - Ссылки]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuLinks_Click()

    Dim strFilePathTemp As String

    strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\" & strPCLangCurrentID & "\Links.html"

    If FileExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPathBackSL & strToolsDocs_Path & "\0409\Links.html"
    End If

    RunUtilsShell strFilePathTemp, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuModulesVersion_Click
'! Description (Описание)  :   [Меню - Версии модулей]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuModulesVersion_Click()
    VerModules
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuOptions_Click
'! Description (Описание)  :   [Меню - Настройки]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuOptions_Click()
    ctlUcStatusBar1.PanelText(1) = strMessages(146)
    ChangeStatusBarText strMessages(146)

    If IsFormLoaded("frmOptions") = False Then
        frmOptions.Show vbModal, Me
    Else
        frmOptions.FormLoadAction
        frmOptions.Show vbModal, Me
    End If

    If mbRestartProgram Then
        ShellExecute Me.hWnd, "open", strAppEXEName, vbNullString, strAppPath, SW_SHOWNORMAL
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuOsZoneNet_Click
'! Description (Описание)  :   [Меню - Форум]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuOsZoneNet_Click()
    RunUtilsShell strUrlOsZoneNetThread, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub mnuReCollectHWID_Click
'! Description (Описание)  :   [Меню - Обновить оборудование]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub mnuReCollectHWID_Click()

    ReCollectHWID
    ' Режим при старте
    SelectStartMode
    ' Показ списка
    lblWait.Visible = False
    lvDevices.Visible = True
    
    FindCheckCountList
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub OpenDeviceProp
'! Description (Описание)  :   [Открыть свойства устройства]
'! Parameters  (Переменные):   strHwid (String)
'!--------------------------------------------------------------------------------
Private Sub OpenDeviceProp(ByVal strHwid As String)

    Dim cmdString       As String
    Dim cmdStringParams As String
    Dim nRetShellEx     As Boolean

    cmdString = "rundll32.exe"
    cmdStringParams = "devmgr.dll,DeviceProperties_RunDLL /DeviceID " & strHwid
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    If mbDebugStandart Then DebugMode "cmdStringParams: " & cmdStringParams
    
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL, cmdStringParams)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optArchCustom_Click
'! Description (Описание)  :   [Выбор режима имени архива]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optArchCustom_Click()

    Dim strTempString As String

    With txtArchName
        .Locked = False
        .Enabled = True
        strTempString = SafeDir(ExpandArchNamebyEnvironment(strArchNameCustom))

        If LenB(SafeDir(strTempString)) Then
            .Text = strTempString
        Else
            .Text = CollectDPName(strCompName)
        End If
    End With
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optArchModelPC_Click
'! Description (Описание)  :   [Выбор режима имени архива]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optArchModelPC_Click()

    With txtArchName
        .Text = CollectDPName(strCompModel)
        .Locked = True
        .Enabled = False
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optArchNamePC_Click
'! Description (Описание)  :   [Выбор режима имени архива]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optArchNamePC_Click()

    With txtArchName
        .Text = CollectDPName(strCompName)
        .Locked = True
        .Enabled = False
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optGrp1_Click
'! Description (Описание)  :   [Выбор режима просмотра списка стройств - Microsoft]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optGrp1_Click()
    If optGrp1.Value Then
        ReNewLVlist
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optGrp2_Click
'! Description (Описание)  :   [Выбор режима просмотра списка стройств - Все]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optGrp2_Click()
    If optGrp2.Value Then
        ReNewLVlist
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optGrp3_Click
'! Description (Описание)  :   [Выбор режима просмотра списка стройств - OEM]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optGrp3_Click()
    If optGrp3.Value Then
        ReNewLVlist
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optGrp4_Click
'! Description (Описание)  :   [Выбор режима просмотра списка стройств - Ничего]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub optGrp4_Click()
Attribute optGrp4_Click.VB_UserMemId = 1610809397
    If optGrp4.Value Then
        ReNewLVlist
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReCollectHWID
'! Description (Описание)  :   [Обновление списка оборудования - читаем реестр заново]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ReCollectHWID()

    BlockControl True
    ChangeStatusBarText strMessages(3)
    ' Очищаем
    lvDevices.ListItems.Clear
    lvDevices.Visible = False
    lblWait.Visible = True
    DoEvents
    
    ' повторно собираем данные
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateInProgress
    frmProgress.Show vbModal, Me
    
    ' А теперь перестраиваем список драйверов
    LoadListbyMode
    ListViewResize
    BlockControl False
    
    ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    
    ChangeStatusBarText strMessages(5)
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReNewLVlist
'! Description (Описание)  :   [Перестроить лист оборудования с заданным фильтром]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ReNewLVlist()
Attribute ReNewLVlist.VB_UserMemId = 1610809398
    lvDevices.Visible = False
    lblWait.Visible = True
    DoEvents
    LoadListbyMode
    lvDevices.Visible = True
    lblWait.Visible = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SearchSectInSect
'! Description (Описание)  :   [Поиск секций в inf-файле]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function SearchSectInSect(ByRef arrZ() As String) As String()

    Dim strFileName      As String
    Dim strFileNameSect  As String
    Dim strFileName_x()  As String
    Dim strSectionList() As String
    Dim d                As Long
    Dim n                As Long
    Dim ii               As Long
    Dim miMaxCountArr    As Long

    miMaxCountArr = 100
    ' максимальное кол-во элементов в массиве
    ReDim strSectionList(miMaxCountArr) As String

    For d = 1 To UBound(arrZ, 1)
        strFileName = TrimNull(arrZ(d, 2))
        strFileNameSect = arrZ(d, 1)

        ' Отбрасываем все после ";"
        If InStr(1, strFileName, ";", vbTextCompare) Then
            strFileName = Trim$(Left$(strFileName, InStrRev(strFileName, ";") - 1))
        End If

        If StrComp(strFileNameSect, "CopyFiles", vbTextCompare) = 0 Then
            strFileName_x = Split(strFileName, ",")

            For ii = 0 To UBound(strFileName_x)

                ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
                If n = miMaxCountArr Then
                    miMaxCountArr = miMaxCountArr + miMaxCountArr
                    ReDim Preserve strSectionList(1, miMaxCountArr)
                End If

                strSectionList(n) = strFileName_x(ii)
                n = n + 1
            Next
        End If

    Next

    ' Переобъявляем массив на реальное кол-во записей
    If n > 0 Then
        ReDim Preserve strSectionList(n - 1)
    Else
        ReDim Preserve strSectionList(0)
    End If

    SearchSectInSect = strSectionList
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectStartArchName
'! Description (Описание)  :   [Режим имени архива при старте]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SelectStartArchName()

    Select Case lngArchNameMode

        Case 0
            If Not optArchCustom.Value Then
                optArchCustom.Value = True
            Else
                optArchCustom_Click
            End If
            
        Case 1
            If Not optArchNamePC.Value Then
                optArchNamePC.Value = True
            Else
                optArchNamePC_Click
            End If
        Case 2
            If Not optArchModelPC.Value Then
                optArchModelPC.Value = True
            Else
                optArchModelPC_Click
            End If
            
        Case Else
            If Not optArchCustom.Value Then
                optArchCustom.Value = True
            Else
                optArchCustom_Click
            End If
            
    End Select
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectStartArchName
'! Description (Описание)  :   [Режим фильтра при старте]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SelectStartMode()

    Select Case miStartMode

        Case 1
            If Not optGrp1.Value Then
                optGrp1.Value = True
            Else
                optGrp1_Click
            End If
        Case 2
            If Not optGrp2.Value Then
                optGrp2.Value = True
            Else
                optGrp2_Click
            End If
            
        Case 3
            If Not optGrp3.Value Then
                optGrp3.Value = True
            Else
                optGrp3_Click
            End If

        Case 4
            If Not optGrp4.Value Then
                optGrp4.Value = True
            Else
                optGrp4_Click
            End If
            
    End Select
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetBtnStyle
'! Description (Описание)  :   [Установка свойств стиля для кнопки]
'! Parameters  (Переменные):   ctlObject (Object)
'!--------------------------------------------------------------------------------
Private Sub SetBtnStyle(ctlObject As Object)
    
    With ctlObject
        .ButtonStyle = lngStatusBtnStyle
        .ColorScheme = lngStatusBtnStyleColor
        
        If lngStatusBtnStyleColor = 3 Then
            .BackColor = lngStatusBtnBackColor
        End If
        
        .ForeColor = lngFontBtn_Color
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub StartBackUp
'! Description (Описание)  :   [Процедура запуска бекапа, по нажатию кнопки]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub StartBackUp()

    Dim destDir               As String
    Dim destDirDialog         As String
    Dim nn                    As Long
    Dim ii                    As Long
    Dim DD                    As Long
    Dim strDest               As String
    Dim arr_Z()               As String
    Dim arr_Z2()              As String
    Dim arr_Z3()              As String
    Dim arr_Z4()              As String
    Dim arr_Z5()              As String
    Dim arr_ZF1()             As FindListStruct
    Dim inf()                 As String
    Dim strInfFileName        As String
    Dim strInfFile2Path       As String
    Dim lngArrCount           As Long
    Dim lvCount               As Long
    Dim lvCountCheck          As Long
    Dim lngTimeScriptRun      As Currency
    Dim lngTimeScriptFinish   As Currency
    Dim strAllTimeScriptRun   As String
    Dim miPBInterval          As Long
    Dim miPBNext              As Long
    Dim strDriverDesc         As String
    Dim strClass              As String
    Dim strInfSection         As String
    Dim numCat                As Long
    Dim mbDoZip               As Boolean
    Dim str7zFileArchivePath  As String
    Dim strStatusMsgTemp      As String
    Dim strSectionName        As String
    Dim strFileList           As String
    Dim strCatFileName4Inf    As String
    Dim strInfFile2Path4Cat   As String
    Dim strDataSHA1           As String
    Dim lngNumFilesFromFolder As String
    Dim strFolderPath         As String
    Dim strFileNameInf        As String
    Dim mbCompare             As Boolean
    Dim mbBackUPedFiles       As Boolean
    Dim lngUBoundZ2           As Long

    If mbDebugDetail Then DebugMode "cmdStartBackUp_Click-Start"
    lngTimeScriptRun = GetTimeStart

    lvCountCheck = FindCheckCountList(True)
    '# Если есть выделенные строки
    If lvCountCheck = 0 Then
        MsgBox strMessages(6), vbInformation + vbOKOnly, strProductName
    Else

        '# Диалог открытия файла
        With New CommonDialog
            If mbIsDriveCDRoom Then
                .InitDir = PathCollect(strAppPathBackSL & "drivers\")
            Else
                .InitDir = PathCollect(DefineFolderBackUp)
            End If
            
            If IsWinXPOrLater Then
                .flags = CdlBIFNewDialogStyle Or CdlBIFUAHint
            Else
                .flags = CdlBIFNewDialogStyle
            End If
            
            .DialogTitle = strMessages(2)
            
            If .ShowFolderBrowser Then
                destDirDialog = .FileName
            End If
    
        End With
    
        If LenB(destDirDialog) = 0 Then
            '# if user cancel #
            Exit Sub
        End If

        If mbDebugStandart Then DebugMode "StartBackUp: Destination=" & destDirDialog

        'Блокируем лист перед бекапом
        If mbBlockListOnBackup Then
            If mbDebugStandart Then DebugMode "BlockListOnBackup: TRUE"
            lvDevices.Enabled = False
        End If

        ' Блокируем элементы от греха подальше
        If mbDebugStandart Then DebugMode "BlockControl: TRUE"
        BlockControl True
        MousePointer = 11
        '# display hourglass cursor while read #
        DoEvents

        'формируем путь каталога назначения бекапа
        If LenB(Trim$(txtArchName)) Then
            destDir = BackslashAdd2Path(destDirDialog) & txtArchName
        Else
            destDir = BackslashAdd2Path(destDirDialog) & CollectDPName(strCompName)
        End If

        If mbDebugStandart Then DebugMode "***StartBackUp: Destination directory: " & destDir

        If PathExists(destDir) Then
            If mbDebugStandart Then DebugMode "***StartBackUp: Clean destination directory: " & destDir
            ChangeStatusBarText strMessages(82)
            DelRecursiveFolder destDir
        End If

        ' Отображаем ProgressBar
        With ctlProgressBar1
            .Value = 0
            .Visible = True
            .SetTaskBarProgressState PrbTaskBarStateInProgress
            .SetTaskBarProgressValue .Value, .Max
            ChangeFrmMainCaption .Value
        End With
        
        If cmbTypeBackUp.ListIndex = 0 Then
            miPBInterval = Round(10000 / lvCountCheck)
        Else
            miPBInterval = Round(9000 / lvCountCheck)
        End If

        miPBNext = 0
        '# loop all drivers in grid #
        nn = -1
        numCat = 1
        lvCount = lvDevices.ListItems.count
        If mbDebugStandart Then DebugMode "***StartBackUp: Count of drivers: " & lvCount
        If mbDebugStandart Then DebugMode "***StartBackUp: Count of checked drivers: " & lvCountCheck

        For ii = 1 To lvCount
            mbBackUPedFiles = False

            With lvDevices.ListItems.item(ii)
                '# Ищем в цикле выделенные строки
                If .Checked Then
                    If mbDebugStandart Then DebugMode "____________________________________________________________________"
                    If mbDebugStandart Then DebugMode "***StartBackUp: DRIVER in List №" & (ii)
                    'Заполняем массив даными
                    strDriverDesc = SafeDir(.Text)
                    strClass = SafeDir(.SubItems(5))
                    strInfFileName = .SubItems(6)
                    strInfSection = .SubItems(7)
                    If mbDebugStandart Then DebugMode "***StartBackUp: DRIVER=" & strDriverDesc & " Inf=" & strInfFileName
    
                    ' Прерываем процесс
                    If mbBreakUpdateDBAll Then
                        MsgBox strMessages(27) & vbNewLine & strDriverDesc, vbInformation, strProductName
                        If mbDebugStandart Then DebugMode "***StartBackUp: BREAK by USER"
                        Exit For
                    End If
    
                    nn = nn + 1
                    strStatusMsgTemp = strMessages(9) & " (" & nn + 1 & " " & strMessages(108) & " " & lvCountCheck & "): " & strDriverDesc & ": "
                    ChangeStatusBarText strStatusMsgTemp
                    ReDim Preserve inf(nn)
                    '# Создаем директорию приемник
                    strDest = BackslashAdd2Path(destDir) & strClass & vbBackslash & strDriverDesc
                    strInfFile2Path = BackslashAdd2Path(strDest) & strInfFileName
                    If mbDebugStandart Then DebugMode "***StartBackUp: DestForDriver=" & strDest
    
                    ' Если исходный inf-файл существует, то продолжаем, если нет пропускаем
                    If FileExists(strInfDir & strInfFileName) Then
    
                        ' Если каталога нет, то создаем
                        If PathExists(strDest) = False Then
                            CreateNewDirectory strDest
                            numCat = 1
                        Else
    
                            ' А если есть, то значит мы уже обрабатывали такой драйвер, делаем его копию
                            If FileExists(strInfFile2Path) = False Then
                                strDest = strDest & "_" & numCat
                                CreateNewDirectory strDest
                                numCat = numCat + 1
                            End If
                        End If
    
                        strInfFile2Path = BackslashAdd2Path(strDest) & strInfFileName
                        '# Копируем инф-файл в каталог назначения
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Copy Inf-File"
                        If mbDebugStandart Then DebugMode strStatusMsgTemp & "Analizing '[SourceDisksFiles]'"
                        CopyFileTo strInfDir & strInfFileName, strInfFile2Path
                        'CopyFileTo "c:\oem6.inf", strInfFile2Path
                        DoEvents
                        '# Копируем cat-файл в каталог назначения
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Search CatalogFile"
                        If mbDebugStandart Then DebugMode strStatusMsgTemp & "Search CatalogFile"
                        strCatFileName4Inf = FindCopyCatFile(strInfFile2Path, strDest)
    
                        ' Если существует cat-файл, то переименовываем inf-файл в имя cat-файла
                        If LenB(strCatFileName4Inf) Then
                            strInfFile2Path4Cat = PathCombine(GetPathNameFromPath(strInfFile2Path), GetFileName_woExt(strCatFileName4Inf) & ".inf")
    
                            If MoveFileTo(strInfFile2Path, strInfFile2Path4Cat) Then
                                strInfFile2Path = strInfFile2Path4Cat
                            End If
                        End If
    
                        DoEvents
                        ' Дополнительно ищем и копируем все файлы из каталога c:\WINDOWS\system32\DRVSTORE\
                        If mbDebugStandart Then DebugMode "***" & strStatusMsgTemp & "Analizing DRVSTORE"
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Analizing DriverStore folder"
    
                        If OSCurrVersionStruct.VerMajor = 5 Then
                            If LenB(strCatFileName4Inf) And IsWinXPOrLater Then
                                If FileExists(BackslashAdd2Path(strDest) & strCatFileName4Inf) Then
                                    ' Сравнение файлов по Hash
                                    strDataSHA1 = CalcHashFile(BackslashAdd2Path(strDest) & strCatFileName4Inf, CAPICOM_HASH_ALGORITHM_SHA1)
    
                                    arr_ZF1 = SearchFoldersInRoot(strSysDirDRVStore, "*" & "_" & UCase$(strDataSHA1) & "*")
    
                                    Dim lngUBoundZF1 As Long
    
                                    lngUBoundZF1 = UBound(arr_ZF1)
    
                                    For DD = 0 To lngUBoundZF1
                                        strFolderPath = arr_ZF1(DD).Path
                                        strFileNameInf = arr_ZF1(DD).Name
    
                                        If LenB(strFolderPath) Then
                                            If LenB(strFileNameInf) Then
                                                strFileNameInf = BackslashAdd2Path(strFolderPath) & strFileNameInf & ".inf"
    
                                                If FileExists(strFileNameInf) Then
    
                                                    'Сравнение файлов но Hash SHA1-сумме
                                                    mbCompare = CompareFilesByHashCAPICOM(strFileNameInf, strInfFile2Path)
    
                                                    If mbCompare Then
                                                        ' Удаляем предыдущий inf, чтобы не было дублей
                                                        DeleteFiles strInfFile2Path
                                                        strInfFile2Path = strFileNameInf
                                                        ' Копируем содержимое архива
                                                        If mbDebugStandart Then DebugMode "******CopyFiles from DrvStore: " & strFolderPath
                                                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Copying files from DriverStore folder"
                                                        lngNumFilesFromFolder = rgbCopyFiles(strFolderPath, strDest, ALL_FILES)
                                                        If mbDebugStandart Then DebugMode "******CopyFiles - count files: " & lngNumFilesFromFolder
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
    
                                    Next
                                End If
                            End If
    
                        Else
                            strFileNameInf = GetInfDriverStorePath(strInfDir & strInfFileName)
    
                            If LenB(strFileNameInf) Then
                                If FileExists(strFileNameInf) Then
    
                                    'Сравнение файлов найденного в DriverStore и Windows\inf по Hash SHA1-сумме
                                    mbCompare = CompareFilesByHashCAPICOM(strFileNameInf, strInfFile2Path)
    
                                    If mbCompare Then
                                        ' Получение пути каталога с драйверами
                                        strFolderPath = GetPathNameFromPath(strFileNameInf)
                                        ' Удаляем предыдущий inf, чтобы не было дублей
                                        DeleteFiles strInfFile2Path
                                        strInfFile2Path = strFileNameInf
                                        ' Копируем содержимое DrvStore в каталог назначения
                                        If mbDebugStandart Then DebugMode "******CopyFiles from DrvStore: " & strFolderPath
                                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Copying files from DriverStore folder"
                                        lngNumFilesFromFolder = rgbCopyFiles(strFolderPath, strDest, ALL_FILES)
                                        If mbDebugStandart Then DebugMode "******CopyFiles - count files: " & lngNumFilesFromFolder
                                    End If
                                End If
                            End If
                        End If
    
                        ' Анализируем секции sourcediskfiles sourcedisknames  и строим массим имен файлов и путей куда их надо копировать
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Collecting path of files information"
                        CollectDestPathFiles strInfFile2Path
                        '#  Читаем INF - для SourceDisksFiles на основе путей DefaultDestDir
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Analyzing '[SourceDisksFiles]'"
                        If mbDebugStandart Then DebugMode "***" & strStatusMsgTemp & "Analizing '[SourceDisksFiles]'"
                        arr_Z = LoadIniSectionKeys("SourceDisksFiles", strInfFile2Path)
                        CopyFile2Dest arr_Z, strDest, "DefaultDestDir", strInfFile2Path
                        DoEvents
                        '#  Читаем INF - из дополнительных секций DefaultDestDir
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Analyzing '[DestinationDirs]'"
                        If mbDebugStandart Then DebugMode "***" & strStatusMsgTemp & "Analizing '[DestinationDirs]'"
                        arr_Z2 = LoadIniSectionKeys("DestinationDirs", strInfFile2Path)
    
                        lngUBoundZ2 = UBound(arr_Z2)
    
                        For lngArrCount = 0 To lngUBoundZ2
                            strSectionName = arr_Z2(lngArrCount)
    
                            If LenB(strSectionName) Then
                                If StrComp(strSectionName, "DefaultDestDir", vbTextCompare) <> 0 Then
                                    arr_Z = LoadIniSectionKeys(strSectionName, strInfFile2Path)
                                    If mbDebugDetail Then DebugMode "***" & strStatusMsgTemp & "Analizing section: " & strSectionName
                                    CopyFile2Dest arr_Z, strDest, strSectionName, strInfFile2Path, True
                                End If
                            End If
    
                        Next
                        DoEvents
                        
                        ' Дополнительный анализ секций на параметр CopyFiles
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Analyzing CopyFiles '" & strInfSection & "'"
                        If mbDebugStandart Then DebugMode "***" & strStatusMsgTemp & "Analizing section by CopyFiles: " & strInfSection
                        arr_Z4 = GetSectionMass(strInfSection, strInfFile2Path, False)
                        arr_Z5 = SearchSectInSect(arr_Z4)
    
                        Dim lngUBoundZ5 As Long
    
                        lngUBoundZ5 = UBound(arr_Z5)
    
                        For lngArrCount = 0 To lngUBoundZ5
                            strSectionName = arr_Z5(lngArrCount)
    
                            If LenB(strSectionName) Then
                                If mbDebugDetail Then DebugMode "***" & strStatusMsgTemp & "Analizing section: " & strSectionName
                                arr_Z = LoadIniSectionKeys(strSectionName, strInfFile2Path)
                                CopyFile2Dest arr_Z, strDest, "DefaultDestDir", strInfFile2Path, True
                            End If
    
                        Next
                        DoEvents
                        
                        ' Дополнительный анализ секций на параметр CopyFiles Секции strInfSection.CoInstallers
                        Erase arr_Z4
                        Erase arr_Z5
                        ChangeStatusBarText strStatusMsgTemp & vbNewLine & "Analyzing CopyFiles '" & strInfSection & ".CoInstallers'"
                        If mbDebugStandart Then DebugMode "***" & strStatusMsgTemp & "Analizing section CoInstallers: " & strInfSection & ".CoInstallers"
                        arr_Z4 = GetSectionMass(strInfSection & ".CoInstallers", strInfFile2Path, False)
                        arr_Z5 = SearchSectInSect(arr_Z4)
                        lngUBoundZ5 = UBound(arr_Z5)
    
                        For lngArrCount = 0 To lngUBoundZ5
                            strSectionName = arr_Z5(lngArrCount)
    
                            If LenB(strSectionName) Then
                                If mbDebugDetail Then DebugMode "***" & strStatusMsgTemp & "Analizing section: " & strSectionName
                                arr_Z = LoadIniSectionKeys(strSectionName, strInfFile2Path)
                                CopyFile2Dest arr_Z, strDest, "DefaultDestDir", strInfFile2Path, True
                            End If
    
                        Next
                        DoEvents
                        
                        ' Ищем файлы в секции откуда ставились дрова
                        arr_Z3 = LoadIniSectionKeys(strInfSection, strInfFile2Path, False)
                        CopyFile2Dest arr_Z3, strDest, "DefaultDestDir", strInfFile2Path
                    Else
                        If mbDebugStandart Then DebugMode "StartBackUp: Inf-File NotExist=" & strInfDir & strInfFileName
                    End If
    
                    '# show progress #
                    miPBNext = miPBNext + miPBInterval
    
                    If cmbTypeBackUp.ListIndex = 0 Then
                        If miPBNext > 10000 Then
                            miPBNext = 10000
                        End If
                    Else
                        If miPBNext > 9000 Then
                            miPBNext = 9000
                        End If
                    End If
    
                    With ctlProgressBar1
                        .Value = miPBNext
                        .SetTaskBarProgressValue .Value, .Max
                        ChangeFrmMainCaption .Value
                    End With
                    
                    mbBackUPedFiles = True
                End If
            End With

            ' Если что-то было забекаплено, то заносим в лог, если включена отладка
            If mbBackUPedFiles And mbDebugStandart Then
                DoEvents
                strFileList = ListingDirectory(strDest, True)
                If mbDebugStandart Then DebugMode "***Content directory after backup: " & strFileList
            End If

            ' очищаю массивы
            Erase arr_Z
            Erase arr_Z2
            Erase arr_Z3
            Erase arr_Z4
            Erase arr_Z5
            Erase arr_ZF1
        Next
        If mbDebugStandart Then DebugMode "***BackUp all Checked drivers finished."
        DoEvents
        lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
        strAllTimeScriptRun = CalculateTime(lngTimeScriptFinish, True)

        ' Если прерван процесс
        If mbBreakUpdateDBAll Then
            mbBreakUpdateDBAll = False
            ChangeStatusBarText strMessages(66) & " " & strAllTimeScriptRun, , True
        Else

            '# type of backup #
            Select Case cmbTypeBackUp.ListIndex

                '# create ZIP #
                Case 1
                    With ctlProgressBar1
                        .Value = 9000
                        .SetTaskBarProgressValue .Value, .Max
                        ChangeFrmMainCaption .Value
                    End With

                    ChangeStatusBarText "Zipping driver files..."
                    str7zFileArchivePath = BackslashAdd2Path(destDirDialog) & txtArchName & ".7z"
                    If mbDebugStandart Then DebugMode "StartBackUp: Zip to File=" & str7zFileArchivePath
                    mbDoZip = DoZip(destDir, str7zFileArchivePath)
                    DoEvents

                    If mbDoZip Then
                        '# delete temp folder #
                        ChangeStatusBarText "Delete temporary files...Please wait"
                        DelFolderBackUp destDir
                    End If

                    MousePointer = 0
                    lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
                    strAllTimeScriptRun = CalculateTime(lngTimeScriptFinish, True)
                    ChangeStatusBarText strMessages(67) & " " & strAllTimeScriptRun, , True
                    With ctlProgressBar1
                        .Value = 10000
                        .SetTaskBarProgressValue .Value, .Max
                        ChangeFrmMainCaption .Value
                    End With

                    If mbDoZip Then
                        MsgBox strMessages(10) & vbNewLine & str7zFileArchivePath, vbInformation + vbOKOnly, strProductName
                    Else
                        MsgBox strMessages(12), vbInformation + vbOKOnly, strProductName
                    End If

                '# create ZIP-SFX with DPInst #
                Case 2
                    With ctlProgressBar1
                        .Value = 9000
                        .SetTaskBarProgressValue .Value, .Max
                        ChangeFrmMainCaption .Value
                    End With
                    ChangeStatusBarText "Zipping driver files..."
                    str7zFileArchivePath = BackslashAdd2Path(destDirDialog) & txtArchName & ".exe"
                    If mbDebugStandart Then DebugMode "StartBackUp: Zip to File=" & str7zFileArchivePath
                    mbDoZip = DoZip(destDir, str7zFileArchivePath)
                    DoEvents

                    If mbDoZip Then
                        '# delete temp folder #
                        ChangeStatusBarText "Delete temporary files...Please wait"
                        DelFolderBackUp destDir
                    End If

                    '# display default cursor #
                    MousePointer = 0
                    lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
                    strAllTimeScriptRun = CalculateTime(lngTimeScriptFinish, True)
                    ChangeStatusBarText strMessages(67) & " " & strAllTimeScriptRun, , True
                    With ctlProgressBar1
                        .Value = 10000
                        .SetTaskBarProgressValue .Value, .Max
                        ChangeFrmMainCaption .Value
                    End With

                    If mbDoZip Then
                        MsgBox strMessages(10) & vbNewLine & str7zFileArchivePath, vbInformation + vbOKOnly, strProductName
                    Else
                        MsgBox strMessages(12), vbInformation + vbOKOnly, strProductName
                    End If

                Case Else
                    ChangeStatusBarText strMessages(67) & " " & strAllTimeScriptRun, , True
                    With ctlProgressBar1
                        .Value = 10000
                        .SetTaskBarProgressValue .Value, .Max
                        ChangeFrmMainCaption .Value
                    End With
                    MsgBox strMessages(10), vbInformation + vbOKOnly, strProductName
            End Select
        End If

        '# show info of end process #
        ctlProgressBar1.Visible = False
        ChangeFrmMainCaption
        ctlProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    End If

    MousePointer = 0
    ' РазБлокируем элементы от греха подальше
    BlockControl False
    If mbDebugStandart Then DebugMode "BlockControl: TRUE"

    'РазБлокируем лист после бекапа
    If mbBlockListOnBackup Then
        lvDevices.Enabled = True
        If mbDebugStandart Then DebugMode "BlockListOnBackup: FALSE"
        lvDevices.Refresh
    End If
        
    If mbDebugStandart Then DebugMode "cmdStartBackUp_Click-Finish"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function StringCleaner
'! Description (Описание)  :   [Чистка строки для пути назначения]
'! Parameters  (Переменные):   strString (String)
'!--------------------------------------------------------------------------------
Private Function StringCleaner(ByVal strString As String) As String

    Dim strString_x() As String

    If InStr(1, strString, ";") Then
        strString_x = Split(strString, ";")
        strString = Trim$(strString_x(0))
    End If

    If InStr(1, strString, ",") Then
        strString_x = Split(strString, ",")
        strString = strString_x(0)
    End If

    If InStr(1, strString, vbNullChar) Then
        strString = TrimNull(strString)
    End If

    If InStr(1, strString, vbTab) Then
        strString = Replace$(strString, vbTab, vbNullString)
    End If

    If InStr(1, strString, strQuotes) Then
        strString = Replace$(strString, strQuotes, vbNullString)
    End If

    StringCleaner = strString
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchName_KeyPress
'! Description (Описание)  :   [Событие нажатия клавиш в тексте имени архивного файла]
'! Parameters  (Переменные):   KeyAscii (Integer)
'!--------------------------------------------------------------------------------
Private Sub txtArchName_KeyPress(KeyAscii As Integer)

    Dim sTemplate As String

    sTemplate = "!@#$%^&*()_+=\/:;?><|[],"

    If InStr(1, sTemplate, Chr$(KeyAscii), vbTextCompare) Then
        KeyAscii = 0
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UnloadAllForms
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FormToIgnore (String = vbNullString)
'!--------------------------------------------------------------------------------
Public Sub UnloadAllForms(Optional FormToIgnore As String = vbNullString)

    Dim objForm As Form

    For Each objForm In Forms

        If Not objForm Is Nothing Then
            If StrComp(objForm.Name, FormToIgnore, vbTextCompare) <> 0 Then
                Unload objForm
                Set objForm = Nothing
            End If
        End If

    Next objForm

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub VerModules
'! Description (Описание)  :   [Отображение версий модулей]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub VerModules()
    MsgBox strMessages(35) & vbNewLine & _
           "DPinst.exe (x86)" & vbTab & GetFileVersionOnly(strDPInstExePath86) & vbNewLine & _
           "DPinst.exe (x64)" & vbTab & GetFileVersionOnly(strDPInstExePath64) & vbNewLine & _
           "7zSD.sfx" & vbTab & GetFileVersionOnly(strArh7zSFXPATH) & vbNewLine & _
           "7za.exe (x86)" & vbTab & GetFileVersionOnly(strArh7zExePath86) & vbNewLine & _
           "7za64.exe (x64)" & vbTab & GetFileVersionOnly(strArh7zExePath64), vbInformation, strProductName
End Sub

