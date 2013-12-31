VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки программы"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13245
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   13245
   StartUpPosition =   1  'CenterOwner
   Begin prjDBS.ctlJCFrames frMain 
      Height          =   5275
      Left            =   3180
      Top             =   25
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      Caption         =   "Основные настройки программы"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDBS.ctlTextInteger txtDebugLogLevel 
         Height          =   255
         Left            =   7680
         TabIndex        =   0
         Top             =   4320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         Caption         =   ""
         Text            =   "0"
         MinValue        =   1
      End
      Begin VB.ComboBox cmbTypeBackUp 
         Height          =   345
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2760
         Width           =   4815
      End
      Begin prjDBS.ctlCheckBoxTVH chkRemoveHistory 
         Height          =   210
         Left            =   495
         TabIndex        =   6
         Top             =   4590
         Width           =   7920
         _ExtentX        =   11245
         _ExtentY        =   370
         Caption         =   "Удалять старый лог-файл при запуске программы"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkRemoveTemp 
         Height          =   210
         Left            =   495
         TabIndex        =   7
         Top             =   3750
         Width           =   7920
         _ExtentX        =   8281
         _ExtentY        =   370
         Caption         =   "Удалять созданные временные файлы при выходе"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkUpdate 
         Height          =   210
         Left            =   495
         TabIndex        =   8
         Top             =   660
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   370
         Caption         =   "Проверять обновление программы"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkHideOtherProcess 
         Height          =   210
         Left            =   495
         TabIndex        =   9
         Top             =   1200
         Width           =   7920
         _ExtentX        =   6350
         _ExtentY        =   370
         Caption         =   "Запускать внешние процессы скрытыми"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkTempPath 
         Height          =   210
         Left            =   495
         TabIndex        =   10
         Top             =   3450
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   370
         Caption         =   "Каталог временных файлов"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkUpdateBeta 
         Height          =   210
         Left            =   3630
         TabIndex        =   11
         Top             =   660
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   370
         Caption         =   "Проверять на наличие beta-версий программы."
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkSilentDll 
         Height          =   210
         Left            =   495
         TabIndex        =   12
         Top             =   930
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   370
         Caption         =   "Тихая регистрация/обновление системных компонентов (dll и ocx)"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkDebug 
         Height          =   210
         Left            =   495
         TabIndex        =   13
         Top             =   4320
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   370
         Caption         =   "Отладочный режим включен"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlUcPickBox ucDebugLogPath 
         Height          =   315
         Left            =   2280
         TabIndex        =   14
         Top             =   4845
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlUcPickBox ucTempPath 
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         Top             =   3390
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         FileFlags       =   3678212
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlOptionBoxTVH optGrp1 
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         Caption         =   "Microsoft"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlOptionBoxTVH optGrp2 
         Height          =   255
         Left            =   2085
         TabIndex        =   17
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         Caption         =   "OEM"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlOptionBoxTVH optGrp3 
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   2100
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         Caption         =   "Все"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlOptionBoxTVH optGrp4 
         Height          =   255
         Left            =   2085
         TabIndex        =   19
         Top             =   2100
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         Caption         =   "Ни одного"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkHideOther 
         Height          =   255
         Left            =   3720
         TabIndex        =   20
         Top             =   1800
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         Caption         =   "Скрывать все кроме выбранной группы"
         Transparent     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   -1  'True
      End
      Begin prjDBS.ctlCheckBoxTVH chkCheckAll 
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Top             =   2040
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         Caption         =   "Выделять всю группу при выборе"
         Transparent     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   -1  'True
      End
      Begin prjDBS.ctlLabelTVH lblDebugLogPath 
         Height          =   255
         Left            =   495
         Top             =   4875
         Width           =   1695
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
         Caption         =   "Путь до log-файла:"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblOptionsTemp 
         Height          =   285
         Left            =   240
         Top             =   3150
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Работа с временными файлами"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblOptionsStart 
         Height          =   285
         Left            =   240
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Действия при запуске программы"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblDebug 
         Height          =   285
         Left            =   240
         Top             =   3990
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Настройки отладочного режима"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblRezim 
         Height          =   285
         Left            =   240
         Top             =   1485
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Режим работы фильтра по умолчанию"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblTypeBackUp 
         Height          =   225
         Left            =   240
         Top             =   2400
         Width           =   8175
         _ExtentX        =   7726
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Режим создания резервных копий по умолчанию"
         AutoSize        =   -1  'True
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblDebugLogLevel 
         Height          =   255
         Left            =   4680
         Top             =   4320
         Width           =   3015
         _ExtentX        =   5318
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
         Caption         =   "Уровень отладки:"
         ShadowStyle     =   0
         Alignment       =   1
      End
   End
   Begin prjDBS.ctlJCFrames frMainTools 
      Height          =   5275
      Left            =   3375
      Top             =   200
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      Caption         =   "Расположение основных утилит (Tools)"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDBS.ctlLabelTVH lblArcSFXConfigEn 
         Height          =   255
         Left            =   150
         Top             =   2730
         Width           =   2280
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
         Caption         =   "7za-SFXConfig (English)"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblArcSFXConfig 
         Height          =   255
         Left            =   150
         Top             =   2250
         Width           =   2280
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
         Caption         =   "7za-SFXConfig"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblArc 
         Height          =   255
         Left            =   150
         Top             =   1350
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "7za"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblDPInst64 
         Height          =   255
         Left            =   150
         Top             =   930
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "DPInst.exe (64-bit)"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblDPInst86 
         Height          =   255
         Left            =   150
         Top             =   510
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "DPInst.exe (32-bit)"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblArcSFX 
         Height          =   255
         Left            =   150
         Top             =   1770
         Width           =   2280
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
         Caption         =   "7za-sfxModule"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlUcPickBox ucDPInst86Path 
         Height          =   315
         Left            =   2535
         TabIndex        =   3
         Top             =   510
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlUcPickBox ucDPInst64Path 
         Height          =   315
         Left            =   2535
         TabIndex        =   4
         Top             =   930
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlUcPickBox ucArchPath 
         Height          =   315
         Left            =   2535
         TabIndex        =   22
         Top             =   1350
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlJCbutton cmdPathDefault 
         Height          =   495
         Left            =   4815
         TabIndex        =   23
         Top             =   3210
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Сбросить настройки расположения утилит"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDBS.ctlUcPickBox ucArchPathSFX 
         Height          =   315
         Left            =   2535
         TabIndex        =   39
         Top             =   1770
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlUcPickBox ucArchPathSFXConfig 
         Height          =   315
         Left            =   2535
         TabIndex        =   40
         Top             =   2250
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlUcPickBox ucArchPathSFXConfigEn 
         Height          =   315
         Left            =   2535
         TabIndex        =   51
         Top             =   2730
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         FileFlags       =   2621446
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
   End
   Begin prjDBS.ctlJCFrames frArchName 
      Height          =   5275
      Left            =   3585
      Top             =   400
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      Caption         =   "Имя архива"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin VB.TextBox txtArchNameShablon 
         Height          =   330
         Left            =   480
         TabIndex        =   62
         Top             =   2205
         Width           =   7635
      End
      Begin VB.TextBox txtMacrosPCName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "%PCNAME%"
         Top             =   3285
         Width           =   1500
      End
      Begin VB.TextBox txtMacrosPCModel 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "%PCMODEL%"
         Top             =   3645
         Width           =   1500
      End
      Begin VB.TextBox txtMacrosOSVER 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "%OSVER%"
         Top             =   4005
         Width           =   1500
      End
      Begin VB.TextBox txtMacrosOSBIT 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "%OSBIT%"
         Top             =   4365
         Width           =   1500
      End
      Begin VB.TextBox txtMacrosDate 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "%DATE%"
         Top             =   4725
         Width           =   1500
      End
      Begin prjDBS.ctlOptionBoxTVH optArchModelPC 
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   1125
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   450
         Caption         =   "Модель компьютера"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlOptionBoxTVH optArchNamePC 
         Height          =   255
         Left            =   480
         TabIndex        =   52
         Top             =   765
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   450
         Caption         =   "Имя компьютера"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlOptionBoxTVH optArchCustom 
         Height          =   255
         Left            =   480
         TabIndex        =   56
         Top             =   1485
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   450
         Caption         =   "По шаблону"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlLabelTVH lblMacrosDate 
         Height          =   375
         Left            =   2400
         Top             =   4725
         Width           =   5775
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
         Caption         =   "Дата создания резервной копии"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblMacrosOSBit 
         Height          =   375
         Left            =   2400
         Top             =   4365
         Width           =   5775
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
         Caption         =   "Архитектура операционной системы, в виде x32[64]"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblMacrosOSVer 
         Height          =   375
         Left            =   2400
         Top             =   4005
         Width           =   5775
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
         Caption         =   "Версия операционной системы в виде wnt5[6]"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblMacrosPCModel 
         Height          =   375
         Left            =   2400
         Top             =   3645
         Width           =   5775
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
         Caption         =   "Модель компьютера/материнской платы"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblMacrosParam 
         Height          =   255
         Left            =   480
         Top             =   2970
         Width           =   1755
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Параметр"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblMacrosDescription 
         Height          =   255
         Left            =   2400
         Top             =   2970
         Width           =   5865
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Описание  параметра"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblMacrosPCName 
         Height          =   375
         Left            =   2400
         Top             =   3285
         Width           =   5775
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
         Caption         =   "Краткое имя компьютера, без доменного суффикса"
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblMacrosType 
         Height          =   285
         Left            =   480
         Top             =   2685
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Доступные макроподстановки:"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblArchShablon 
         Height          =   285
         Left            =   240
         Top             =   1845
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Шаблон имени архива"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblArchNameStart 
         Height          =   285
         Left            =   240
         Top             =   405
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Имя архива по умолчанию"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
   End
   Begin prjDBS.ctlJCFrames frOS 
      Height          =   5275
      Left            =   3780
      Top             =   600
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      Caption         =   "Поддерживаемые ОС"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDBS.ctlJCbutton cmdAddOS 
         Height          =   750
         Left            =   120
         TabIndex        =   53
         Top             =   4400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1323
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Добавить"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDBS.ctlJCbutton cmdEditOS 
         Height          =   750
         Left            =   2160
         TabIndex        =   54
         Top             =   4400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1323
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Изменить"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDBS.ctlJCbutton cmdDelOS 
         Height          =   750
         Left            =   4200
         TabIndex        =   55
         Top             =   4400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1323
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Удалить"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
   End
   Begin prjDBS.ctlJCFrames frDesign 
      Height          =   5275
      Left            =   3960
      Top             =   800
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      Caption         =   "Оформление"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin VB.ComboBox cmbImageMain 
         Height          =   345
         ItemData        =   "frmOptions.frx":000C
         Left            =   405
         List            =   "frmOptions.frx":0019
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   3075
         Width           =   3000
      End
      Begin prjDBS.ctlUcPickBox ucColorButton 
         Height          =   315
         Left            =   405
         TabIndex        =   43
         Top             =   2340
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         DefaultExt      =   ""
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlUcPickBox ucFontButton 
         Height          =   315
         Left            =   405
         TabIndex        =   44
         Top             =   1935
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         Color           =   0
         DefaultExt      =   ""
         DialogType      =   2
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         ToolTipText3    =   "Click Here to Locate File"
      End
      Begin prjDBS.ctlCheckBoxTVH chkButtonDisable 
         Height          =   450
         Left            =   5790
         TabIndex        =   45
         Top             =   1935
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   794
         Caption         =   "Активность кнопки (только для тестирования)"
         Checked         =   -1  'True
      End
      Begin prjDBS.ctlCheckBoxTVH chkFormMaximaze 
         Height          =   210
         Left            =   3285
         TabIndex        =   46
         Top             =   795
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   370
         Caption         =   "Запускать развернутым на весь экран"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlTextInteger txtFormHeight 
         Height          =   255
         Left            =   1245
         TabIndex        =   47
         Top             =   795
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   ""
         Text            =   "0"
      End
      Begin prjDBS.ctlTextInteger txtFormWidth 
         Height          =   255
         Left            =   1245
         TabIndex        =   48
         Top             =   1140
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Caption         =   ""
         Text            =   "0"
      End
      Begin prjDBS.ctlCheckBoxTVH chkFormSizeSave 
         Height          =   210
         Left            =   3285
         TabIndex        =   49
         Top             =   1140
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   370
         Caption         =   "Сохранять размеры формы при выходе"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlLabelTVH lblFormWidthMin 
         Height          =   930
         Left            =   135
         Top             =   3600
         Width           =   8370
         _ExtentX        =   0
         _ExtentY        =   132
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   $"frmOptions.frx":0051
         WordWrap        =   -1  'True
         ForeColor       =   255
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblImageMain 
         Height          =   255
         Left            =   135
         Top             =   2775
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Основные картинки"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
         ShadowColorStart=   0
      End
      Begin prjDBS.ctlLabelTVH lblFormWidth 
         Height          =   210
         Left            =   405
         Top             =   1140
         Width           =   645
         _ExtentX        =   1270
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
         Caption         =   "Ширина:"
         AutoSize        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblFormHeight 
         Height          =   210
         Left            =   405
         Top             =   795
         Width           =   630
         _ExtentX        =   1191
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
         Caption         =   "Высота:"
         AutoSize        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblSizeForm 
         Height          =   255
         Left            =   135
         Top             =   495
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Размеры основного окна"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlJCbutton cmdFutureButton 
         Height          =   510
         Left            =   3390
         TabIndex        =   50
         Top             =   1935
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12244692
         Caption         =   "Твоя будущая кнопка"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
         ColorScheme     =   1
      End
      Begin prjDBS.ctlLabelTVH lblSizeButton 
         Height          =   255
         Left            =   135
         Top             =   1575
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Свойства кнопок"
         ForeColor       =   -2147483635
         ShadowStyle     =   0
      End
   End
   Begin prjDBS.ctlJCFrames frOptions 
      Height          =   5275
      Left            =   50
      Top             =   25
      Width           =   3000
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      Caption         =   "Настройки"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDBS.ctlJCbutton cmdOK 
         Height          =   750
         Left            =   75
         TabIndex        =   1
         Top             =   3500
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1323
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Сохранить изменения и выйти"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin prjDBS.ctlJCbutton cmdExit 
         Height          =   735
         Left            =   75
         TabIndex        =   2
         Top             =   4400
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1296
         ButtonStyle     =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Выход без сохранения"
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
   End
   Begin prjDBS.ctlJCFrames frOther 
      Height          =   5275
      Left            =   4185
      Top             =   1000
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
   End
   Begin prjDBS.ctlJCFrames frDpInstParam 
      Height          =   5275
      Left            =   4380
      Top             =   1200
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      RoundedCorner   =   0   'False
      Caption         =   "Параметры запуска DPInst"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin VB.TextBox txtCmdStringDPInst 
         Height          =   330
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   4845
         Width           =   5535
      End
      Begin prjDBS.ctlXpButton cmdLegacyMode 
         Height          =   210
         Left            =   2670
         TabIndex        =   25
         ToolTipText     =   "More on MSDN..."
         Top             =   660
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   370
         Caption         =   "?"
         PicturePosition =   0
         ButtonStyle     =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   0
         XPColor_Hover   =   0
         XPDefaultColors =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin prjDBS.ctlXpButton cmdPromptIfDriverIsNotBetter 
         Height          =   210
         Left            =   2655
         TabIndex        =   26
         ToolTipText     =   "More on MSDN..."
         Top             =   1215
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   370
         Caption         =   "?"
         PicturePosition =   0
         ButtonStyle     =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   0
         XPColor_Hover   =   0
         XPDefaultColors =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin prjDBS.ctlXpButton cmdForceIfDriverIsNotBetter 
         Height          =   210
         Left            =   2670
         TabIndex        =   27
         ToolTipText     =   "More on MSDN..."
         Top             =   1815
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   370
         Caption         =   "?"
         PicturePosition =   0
         ButtonStyle     =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   0
         XPColor_Hover   =   0
         XPDefaultColors =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin prjDBS.ctlXpButton cmdSuppressAddRemovePrograms 
         Height          =   210
         Left            =   2670
         TabIndex        =   28
         ToolTipText     =   "More on MSDN..."
         Top             =   2370
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   370
         Caption         =   "?"
         PicturePosition =   0
         ButtonStyle     =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   0
         XPColor_Hover   =   0
         XPDefaultColors =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin prjDBS.ctlXpButton cmdSuppressWizard 
         Height          =   210
         Left            =   2670
         TabIndex        =   29
         ToolTipText     =   "More on MSDN..."
         Top             =   2865
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   370
         Caption         =   "?"
         PicturePosition =   0
         ButtonStyle     =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   0
         XPColor_Hover   =   0
         XPDefaultColors =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin prjDBS.ctlXpButton cmdQuietInstall 
         Height          =   210
         Left            =   2670
         TabIndex        =   30
         ToolTipText     =   "More on MSDN..."
         Top             =   3420
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   370
         Caption         =   "?"
         PicturePosition =   0
         ButtonStyle     =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   0
         XPColor_Hover   =   0
         XPDefaultColors =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin prjDBS.ctlCheckBoxTVH chkLegacyMode 
         Height          =   210
         Left            =   135
         TabIndex        =   31
         Top             =   705
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Caption         =   "LegacyMode"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkPromptIfDriverIsNotBetter 
         Height          =   210
         Left            =   135
         TabIndex        =   32
         Top             =   1215
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   370
         Caption         =   "PromptIfDriverIsNotBetter"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkForceIfDriverIsNotBetter 
         Height          =   210
         Left            =   135
         TabIndex        =   33
         Top             =   1815
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   370
         Caption         =   "ForceIfDriverIsNotBetter"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkSuppressAddRemovePrograms 
         CausesValidation=   0   'False
         Height          =   210
         Left            =   135
         TabIndex        =   34
         Top             =   2370
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         Caption         =   "SuppressAddRemovePrograms"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkSuppressWizard 
         Height          =   210
         Left            =   135
         TabIndex        =   35
         Top             =   2865
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   370
         Caption         =   "SuppressWizard"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkQuietInstall 
         Height          =   210
         Left            =   135
         TabIndex        =   36
         Top             =   3420
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Caption         =   "QuietInstall"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlCheckBoxTVH chkScanHardware 
         Height          =   210
         Left            =   135
         TabIndex        =   37
         Top             =   3960
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   370
         Caption         =   "ScanHardware"
         Transparent     =   -1  'True
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
      Begin prjDBS.ctlXpButton cmdScanHardware 
         Height          =   210
         Left            =   2670
         TabIndex        =   38
         ToolTipText     =   "More on MSDN..."
         Top             =   3915
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   370
         Caption         =   "?"
         PicturePosition =   0
         ButtonStyle     =   0
         PictureWidth    =   0
         PictureHeight   =   0
         ShowFocusRect   =   0   'False
         XPColor_Pressed =   0
         XPColor_Hover   =   0
         XPDefaultColors =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin prjDBS.ctlLabelTVH lblCmdStringDPInst 
         Height          =   285
         Left            =   150
         Top             =   4845
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Итоговые параметры запуска "
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblDescription 
         Height          =   255
         Left            =   2745
         Top             =   390
         Width           =   5505
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Описание  параметра"
         ShadowStyle     =   0
         Alignment       =   2
      End
      Begin prjDBS.ctlLabelTVH lblParam 
         Height          =   255
         Left            =   135
         Top             =   390
         Width           =   2595
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Параметр"
         ShadowStyle     =   0
         Alignment       =   2
      End
      Begin prjDBS.ctlLabelTVH lblPromptIfDriverIsNotBetter 
         Height          =   540
         Left            =   2940
         Top             =   1215
         Width           =   5505
         _ExtentX        =   9710
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
         Caption         =   "display a dialog box if a new driver is not a better match to a device than a driver that is currently installed on the device"
         WordWrap        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblLegacyMode 
         Height          =   450
         Left            =   2940
         Top             =   705
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "install unsigned drivers and driver packages that have missing files"
         WordWrap        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblForceIfDriverIsNotBetter 
         Height          =   540
         Left            =   2940
         Top             =   1815
         Width           =   5505
         _ExtentX        =   9710
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
         Caption         =   "install a driver on a device even if the driver that is currently installed on the device is a better match than the new driver"
         WordWrap        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblSuppressAddRemovePrograms 
         Height          =   540
         Left            =   2940
         Top             =   2370
         Width           =   5505
         _ExtentX        =   9710
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
         Caption         =   "suppress the addition of Add or Remove Programs entries that represent the drivers and driver package"
         WordWrap        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblSuppressWizard 
         Height          =   540
         Left            =   2940
         Top             =   2865
         Width           =   5505
         _ExtentX        =   9710
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
         Caption         =   "configures DPInst to suppress the display of wizard pages and other user messages that DPInst generates."
         WordWrap        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblQuietInstall 
         Height          =   540
         Left            =   2940
         Top             =   3420
         Width           =   5550
         _ExtentX        =   9790
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
         Caption         =   "configures DPInst to suppress the display of wizard pages and most other user messages."
         WordWrap        =   -1  'True
         ShadowStyle     =   0
      End
      Begin prjDBS.ctlLabelTVH lblScanHardware 
         Height          =   960
         Left            =   2940
         Top             =   3915
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   1693
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   $"frmOptions.frx":010C
         WordWrap        =   -1  'True
         ShadowStyle     =   0
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents lvOS            As cListView
Attribute lvOS.VB_VarHelpID = -1
Public WithEvents lvOptions       As cListView
Attribute lvOptions.VB_VarHelpID = -1

Private strItemOptions1           As String
Private strItemOptions2           As String
Private strItemOptions3           As String
Private strItemOptions4           As String
Private strItemOptions5           As String
Private strItemOptions6           As String
Private strTableOSHeader1         As String
Private strTableOSHeader2         As String
Private strTableOSHeader3         As String
Private cmbListTypeBackupElement1 As String
Private cmbListTypeBackupElement2 As String
Private cmbListTypeBackupElement3 As String

Private Sub ChangeButtonProperties()

    SetButtonProperties , cmdFutureButton, True
    ucFontButton.FontColor = cmdFutureButton.ForeColor
End Sub

Private Sub chkButtonDisable_Click()

    cmdFutureButton.Enabled = chkButtonDisable.Checked
End Sub

Private Sub chkDebug_Click()

    DebugCtlEnable chkDebug.Checked
End Sub

Private Sub chkForceIfDriverIsNotBetter_Click()

    mboolDpInstForceIfDriverIsNotBetter = chkForceIfDriverIsNotBetter.Checked
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkFormMaximaze_Click()

    If chkFormMaximaze.Checked Then
        chkFormSizeSave.Checked = False
    End If
End Sub

Private Sub chkFormSizeSave_Click()

    If chkFormSizeSave.Checked Then
        chkFormMaximaze.Checked = False
    End If
End Sub

Private Sub chkHideOther_Click()

    chkCheckAll.Enabled = chkHideOther.Checked
End Sub

Private Sub chkLegacyMode_Click()

    mboolDpInstLegacyMode = chkLegacyMode.Checked
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkPromptIfDriverIsNotBetter_Click()

    mboolDpInstPromptIfDriverIsNotBetter = chkPromptIfDriverIsNotBetter.Checked
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkQuietInstall_Click()

    mboolDpInstQuietInstall = chkQuietInstall.Checked
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkScanHardware_Click()

    mboolDpInstScanHardware = chkScanHardware.Checked
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkSuppressAddRemovePrograms_Click()

    mboolDpInstSuppressAddRemovePrograms = chkSuppressAddRemovePrograms.Checked
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkSuppressWizard_Click()

    mboolDpInstSuppressWizard = chkSuppressWizard.Checked
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkTempPath_Click()

    TempCtlEnable chkTempPath.Checked
End Sub

Private Sub chkUpdate_Click()

    UpdateCtlEnable chkUpdate.Checked
End Sub

Private Sub cmbImageMain_Click()

    If PathFileExists(strPathImageMain & cmbImageMain.Text) = 0 Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If
End Sub

Private Sub cmbImageMain_GotFocus()

    HighlightActiveControl Me, cmbImageMain, True
End Sub

Private Sub cmbImageMain_LostFocus()

    If PathFileExists(strPathImageMain & cmbImageMain.Text) = 0 Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If

    HighlightActiveControl Me, cmbImageMain, False
End Sub

Private Sub cmbTypeBackUp_GotFocus()

    HighlightActiveControl Me, cmbTypeBackUp, True
End Sub

Private Sub cmbTypeBackUp_LostFocus()

    HighlightActiveControl Me, cmbTypeBackUp, False
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdAddOS_Click
'!  Переменные  :
'!  Описание    :  кнопка добавления ОС
'! -----------------------------------------------------------
Private Sub cmdAddOS_Click()

    mboolAddInList = True
    frmOSEdit.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdDelOS_Click
'!  Переменные  :
'!  Описание    :  кнопка удаление ОС
'! -----------------------------------------------------------
Private Sub cmdDelOS_Click()

    Dim i As Long

    With lvOS

        If .Count > 0 Then
            i = .SelectedItem
            .RemoveItem (i)
            LastIdOS = LastIdOS - 1
        End If
    End With

    'LVOS
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdEditOS_Click
'!  Переменные  :
'!  Описание    :  кнопка редактирование ОС
'! -----------------------------------------------------------
Private Sub cmdEditOS_Click()

    TransferOSData
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdExit_Click
'!  Переменные  :
'!  Описание    : Нажатие кнопки Выход. Выход без сохранения
'! -----------------------------------------------------------
Private Sub cmdExit_Click()

    Unload Me
End Sub

Private Sub cmdForceIfDriverIsNotBetter_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = kavichki & "http://msdn.microsoft.com/en-us/library/ms793551.aspx" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdLegacyMode_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = kavichki & "http://msdn.microsoft.com/en-us/library/ms794322.aspx" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdOK_Click
'!  Переменные  :
'!  Описание    :  Нажатие кнопки ОК. Применение настроек
'! -----------------------------------------------------------
Private Sub cmdOK_Click()

    Dim MsgRet As Long

    If mboolIsDriveCDRoom And mboolLoadIniTmpAfterRestart Then
        SaveOptions
        MsgRet = MsgBox(strMessages(36), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mboolRestartProgram = MsgRet = vbYes
    ElseIf Not FileisReadOnly(strSysIni) Then
        SaveOptions
        MsgRet = MsgBox(strMessages(36), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mboolRestartProgram = MsgRet = vbYes
    End If

    Unload Me
End Sub

Private Sub cmdPathDefault_Click()

    'Секция DPInst
    ucDPInst86Path.Path = "Tools\DPInst\DPInst.exe"
    ucDPInst64Path.Path = "Tools\DPInst\DPInst64.exe"
    'Секция Arc
    ucArchPath.Path = "Tools\Arc\7za.exe"
    ucArchPathSFX.Path = "Tools\Arc\sfx\7zSD.sfx"
    ucArchPathSFXConfig.Path = "Tools\Arc\sfx\config.txt"
    ucArchPathSFXConfigEn.Path = "Tools\Arc\sfx\config_en.txt"
End Sub

Private Sub cmdPromptIfDriverIsNotBetter_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = kavichki & "http://msdn.microsoft.com/en-us/library/ms793530.aspx" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdQuietInstall_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = kavichki & "http://msdn.microsoft.com/en-us/library/ms794300.aspx" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdScanHardware_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = kavichki & "http://msdn.microsoft.com/en-us/library/ms794295.aspx" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdSuppressAddRemovePrograms_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = kavichki & "http://msdn.microsoft.com/en-us/library/ms794270.aspx" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdSuppressWizard_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = kavichki & "http://msdn.microsoft.com/en-us/library/ms791062.aspx" & kavichki
    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub DebugCtlEnable(ByVal mboolEnable As Boolean)

    chkRemoveHistory.Enabled = mboolEnable
    ucDebugLogPath.Enabled = mboolEnable
End Sub

Private Sub FontCharsetChange()

    ' Выставляем шрифт
    Me.Font.Name = strOtherForm_FontName
    Me.Font.Size = lngOtherForm_FontSize
    Me.Font.Charset = lngDialog_Charset
End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_KeyDown
'!  Переменные  :  KeyCode As Integer, Shift As Integer
'!  Описание    :  Обработка нажатий клавиш клавиатуры сначала на форме
'! -----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        If MsgBox(strMessages(37), vbQuestion + vbYesNo, strProductName) = vbYes Then
            Unload Me
        End If
    End If
End Sub

'! -----------------------------------------------------------
'!  Функция     :  Form_Load
'!  Переменные  :
'!  Описание    :  Загрузка формы
'! -----------------------------------------------------------
Private Sub Form_Load()

    'SetSmallIcon Me.hWnd
    Call SetIcon(Me.hwnd, "FRMOPTIONS", False)
    
    Me.Height = 5825
    Me.Width = 11900
    'Top
    frOptions.Top = 50
    frMain.Top = 50
    frMainTools.Top = 50
    frArchName.Top = 50
    frOS.Top = 50
    frDesign.Top = 50
    frOther.Top = 50
    frDpInstParam.Top = 50
    'Left
    frMain.Left = 3100
    frMainTools.Left = 3100
    frArchName.Left = 3100
    frOS.Left = 3100
    frDesign.Left = 3100
    frOther.Left = 3100
    frDpInstParam.Left = 3100
    ' Устанавливаем минимальные значения
    txtFormHeight.MinValue = MainFormHeightMin
    txtFormWidth.MinValue = MainFormWidthMin
    ' Устанавливаем картинки кнопок и убираем описание кнопок
    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2BtnJC cmdAddOS, "BTN_ADD", strPathImageMainWork
    LoadIconImage2BtnJC cmdEditOS, "BTN_EDIT", strPathImageMainWork
    LoadIconImage2BtnJC cmdDelOS, "BTN_DELETE", strPathImageMainWork
    LoadIconImage2BtnJC cmdFutureButton, "BTN_STARTBACKUP", strPathImageMainWork

    ' Локализациz приложения
    If mboolLanguageChange Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    ' загрузить список опций
    tvOptionsLoad
    ' Заполнить опции
    ReadOptions
    ' установить опции шрифта и цвета
    'SetButtonProperties cmdChooseFont
    'cmdColorButton.Value = lngDialog_Color
    ' установить опции шрифта и цвета
    SetButtonProperties , cmdFutureButton, True
    ' Выставляем основные настройки
    frMain.ZOrder 0
    lvOptions.ItemSelected(1) = True
    DoEvents
    ucColorButton.Locked = True
    ucFontButton.Locked = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Выгружаем из памяти форму и другие компоненты
    lvOS.Destroy
    Set lvOS = Nothing
    lvOptions.Destroy
    Set lvOptions = Nothing
    Set frmOptions = Nothing
End Sub

Private Sub Form_Terminate()

    If Forms.Count = 0 Then
        UnloadApp
    End If
End Sub

Private Sub InitializeObjectProperties()

    ' изменение шрифта и текста
    ChangeButtonProperties
    ucFontButton.FontFlags = ScreenFonts Or InitToLogFontStruct
End Sub

'заполнение списка типами создания резервных копий
Private Sub LoadComboList()

    Dim strFormName As String

    strFormName = CStr(frmMain.Name)
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

    'CMBTYPEBACKUP
End Sub

'! -----------------------------------------------------------
'!  Функция     :  LoadList_OS
'!  Переменные  :
'!  Описание    :  Построение спиcка ОС
'! -----------------------------------------------------------
Private Sub LoadList_OS()

    Dim i As Long

    Set lvOS = New cListView

    With lvOS
        .Create frOS.hwnd, LVS_REPORT Or LVS_AUTOARRANGE, 10, 29, 550, 180, , WS_EX_STATICEDGE
        .SetStyleEx LVS_EX_FLATSB Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES
        .AddColumn 1, strTableOSHeader1, 150
        .AddColumn 2, strTableOSHeader2, 50
        .AddColumn 3, strTableOSHeader3, 300

        For i = 0 To OSCount - 1
            .AddItem arrOSList(i, 0), , i
            .ItemText(1, i) = arrOSList(i, 1)
            .ItemText(2, i) = arrOSList(i, 2)
            .ItemText(3, i) = arrOSList(i, 3)
        Next
        .AutoArrange = True
    End With

    LastIdOS = OSCount
    '
    lvOS_ReSize
End Sub

Private Sub LoadListCombo(cmbName As ComboBox, strImagePath As String)

    Dim strListFolderTemp() As String
    Dim i                   As Integer

    strListFolderTemp = GetAllFolderInFolder(strImagePath)

    With cmbName
        .Clear

        For i = LBound(strListFolderTemp) To UBound(strListFolderTemp)
            .AddItem strListFolderTemp(i), i
        Next
    End With
End Sub

Private Sub LoadStartMode()

    Dim strFormName As String

    strFormName = CStr(frmMain.Name)
    optGrp1.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "optGrp1", optGrp1.Caption)
    optGrp2.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "optGrp2", optGrp2.Caption)
    optGrp3.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "optGrp3", optGrp3.Caption)
    optGrp4.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "optGrp4", optGrp4.Caption)
    chkHideOther.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "chkHideOther", chkHideOther.Caption)
    chkCheckAll.Caption = LocaliseString(strPCLangCurrentPath, strFormName, "chkCheckAll", chkCheckAll.Caption)
    ' Режим выделения при старте
    SelectStartMode
End Sub

Private Sub Localise(strPathFile As String)

    Dim strFormName     As String
    Dim strFormNameMain As String

    strFormName = CStr(Me.Name)
    strFormNameMain = CStr(frmMain.Name)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    frOptions.Caption = LocaliseString(strPathFile, strFormName, "frOptions", frOptions.Caption)
    ' Описание режимов
    strItemOptions1 = LocaliseString(strPathFile, strFormName, "ItemOptions1", "Основные настройки")
    strItemOptions2 = LocaliseString(strPathFile, strFormName, "ItemOptions2", "Поддерживаемые ОС")
    strItemOptions3 = LocaliseString(strPathFile, strFormName, "ItemOptions3", "Рабочие утилиты")
    strItemOptions4 = LocaliseString(strPathFile, strFormName, "ItemOptions4", "Имя Архива")
    strItemOptions5 = LocaliseString(strPathFile, strFormName, "ItemOptions5", "Оформление программы")
    strItemOptions6 = LocaliseString(strPathFile, strFormName, "ItemOptions6", "Параметры запуска DPInst")
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    frMain.Caption = LocaliseString(strPathFile, strFormName, "frMain", frMain.Caption)
    lblOptionsStart.Caption = LocaliseString(strPathFile, strFormName, "lblOptionsStart", lblOptionsStart.Caption)
    chkUpdate.Caption = LocaliseString(strPathFile, strFormName, "chkUpdate", chkUpdate.Caption)
    chkUpdateBeta.Caption = LocaliseString(strPathFile, strFormName, "chkUpdateBeta", chkUpdateBeta.Caption)
    chkHideOtherProcess.Caption = LocaliseString(strPathFile, strFormName, "chkHideOtherProcess", chkHideOtherProcess.Caption)
    lblOptionsTemp.Caption = LocaliseString(strPathFile, strFormName, "lblOptionsTemp", lblOptionsTemp.Caption)
    chkTempPath.Caption = LocaliseString(strPathFile, strFormName, "chkTempPath", chkTempPath.Caption)
    chkRemoveTemp.Caption = LocaliseString(strPathFile, strFormName, "chkRemoveTemp", chkRemoveTemp.Caption)
    lblDebug.Caption = LocaliseString(strPathFile, strFormName, "lblDebug", lblDebug.Caption)
    chkDebug.Caption = LocaliseString(strPathFile, strFormName, "chkDebug", chkDebug.Caption)
    chkRemoveHistory.Caption = LocaliseString(strPathFile, strFormName, "chkRemoveHistory", chkRemoveHistory.Caption)
    lblRezim.Caption = LocaliseString(strPathFile, strFormName, "lblRezim", lblRezim.Caption)
    lblDebugLogPath.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogPath", lblDebugLogPath.Caption)
    frMainTools.Caption = LocaliseString(strPathFile, strFormName, "frMainTools", frMainTools.Caption)
    cmdPathDefault.Caption = LocaliseString(strPathFile, strFormName, "cmdPathDefault", cmdPathDefault.Caption)
    frOS.Caption = LocaliseString(strPathFile, strFormName, "frOS", frOS.Caption)
    cmdAddOS.Caption = LocaliseString(strPathFile, strFormName, "cmdAddOS", cmdAddOS.Caption)
    cmdEditOS.Caption = LocaliseString(strPathFile, strFormName, "cmdEditOS", cmdEditOS.Caption)
    cmdDelOS.Caption = LocaliseString(strPathFile, strFormName, "cmdDelOS", cmdDelOS.Caption)
    frDesign.Caption = LocaliseString(strPathFile, strFormName, "frDesign", frDesign.Caption)
    lblSizeForm.Caption = LocaliseString(strPathFile, strFormName, "lblSizeForm", lblSizeForm.Caption)
    lblFormHeight.Caption = LocaliseString(strPathFile, strFormName, "lblFormHeight", lblFormHeight.Caption)
    lblFormWidth.Caption = LocaliseString(strPathFile, strFormName, "lblFormWidth", lblFormWidth.Caption)
    chkFormMaximaze.Caption = LocaliseString(strPathFile, strFormName, "chkFormMaximaze", chkFormMaximaze.Caption)
    chkFormSizeSave.Caption = LocaliseString(strPathFile, strFormName, "chkFormSizeSave", chkFormSizeSave.Caption)
    lblSizeButton.Caption = LocaliseString(strPathFile, strFormName, "lblSizeButton", lblSizeButton.Caption)
    lblImageMain.Caption = LocaliseString(strPathFile, strFormName, "lblImageMain", lblImageMain.Caption)
    lblFormWidthMin.Caption = LocaliseString(strPathFile, strFormName, "lblFormWidthMin", lblFormWidthMin.Caption)
    ucColorButton.DialogMsg(ucColor) = LocaliseString(strPathFile, strFormName, "ButtonColor", ucColorButton.DialogMsg(ucColor))
    ucFontButton.DialogMsg(ucFont) = LocaliseString(strPathFile, strFormName, "ButtonFont", ucFontButton.DialogMsg(ucFont))
    frDpInstParam.Caption = LocaliseString(strPathFile, strFormName, "frDpInstParam", frDpInstParam.Caption)
    lblParam.Caption = LocaliseString(strPathFile, strFormName, "lblParam", lblParam.Caption)
    lblDescription.Caption = LocaliseString(strPathFile, strFormName, "lblDescription", lblDescription.Caption)
    lblLegacyMode.Caption = LocaliseString(strPathFile, strFormName, "lblLegacyMode", lblLegacyMode.Caption)
    lblPromptIfDriverIsNotBetter.Caption = LocaliseString(strPathFile, strFormName, "lblPromptIfDriverIsNotBetter", lblPromptIfDriverIsNotBetter.Caption)
    lblForceIfDriverIsNotBetter.Caption = LocaliseString(strPathFile, strFormName, "lblForceIfDriverIsNotBetter", lblForceIfDriverIsNotBetter.Caption)
    lblSuppressAddRemovePrograms.Caption = LocaliseString(strPathFile, strFormName, "lblSuppressAddRemovePrograms", lblSuppressAddRemovePrograms.Caption)
    lblSuppressWizard.Caption = LocaliseString(strPathFile, strFormName, "lblSuppressWizard", lblSuppressWizard.Caption)
    lblQuietInstall.Caption = LocaliseString(strPathFile, strFormName, "lblQuietInstall", lblQuietInstall.Caption)
    lblScanHardware.Caption = LocaliseString(strPathFile, strFormName, "lblScanHardware", lblScanHardware.Caption)
    lblCmdStringDPInst.Caption = LocaliseString(strPathFile, strFormName, "lblCmdStringDPInst", lblCmdStringDPInst.Caption)
    strTableOSHeader1 = LocaliseString(strPathFile, strFormName, "TableOSHeader1", "Версия")
    strTableOSHeader2 = LocaliseString(strPathFile, strFormName, "TableOSHeader2", "x64")
    strTableOSHeader3 = LocaliseString(strPathFile, strFormName, "TableOSHeader3", "Путь")
    chkSilentDll.Caption = LocaliseString(strPathFile, strFormName, "chkSilentDll", chkSilentDll.Caption)
    frArchName.Caption = LocaliseString(strPathFile, strFormName, "frArchName", frArchName.Caption)
    lblArchNameStart.Caption = LocaliseString(strPathFile, strFormName, "lblArchNameStart", lblArchNameStart.Caption)
    optArchNamePC.Caption = LocaliseString(strPathFile, strFormName, "optArchNamePC", optArchNamePC.Caption)
    optArchModelPC.Caption = LocaliseString(strPathFile, strFormName, "optArchModelPC", optArchModelPC.Caption)
    optArchCustom.Caption = LocaliseString(strPathFile, strFormName, "optArchCustom", optArchCustom.Caption)
    lblArchShablon.Caption = LocaliseString(strPathFile, strFormName, "lblArchShablon", lblArchShablon.Caption)
    lblMacrosType.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosType", lblMacrosType.Caption)
    lblMacrosParam.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosParam", lblMacrosParam.Caption)
    lblMacrosDescription.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosDescription", lblMacrosDescription.Caption)
    lblMacrosPCName.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosPCName", lblMacrosPCName.Caption)
    lblMacrosPCModel.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosPCModel", lblMacrosPCModel.Caption)
    lblMacrosOSVer.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosOSVer", lblMacrosOSVer.Caption)
    lblMacrosOSBit.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosOSBit", lblMacrosOSBit.Caption)
    lblMacrosDate.Caption = LocaliseString(strPathFile, strFormName, "lblMacrosDate", lblMacrosDate.Caption)
    lblTypeBackUp.Caption = LocaliseString(strPathFile, strFormName, "lblTypeBackUp", lblTypeBackUp.Caption)
    cmdFutureButton.Caption = LocaliseString(strPathFile, strFormName, "cmdFutureButton", cmdFutureButton.Caption)
    chkButtonDisable.Caption = LocaliseString(strPathFile, strFormName, "chkButtonDisable", chkButtonDisable.Caption)
    lblDebugLogLevel.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogLevel", lblDebugLogLevel.Caption)
End Sub

'! -----------------------------------------------------------
'!  Функция     :  lvOptions_ItemChanged
'!  Переменные  :
'!  Описание    :  При выборе опции происходит отображение соответсвующего окна
'! -----------------------------------------------------------
Private Sub lvOptions_ItemChanged(ByVal iIndex As Long)

    Select Case lvOptions.ItemCaption(iIndex)

        Case strItemOptions1
            frMain.ZOrder 0
            cmbTypeBackUp.SetFocus

        Case strItemOptions3
            frMainTools.ZOrder 0
            ucDPInst86Path.SetFocus

        Case strItemOptions4
            frArchName.ZOrder 0
            txtArchNameShablon.SetFocus

        Case strItemOptions2
            frOS.ZOrder 0

        Case strItemOptions5
            frDesign.ZOrder 0
            cmbImageMain.SetFocus

        Case strItemOptions6
            frDpInstParam.ZOrder 0
            txtCmdStringDPInst.SetFocus

        Case Else
            frOther.ZOrder 0
    End Select
End Sub

'! -----------------------------------------------------------
'!  Функция     :  lvOS_DblClick
'!  Переменные  :
'!  Описание    :  Двойнок клик по элементу списка вызывает форму редактирования
'! -----------------------------------------------------------
Private Sub lvOS_DblClick(ByVal iItem As Long, ByVal Button As MouseButtonConstants)

    TransferOSData
End Sub

'! -----------------------------------------------------------
'!  Функция     :  lvOS_Size
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Private Sub lvOS_ReSize()

    Dim lngLVHeight As Long
    Dim lngLVWidht  As Long
    Dim lngLVTop    As Long
    Dim lngLVLeft   As Long

    lngLVTop = 29
    lngLVLeft = (cmdAddOS.Left / Screen.TwipsPerPixelX)
    lngLVHeight = (cmdAddOS.Top / Screen.TwipsPerPixelY) - lngLVTop - 10
    lngLVWidht = (frOS.Width / Screen.TwipsPerPixelX) - 10 - lngLVLeft

    If Not (lvOS Is Nothing) Then
        lvOS.Move lngLVLeft, lngLVTop, lngLVWidht, lngLVHeight
        lvOS.Refresh
    End If
End Sub

'! -----------------------------------------------------------
'!  Функция     :  ReadOptions
'!  Переменные  :
'!  Описание    :  Читаем настройки программы и заполняем поля
'! -----------------------------------------------------------
Private Sub ReadOptions()

    ' загрузить список ОС
    LoadList_OS
    ' Остальные параметры
    chkUpdate.Checked = mboolUpdateCheck
    chkUpdateBeta.Checked = mboolUpdateCheckBeta
    chkSilentDll.Checked = mboolSilentDLL
    chkRemoveTemp.Checked = mboolDelTmpAfterClose
    chkDebug.Checked = mboolDebugEnable
    chkRemoveHistory.Checked = mboolCleanHistory
    chkFormMaximaze.Checked = mboolStartMaximazed
    chkFormSizeSave.Checked = mboolSaveSizeOnExit
    chkTempPath.Checked = mboolTempPath
    ucTempPath.Path = strAlternativeTempPath
    chkHideOtherProcess.Checked = mboolHideOtherProcess
    ucDebugLogPath.Path = strDebugLogPath
    txtDebugLogLevel.Text = lngDetailMode
    ' Режим при старте
    LoadComboList
    LoadStartMode
    'MainForm
    txtFormHeight.Text = MainFormHeight
    txtFormWidth.Text = MainFormWidth

    'Пути к программам
    If mboolPatnAbs Then
        'Секция DPInst
        ucDPInst86Path.Path = strDPInstExePath86
        ucDPInst64Path.Path = strDPInstExePath64
        'Секция Arc
        ucArchPath.Path = strArh7zExePATH
        ucArchPathSFX.Path = strArh7zSFXPATH
        ucArchPathSFXConfig.Path = strArh7zSFXConfigPath
        ucArchPathSFXConfigEn.Path = strArh7zSFXConfigPathEn
    Else
        'Секция DPInst
        ucDPInst86Path.Path = Replace$(strDPInstExePath86, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDPInst64Path.Path = Replace$(strDPInstExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        'Секция Arc
        ucArchPath.Path = Replace$(strArh7zExePATH, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFX.Path = Replace$(strArh7zSFXPATH, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFXConfig.Path = Replace$(strArh7zSFXConfigPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFXConfigEn.Path = Replace$(strArh7zSFXConfigPathEn, strAppPathBackSL, vbNullString, , , vbTextCompare)
    End If

    ' Настройки DpInst
    chkLegacyMode.Checked = mboolDpInstLegacyMode
    chkPromptIfDriverIsNotBetter.Checked = mboolDpInstPromptIfDriverIsNotBetter
    chkForceIfDriverIsNotBetter.Checked = mboolDpInstForceIfDriverIsNotBetter
    chkSuppressAddRemovePrograms.Checked = mboolDpInstSuppressAddRemovePrograms
    chkSuppressWizard.Checked = mboolDpInstSuppressWizard
    chkQuietInstall.Checked = mboolDpInstQuietInstall
    chkScanHardware.Checked = mboolDpInstScanHardware
    ' Другие настройки
    'txtCmdStringDPInst = CollectCmdString
    ' Загрузка списка скинов
    LoadListCombo cmbImageMain, strPathImageMain
    cmbImageMain.Text = strImageMainName
    ' изменение активности элементов
    DebugCtlEnable chkDebug.Checked
    TempCtlEnable chkTempPath.Checked
    UpdateCtlEnable chkUpdate.Checked
    ' Имя архива при старте
    SelectStartArchName
    txtArchNameShablon.Text = strArchNameCustom
    ' Инициализация параметров для изменения шрифта и цвета
    InitializeObjectProperties
End Sub

'! -----------------------------------------------------------
'!  Функция     :  SaveOptions
'!  Переменные  :
'!  Описание    :  Сохранение настроек в ини-файл
'! -----------------------------------------------------------
Private Sub SaveOptions()

    Dim miRezim       As Long
    Dim miArchName    As Long
    Dim cnt           As Long
    Dim OSCountNew    As Long
    Dim strSysIniTemp As String

    If mboolIsDriveCDRoom And Not mboolLoadIniTmpAfterRestart Then
        If strSysIni <> strWorkTemp & "\DriversBackuper.ini" Then
            MsgBox strMessages(38), vbInformation + vbApplicationModal, strProductName
            Exit Sub
        End If

    ElseIf mboolIsDriveCDRoom And mboolLoadIniTmpAfterRestart Then
        strSysIniTemp = strWinTemp & "Settings_DBS_TMP.ini"
        SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP", True
        SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP_PATH", strSysIniTemp
    Else
        strSysIniTemp = strSysIni
    End If

    '**************************************************
    '***************** Запись настроек ****************
    '**************************************************
    ' Секция MAIN
    'Удаление TEMP при выходе
    IniWriteStrPrivate "Main", "DelTmpAfterClose", CStr(Abs(chkRemoveTemp.Checked)), strSysIniTemp
    ' Автообновление
    IniWriteStrPrivate "Main", "UpdateCheck", CStr(Abs(chkUpdate.Checked)), strSysIniTemp
    ' Автообновление Beta
    IniWriteStrPrivate "Main", "UpdateCheckBeta", CStr(Abs(chkUpdateBeta.Checked)), strSysIniTemp
    ' Режим запуска
    IniWriteStrPrivate "Main", "CheckAllGroup", CStr(Abs(chkCheckAll.Checked)), strSysIniTemp
    IniWriteStrPrivate "Main", "ListOnlyGroup", CStr(Abs(chkHideOther.Checked)), strSysIniTemp

    If optGrp1.Checked Then
        miRezim = 1
    ElseIf optGrp2.Checked Then
        miRezim = 2
    ElseIf optGrp3.Checked Then
        miRezim = 3
    Else
        miRezim = 4
    End If

    IniWriteStrPrivate "Main", "StartMode", CStr(miRezim), strSysIniTemp
    'IniWriteStrPrivate "Main", "EULAAgree", CStr(Abs(mboolEULAAgree)), strSysIniTemp
    IniWriteStrPrivate "Main", "HideOtherProcess", CStr(Abs(chkHideOtherProcess.Checked)), strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTemp", CStr(Abs(chkTempPath.Checked)), strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTempPath", ucTempPath.Path, strSysIniTemp
    IniWriteStrPrivate "Main", "IconMainSkin", cmbImageMain.Text, strSysIniTemp
    IniWriteStrPrivate "Main", "SilentDLL", CStr(Abs(chkSilentDll.Checked)), strSysIniTemp
    IniWriteStrPrivate "Main", "ArchMode", CStr(cmbTypeBackUp.ListIndex), strSysIni

    If mboolLoadIniTmpAfterRestart Then
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", 1, strSysIniTemp
    End If

    IniWriteStrPrivate "Main", "DisableDEP", CStr(Abs(mboolDisableDEP)), strSysIniTemp
    ' Секция Debug
    IniWriteStrPrivate "Debug", "DebugEnable", CStr(Abs(chkDebug.Checked)), strSysIniTemp
    ' Очистка истории:
    IniWriteStrPrivate "Debug", "CleenHistory", CStr(Abs(chkRemoveHistory.Checked)), strSysIniTemp
    ' Путь до лог-файла
    IniWriteStrPrivate "Debug", "DebugLogPath", ucDebugLogPath.Path, strSysIniTemp
    IniWriteStrPrivate "Debug", "Detailmode", CStr(txtDebugLogLevel.Text), strSysIniTemp
    'Секция DPInst
    IniWriteStrPrivate "DPInst", "PathExe", ucDPInst86Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PathExe64", ucDPInst64Path.Path, strSysIniTemp
    'IniWriteStrPrivate "DPInst", "LegacyMode", CStr(Abs(chkLegacyMode.Checked)), strSysIniTemp
    'IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", CStr(Abs(chkPromptIfDriverIsNotBetter.Checked)), strSysIniTemp
    'IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", CStr(Abs(chkForceIfDriverIsNotBetter.Checked)), strSysIniTemp
    'IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", CStr(Abs(chkSuppressAddRemovePrograms.Checked)), strSysIniTemp
    'IniWriteStrPrivate "DPInst", "SuppressWizard", CStr(Abs(chkSuppressWizard.Checked)), strSysIniTemp
    'IniWriteStrPrivate "DPInst", "QuietInstall", CStr(Abs(chkQuietInstall.Checked)), strSysIniTemp
    'IniWriteStrPrivate "DPInst", "ScanHardware", CStr(Abs(chkScanHardware.Checked)), strSysIniTemp
    'Секция Arc
    IniWriteStrPrivate "Arc", "PathExe", ucArchPath.Path, strSysIniTemp
    IniWriteStrPrivate "Arc", "CompressParam1", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 *.ini -ir!*.inf", strSysIni
    IniWriteStrPrivate "Arc", "CompressParam2", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 -xr!*.inf -x!*.ini", strSysIni
    IniWriteStrPrivate "Arc", "PathSFX", ucArchPathSFX.Path, strSysIni
    IniWriteStrPrivate "Arc", "PathSFXConfig", ucArchPathSFXConfig.Path, strSysIni
    IniWriteStrPrivate "Arc", "PathSFXConfigEn", ucArchPathSFXConfigEn.Path, strSysIni

    '[ARCName]
    If optArchNamePC.Checked Then
        miArchName = 1
    ElseIf optArchModelPC.Checked Then
        miArchName = 2
    Else
        miArchName = 0
    End If

    IniWriteStrPrivate "ARCName", "StartMode", miArchName, strSysIni
    IniWriteStrPrivate "ARCName", "CustomName", txtArchNameShablon, strSysIni
    'Секция OS
    OSCountNew = lvOS.Count
    IniWriteStrPrivate "OS", "OSCount", CStr(OSCountNew), strSysIniTemp

    'Заполяем в цикле подсекции ОС
    For cnt = 1 To OSCountNew

        'Секция OS_N
        With lvOS
            IniWriteStrPrivate "OS_" & cnt, "Ver", .ItemCaption(cnt - 1), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "is64bit", .SubItemCaption(cnt - 1, 1), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "drpFolder", .SubItemCaption(cnt - 1, 2), strSysIniTemp
        End With

    Next
    'Секция MainForm
    IniWriteStrPrivate "MainForm", "Width", txtFormWidth.Text, strSysIniTemp
    IniWriteStrPrivate "MainForm", "Height", txtFormHeight.Text, strSysIniTemp
    IniWriteStrPrivate "MainForm", "StartMaximazed", CStr(Abs(chkFormMaximaze.Checked)), strSysIniTemp
    mboolSaveSizeOnExit = chkFormSizeSave.Checked
    IniWriteStrPrivate "MainForm", "SaveSizeOnExit", CStr(Abs(chkFormSizeSave.Checked)), strSysIniTemp
    IniWriteStrPrivate "MainForm", "HighlightColor", CStr(glHighlightColor), strSysIniTemp
    'Секция Buttons
    IniWriteStrPrivate "Button", "FontName", strDialog_FontName, strSysIniTemp
    IniWriteStrPrivate "Button", "FontSize", CStr(miDialog_FontSize), strSysIniTemp
    IniWriteStrPrivate "Button", "FontUnderline", CStr(Abs(mboolDialog_Underline)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontStrikethru", CStr(Abs(mboolDialog_Strikethru)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontItalic", CStr(Abs(mboolDialog_Italic)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontBold", CStr(Abs(mboolDialog_Bold)), strSysIniTemp
    IniWriteStrPrivate "Button", "FontColor", CStr(cmdFutureButton.ForeColor), strSysIniTemp
    ' Приводим Ini файл к читабельному виду
    NormIniFile strSysIniTemp
End Sub

' Режим при старте
Private Sub SelectStartArchName()

    Select Case lngArchNameMode

        Case 0
            optArchCustom.ClearChecks
            optArchCustom.Checked = True

            'optArchCustom_Click
        Case 1
            optArchNamePC.ClearChecks
            optArchNamePC.Checked = True

            'optArchNamePC_Click
        Case 2
            optArchModelPC.ClearChecks
            optArchModelPC.Checked = True

            'optArchModelPC_Click
        Case Else
            optArchCustom.ClearChecks
            optArchCustom.Checked = True
            'optArchCustom_Click
    End Select
End Sub

' Режим при старте
Private Sub SelectStartMode()

    Select Case miStartMode

        Case 1
            optGrp1.ClearChecks
            optGrp1.Checked = True

        Case 2
            optGrp2.ClearChecks
            optGrp2.Checked = True

        Case 3
            optGrp3.ClearChecks
            optGrp3.Checked = True

        Case 4
            optGrp4.ClearChecks
            optGrp4.Checked = True
    End Select
End Sub

Private Sub TempCtlEnable(ByVal mboolEnable As Boolean)

    ucTempPath.Enabled = mboolEnable
End Sub

'! -----------------------------------------------------------
'!  Функция     :  TransferOSData
'!  Переменные  :
'!  Описание    :  Передача параметров ОС из спика в форму редактирования
'! -----------------------------------------------------------
Private Sub TransferOSData()

    Dim i As Long

    With lvOS
        i = .SelectedItem

        If i = -1 Then
            Exit Sub
        End If

        frmOSEdit.txtOSVer.Text = .ItemCaption(i)
        frmOSEdit.ucPathDRP.Path = .SubItemCaption(i, 2)
        frmOSEdit.chk64bit.Checked = CBool(.SubItemCaption(i, 1))
    End With

    'LVOS
    frmOSEdit.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  tvOptionsLoad
'!  Переменные  :
'!  Описание    :  Построение дерева настроек
'! -----------------------------------------------------------
Private Sub tvOptionsLoad()

    Set lvOptions = New cListView

    With lvOptions
        .Create frOptions.hwnd, LVS_LIST Or LVS_SINGLESEL Or LVS_SHOWSELALWAYS, 5, 29, 190, 198, , WS_EX_STATICEDGE
        .SetStyleEx LVS_EX_ONECLICKACTIVATE Or LVS_EX_UNDERLINEHOT
        .AddItem strItemOptions1, , 0
        .AddItem strItemOptions2, , 1
        .AddItem strItemOptions3, , 2
        .AddItem strItemOptions4, , 3
        .AddItem strItemOptions5, , 4
        '.AddItem strItemOptions6, , 4
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_MAIN", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_OSLIST", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_TOOLS_MAIN", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_ARCHNAME", strPathImageMainWork)
        .ImgLst_AddIcon LoadIconImageFromPath("OPT_DESIGN", strPathImageMainWork)
        '.ImgLst_AddIcon LoadIconImageFromPath("OPT_DPINST", strPathImageMainWork)
    End With

    'LVOPTIONS
End Sub

Private Sub txtArchNameShablon_GotFocus()

    HighlightActiveControl Me, txtArchNameShablon, True
End Sub

Private Sub txtArchNameShablon_LostFocus()

    HighlightActiveControl Me, txtArchNameShablon, False
End Sub

Private Sub txtCmdStringDPInst_GotFocus()

    HighlightActiveControl Me, txtCmdStringDPInst, True
End Sub

Private Sub txtCmdStringDPInst_LostFocus()

    HighlightActiveControl Me, txtCmdStringDPInst, False
End Sub

Private Sub txtMacrosDate_DblClick()

    txtMacrosDate.SelStart = 0
    txtMacrosDate.SelLength = Len(txtMacrosDate.Text)
End Sub

Private Sub txtMacrosOSBit_DblClick()

    txtMacrosOSBIT.SelStart = 0
    txtMacrosOSBIT.SelLength = Len(txtMacrosOSBIT.Text)
End Sub

Private Sub txtMacrosOSVer_DblClick()

    txtMacrosOSVER.SelStart = 0
    txtMacrosOSVER.SelLength = Len(txtMacrosOSVER.Text)
End Sub

Private Sub txtMacrosPCModel_DblClick()

    txtMacrosPCModel.SelStart = 0
    txtMacrosPCModel.SelLength = Len(txtMacrosPCModel.Text)
End Sub

Private Sub txtMacrosPCName_DblClick()

    txtMacrosPCName.SelStart = 0
    txtMacrosPCName.SelLength = Len(txtMacrosPCName.Text)
End Sub

Private Sub ucArchPath_Click()

    Dim strTempPath As String

    If ucArchPath.FileCount > 0 Then
        strTempPath = ucArchPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucArchPath.Path = strTempPath
    End If
End Sub

Private Sub ucArchPath_GotFocus()

    HighlightActiveControl Me, ucArchPath, True
End Sub

Private Sub ucArchPath_LostFocus()

    HighlightActiveControl Me, ucArchPath, False
End Sub

Private Sub ucArchPathSFX_Click()

    Dim strTempPath As String

    If ucArchPathSFX.FileCount > 0 Then
        strTempPath = ucArchPathSFX.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucArchPathSFX.Path = strTempPath
    End If
End Sub

Private Sub ucArchPathSFX_GotFocus()

    HighlightActiveControl Me, ucArchPathSFX, True
End Sub

Private Sub ucArchPathSFX_LostFocus()

    HighlightActiveControl Me, ucArchPathSFX, False
End Sub

Private Sub ucArchPathSFXConfig_Click()

    Dim strTempPath As String

    If ucArchPathSFXConfig.FileCount > 0 Then
        strTempPath = ucArchPathSFXConfig.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucArchPathSFXConfig.Path = strTempPath
    End If
End Sub

Private Sub ucArchPathSFXConfig_GotFocus()

    HighlightActiveControl Me, ucArchPathSFXConfig, True
End Sub

Private Sub ucArchPathSFXConfig_LostFocus()

    HighlightActiveControl Me, ucArchPathSFXConfig, False
End Sub

Private Sub ucArchPathSFXConfigEn_Click()

    Dim strTempPath As String

    If ucArchPathSFXConfigEn.FileCount > 0 Then
        strTempPath = ucArchPathSFXConfigEn.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucArchPathSFXConfigEn.Path = strTempPath
    End If
End Sub

Private Sub ucArchPathSFXConfigEn_GotFocus()

    HighlightActiveControl Me, ucArchPathSFXConfigEn, True
End Sub

Private Sub ucArchPathSFXConfigEn_LostFocus()

    HighlightActiveControl Me, ucArchPathSFXConfigEn, False
End Sub

Private Sub ucColorButton_Click()

    lngDialog_Color = ucColorButton.Color
    SetButtonProperties , cmdFutureButton, True
End Sub

Private Sub ucColorButton_GotFocus()

    HighlightActiveControl Me, ucColorButton, True
End Sub

Private Sub ucColorButton_LostFocus()

    HighlightActiveControl Me, ucColorButton, False
End Sub

'! -----------------------------------------------------------
'!  Функция     :  ucDebugLogPath_Click
'!  Переменные  :
'!  Описание    :  выбор каталога или файла
'! -----------------------------------------------------------
Private Sub ucDebugLogPath_Click()

    Dim strTempPath As String

    If ucDebugLogPath.FileCount > 0 Then
        strTempPath = ucDebugLogPath.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucDebugLogPath.Path = strTempPath
    End If
End Sub

Private Sub ucDebugLogPath_GotFocus()

    HighlightActiveControl Me, ucDebugLogPath, True
End Sub

Private Sub ucDebugLogPath_LostFocus()

    HighlightActiveControl Me, ucDebugLogPath, False
End Sub

'! -----------------------------------------------------------
'!  Функция     :  ucDPInst64Path_Click
'!  Переменные  :
'!  Описание    :  выбор каталога или файла
'! -----------------------------------------------------------
Private Sub ucDPInst64Path_Click()

    Dim strTempPath As String

    If ucDPInst64Path.FileCount > 0 Then
        strTempPath = ucDPInst64Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucDPInst64Path.Path = strTempPath
    End If
End Sub

Private Sub ucDPInst64Path_GotFocus()

    HighlightActiveControl Me, ucDPInst64Path, True
End Sub

Private Sub ucDPInst64Path_LostFocus()

    HighlightActiveControl Me, ucDPInst64Path, False
End Sub

'! -----------------------------------------------------------
'!  Функция     :  ucDPInst86Path_Click
'!  Переменные  :
'!  Описание    :  выбор каталога или файла
'! -----------------------------------------------------------
Private Sub ucDPInst86Path_Click()

    Dim strTempPath As String

    If ucDPInst86Path.FileCount > 0 Then
        strTempPath = ucDPInst86Path.FileName

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucDPInst86Path.Path = strTempPath
    End If
End Sub

Private Sub ucDPInst86Path_GotFocus()

    HighlightActiveControl Me, ucDPInst86Path, True
End Sub

Private Sub ucDPInst86Path_LostFocus()

    HighlightActiveControl Me, ucDPInst86Path, False
End Sub

'! -----------------------------------------------------------
'!  Функция     :  ucFontButton_Click
'!  Переменные  :
'!  Описание    :  выбор шрифта кнопки
'! -----------------------------------------------------------
Private Sub ucFontButton_Click()

    Dim NewFontButton As StdFont

    Set NewFontButton = ucFontButton.Font

    If Not NewFontButton Is Nothing Then
        strDialog_FontName = NewFontButton.Name
        miDialog_FontSize = NewFontButton.Size
        mboolDialog_Underline = NewFontButton.Underline
        mboolDialog_Strikethru = NewFontButton.Strikethrough
        mboolDialog_Bold = NewFontButton.Bold
        mboolDialog_Italic = NewFontButton.Italic
        'lngDialog_Language = NewFontButton.Charset
        'lngDialog_Color = ucFontButton.Color
        'cmdFutureButton.Refresh
        'cmdFutureButton.Font.Charset = NewFont.Charset
        'cmdFutureButton.Font.Weight = NewFont.Weight
    End If

    SetButtonProperties , cmdFutureButton, True
End Sub

Private Sub ucFontButton_GotFocus()

    HighlightActiveControl Me, ucFontButton, True
End Sub

Private Sub ucFontButton_LostFocus()

    HighlightActiveControl Me, ucFontButton, False
End Sub

'! -----------------------------------------------------------
'!  Функция     :  ucTempPath_Click
'!  Переменные  :
'!  Описание    :  выбор каталога или файла
'! -----------------------------------------------------------
Private Sub ucTempPath_Click()

    Dim strTempPath As String

    If ucTempPath.FileCount > 0 Then
        strTempPath = ucTempPath.Path

        If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
            strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
        End If
    End If

    If LenB(strTempPath) > 0 Then
        ucTempPath.Path = strTempPath
    End If
End Sub

Private Sub ucTempPath_GotFocus()

    HighlightActiveControl Me, ucTempPath, True
End Sub

Private Sub ucTempPath_LostFocus()

    HighlightActiveControl Me, ucTempPath, False
End Sub

Private Sub UpdateCtlEnable(ByVal mboolEnable As Boolean)

    chkUpdateBeta.Enabled = mboolEnable
End Sub
