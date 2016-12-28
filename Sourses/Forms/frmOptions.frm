VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки программы"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   13725
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   8145
   ScaleWidth      =   13725
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.ctlJCFrames frMain 
      Height          =   5300
      Left            =   3105
      Top             =   25
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Основные настройки программы"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ComboBoxW cmbTypeBackUp 
         Height          =   315
         Left            =   480
         TabIndex        =   4
         Top             =   2760
         Width           =   4815
         _ExtentX        =   8493
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
      Begin prjDIADBS.CheckBoxW chkRemoveTemp 
         Height          =   210
         Left            =   495
         TabIndex        =   5
         Top             =   3750
         Width           =   7920
         _ExtentX        =   8281
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":000C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdate 
         Height          =   210
         Left            =   495
         TabIndex        =   6
         Top             =   660
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0084
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkHideOtherProcess 
         Height          =   210
         Left            =   495
         TabIndex        =   7
         Top             =   1200
         Width           =   7920
         _ExtentX        =   6350
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":00E0
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkTempPath 
         Height          =   210
         Left            =   495
         TabIndex        =   8
         Top             =   3450
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0146
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdateBeta 
         Height          =   210
         Left            =   3630
         TabIndex        =   9
         Top             =   660
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0196
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSilentDll 
         Height          =   210
         Left            =   495
         TabIndex        =   10
         Top             =   930
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":020C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucTempPath 
         Height          =   315
         Left            =   3840
         TabIndex        =   11
         Top             =   3390
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.OptionButtonW optGrp1 
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
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
         Caption         =   "frmOptions.frx":02A8
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optGrp2 
         Height          =   255
         Left            =   2085
         TabIndex        =   13
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
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
         Caption         =   "frmOptions.frx":02DA
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optGrp3 
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   2100
         Width           =   1500
         _ExtentX        =   2646
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
         Caption         =   "frmOptions.frx":0300
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optGrp4 
         Height          =   255
         Left            =   2085
         TabIndex        =   15
         Top             =   2100
         Width           =   1500
         _ExtentX        =   2646
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
         Caption         =   "frmOptions.frx":0326
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkHideOther 
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   1800
         Width           =   4575
         _ExtentX        =   8070
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
         Caption         =   "frmOptions.frx":0358
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkCheckAll 
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   2040
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":03BE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblOptionsTemp 
         Height          =   285
         Left            =   240
         TabIndex        =   51
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Работа с временными файлами"
      End
      Begin prjDIADBS.LabelW lblOptionsStart 
         Height          =   285
         Left            =   240
         TabIndex        =   52
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Действия при запуске программы"
      End
      Begin prjDIADBS.LabelW lblRezim 
         Height          =   285
         Left            =   240
         TabIndex        =   53
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Режим работы фильтра по умолчанию"
      End
      Begin prjDIADBS.LabelW lblTypeBackUp 
         Height          =   225
         Left            =   240
         TabIndex        =   54
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Режим создания резервных копий по умолчанию"
         AutoSize        =   -1  'True
      End
   End
   Begin prjDIADBS.ctlJCFrames frOptions 
      Height          =   5300
      Left            =   50
      Top             =   25
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Настройки"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlJCbutton cmdOK 
         Height          =   645
         Left            =   75
         TabIndex        =   0
         Top             =   3735
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Сохранить изменения и выйти"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdExit 
         Default         =   -1  'True
         Height          =   645
         Left            =   75
         TabIndex        =   1
         Top             =   4515
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Выход без сохранения"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ListView lvOptions 
         Height          =   3195
         Left            =   120
         TabIndex        =   87
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         View            =   2
         Arrange         =   3
         LabelEdit       =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         ClickableColumnHeaders=   0   'False
         TrackSizeColumnHeaders=   0   'False
         ResizableColumnHeaders=   0   'False
      End
   End
   Begin prjDIADBS.ctlJCFrames frMainTools 
      Height          =   5280
      Left            =   3375
      Top             =   375
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Расположение основных утилит (Tools)"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlUcPickBox ucDPInst86Path 
         Height          =   315
         Left            =   2535
         TabIndex        =   2
         Top             =   510
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlUcPickBox ucDPInst64Path 
         Height          =   315
         Left            =   2535
         TabIndex        =   3
         Top             =   930
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPath 
         Height          =   315
         Left            =   2535
         TabIndex        =   18
         Top             =   1350
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlJCbutton cmdPathDefault 
         Height          =   495
         Left            =   4815
         TabIndex        =   19
         Top             =   3210
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   10
         BackColor       =   16765357
         Caption         =   "Сбросить настройки расположения утилит"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPathSFX 
         Height          =   315
         Left            =   2535
         TabIndex        =   29
         Top             =   1770
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPathSFXConfig 
         Height          =   315
         Left            =   2535
         TabIndex        =   30
         Top             =   2250
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPathSFXConfigEn 
         Height          =   315
         Left            =   2535
         TabIndex        =   39
         Top             =   2730
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   3
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
      End
      Begin prjDIADBS.LabelW lblArcSFXConfigEn 
         Height          =   255
         Left            =   150
         TabIndex        =   62
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
         BackStyle       =   0
         Caption         =   "7za-SFXConfig (English)"
      End
      Begin prjDIADBS.LabelW lblArcSFXConfig 
         Height          =   255
         Left            =   150
         TabIndex        =   63
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
         BackStyle       =   0
         Caption         =   "7za-SFXConfig"
      End
      Begin prjDIADBS.LabelW lblArc 
         Height          =   255
         Left            =   150
         TabIndex        =   64
         Top             =   1350
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
         BackStyle       =   0
         Caption         =   "7za"
      End
      Begin prjDIADBS.LabelW lblDPInst64 
         Height          =   255
         Left            =   150
         TabIndex        =   65
         Top             =   930
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
         BackStyle       =   0
         Caption         =   "DPInst.exe (64-bit)"
      End
      Begin prjDIADBS.LabelW lblDPInst86 
         Height          =   255
         Left            =   150
         TabIndex        =   66
         Top             =   510
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
         BackStyle       =   0
         Caption         =   "DPInst.exe (32-bit)"
      End
      Begin prjDIADBS.LabelW lblArcSFX 
         Height          =   255
         Left            =   150
         TabIndex        =   67
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
         BackStyle       =   0
         Caption         =   "7za-sfxModule"
      End
   End
   Begin prjDIADBS.ctlJCFrames frArchName 
      Height          =   5280
      Left            =   3645
      Top             =   705
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Имя архива"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.TextBoxW txtArchNameShablon 
         Height          =   330
         Left            =   480
         TabIndex        =   50
         Top             =   2205
         Width           =   7635
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
      Begin prjDIADBS.TextBoxW txtMacrosPCName 
         Height          =   255
         Left            =   480
         TabIndex        =   49
         Top             =   3285
         Width           =   1500
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
         BackColor       =   -2147483633
         BorderStyle     =   0
         Text            =   "frmOptions.frx":041A
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosPCModel 
         Height          =   255
         Left            =   480
         TabIndex        =   48
         Top             =   3645
         Width           =   1500
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
         BackColor       =   -2147483633
         BorderStyle     =   0
         Text            =   "frmOptions.frx":044A
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSVER 
         Height          =   255
         Left            =   480
         TabIndex        =   47
         Top             =   4005
         Width           =   1500
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
         BackColor       =   -2147483633
         BorderStyle     =   0
         Text            =   "frmOptions.frx":047C
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSBIT 
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   4365
         Width           =   1500
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
         BackColor       =   -2147483633
         BorderStyle     =   0
         Text            =   "frmOptions.frx":04AA
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosDate 
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   4725
         Width           =   1500
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
         BackColor       =   -2147483633
         BorderStyle     =   0
         Text            =   "frmOptions.frx":04D8
         Locked          =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optArchModelPC 
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   1125
         Width           =   7635
         _ExtentX        =   13467
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
         Caption         =   "frmOptions.frx":0504
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optArchNamePC 
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   765
         Width           =   7635
         _ExtentX        =   13467
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
         Caption         =   "frmOptions.frx":0546
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optArchCustom 
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   1485
         Width           =   7635
         _ExtentX        =   13467
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
         Caption         =   "frmOptions.frx":0582
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblMacrosDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   68
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
         BackStyle       =   0
         Caption         =   "Дата создания резервной копии"
      End
      Begin prjDIADBS.LabelW lblMacrosOSBit 
         Height          =   375
         Left            =   2400
         TabIndex        =   69
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
         BackStyle       =   0
         Caption         =   "Архитектура операционной системы, в виде x32[64]"
      End
      Begin prjDIADBS.LabelW lblMacrosOSVer 
         Height          =   375
         Left            =   2400
         TabIndex        =   70
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
         BackStyle       =   0
         Caption         =   "Версия операционной системы в виде wnt5[6]"
      End
      Begin prjDIADBS.LabelW lblMacrosPCModel 
         Height          =   375
         Left            =   2400
         TabIndex        =   71
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
         BackStyle       =   0
         Caption         =   "Модель компьютера/материнской платы"
      End
      Begin prjDIADBS.LabelW lblMacrosParam 
         Height          =   255
         Left            =   480
         TabIndex        =   72
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
         BackStyle       =   0
         Caption         =   "Параметр"
      End
      Begin prjDIADBS.LabelW lblMacrosDescription 
         Height          =   255
         Left            =   2400
         TabIndex        =   73
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
         BackStyle       =   0
         Caption         =   "Описание  параметра"
      End
      Begin prjDIADBS.LabelW lblMacrosPCName 
         Height          =   375
         Left            =   2400
         TabIndex        =   74
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
         BackStyle       =   0
         Caption         =   "Краткое имя компьютера, без доменного суффикса"
      End
      Begin prjDIADBS.LabelW lblMacrosType 
         Height          =   285
         Left            =   480
         TabIndex        =   75
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
         ForeColor       =   -2147483635
         Caption         =   "Доступные макроподстановки:"
      End
      Begin prjDIADBS.LabelW lblArchShablon 
         Height          =   285
         Left            =   240
         TabIndex        =   76
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
         ForeColor       =   -2147483635
         Caption         =   "Шаблон имени архива"
      End
      Begin prjDIADBS.LabelW lblArchNameStart 
         Height          =   285
         Left            =   240
         TabIndex        =   77
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
         ForeColor       =   -2147483635
         Caption         =   "Имя архива по умолчанию"
      End
   End
   Begin prjDIADBS.ctlJCFrames frOS 
      Height          =   5280
      Left            =   3900
      Top             =   1020
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Поддерживаемые ОС"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ListView lvOS 
         Height          =   3795
         Left            =   120
         TabIndex        =   88
         Top             =   480
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6694
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
      Begin prjDIADBS.ctlJCbutton cmdAddOS 
         Height          =   750
         Left            =   120
         TabIndex        =   41
         Top             =   4400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1323
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   10
         BackColor       =   16765357
         Caption         =   "Добавить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdEditOS 
         Height          =   750
         Left            =   2160
         TabIndex        =   42
         Top             =   4400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1323
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   10
         BackColor       =   16765357
         Caption         =   "Изменить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdDelOS 
         Height          =   750
         Left            =   4200
         TabIndex        =   43
         Top             =   4400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1323
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   10
         BackColor       =   16765357
         Caption         =   "Удалить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
   End
   Begin prjDIADBS.ctlJCFrames frDesign 
      Height          =   5280
      Left            =   4140
      Top             =   1320
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Оформление"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ComboBoxW cmbImageMain 
         Height          =   315
         Left            =   405
         TabIndex        =   32
         Top             =   3075
         Width           =   3000
         _ExtentX        =   5292
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
         Sorted          =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkButtonDisable 
         Height          =   450
         Left            =   5790
         TabIndex        =   33
         Top             =   1935
         Width           =   2400
         _ExtentX        =   4233
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
         Caption         =   "frmOptions.frx":05B6
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkFormMaximaze 
         Height          =   210
         Left            =   3285
         TabIndex        =   34
         Top             =   795
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":062C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtFormHeight 
         Height          =   255
         Left            =   1245
         TabIndex        =   35
         Top             =   795
         Width           =   1575
         _ExtentX        =   2778
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
         Min             =   2000
         Max             =   25000
         Value           =   2000
      End
      Begin prjDIADBS.SpinBox txtFormWidth 
         Height          =   255
         Left            =   1245
         TabIndex        =   36
         Top             =   1140
         Width           =   1575
         _ExtentX        =   2778
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
         Min             =   2000
         Max             =   25000
         Value           =   2000
      End
      Begin prjDIADBS.CheckBoxW chkFormSizeSave 
         Height          =   210
         Left            =   3285
         TabIndex        =   37
         Top             =   1140
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0692
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFutureButton 
         Height          =   510
         Left            =   3390
         TabIndex        =   38
         Top             =   1935
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   900
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   10
         BackColor       =   12244692
         Caption         =   "Твоя будущая кнопка"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
         ColorScheme     =   1
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorButton 
         Height          =   795
         Left            =   240
         TabIndex        =   86
         Top             =   1920
         Width           =   2445
         _ExtentX        =   5027
         _ExtentY        =   1402
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
         BackColor       =   16765357
         Caption         =   "Установить цвет и шрифт текста кнопки"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.LabelW lblFormWidthMin 
         Height          =   930
         Left            =   135
         TabIndex        =   56
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
         ForeColor       =   255
         Caption         =   $"frmOptions.frx":06F6
      End
      Begin prjDIADBS.LabelW lblImageMain 
         Height          =   255
         Left            =   135
         TabIndex        =   57
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Основные картинки"
      End
      Begin prjDIADBS.LabelW lblFormWidth 
         Height          =   210
         Left            =   405
         TabIndex        =   58
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
      End
      Begin prjDIADBS.LabelW lblFormHeight 
         Height          =   210
         Left            =   405
         TabIndex        =   59
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
      End
      Begin prjDIADBS.LabelW lblSizeForm 
         Height          =   255
         Left            =   135
         TabIndex        =   60
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Размеры основного окна"
      End
      Begin prjDIADBS.LabelW lblSizeButton 
         Height          =   255
         Left            =   135
         TabIndex        =   61
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Свойства кнопок"
      End
   End
   Begin prjDIADBS.ctlJCFrames frDpInstParam 
      Height          =   5280
      Left            =   4380
      Top             =   1620
      Width           =   8655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Параметры запуска DPInst"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin VB.CommandButton cmdLegacyMode 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   97
         ToolTipText     =   "More on MSDN..."
         Top             =   660
         Width           =   255
      End
      Begin VB.CommandButton cmdPromptIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2640
         TabIndex        =   100
         ToolTipText     =   "More on MSDN..."
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton cmdForceIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   103
         ToolTipText     =   "More on MSDN..."
         Top             =   1905
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressAddRemovePrograms 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   106
         ToolTipText     =   "More on MSDN..."
         Top             =   2460
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressWizard 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   109
         ToolTipText     =   "More on MSDN..."
         Top             =   2955
         Width           =   255
      End
      Begin VB.CommandButton cmdQuietInstall 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   20
         ToolTipText     =   "More on MSDN..."
         Top             =   3510
         Width           =   255
      End
      Begin VB.CommandButton cmdScanHardware 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   21
         ToolTipText     =   "More on MSDN..."
         Top             =   4005
         Width           =   255
      End
      Begin prjDIADBS.TextBoxW txtCmdStringDPInst 
         Height          =   330
         Left            =   2895
         TabIndex        =   22
         Top             =   4875
         Width           =   5535
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
         Locked          =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkLegacyMode 
         Height          =   210
         Left            =   120
         TabIndex        =   96
         Top             =   660
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":07B1
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkPromptIfDriverIsNotBetter 
         Height          =   210
         Left            =   120
         TabIndex        =   99
         Top             =   1305
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":07E5
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkForceIfDriverIsNotBetter 
         Height          =   210
         Left            =   120
         TabIndex        =   102
         Top             =   1905
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0837
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressAddRemovePrograms 
         CausesValidation=   0   'False
         Height          =   210
         Left            =   120
         TabIndex        =   105
         Top             =   2460
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0887
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressWizard 
         Height          =   210
         Left            =   120
         TabIndex        =   108
         Top             =   2955
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":08D9
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkQuietInstall 
         Height          =   210
         Left            =   120
         TabIndex        =   23
         Top             =   3510
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0915
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkScanHardware 
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   4005
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":094D
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblCmdStringDPInst 
         Height          =   210
         Left            =   135
         TabIndex        =   25
         Top             =   4875
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   370
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
         Caption         =   "Итоговые параметры запуска "
      End
      Begin prjDIADBS.LabelW lblDescription 
         Height          =   255
         Left            =   2865
         TabIndex        =   95
         Top             =   350
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Описание  параметра"
      End
      Begin prjDIADBS.LabelW lblParam 
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   350
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Параметр"
      End
      Begin prjDIADBS.LabelW lblPromptIfDriverIsNotBetter 
         Height          =   570
         Left            =   2925
         TabIndex        =   101
         Top             =   1305
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   1005
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
         Caption         =   "display a dialog box if a new driver is not a better match to a device than a driver that is currently installed on the device"
      End
      Begin prjDIADBS.LabelW lblLegacyMode 
         Height          =   645
         Left            =   2925
         TabIndex        =   98
         Top             =   660
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   1138
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
         Caption         =   "install unsigned drivers and driver packages that have missing files"
      End
      Begin prjDIADBS.LabelW lblForceIfDriverIsNotBetter 
         Height          =   510
         Left            =   2925
         TabIndex        =   104
         Top             =   1905
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   900
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
         Caption         =   "install a driver on a device even if the driver that is currently installed on the device is a better match than the new driver"
      End
      Begin prjDIADBS.LabelW lblSuppressAddRemovePrograms 
         Height          =   450
         Left            =   2925
         TabIndex        =   107
         Top             =   2460
         Width           =   5580
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
         BackStyle       =   0
         Caption         =   "suppress the addition of Add or Remove Programs entries that represent the drivers and driver package"
      End
      Begin prjDIADBS.LabelW lblSuppressWizard 
         Height          =   450
         Left            =   2925
         TabIndex        =   110
         Top             =   2955
         Width           =   5550
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
         BackStyle       =   0
         Caption         =   "configures DPInst to suppress the display of wizard pages and other user messages that DPInst generates."
      End
      Begin prjDIADBS.LabelW lblQuietInstall 
         Height          =   450
         Left            =   2925
         TabIndex        =   26
         Top             =   3510
         Width           =   5550
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
         BackStyle       =   0
         Caption         =   "configures DPInst to suppress the display of wizard pages and most other user messages."
      End
      Begin prjDIADBS.LabelW lblScanHardware 
         Height          =   900
         Left            =   2925
         TabIndex        =   27
         Top             =   4005
         Width           =   5550
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
         BackStyle       =   0
         Caption         =   $"frmOptions.frx":0985
      End
   End
   Begin prjDIADBS.ctlJCFrames frDebug 
      Height          =   5295
      Left            =   4620
      Top             =   1920
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      Caption         =   "Отладочный режим"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.TextBoxW txtDebugLogName 
         Height          =   315
         Left            =   480
         TabIndex        =   28
         Top             =   2520
         Width           =   7815
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
      Begin prjDIADBS.ctlUcPickBox ucDebugLogPath 
         Height          =   315
         Left            =   480
         TabIndex        =   55
         Top             =   1890
         Width           =   7845
         _ExtentX        =   10821
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         DefaultExt      =   ""
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         UseDialogText   =   0   'False
         Locked          =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtDebugLogLevel 
         Height          =   255
         Left            =   7680
         TabIndex        =   78
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
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
         Min             =   1
         Value           =   1
      End
      Begin prjDIADBS.TextBoxW txtMacrosDateDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   111
         Top             =   4875
         Width           =   1500
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
         Text            =   "frmOptions.frx":0A83
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSBITDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   112
         Top             =   4515
         Width           =   1500
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
         Text            =   "frmOptions.frx":0AAF
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosOSVERDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   113
         Top             =   4155
         Width           =   1500
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
         Text            =   "frmOptions.frx":0ADD
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosPCModelDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   114
         Top             =   3795
         Width           =   1500
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
         Text            =   "frmOptions.frx":0B0B
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtMacrosPCNameDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   115
         Top             =   3435
         Width           =   1500
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
         Text            =   "frmOptions.frx":0B3D
         Locked          =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDebug 
         Height          =   210
         Left            =   495
         TabIndex        =   116
         Top             =   720
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0B6D
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDebugLog2AppPath 
         Height          =   210
         Left            =   495
         TabIndex        =   117
         Top             =   1320
         Width           =   7920
         _ExtentX        =   11245
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0BBD
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDebugTime2File 
         Height          =   210
         Left            =   495
         TabIndex        =   118
         Top             =   1020
         Width           =   7920
         _ExtentX        =   11245
         _ExtentY        =   370
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmOptions.frx":0C3D
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblDebugLogLevel 
         Height          =   255
         Left            =   4680
         TabIndex        =   79
         Top             =   720
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "Уровень отладки:"
      End
      Begin prjDIADBS.LabelW lblMacrosDateDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   80
         Top             =   4905
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
         BackStyle       =   0
         Caption         =   "Дата и время создания лог-файла"
      End
      Begin prjDIADBS.LabelW lblMacrosOSBitDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   81
         Top             =   4545
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
         BackStyle       =   0
         Caption         =   "Архитектура операционной системы, в виде x32[64]"
      End
      Begin prjDIADBS.LabelW lblMacrosOSVerDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   82
         Top             =   4185
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
         BackStyle       =   0
         Caption         =   "Версия операционной системы в виде wnt5[6]"
      End
      Begin prjDIADBS.LabelW lblMacrosPCModelDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   83
         Top             =   3825
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
         BackStyle       =   0
         Caption         =   "Модель компьютера/материнской платы"
      End
      Begin prjDIADBS.LabelW lblMacrosParamDebug 
         Height          =   255
         Left            =   480
         TabIndex        =   84
         Top             =   3150
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
         BackStyle       =   0
         Caption         =   "Параметр"
      End
      Begin prjDIADBS.LabelW lblMacrosDescriptionDebug 
         Height          =   255
         Left            =   2400
         TabIndex        =   85
         Top             =   3150
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
         BackStyle       =   0
         Caption         =   "Описание  параметра"
      End
      Begin prjDIADBS.LabelW lblMacrosPCNameDebug 
         Height          =   375
         Left            =   2400
         TabIndex        =   89
         Top             =   3465
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
         BackStyle       =   0
         Caption         =   "Краткое имя компьютера, без доменного суффикса"
      End
      Begin prjDIADBS.LabelW lblMacrosTypeDebug 
         Height          =   285
         Left            =   480
         TabIndex        =   90
         Top             =   2865
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Доступные макроподстановки для имени лог-файла:"
      End
      Begin prjDIADBS.LabelW lblDebugLogPath 
         Height          =   285
         Left            =   480
         TabIndex        =   91
         Top             =   1575
         Width           =   7845
         _ExtentX        =   13838
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Каталог для создания log-файлов:"
      End
      Begin prjDIADBS.LabelW lblDebug 
         Height          =   270
         Left            =   240
         TabIndex        =   92
         Top             =   420
         Width           =   8295
         _ExtentX        =   14631
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Настройки отладочного режима"
      End
      Begin prjDIADBS.LabelW lblDebugLogName 
         Height          =   285
         Left            =   495
         TabIndex        =   93
         Top             =   2225
         Width           =   7845
         _ExtentX        =   13838
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
         ForeColor       =   -2147483635
         BackStyle       =   0
         Caption         =   "Каталог для создания log-файлов:"
      End
   End
   Begin prjDIADBS.ctlJCFrames frOther 
      Height          =   5295
      Left            =   4890
      Top             =   2235
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14016736
      FillColor       =   14016736
      RoundedCorner   =   0   'False
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private strFormName               As String

Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

Private Sub ChangeButtonProperties()

    SetBtnFontProperties cmdFutureButton
    'ucFontButton.FontColor = cmdFutureButton.ForeColor
End Sub

Private Sub chkButtonDisable_Click()

    cmdFutureButton.Enabled = chkButtonDisable.Value
End Sub

Private Sub chkDebug_Click()

    DebugCtlEnable chkDebug.Value
End Sub

Private Sub chkForceIfDriverIsNotBetter_Click()

    mbDpInstForceIfDriverIsNotBetter = chkForceIfDriverIsNotBetter.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkFormMaximaze_Click()

    If chkFormMaximaze.Value Then
        chkFormSizeSave.Value = False
    End If
End Sub

Private Sub chkFormSizeSave_Click()

    If chkFormSizeSave.Value Then
        chkFormMaximaze.Value = False
    End If
End Sub

Private Sub chkHideOther_Click()

    chkCheckAll.Enabled = CBool(chkHideOther.Value)
End Sub

Private Sub chkLegacyMode_Click()

    mbDpInstLegacyMode = chkLegacyMode.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkPromptIfDriverIsNotBetter_Click()

    mbDpInstPromptIfDriverIsNotBetter = chkPromptIfDriverIsNotBetter.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkQuietInstall_Click()

    mbDpInstQuietInstall = chkQuietInstall.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkScanHardware_Click()

    mbDpInstScanHardware = chkScanHardware.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkSuppressAddRemovePrograms_Click()

    mbDpInstSuppressAddRemovePrograms = chkSuppressAddRemovePrograms.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkSuppressWizard_Click()

    mbDpInstSuppressWizard = chkSuppressWizard.Value
    txtCmdStringDPInst = CollectCmdString
End Sub

Private Sub chkTempPath_Click()

    TempCtlEnable chkTempPath.Value
End Sub

Private Sub chkUpdate_Click()

    UpdateCtlEnable chkUpdate.Value
End Sub

Private Sub cmbImageMain_Click()

    If PathExists(strPathImageMain & cmbImageMain.Text) = False Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If
End Sub

Private Sub cmbImageMain_GotFocus()

    HighlightActiveControl Me, cmbImageMain, True
End Sub

Private Sub cmbImageMain_LostFocus()

    If PathExists(strPathImageMain & cmbImageMain.Text) = False Then
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

    mbAddInList = True
    frmOSEdit.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdDelOS_Click
'!  Переменные  :
'!  Описание    :  кнопка удаление ОС
'! -----------------------------------------------------------
Private Sub cmdDelOS_Click()

    Dim ii As Long
    'Dim ii As LvwListItem

    With lvOS

        If .ListItems.count > 0 Then
            'ii = .SelectedItem
            'ii.ListSubItems.Remove
            '.RemoveItem (ii)
            lngLastIdOS = lngLastIdOS - 1
        End If
    End With

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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdFontColorButton_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorButton_Click()

    With frmFontDialog
        .optControl(0).Value = True
        .txtFont.Font.Name = strFontBtn_Name
        .txtFont.Font.Size = miFontBtn_Size
        .txtFont.Font.Bold = mbFontBtn_Bold
        .txtFont.Font.Italic = mbFontBtn_Italic
        .txtFont.Font.Underline = mbFontBtn_Underline
        .txtFont.Font.Charset = lngFont_Charset
        .txtFont.ForeColor = lngFontBtn_Color
        .Show vbModal, Me
    End With

End Sub


Private Sub cmdForceIfDriverIsNotBetter_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = strQuotes & "http://msdn.microsoft.com/en-us/library/ms793551.aspx" & strQuotes
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdLegacyMode_Click()

    Dim cmdString   As String
    Dim nRetShellEx As Boolean

    cmdString = strQuotes & "http://msdn.microsoft.com/en-us/library/ms794322.aspx" & strQuotes
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

'! -----------------------------------------------------------
'!  Функция     :  cmdOK_Click
'!  Переменные  :
'!  Описание    :  Нажатие кнопки ОК. Применение настроек
'! -----------------------------------------------------------
Private Sub cmdOK_Click()

    Dim MsgRet As Long

    If mbIsDriveCDRoom And mbLoadIniTmpAfterRestart Then
        SaveOptions
        MsgRet = MsgBox(strMessages(36), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mbRestartProgram = MsgRet = vbYes
    ElseIf Not FileisReadOnly(strSysIni) Then
        SaveOptions
        MsgRet = MsgBox(strMessages(36), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mbRestartProgram = MsgRet = vbYes
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

    cmdString = strQuotes & "http://msdn.microsoft.com/en-us/library/ms793530.aspx" & strQuotes
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdQuietInstall_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = strQuotes & "http://msdn.microsoft.com/en-us/library/ms794300.aspx" & strQuotes
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdScanHardware_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = strQuotes & "http://msdn.microsoft.com/en-us/library/ms794295.aspx" & strQuotes
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdSuppressAddRemovePrograms_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = strQuotes & "http://msdn.microsoft.com/en-us/library/ms794270.aspx" & strQuotes
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub cmdSuppressWizard_Click()

    Dim cmdString   As String
    Dim nRetShellEx As String

    cmdString = strQuotes & "http://msdn.microsoft.com/en-us/library/ms791062.aspx" & strQuotes
    If mbDebugStandart Then DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    If mbDebugStandart Then DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub DebugCtlEnable(ByVal mbEnable As Boolean)

    chkRemoveHistory.Enabled = mbEnable
    ucDebugLogPath.Enabled = mbEnable
End Sub

Private Sub FontCharsetChange()

    ' Выставляем шрифт
    Me.Font.Name = strFontOtherForm_Name
    Me.Font.Size = lngFontOtherForm_Size
    Me.Font.Charset = lngFont_Charset
    
    frArchName.Font.Charset = lngFont_Charset
    frDesign.Font.Charset = lngFont_Charset
    frDpInstParam.Font.Charset = lngFont_Charset
    frMain.Font.Charset = lngFont_Charset
    frMainTools.Font.Charset = lngFont_Charset
    frOptions.Font.Charset = lngFont_Charset
    frOS.Font.Charset = lngFont_Charset
    frOther.Font.Charset = lngFont_Charset
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
    
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, strFormName, False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
        .Height = 5825
        .Width = 11900
    End With
    
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
    txtFormHeight.Min = lngMainFormHeightMin
    txtFormWidth.Min = lngMainFormWidthMin
    ' Устанавливаем картинки кнопок и убираем описание кнопок
    LoadIconImage2Object cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork
    LoadIconImage2Object cmdAddOS, "BTN_ADD", strPathImageMainWork
    LoadIconImage2Object cmdEditOS, "BTN_EDIT", strPathImageMainWork
    LoadIconImage2Object cmdDelOS, "BTN_DELETE", strPathImageMainWork
    LoadIconImage2Object cmdFontColorButton, "BTN_FONT", strPathImageMainWork
    LoadIconImage2Object cmdFutureButton, "BTN_STARTBACKUP", strPathImageMainWork

    ' Локализация приложения
    If mbMultiLanguage Then
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
    'SetBtnFontProperties cmdChooseFont
    'cmdColorButton.Value = lngDialog_Color
    ' установить опции шрифта и цвета
    SetBtnFontProperties cmdFutureButton
    ' Выставляем основные настройки
    frMain.ZOrder 0
    lvOptions.ItemSelected(1) = True
    DoEvents
    'ucColorButton.Locked = True
    'ucFontButton.Locked = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Выгружаем из памяти форму и другие компоненты
    'lvOS.Destroy
    'Set lvOS = Nothing
    'lvOptions.Destroy
    'Set lvOptions = Nothing
    Set frmOptions = Nothing
End Sub

Private Sub InitializeObjectProperties()

    ' изменение шрифта и текста
    ChangeButtonProperties
    'ucFontButton.FontFlags = ScreenFonts Or InitToLogFontStruct
End Sub

'заполнение списка типами создания резервных копий
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

'! -----------------------------------------------------------
'!  Функция     :  LoadList_OS
'!  Переменные  :
'!  Описание    :  Построение спиcка ОС
'! -----------------------------------------------------------
Private Sub LoadList_OS()

    Dim ii As Long

    Set lvOS = New cListView

    With lvOS
        .Create frOS.hWnd, LVS_REPORT Or LVS_AUTOARRANGE, 10, 29, 550, 180, , WS_EX_STATICEDGE
        .SetStyleEx LVS_EX_FLATSB Or LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES
        .AddColumn 1, strTableOSHeader1, 150
        .AddColumn 2, strTableOSHeader2, 50
        .AddColumn 3, strTableOSHeader3, 300

        For ii = 0 To OSCount - 1
            .AddItem arrOSList(I, 0), , ii
            .ItemText(1, ii) = arrOSList(ii, 1)
            .ItemText(2, ii) = arrOSList(ii, 2)
            .ItemText(3, ii) = arrOSList(ii, 3)
        Next
        .AutoArrange = True
    End With

    lngLastIdOS = OSCount
    lvOS_ReSize
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadListCombo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   cmbName (ComboBox)
'                              strImagePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadListCombo(cmbName As ComboBox, strImagePath As String)

    Dim strListFolderTemp() As String
    Dim ii                  As Integer

    strListFolderTemp = SearchFoldersInRoot(strImagePath, "*")

    With cmbName
        .Clear

        For ii = LBound(strListFolderTemp, 2) To UBound(strListFolderTemp, 2)
            .AddItem strListFolderTemp(1, ii), ii
        Next

    End With

End Sub

Private Sub LoadStartMode()
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

    Dim strFormNameMain As String

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
    cmdFontColorButton.Caption = LocaliseString(strPathFile, strFormName, "cmdFontColorButton", cmdFontColorButton.Caption)
    'ucColorButton.DialogMsg(ucColor) = LocaliseString(StrPathFile, strFormName, "ButtonColor", ucColorButton.DialogMsg(ucColor))
    'ucFontButton.DialogMsg(ucFont) = LocaliseString(StrPathFile, strFormName, "ButtonFont", ucFontButton.DialogMsg(ucFont))
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
'Private Sub lvOS_DblClick(ByVal iItem As Long, ByVal Button As MouseButtonConstants)
'
'    TransferOSData
'End Sub
'
''! -----------------------------------------------------------
''!  Функция     :  lvOS_Size
''!  Переменные  :
''!  Описание    :
''! -----------------------------------------------------------
'Private Sub lvOS_ReSize()
'
'    Dim lngLVHeight As Long
'    Dim lngLVWidht  As Long
'    Dim lngLVTop    As Long
'    Dim lngLVLeft   As Long
'
'    lngLVTop = 29
'    lngLVLeft = (cmdAddOS.Left / Screen.TwipsPerPixelX)
'    lngLVHeight = (cmdAddOS.Top / Screen.TwipsPerPixelY) - lngLVTop - 10
'    lngLVWidht = (frOS.Width / Screen.TwipsPerPixelX) - 10 - lngLVLeft
'
'    If Not (lvOS Is Nothing) Then
'        lvOS.Move lngLVLeft, lngLVTop, lngLVWidht, lngLVHeight
'        lvOS.Refresh
'    End If
'End Sub

'! -----------------------------------------------------------
'!  Функция     :  ReadOptions
'!  Переменные  :
'!  Описание    :  Читаем настройки программы и заполняем поля
'! -----------------------------------------------------------
Private Sub ReadOptions()

    ' загрузить список ОС
    LoadList_OS
    ' Остальные параметры
    chkUpdate.Value = Abs(mbUpdateCheck)
    chkUpdateBeta.Value = Abs(mbUpdateCheckBeta)
    chkSilentDll.Value = Abs(mbSilentDLL)
    chkRemoveTemp.Value = Abs(mbDelTmpAfterClose)
    chkDebug.Value = Abs(mbDebugStandart)
    chkRemoveHistory.Value = Abs(mbCleanHistory)
    chkFormMaximaze.Value = Abs(mbStartMaximazed)
    chkFormSizeSave.Value = Abs(mbSaveSizeOnExit)
    chkTempPath.Value = Abs(mbTempPath)
    ucTempPath.Path = strAlternativeTempPath
    chkHideOtherProcess.Value = Abs(mbHideOtherProcess)
    ucDebugLogPath.Path = strDebugLogPath
    txtDebugLogLevel.Text = lngDetailMode
    ' Режим при старте
    LoadComboList
    LoadStartMode
    'MainForm
    txtFormHeight.Text = lngMainFormHeight
    txtFormWidth.Text = lngMainFormWidth

    'Пути к программам
    If mbPatnAbs Then
        'Секция DPInst
        ucDPInst86Path.Path = strDPInstExePath86
        ucDPInst64Path.Path = strDPInstExePath64
        'Секция Arc
        ucArchPath.Path = strArh7zExePath
        ucArchPathSFX.Path = strArh7zSFXPATH
        ucArchPathSFXConfig.Path = strArh7zSFXConfigPath
        ucArchPathSFXConfigEn.Path = strArh7zSFXConfigPathEn
    Else
        'Секция DPInst
        ucDPInst86Path.Path = Replace$(strDPInstExePath86, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDPInst64Path.Path = Replace$(strDPInstExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        'Секция Arc
        ucArchPath.Path = Replace$(strArh7zExePath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFX.Path = Replace$(strArh7zSFXPATH, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFXConfig.Path = Replace$(strArh7zSFXConfigPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFXConfigEn.Path = Replace$(strArh7zSFXConfigPathEn, strAppPathBackSL, vbNullString, , , vbTextCompare)
    End If

    ' Настройки DpInst
    chkLegacyMode.Value = Abs(mbDpInstLegacyMode)
    chkPromptIfDriverIsNotBetter.Value = Abs(mbDpInstPromptIfDriverIsNotBetter)
    chkForceIfDriverIsNotBetter.Value = Abs(mbDpInstForceIfDriverIsNotBetter)
    chkSuppressAddRemovePrograms.Value = Abs(mbDpInstSuppressAddRemovePrograms)
    chkSuppressWizard.Value = Abs(mbDpInstSuppressWizard)
    chkQuietInstall.Value = Abs(mbDpInstQuietInstall)
    chkScanHardware.Value = Abs(mbDpInstScanHardware)
    ' Другие настройки
    'txtCmdStringDPInst = CollectCmdString
    ' Загрузка списка скинов
    LoadListCombo cmbImageMain, strPathImageMain
    cmbImageMain.Text = strImageMainName
    ' изменение активности элементов
    DebugCtlEnable CBool(chkDebug.Value)
    TempCtlEnable CBool(chkTempPath.Value)
    UpdateCtlEnable CBool(chkUpdate.Value)
    ' Имя архива при старте
    SelectStartArchName
    txtArchNameShablon.Text = strArchNameCustom
    ' Инициализация параметров для изменения шрифта и цвета
    InitializeObjectProperties
    
    'ucFontButton.
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

    If mbIsDriveCDRoom And Not mbLoadIniTmpAfterRestart Then
        If strSysIni <> strWorkTemp & "\DriversBackuper.ini" Then
            MsgBox strMessages(38), vbInformation + vbApplicationModal, strProductName
            Exit Sub
        End If

    ElseIf mbIsDriveCDRoom And mbLoadIniTmpAfterRestart Then
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
    IniWriteStrPrivate "Main", "DelTmpAfterClose", chkRemoveTemp.Value, strSysIniTemp
    ' Автообновление
    IniWriteStrPrivate "Main", "UpdateCheck", chkUpdate.Value, strSysIniTemp
    ' Автообновление Beta
    IniWriteStrPrivate "Main", "UpdateCheckBeta", chkUpdateBeta.Value, strSysIniTemp
    ' Режим запуска
    IniWriteStrPrivate "Main", "CheckAllGroup", chkCheckAll.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "ListOnlyGroup", chkHideOther.Value, strSysIniTemp

    If optGrp1.Value Then
        miRezim = 1
    ElseIf optGrp2.Value Then
        miRezim = 2
    ElseIf optGrp3.Value Then
        miRezim = 3
    Else
        miRezim = 4
    End If

    IniWriteStrPrivate "Main", "StartMode", miRezim, strSysIniTemp
    IniWriteStrPrivate "Main", "EULAAgree", Abs(mbEULAAgree), strSysIniTemp
    IniWriteStrPrivate "Main", "HideOtherProcess", chkHideOtherProcess.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTemp", chkTempPath.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "AlternativeTempPath", ucTempPath.Path, strSysIniTemp
    IniWriteStrPrivate "Main", "IconMainSkin", cmbImageMain.Text, strSysIniTemp
    IniWriteStrPrivate "Main", "SilentDLL", chkSilentDll.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "ArchMode", cmbTypeBackUp.ListIndex, strSysIni

    If mbLoadIniTmpAfterRestart Then
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", 1, strSysIniTemp
    End If

    IniWriteStrPrivate "Main", "DisableDEP", Abs(mbDisableDEP), strSysIniTemp
    ' Секция Debug
    IniWriteStrPrivate "Debug", "DebugEnable", chkDebug.Value, strSysIniTemp
    ' Очистка истории:
    IniWriteStrPrivate "Debug", "CleenHistory", chkRemoveHistory.Value, strSysIniTemp
    ' Путь до лог-файла
    IniWriteStrPrivate "Debug", "DebugLogPath", ucDebugLogPath.Path, strSysIniTemp
    IniWriteStrPrivate "Debug", "Detailmode", txtDebugLogLevel.Text, strSysIniTemp
    'Секция DPInst
    IniWriteStrPrivate "DPInst", "PathExe", ucDPInst86Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PathExe64", ucDPInst64Path.Path, strSysIniTemp
    'Секция Arc
    IniWriteStrPrivate "Arc", "PathExe", ucArchPath.Path, strSysIniTemp
    IniWriteStrPrivate "Arc", "CompressParam1", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 *.ini -ir!*.inf", strSysIni
    IniWriteStrPrivate "Arc", "CompressParam2", "-mmt=off -m0=BCJ2 -m1=LZMA2:d32m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 -xr!*.inf -x!*.ini", strSysIni
    IniWriteStrPrivate "Arc", "PathSFX", ucArchPathSFX.Path, strSysIni
    IniWriteStrPrivate "Arc", "PathSFXConfig", ucArchPathSFXConfig.Path, strSysIni
    IniWriteStrPrivate "Arc", "PathSFXConfigEn", ucArchPathSFXConfigEn.Path, strSysIni

    '[ARCName]
    If optArchNamePC.Value Then
        miArchName = 1
    ElseIf optArchModelPC.Value Then
        miArchName = 2
    Else
        miArchName = 0
    End If

    IniWriteStrPrivate "ARCName", "StartMode", miArchName, strSysIni
    IniWriteStrPrivate "ARCName", "CustomName", txtArchNameShablon, strSysIni
    'Секция OS
    OSCountNew = lvOS.count
    IniWriteStrPrivate "OS", "OSCount", OSCountNew, strSysIniTemp

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
    IniWriteStrPrivate "MainForm", "StartMaximazed", chkFormMaximaze.Value, strSysIniTemp
    mbSaveSizeOnExit = chkFormSizeSave.Value
    IniWriteStrPrivate "MainForm", "SaveSizeOnExit", chkFormSizeSave.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "HighlightColor", CStr(glHighlightColor), strSysIniTemp
    'Секция Buttons
    IniWriteStrPrivate "Button", "FontName", strFontBtn_Name, strSysIniTemp
    IniWriteStrPrivate "Button", "FontSize", miFontBtn_Size, strSysIniTemp
    IniWriteStrPrivate "Button", "FontUnderline", Abs(mbFontBtn_Underline), strSysIniTemp
    IniWriteStrPrivate "Button", "FontStrikethru", Abs(mbFontBtn_Strikethru), strSysIniTemp
    IniWriteStrPrivate "Button", "FontItalic", Abs(mbFontBtn_Italic), strSysIniTemp
    IniWriteStrPrivate "Button", "FontBold", Abs(mbFontBtn_Bold), strSysIniTemp
    IniWriteStrPrivate "Button", "FontColor", CStr(cmdFutureButton.ForeColor), strSysIniTemp
    ' Приводим Ini файл к читабельному виду
    NormIniFile strSysIniTemp
End Sub

' Режим при старте
Private Sub SelectStartArchName()

    Select Case lngArchNameMode

        Case 0
            optArchCustom.Value = True

        Case 1
            optArchNamePC.Value = True
            
        Case 2
            optArchModelPC.Value = True

        Case Else
            optArchCustom.Value = True
    End Select
End Sub

' Режим при старте
Private Sub SelectStartMode()

    Select Case miStartMode

        Case 1
            optGrp1.Value = True

        Case 2
            optGrp2.Value = True

        Case 3
            optGrp3.Value = True

        Case 4
            optGrp4.Value = True
    End Select
End Sub

Private Sub TempCtlEnable(ByVal mbEnable As Boolean)

    ucTempPath.Enabled = mbEnable
End Sub

'! -----------------------------------------------------------
'!  Функция     :  TransferOSData
'!  Переменные  :
'!  Описание    :  Передача параметров ОС из спика в форму редактирования
'! -----------------------------------------------------------
Private Sub TransferOSData()

    Dim ii As Long

    With lvOS
        ii = .SelectedItem

        If ii = -1 Then
            Exit Sub
        End If

        frmOSEdit.txtOSVer.Text = .ItemCaption(ii)
        frmOSEdit.ucPathDRP.Path = .SubItemCaption(ii, 2)
        frmOSEdit.chk64bit.Value = CBool(.SubItemCaption(ii, 1))
    End With

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
        .Create frOptions.hWnd, LVS_LIST Or LVS_SINGLESEL Or LVS_SHOWSELALWAYS, 5, 29, 190, 198, , WS_EX_STATICEDGE
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
'Private Sub ucFontButton_Click()
'
'    Dim NewFontButton As StdFont
'
'    Set NewFontButton = ucFontButton.Font
'
'    If Not NewFontButton Is Nothing Then
'        strFontBtn_Name = NewFontButton.Name
'        miFontBtn_Size = NewFontButton.Size
'        mbFontBtn_Underline = NewFontButton.Underline
'        mbFontBtn_Strikethru = NewFontButton.Strikethrough
'        mbFontBtn_Bold = NewFontButton.Bold
'        mbFontBtn_Italic = NewFontButton.Italic
'        'lngDialog_Language = NewFontButton.Charset
'        'lngDialog_Color = ucFontButton.Color
'        'cmdFutureButton.Refresh
'        'cmdFutureButton.Font.Charset = NewFont.Charset
'        'cmdFutureButton.Font.Weight = NewFont.Weight
'    End If
'
'    SetBtnFontProperties cmdFutureButton
'End Sub

'Private Sub ucFontButton_GotFocus()
'
'    HighlightActiveControl Me, ucFontButton, True
'End Sub
'
'Private Sub ucFontButton_LostFocus()
'
'    HighlightActiveControl Me, ucFontButton, False
'End Sub

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

Private Sub UpdateCtlEnable(ByVal mbEnable As Boolean)

    chkUpdateBeta.Enabled = mbEnable
End Sub
