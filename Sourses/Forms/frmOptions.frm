VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Настройки программы"
   ClientHeight    =   7740
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
   ScaleHeight     =   7740
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
         Left            =   500
         TabIndex        =   4
         Top             =   2865
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
         Height          =   255
         Left            =   500
         TabIndex        =   5
         Top             =   3900
         Width           =   7920
         _ExtentX        =   13970
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
         Caption         =   "frmOptions.frx":000C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdate 
         Height          =   255
         Left            =   500
         TabIndex        =   6
         Top             =   650
         Width           =   3200
         _ExtentX        =   5636
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
         Caption         =   "frmOptions.frx":0084
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkHideOtherProcess 
         Height          =   255
         Left            =   500
         TabIndex        =   7
         Top             =   1250
         Width           =   7920
         _ExtentX        =   13970
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
         Caption         =   "frmOptions.frx":00E0
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkTempPath 
         Height          =   255
         Left            =   500
         TabIndex        =   8
         Top             =   3555
         Width           =   3255
         _ExtentX        =   5741
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
         Caption         =   "frmOptions.frx":0146
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkUpdateBeta 
         Height          =   255
         Left            =   3735
         TabIndex        =   9
         Top             =   650
         Width           =   4680
         _ExtentX        =   8255
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
         Caption         =   "frmOptions.frx":0196
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSilentDll 
         Height          =   255
         Left            =   500
         TabIndex        =   10
         Top             =   950
         Width           =   7935
         _ExtentX        =   13996
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
         Caption         =   "frmOptions.frx":020C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucTempPath 
         Height          =   315
         Left            =   3840
         TabIndex        =   11
         Top             =   3555
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         UseAutoForeColor=   0   'False
         DefaultExt      =   ""
         Enabled         =   0   'False
         FileFlags       =   524288
         Filters         =   "Supported files|*.*|All Files (*.*)"
         UseDialogText   =   0   'False
         Locked          =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optGrp1 
         Height          =   255
         Left            =   500
         TabIndex        =   12
         Top             =   1905
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
         Top             =   1905
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
         Left            =   500
         TabIndex        =   14
         Top             =   2205
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
         Top             =   2205
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
         Left            =   3735
         TabIndex        =   16
         Top             =   2205
         Width           =   4695
         _ExtentX        =   8281
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
         Height          =   255
         Left            =   3735
         TabIndex        =   17
         Top             =   1905
         Width           =   4695
         _ExtentX        =   8281
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
         Caption         =   "frmOptions.frx":03BE
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblOptionsTemp 
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   3255
         Width           =   8175
         _ExtentX        =   14420
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
         Caption         =   "Работа с временными файлами"
      End
      Begin prjDIADBS.LabelW lblOptionsStart 
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   400
         Width           =   8175
         _ExtentX        =   14420
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
         Caption         =   "Действия при запуске программы"
      End
      Begin prjDIADBS.LabelW lblRezim 
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1605
         Width           =   8175
         _ExtentX        =   14420
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
         Caption         =   "Режим работы фильтра по умолчанию"
      End
      Begin prjDIADBS.LabelW lblTypeBackUp 
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   2550
         Width           =   8175
         _ExtentX        =   14420
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   0
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
         SmallIcons      =   "ImageListOptions"
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
   Begin prjDIADBS.ImageList ImageListOptions 
      Left            =   120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      InitListImages  =   "frmOptions.frx":041A
   End
   Begin prjDIADBS.ctlJCFrames frMainTools 
      Height          =   5300
      Left            =   3375
      Top             =   375
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
      Caption         =   "Расположение основных утилит (Tools)"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ctlUcPickBox ucDPInst86Path 
         Height          =   315
         Left            =   2535
         TabIndex        =   84
         Top             =   510
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
         UseDialogText   =   0   'False
         Locked          =   -1  'True
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
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
         UseDialogText   =   0   'False
         Locked          =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucArch86Path 
         Height          =   315
         Left            =   2535
         TabIndex        =   18
         Top             =   1350
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
         UseDialogText   =   0   'False
         Locked          =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdPathDefault 
         Height          =   495
         Left            =   4815
         TabIndex        =   19
         Top             =   3690
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
         ButtonStyle     =   8
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
         Top             =   2250
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         UseDialogText   =   0   'False
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPathSFXConfig 
         Height          =   315
         Left            =   2535
         TabIndex        =   30
         Top             =   2730
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         UseDialogText   =   0   'False
      End
      Begin prjDIADBS.ctlUcPickBox ucArchPathSFXConfigEn 
         Height          =   315
         Left            =   2535
         TabIndex        =   38
         Top             =   3210
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.*|All Files (*.*)"
         UseDialogText   =   0   'False
      End
      Begin prjDIADBS.ctlUcPickBox ucArch64Path 
         Height          =   315
         Left            =   2535
         TabIndex        =   116
         Top             =   1800
         Width           =   5895
         _ExtentX        =   10239
         _ExtentY        =   556
         DefaultExt      =   ""
         DialogType      =   1
         Enabled         =   0   'False
         Filters         =   "Supported files|*.exe|EXE Files (*.exe)"
         UseDialogText   =   0   'False
         Locked          =   -1  'True
      End
      Begin prjDIADBS.LabelW lblArc64 
         Height          =   255
         Left            =   150
         TabIndex        =   117
         Top             =   1800
         Width           =   2280
         _ExtentX        =   4022
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
         Caption         =   "7za.exe (64-bit)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblArcSFXConfigEn 
         Height          =   255
         Left            =   150
         TabIndex        =   59
         Top             =   3210
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "7za-SFXConfig (English)"
      End
      Begin prjDIADBS.LabelW lblArcSFXConfig 
         Height          =   255
         Left            =   150
         TabIndex        =   60
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "7za-SFXConfig"
      End
      Begin prjDIADBS.LabelW lblArc86 
         Height          =   255
         Left            =   150
         TabIndex        =   61
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "7za.exe (32-bit)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblDPInst64 
         Height          =   255
         Left            =   150
         TabIndex        =   62
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "DPInst.exe (64-bit)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblDPInst86 
         Height          =   255
         Left            =   150
         TabIndex        =   63
         Top             =   510
         Width           =   2280
         _ExtentX        =   4022
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
         Caption         =   "DPInst.exe (32-bit)"
         WordWrap        =   0   'False
      End
      Begin prjDIADBS.LabelW lblArcSFX 
         Height          =   255
         Left            =   150
         TabIndex        =   64
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
         Alignment       =   1
         BackStyle       =   0
         Caption         =   "7za-sfxModule"
      End
   End
   Begin prjDIADBS.ctlJCFrames frArchName 
      Height          =   5300
      Left            =   3645
      Top             =   705
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
      Caption         =   "Имя архива"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.TextBoxW txtArchNameShablon 
         Height          =   330
         Left            =   480
         TabIndex        =   49
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
      Begin prjDIADBS.TextBoxW txtArchMacrosPCName 
         Height          =   255
         Left            =   480
         TabIndex        =   48
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
         Text            =   "frmOptions.frx":043A
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtArchMacrosPCModel 
         Height          =   255
         Left            =   480
         TabIndex        =   47
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
         Text            =   "frmOptions.frx":046A
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtArchMacrosOSVER 
         Height          =   255
         Left            =   480
         TabIndex        =   46
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
         Text            =   "frmOptions.frx":049C
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtArchMacrosOSBIT 
         Height          =   255
         Left            =   480
         TabIndex        =   45
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
         Text            =   "frmOptions.frx":04CA
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtArchMacrosDate 
         Height          =   255
         Left            =   480
         TabIndex        =   44
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
         Text            =   "frmOptions.frx":04F8
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
         Caption         =   "frmOptions.frx":0524
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optArchNamePC 
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   735
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
         Caption         =   "frmOptions.frx":0566
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.OptionButtonW optArchCustom 
         Height          =   255
         Left            =   480
         TabIndex        =   43
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
         Caption         =   "frmOptions.frx":05A2
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.LabelW lblArchMacrosDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   65
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
      Begin prjDIADBS.LabelW lblArchMacrosOSBit 
         Height          =   375
         Left            =   2400
         TabIndex        =   66
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
      Begin prjDIADBS.LabelW lblArchMacrosOSVer 
         Height          =   375
         Left            =   2400
         TabIndex        =   67
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
      Begin prjDIADBS.LabelW lblArchMacrosPCModel 
         Height          =   375
         Left            =   2400
         TabIndex        =   68
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
      Begin prjDIADBS.LabelW lblArchMacrosParam 
         Height          =   255
         Left            =   480
         TabIndex        =   69
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
      Begin prjDIADBS.LabelW lblArchMacrosDescription 
         Height          =   255
         Left            =   2400
         TabIndex        =   70
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
      Begin prjDIADBS.LabelW lblArchMacrosPCName 
         Height          =   375
         Left            =   2400
         TabIndex        =   71
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
      Begin prjDIADBS.LabelW lblArchMacrosType 
         Height          =   255
         Left            =   480
         TabIndex        =   72
         Top             =   2685
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   450
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
         Caption         =   "Доступные макроподстановки:"
      End
      Begin prjDIADBS.LabelW lblArchShablon 
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   1845
         Width           =   8175
         _ExtentX        =   14420
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
         Caption         =   "Шаблон имени архива"
      End
      Begin prjDIADBS.LabelW lblArchNameStart 
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   405
         Width           =   8100
         _ExtentX        =   14288
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
         Caption         =   "Имя архива по умолчанию"
      End
   End
   Begin prjDIADBS.ctlJCFrames frOS 
      Height          =   5300
      Left            =   3900
      Top             =   1020
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
      Caption         =   "Поддерживаемые ОС"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.ListView lvOS 
         Height          =   3795
         Left            =   120
         TabIndex        =   85
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
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
         View            =   3
         Arrange         =   1
         AllowColumnReorder=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         LabelEdit       =   2
         HideSelection   =   0   'False
         ShowLabelTips   =   -1  'True
         HoverSelection  =   -1  'True
         HotTracking     =   -1  'True
         HighlightHot    =   -1  'True
         TextBackground  =   1
      End
      Begin prjDIADBS.ctlJCbutton cmdAddOS 
         Height          =   750
         Left            =   120
         TabIndex        =   40
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
         ButtonStyle     =   8
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
         ButtonStyle     =   8
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
         ButtonStyle     =   8
         BackColor       =   16765357
         Caption         =   "Удалить"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
   End
   Begin prjDIADBS.ctlJCFrames frDesign 
      Height          =   5300
      Left            =   4140
      Top             =   1320
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
      Caption         =   "Оформление"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin prjDIADBS.CheckBoxW chkButtonDisable 
         Height          =   240
         Left            =   500
         TabIndex        =   32
         Top             =   2745
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "frmOptions.frx":05D6
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkFormMaximaze 
         Height          =   255
         Left            =   3300
         TabIndex        =   33
         Top             =   800
         Width           =   5040
         _ExtentX        =   8890
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
         Caption         =   "frmOptions.frx":064C
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtFormHeight 
         Height          =   255
         Left            =   1410
         TabIndex        =   34
         Top             =   795
         Width           =   1395
         _ExtentX        =   2461
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
         Min             =   6000
         Max             =   25000
         Value           =   6000
      End
      Begin prjDIADBS.SpinBox txtFormWidth 
         Height          =   255
         Left            =   1410
         TabIndex        =   35
         Top             =   1155
         Width           =   1395
         _ExtentX        =   2461
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
         Min             =   13000
         Max             =   25000
         Value           =   13000
      End
      Begin prjDIADBS.CheckBoxW chkFormSizeSave 
         Height          =   255
         Left            =   3300
         TabIndex        =   36
         Top             =   1155
         Width           =   5040
         _ExtentX        =   8890
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
         Caption         =   "frmOptions.frx":06B2
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFutureButton 
         Height          =   600
         Left            =   500
         TabIndex        =   37
         Top             =   1900
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1058
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
         Caption         =   "Твоя будущая кнопка"
         CaptionEffects  =   0
         PictureAlign    =   0
         PicturePushOnHover=   -1  'True
         PictureShadow   =   -1  'True
      End
      Begin prjDIADBS.ctlJCbutton cmdFontColorButton 
         Height          =   600
         Left            =   3300
         TabIndex        =   83
         Top             =   1900
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1058
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
      Begin prjDIADBS.ComboBoxW cmbImageMain 
         Height          =   315
         Left            =   500
         TabIndex        =   118
         Top             =   4440
         Width           =   2595
         _ExtentX        =   4577
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
         Text            =   "frmOptions.frx":0716
         Sorted          =   -1  'True
      End
      Begin prjDIADBS.ComboBoxW cmbButtonStyle 
         Height          =   315
         Left            =   500
         TabIndex        =   121
         Top             =   3400
         Width           =   2600
         _ExtentX        =   4233
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
         Sorted          =   -1  'True
      End
      Begin prjDIADBS.ctlColorButton ctlStatusBtnBackColor 
         Height          =   330
         Left            =   6225
         TabIndex        =   123
         Top             =   3400
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   582
         Icon            =   "frmOptions.frx":074E
         BackColor       =   14016736
      End
      Begin prjDIADBS.ComboBoxW cmbButtonStyleColor 
         Height          =   315
         Left            =   3465
         TabIndex        =   124
         Top             =   3400
         Width           =   2595
         _ExtentX        =   4233
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
         Sorted          =   -1  'True
      End
      Begin prjDIADBS.LabelW lblButtonStyleColor 
         Height          =   255
         Left            =   3465
         TabIndex        =   125
         Top             =   3100
         Width           =   3495
         _ExtentX        =   6165
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
         BackStyle       =   0
         Caption         =   "Стиль оформления кнопки"
      End
      Begin prjDIADBS.LabelW lblButtonStyle 
         Height          =   255
         Left            =   500
         TabIndex        =   122
         Top             =   3100
         Width           =   3000
         _ExtentX        =   5292
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
         BackStyle       =   0
         Caption         =   "Стиль оформления кнопки"
      End
      Begin prjDIADBS.LabelW lblImageMain 
         Height          =   255
         Left            =   500
         TabIndex        =   120
         Top             =   4150
         Width           =   3000
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
         Caption         =   "Основные картинки"
      End
      Begin prjDIADBS.LabelW lblTheme 
         Height          =   255
         Left            =   240
         TabIndex        =   119
         Top             =   3850
         Width           =   8200
         _ExtentX        =   14473
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
         Caption         =   "Набор оформления программы (изменение основных иконок, и иконок статуса кнопок)"
      End
      Begin prjDIADBS.LabelW lblFormWidth 
         Height          =   255
         Left            =   495
         TabIndex        =   55
         Top             =   1155
         Width           =   900
         _ExtentX        =   1588
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
         BackStyle       =   0
         Caption         =   "Ширина:"
      End
      Begin prjDIADBS.LabelW lblFormHeight 
         Height          =   255
         Left            =   495
         TabIndex        =   56
         Top             =   795
         Width           =   900
         _ExtentX        =   1588
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
         BackStyle       =   0
         Caption         =   "Высота:"
      End
      Begin prjDIADBS.LabelW lblSizeForm 
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   495
         Width           =   8200
         _ExtentX        =   14473
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
         Left            =   240
         TabIndex        =   58
         Top             =   1575
         Width           =   8200
         _ExtentX        =   14473
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
      Height          =   5300
      Left            =   4380
      Top             =   1620
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
      Caption         =   "Параметры запуска DPInst"
      TextBoxHeight   =   18
      ThemeColor      =   2
      HeaderStyle     =   1
      Begin VB.CommandButton cmdLegacyMode 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   94
         ToolTipText     =   "More on MSDN..."
         Top             =   660
         Width           =   255
      End
      Begin VB.CommandButton cmdPromptIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2640
         TabIndex        =   97
         ToolTipText     =   "More on MSDN..."
         Top             =   1305
         Width           =   255
      End
      Begin VB.CommandButton cmdForceIfDriverIsNotBetter 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   100
         ToolTipText     =   "More on MSDN..."
         Top             =   1905
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressAddRemovePrograms 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   103
         ToolTipText     =   "More on MSDN..."
         Top             =   2460
         Width           =   255
      End
      Begin VB.CommandButton cmdSuppressWizard 
         Caption         =   "?"
         Height          =   255
         Left            =   2655
         TabIndex        =   106
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
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   660
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":0CD4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkPromptIfDriverIsNotBetter 
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   1305
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":0D08
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkForceIfDriverIsNotBetter 
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   1905
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":0D5A
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressAddRemovePrograms 
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   2460
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":0DAA
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkSuppressWizard 
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   2955
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":0DFC
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkQuietInstall 
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3510
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":0E38
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkScanHardware 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4005
         Width           =   2520
         _ExtentX        =   4445
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
         Caption         =   "frmOptions.frx":0E70
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
         TabIndex        =   92
         Top             =   370
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
         TabIndex        =   91
         Top             =   370
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
         TabIndex        =   98
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
         TabIndex        =   95
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
         TabIndex        =   101
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
         TabIndex        =   104
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
         TabIndex        =   107
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
         Caption         =   $"frmOptions.frx":0EA8
      End
   End
   Begin prjDIADBS.ctlJCFrames frDebug 
      Height          =   5300
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
         Left            =   500
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
         Text            =   "frmOptions.frx":0FA6
      End
      Begin prjDIADBS.TextBoxW txtDebugMacrosDate 
         Height          =   255
         Left            =   500
         TabIndex        =   108
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
         Text            =   "frmOptions.frx":0FEA
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtDebugMacrosOSBIT 
         Height          =   255
         Left            =   500
         TabIndex        =   109
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
         Text            =   "frmOptions.frx":1016
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtDebugMacrosOSVER 
         Height          =   255
         Left            =   500
         TabIndex        =   110
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
         Text            =   "frmOptions.frx":1044
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtDebugMacrosPCModel 
         Height          =   255
         Left            =   500
         TabIndex        =   111
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
         Text            =   "frmOptions.frx":1072
         Locked          =   -1  'True
      End
      Begin prjDIADBS.TextBoxW txtDebugMacrosPCName 
         Height          =   255
         Left            =   500
         TabIndex        =   112
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
         Text            =   "frmOptions.frx":10A4
         Locked          =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDebug 
         Height          =   255
         Left            =   500
         TabIndex        =   113
         Top             =   700
         Width           =   4440
         _ExtentX        =   7832
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
         Caption         =   "frmOptions.frx":10D4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.ctlUcPickBox ucDebugLogPath 
         Height          =   315
         Left            =   500
         TabIndex        =   54
         Top             =   1900
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
      Begin prjDIADBS.CheckBoxW chkDebugLog2AppPath 
         Height          =   255
         Left            =   500
         TabIndex        =   114
         Top             =   1300
         Width           =   7920
         _ExtentX        =   13970
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
         Caption         =   "frmOptions.frx":1124
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.CheckBoxW chkDebugTime2File 
         Height          =   255
         Left            =   500
         TabIndex        =   115
         Top             =   1000
         Width           =   7920
         _ExtentX        =   13970
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
         Caption         =   "frmOptions.frx":11A4
         Transparent     =   -1  'True
      End
      Begin prjDIADBS.SpinBox txtDebugLogLevel 
         Height          =   255
         Left            =   7680
         TabIndex        =   75
         Top             =   700
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
      Begin prjDIADBS.LabelW lblDebugLogLevel 
         Height          =   255
         Left            =   4995
         TabIndex        =   76
         Top             =   700
         Width           =   2595
         _ExtentX        =   4577
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
      Begin prjDIADBS.LabelW lblDebugMacrosDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   77
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
      Begin prjDIADBS.LabelW lblDebugMacrosOSBit 
         Height          =   375
         Left            =   2400
         TabIndex        =   78
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
      Begin prjDIADBS.LabelW lblDebugMacrosOSVer 
         Height          =   375
         Left            =   2400
         TabIndex        =   79
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
      Begin prjDIADBS.LabelW lblDebugMacrosPCModel 
         Height          =   375
         Left            =   2400
         TabIndex        =   80
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
      Begin prjDIADBS.LabelW lblDebugMacrosParam 
         Height          =   255
         Left            =   500
         TabIndex        =   81
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
      Begin prjDIADBS.LabelW lblDebugMacrosDescription 
         Height          =   255
         Left            =   2400
         TabIndex        =   82
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
      Begin prjDIADBS.LabelW lblDebugMacrosPCName 
         Height          =   375
         Left            =   2400
         TabIndex        =   86
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
      Begin prjDIADBS.LabelW lblDebugMacrosType 
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   2900
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   450
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
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   1600
         Width           =   7845
         _ExtentX        =   13838
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
         Caption         =   "Каталог для создания log-файлов:"
      End
      Begin prjDIADBS.LabelW lblDebug 
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   400
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
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   2250
         Width           =   7845
         _ExtentX        =   13838
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
         Caption         =   "Каталог для создания log-файлов:"
      End
   End
   Begin prjDIADBS.ctlJCFrames frOther 
      Height          =   5300
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

Private strItemOptions1           As String 'Основные настройки
Private strItemOptions2           As String 'Поддерживаемые ОС
Private strItemOptions3           As String 'Рабочие утилиты
Private strItemOptions4           As String 'Имя Архива
Private strItemOptions5           As String 'Оформление программы
Private strItemOptions6           As String 'Параметры запуска DPInst
Private strItemOptions7           As String 'Отладочный режим
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkButtonDisable_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkButtonDisable_Click()
    cmdFutureButton.Enabled = CBool(chkButtonDisable.Value)
End Sub


'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkDebugLog2AppPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkDebugLog2AppPath_Click()
    DebugCtlEnableLog2App Not CBool(chkDebugLog2AppPath.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkDebug_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkDebug_Click()
    DebugCtlEnable CBool(chkDebug.Value)
    DebugCtlEnableLog2App Not CBool(chkDebugLog2AppPath.Value)

    If Not CBool(chkDebug.Value) Then
        If Not CBool(chkDebugLog2AppPath.Value) Then
            ucDebugLogPath.Enabled = False
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkForceIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkForceIfDriverIsNotBetter_Click()
    mbDpInstForceIfDriverIsNotBetter = CBool(chkForceIfDriverIsNotBetter.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkFormMaximaze_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkFormMaximaze_Click()

    If chkFormMaximaze.Value Then
        chkFormSizeSave.Value = vbUnchecked
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkFormSizeSave_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkFormSizeSave_Click()

    If chkFormSizeSave.Value Then
        chkFormMaximaze.Value = vbUnchecked
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
'! Procedure   (Функция)   :   Sub chkLegacyMode_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkLegacyMode_Click()
    mbDpInstLegacyMode = CBool(chkLegacyMode.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkPromptIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkPromptIfDriverIsNotBetter_Click()
    mbDpInstPromptIfDriverIsNotBetter = CBool(chkPromptIfDriverIsNotBetter.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkQuietInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkQuietInstall_Click()
    mbDpInstQuietInstall = CBool(chkQuietInstall.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkScanHardware_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkScanHardware_Click()
    mbDpInstScanHardware = CBool(chkScanHardware.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkSuppressAddRemovePrograms_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressAddRemovePrograms_Click()
    mbDpInstSuppressAddRemovePrograms = CBool(chkSuppressAddRemovePrograms.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkSuppressWizard_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkSuppressWizard_Click()
    mbDpInstSuppressWizard = CBool(chkSuppressWizard.Value)
    txtCmdStringDPInst = CollectCmdString
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkTempPath_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkTempPath_Click()
    TempCtlEnable CBool(chkTempPath.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkUpdate_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkUpdate_Click()
    UpdateCtlEnable CBool(chkUpdate.Value)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyleColor_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyleColor_Click()
Dim lngIndex As Long

    lngIndex = cmbButtonStyleColor.ListIndex
    If lngIndex > -1 Then
        cmdFutureButton.ColorScheme = lngIndex
        If lngIndex < 3 Then
            ctlStatusBtnBackColor.Visible = False
        Else
            ctlStatusBtnBackColor.Visible = True
            cmdFutureButton.BackColor = ctlStatusBtnBackColor.Value
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyleColor_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyleColor_GotFocus()
    HighlightActiveControl Me, cmbButtonStyleColor, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyleColor_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyleColor_LostFocus()
    HighlightActiveControl Me, cmbButtonStyleColor, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyle_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyle_Click()
Dim lngIndex As Long
    
    lngIndex = cmbButtonStyle.ListIndex
    If lngIndex > -1 Then
        cmdFutureButton.ButtonStyle = lngIndex
        Select Case lngIndex
            Case 0, 1, 4, 6, 7, 11, 12
                cmbButtonStyleColor.Enabled = False
                ctlStatusBtnBackColor.Visible = True
                cmbButtonStyleColor.ListIndex = 3
            Case 2, 3, 5, 8, 9, 10
                cmbButtonStyleColor.ListIndex = cmdFutureButton.ColorScheme
                cmbButtonStyleColor.Enabled = True
                cmbButtonStyleColor_Click
            Case 13
                cmbButtonStyleColor.Enabled = False
                ctlStatusBtnBackColor.Visible = False
        End Select
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyle_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyle_GotFocus()
    HighlightActiveControl Me, cmbButtonStyle, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbButtonStyle_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbButtonStyle_LostFocus()
    HighlightActiveControl Me, cmbButtonStyle, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageMain_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_Click()

    If PathExists(strPathImageMain & cmbImageMain.Text) = False Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageMain_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_GotFocus()
    HighlightActiveControl Me, cmbImageMain, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbImageMain_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbImageMain_LostFocus()

    If PathExists(strPathImageMain & cmbImageMain.Text) = False Then
        cmbImageMain.BackColor = vbRed
    Else
        cmbImageMain.BackColor = &H80000005
    End If

    HighlightActiveControl Me, cmbImageMain, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbTypeBackUp_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbTypeBackUp_GotFocus()

    HighlightActiveControl Me, cmbTypeBackUp, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmbTypeBackUp_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmbTypeBackUp_LostFocus()

    HighlightActiveControl Me, cmbTypeBackUp, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdAddOS_Click
'! Description (Описание)  :   [кнопка добавления ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdAddOS_Click()
    mbAddInList = True
    frmOSEdit.Show vbModal, Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdDelOS_Click
'! Description (Описание)  :   [кнопка удаление ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdDelOS_Click()

    With lvOS

        If .ListItems.count Then
            .ListItems.Remove (.SelectedItem.Index)
            lngLastIdOS = lngLastIdOS - 1
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdEditOS_Click
'! Description (Описание)  :   [кнопка редактирование ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdEditOS_Click()
    TransferOSData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [Нажатие кнопки Выход. Выход без сохранения]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Me.Hide
    ChangeStatusBarText cmdExit.Caption
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdFontColorButton_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdFontColorButton_Click()

    With frmFontDialog
        .optControl(3).Value = True
        .txtFont.Font.Name = strFontBtn_Name
        .txtFont.Font.Size = miFontBtn_Size
        .txtFont.Font.Bold = mbFontBtn_Bold
        .txtFont.Font.Italic = mbFontBtn_Italic
        .txtFont.Font.Underline = mbFontBtn_Underline
        .txtFont.Font.Charset = lngFont_Charset
        .txtFont.ForeColor = lngFontBtn_Color
        .ctlFontColor.Value = lngFontBtn_Color
        .Show vbModal, Me
    End With
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdForceIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdForceIfDriverIsNotBetter_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff544948.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdLegacyMode_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdLegacyMode_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff548635.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [Нажатие кнопки ОК. Применение настроек]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()

    Dim lngMsgRet As Long

    If mbIsDriveCDRoom And mbLoadIniTmpAfterRestart Then
        SaveOptions
        ChangeStatusBarText strMessages(36)
        lngMsgRet = MsgBox(strMessages(36) & strMessages(147), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mbRestartProgram = lngMsgRet = vbYes
    ElseIf FileExists(strSysIni) Then
        If Not FileisReadOnly(strSysIni) Then
            SaveOptions
            ChangeStatusBarText strMessages(36)
            lngMsgRet = MsgBox(strMessages(36) & strMessages(147), vbInformation + vbApplicationModal + vbYesNo, strProductName)
            mbRestartProgram = lngMsgRet = vbYes
        End If
    Else
        SaveOptions
        ChangeStatusBarText strMessages(36)
        lngMsgRet = MsgBox(strMessages(36) & strMessages(147), vbInformation + vbApplicationModal + vbYesNo, strProductName)
        mbRestartProgram = lngMsgRet = vbYes
    End If

    Me.Hide
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPathDefault_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdPathDefault_Click()
    'Секция DPInst
    ucDPInst86Path.Path = "Tools\DPInst\DPInst.exe"
    ucDPInst64Path.Path = "Tools\DPInst\DPInst64.exe"
    'Секция Arc
    ucArch86Path.Path = "Tools\Arc\7za.exe"
    ucArch64Path.Path = "Tools\Arc\7za64.exe"
    ucArchPathSFX.Path = "Tools\Arc\sfx\7zSD.sfx"
    ucArchPathSFXConfig.Path = "Tools\Arc\sfx\config.txt"
    ucArchPathSFXConfigEn.Path = "Tools\Arc\sfx\config_en.txt"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdPromptIfDriverIsNotBetter_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdPromptIfDriverIsNotBetter_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff549759.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdQuietInstall_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdQuietInstall_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff549799.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdScanHardware_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdScanHardware_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff550761.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdSuppressAddRemovePrograms_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdSuppressAddRemovePrograms_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff553404.aspx", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdSuppressWizard_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdSuppressWizard_Click()
    RunUtilsShell "http://msdn.microsoft.com/en-us/library/ff550803.aspx#setting_the_suppresswizard_flag", False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlStatusBtnBackColor_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ctlStatusBtnBackColor_Click()
    If cmbButtonStyleColor.ListIndex = 3 Then
        cmdFutureButton.BackColor = ctlStatusBtnBackColor.Value
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DebugCtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub DebugCtlEnable(ByVal mbEnable As Boolean)
    chkDebugTime2File.Enabled = mbEnable
    txtDebugLogName.Enabled = mbEnable
    ucDebugLogPath.Enabled = mbEnable
    chkDebugLog2AppPath.Enabled = mbEnable
    txtDebugLogLevel.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DebugCtlEnableLog2App
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub DebugCtlEnableLog2App(ByVal mbEnable As Boolean)
    ucDebugLogPath.Enabled = mbEnable
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
'! Procedure   (Функция)   :   Sub FormLoadAction
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub FormLoadAction()

    ' Локализация приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    ' загрузить список опций
    LoadList_lvOptions
    ' Заполнить опции
    ReadOptions
    ' установить опции шрифта и цвета для будущей кнопки
    SetBtnStatusFontProperties cmdFutureButton
    ' установить опции стиля для будущей кнопки
    SetBtnStyle cmdFutureButton

    DoEvents
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [Обработка нажатий клавиш клавиатуры сначала на форме]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        If MsgBox(strMessages(37), vbQuestion + vbYesNo, strProductName) = vbYes Then
            cmdExit_Click
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Load
'! Description (Описание)  :   [Загрузка формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, strFormName, False
        .Height = 5850
        .Width = 11900
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    'Top frame position
    frArchName.Top = 50
    frOptions.Top = 50
    frDesign.Top = 50
    frDpInstParam.Top = 50
    frMain.Top = 50
    frMainTools.Top = 50

    frOS.Top = 50

    frOther.Top = 50
    frDebug.Top = 50

    'Left frame position
    frArchName.Left = 3100
    frDesign.Left = 3100
    frDpInstParam.Left = 3100

    frMain.Left = 3100
    frMainTools.Left = 3100

    frOS.Left = 3100
    frOther.Left = 3100
    frDebug.Left = 3100
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

    'загружаем список стилей кнопок
    LoadComboBtnStyle
    ' Действия при загрузке формы
    FormLoadAction
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_QueryUnload
'! Description (Описание)  :   [Выгружаем из памяти форму и другие компоненты]
'! Parameters  (Переменные):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
        ChangeStatusBarText cmdExit.Caption
    Else
        Set frmOptions = Nothing
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Resize
'! Description (Описание)  :   [Изменение размеров формы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState <> vbMinimized Then
        SetTrayIcon NIM_DELETE, Me.hWnd, 0&, vbNullString
    Else
        SetTrayIcon NIM_ADD, Me.hWnd, Me.Icon, App.ProductName
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub InitializeObjectProperties
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub InitializeObjectProperties()

    cmbButtonStyle.ListIndex = lngStatusBtnStyle
    cmbButtonStyleColor.ListIndex = lngStatusBtnStyleColor
    ctlStatusBtnBackColor.Value = lngStatusBtnBackColor

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadComboBtnStyle
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadComboBtnStyle()
    
    With cmbButtonStyle
        .Clear
        .AddItem "Standard", 0
        .AddItem "Flat", 1
        .AddItem "WindowsXP", 2
        .AddItem "VistaAero", 3
        .AddItem "OfficeXP", 4
        .AddItem "Office2003", 5
        .AddItem "XPToolbar", 6
        .AddItem "VistaToolbar", 7
        .AddItem "Outlook2007", 8
        .AddItem "InstallShield", 9
        .AddItem "GelButton", 10
        .AddItem "3DHover", 11
        .AddItem "FlatHover", 12
        .AddItem "WindowsTheme", 13
    End With
    
    With cmbButtonStyleColor
        .Clear
        .AddItem "Blue", 0
        .AddItem "OliveGreen", 1
        .AddItem "Silver", 2
        .AddItem "Custom", 3
    End With
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
'! Procedure   (Функция)   :   Sub LoadListImage
'! Description (Описание)  :   [Загрузка картинок в ListImage]
'! Parameters  (Переменные):   strImageName as string
'!                             lngImageIndex as Long
'!--------------------------------------------------------------------------------
Private Sub LoadListImage(ByVal strImageName As String, ByVal lngImageIndex As Long)
Dim objPicTmp As StdPicture

    Set objPicTmp = GetImageFromFile(strImageName, strPathImageMainWork)
    If Not objPicTmp Is Nothing Then
        ImageListOptions.ListImages.Add lngImageIndex, , objPicTmp
    End If
    
    Set objPicTmp = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_lvOptions
'! Description (Описание)  :   [Построение дерева настроек]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadList_lvOptions()

    ' Загружаем картинки в ImageList
    If ImageListOptions.ListImages.count = 0 Then
        LoadListImage "OPT_MAIN", 1
        LoadListImage "OPT_OSLIST", 2
        LoadListImage "OPT_TOOLS_MAIN", 3
        LoadListImage "OPT_ARCHNAME", 4
        LoadListImage "OPT_DESIGN", 5
        LoadListImage "OPT_DPINST", 6
        LoadListImage "OPT_DEVPARSER", 7
    End If
           
    ' Заполняем ListView названием опций программы
    With lvOptions
        With .ListItems
            If .count = 0 Then
                .Add 1, , strItemOptions1, , 1
                .Add 2, , strItemOptions2, , 2
                .Add 3, , strItemOptions3, , 3
                .Add 4, , strItemOptions4, , 4
                .Add 5, , strItemOptions5, , 5
                .Add 6, , strItemOptions6, , 6
                .Add 7, , strItemOptions7, , 7
            Else
                .item(1).Text = strItemOptions1
                .item(2).Text = strItemOptions2
                .item(3).Text = strItemOptions3
                .item(4).Text = strItemOptions4
                .item(5).Text = strItemOptions5
                .item(6).Text = strItemOptions6
                .item(7).Text = strItemOptions7
            End If
        End With
    
        .ColumnWidth = .Width - 100
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadList_lvOS
'! Description (Описание)  :   [Построение спиcка ОС]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub LoadList_lvOS()

    Dim ii As Long

    With lvOS
        .ListItems.Clear
        .ColumnHeaders.Clear

        If .ColumnHeaders.count = 0 Then
            .ColumnHeaders.Add 1, , strTableOSHeader1, 150 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 2, , strTableOSHeader2, 60 * Screen.TwipsPerPixelX
            .ColumnHeaders.Add 3, , strTableOSHeader3, 150 * Screen.TwipsPerPixelX
        End If

        For ii = 0 To lngOSCount - 1

            With .ListItems.Add(, , arrOSList(ii).Ver)
                .SubItems(1) = arrOSList(ii).is64bit
                .SubItems(2) = arrOSList(ii).drpFolder
            End With

        Next

    End With

    lngLastIdOS = lngOSCount
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadStartMode
'! Description (Описание)  :   [Режимы работы при старте]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadSkinListCombo
'! Description (Описание)  :   [Загрузка списка скинов]
'! Parameters  (Переменные):   cmbName (ComboBox)
'                              strImagePath (String)
'!--------------------------------------------------------------------------------
Private Sub LoadSkinListCombo(cmbName As Object, strImagePath As String)

    Dim strListFolder_x() As FindListStruct
    Dim ii                As Integer

    strListFolder_x = SearchFoldersInRoot(strImagePath, "*")

    With cmbName
        .Clear

        For ii = 0 To UBound(strListFolder_x)
            .AddItem strListFolder_x(ii).Name, ii
        Next ii

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)

    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    ' Описание режимов
    frOptions.Caption = LocaliseString(strPathFile, strFormName, "frOptions", frOptions.Caption)
    strItemOptions1 = LocaliseString(strPathFile, strFormName, "ItemOptions1", "Основные настройки")
    strItemOptions2 = LocaliseString(strPathFile, strFormName, "ItemOptions2", "Поддерживаемые ОС")
    strItemOptions3 = LocaliseString(strPathFile, strFormName, "ItemOptions3", "Рабочие утилиты")
    strItemOptions4 = LocaliseString(strPathFile, strFormName, "ItemOptions4", "Имя Архива")
    strItemOptions5 = LocaliseString(strPathFile, strFormName, "ItemOptions5", "Оформление программы")
    strItemOptions6 = LocaliseString(strPathFile, strFormName, "ItemOptions6", "Параметры запуска DPInst")
    strItemOptions7 = LocaliseString(strPathFile, strFormName, "ItemOptions7", "Отладочный режим")
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    'frMain
    frMain.Caption = LocaliseString(strPathFile, strFormName, "frMain", frMain.Caption)
    lblOptionsStart.Caption = LocaliseString(strPathFile, strFormName, "lblOptionsStart", lblOptionsStart.Caption)
    chkUpdate.Caption = LocaliseString(strPathFile, strFormName, "chkUpdate", chkUpdate.Caption)
    chkUpdateBeta.Caption = LocaliseString(strPathFile, strFormName, "chkUpdateBeta", chkUpdateBeta.Caption)
    chkHideOtherProcess.Caption = LocaliseString(strPathFile, strFormName, "chkHideOtherProcess", chkHideOtherProcess.Caption)
    chkSilentDll.Caption = LocaliseString(strPathFile, strFormName, "chkSilentDll", chkSilentDll.Caption)
    lblOptionsTemp.Caption = LocaliseString(strPathFile, strFormName, "lblOptionsTemp", lblOptionsTemp.Caption)
    chkTempPath.Caption = LocaliseString(strPathFile, strFormName, "chkTempPath", chkTempPath.Caption)
    chkRemoveTemp.Caption = LocaliseString(strPathFile, strFormName, "chkRemoveTemp", chkRemoveTemp.Caption)
    lblRezim.Caption = LocaliseString(strPathFile, strFormName, "lblRezim", lblRezim.Caption)
    lblTypeBackUp.Caption = LocaliseString(strPathFile, strFormName, "lblTypeBackUp", lblTypeBackUp.Caption)
    'frDebug
    frDebug.Caption = LocaliseString(strPathFile, strFormName, "frDebug", frDebug.Caption)
    lblDebug.Caption = LocaliseString(strPathFile, strFormName, "lblDebug", lblDebug.Caption)
    chkDebug.Caption = LocaliseString(strPathFile, strFormName, "chkDebug", chkDebug.Caption)
    chkDebugLog2AppPath.Caption = LocaliseString(strPathFile, strFormName, "chkDebugLog2AppPath", chkDebugLog2AppPath.Caption)
    chkDebugTime2File.Caption = LocaliseString(strPathFile, strFormName, "chkDebugTime2File", chkDebugTime2File.Caption)
    lblDebugLogName.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogName", lblDebugLogName.Caption)
    lblDebugLogLevel.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogLevel", lblDebugLogLevel.Caption)
    lblDebugLogPath.Caption = LocaliseString(strPathFile, strFormName, "lblDebugLogPath", lblDebugLogPath.Caption)
    lblDebugMacrosType.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosType", lblDebugMacrosType.Caption)
    lblDebugMacrosParam.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosParam", lblDebugMacrosParam.Caption)
    lblDebugMacrosDescription.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosDescription", lblDebugMacrosDescription.Caption)
    lblDebugMacrosPCName.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosPCName", lblDebugMacrosPCName.Caption)
    lblDebugMacrosPCModel.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosPCModel", lblDebugMacrosPCModel.Caption)
    lblDebugMacrosOSVer.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosOSVer", lblDebugMacrosOSVer.Caption)
    lblDebugMacrosOSBit.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosOSBit", lblDebugMacrosOSBit.Caption)
    lblDebugMacrosDate.Caption = LocaliseString(strPathFile, strFormName, "lblDebugMacrosDate", lblDebugMacrosDate.Caption)
    'frOS
    frOS.Caption = LocaliseString(strPathFile, strFormName, "frOS", frOS.Caption)
    cmdAddOS.Caption = LocaliseString(strPathFile, strFormName, "cmdAddOS", cmdAddOS.Caption)
    cmdEditOS.Caption = LocaliseString(strPathFile, strFormName, "cmdEditOS", cmdEditOS.Caption)
    cmdDelOS.Caption = LocaliseString(strPathFile, strFormName, "cmdDelOS", cmdDelOS.Caption)
    strTableOSHeader1 = LocaliseString(strPathFile, strFormName, "TableOSHeader1", "Версия")
    strTableOSHeader2 = LocaliseString(strPathFile, strFormName, "TableOSHeader2", "x64")
    strTableOSHeader3 = LocaliseString(strPathFile, strFormName, "TableOSHeader3", "Путь")
    'frDesign
    frDesign.Caption = LocaliseString(strPathFile, strFormName, "frDesign", frDesign.Caption)
    lblSizeForm.Caption = LocaliseString(strPathFile, strFormName, "lblSizeForm", lblSizeForm.Caption)
    lblFormHeight.Caption = LocaliseString(strPathFile, strFormName, "lblFormHeight", lblFormHeight.Caption)
    lblFormWidth.Caption = LocaliseString(strPathFile, strFormName, "lblFormWidth", lblFormWidth.Caption)
    chkFormMaximaze.Caption = LocaliseString(strPathFile, strFormName, "chkFormMaximaze", chkFormMaximaze.Caption)
    chkFormSizeSave.Caption = LocaliseString(strPathFile, strFormName, "chkFormSizeSave", chkFormSizeSave.Caption)
    lblSizeButton.Caption = LocaliseString(strPathFile, strFormName, "lblSizeButton", lblSizeButton.Caption)
    cmdFutureButton.Caption = LocaliseString(strPathFile, strFormName, "cmdFutureButton", cmdFutureButton.Caption)
    lblImageMain.Caption = LocaliseString(strPathFile, strFormName, "lblImageMain", lblImageMain.Caption)
    cmdFontColorButton.Caption = LocaliseString(strPathFile, strFormName, "cmdFontColorButton", cmdFontColorButton.Caption)
    cmdFutureButton.Caption = LocaliseString(strPathFile, strFormName, "cmdFutureButton", cmdFutureButton.Caption)
    chkButtonDisable.Caption = LocaliseString(strPathFile, strFormName, "chkButtonDisable", chkButtonDisable.Caption)
    lblTheme.Caption = LocaliseString(strPathFile, strFormName, "lblTheme", lblTheme.Caption)
    ctlStatusBtnBackColor.DropDownCaption = LocaliseString(strPathFile, strFormName, "ctlStatusBtnBackColor", ctlStatusBtnBackColor.DropDownCaption)
    lblButtonStyle.Caption = LocaliseString(strPathFile, strFormName, "lblButtonStyle", lblButtonStyle.Caption)
    lblButtonStyleColor.Caption = LocaliseString(strPathFile, strFormName, "lblButtonStyleColor", lblButtonStyleColor.Caption)
    'frDpInstParam
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
    'frArchName
    frArchName.Caption = LocaliseString(strPathFile, strFormName, "frArchName", frArchName.Caption)
    lblArchNameStart.Caption = LocaliseString(strPathFile, strFormName, "lblArchNameStart", lblArchNameStart.Caption)
    optArchNamePC.Caption = LocaliseString(strPathFile, strFormName, "optArchNamePC", optArchNamePC.Caption)
    optArchModelPC.Caption = LocaliseString(strPathFile, strFormName, "optArchModelPC", optArchModelPC.Caption)
    optArchCustom.Caption = LocaliseString(strPathFile, strFormName, "optArchCustom", optArchCustom.Caption)
    lblArchShablon.Caption = LocaliseString(strPathFile, strFormName, "lblArchShablon", lblArchShablon.Caption)
    lblArchMacrosType.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosType", lblArchMacrosType.Caption)
    lblArchMacrosParam.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosParam", lblArchMacrosParam.Caption)
    lblArchMacrosDescription.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosDescription", lblArchMacrosDescription.Caption)
    lblArchMacrosPCName.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosPCName", lblArchMacrosPCName.Caption)
    lblArchMacrosPCModel.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosPCModel", lblArchMacrosPCModel.Caption)
    lblArchMacrosOSVer.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosOSVer", lblArchMacrosOSVer.Caption)
    lblArchMacrosOSBit.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosOSBit", lblArchMacrosOSBit.Caption)
    lblArchMacrosDate.Caption = LocaliseString(strPathFile, strFormName, "lblArchMacrosDate", lblArchMacrosDate.Caption)
    'frMainTools
    frMainTools.Caption = LocaliseString(strPathFile, strFormName, "frMainTools", frMainTools.Caption)
    cmdPathDefault.Caption = LocaliseString(strPathFile, strFormName, "cmdPathDefault", cmdPathDefault.Caption)
    ' Сообщения диалогов выбора файлов и каталогов
    ucArch86Path.ToolTipTexts(ucOpen) = strMessages(151)
    ucArch86Path.DialogMsg(ucOpen) = strMessages(151)
    ucArch64Path.ToolTipTexts(ucOpen) = strMessages(151)
    ucArch64Path.DialogMsg(ucOpen) = strMessages(151)
    ucDebugLogPath.ToolTipTexts(ucFolder) = strMessages(152)
    ucDebugLogPath.DialogMsg(ucFolder) = strMessages(152)
    ucDPInst64Path.ToolTipTexts(ucOpen) = strMessages(151)
    ucDPInst64Path.DialogMsg(ucOpen) = strMessages(151)
    ucDPInst86Path.ToolTipTexts(ucOpen) = strMessages(151)
    ucDPInst86Path.DialogMsg(ucOpen) = strMessages(151)
    ucTempPath.ToolTipTexts(ucFolder) = strMessages(152)
    ucTempPath.DialogMsg(ucFolder) = strMessages(152)
    
End Sub

''!--------------------------------------------------------------------------------
''! Procedure   (Функция)   :   Sub lvOptions_ItemSelect
''! Description (Описание)  :   [При выборе опции происходит отображение соответсвующего окна]
''! Parameters  (Переменные):   item (LvwListItem), Selected (Boolean)
''!--------------------------------------------------------------------------------
Private Sub lvOptions_ItemSelect(ByVal item As LvwListItem, ByVal Selected As Boolean)

    If Selected Then
        Select Case item.Index
        
            Case 1
            'ItemOptions1=Основные настройки
                frMain.ZOrder 0
    
            Case 2
            ' ItemOptions2=Поддерживаемые ОС
                frOS.ZOrder 0
    
            Case 3
            'ItemOptions3=Рабочие утилиты
                frMainTools.ZOrder 0
    
            Case 4
            'ItemOptions4=Имя архива
                frArchName.ZOrder 0
    
            Case 5
            'ItemOptions5=Оформление программы
                frDesign.ZOrder 0
    
            Case 6
            'ItemOptions6=Параметры запуска DPInst
                frDpInstParam.ZOrder 0
                
            Case 7
            'ItemOptions7=Параметры запуска DPInst
                frDebug.ZOrder 0
    
            Case Else
                frOther.ZOrder 0
        End Select
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub lvOS_ColumnClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   ColumnHeader (LvwColumnHeader)
'!--------------------------------------------------------------------------------
Private Sub lvOS_ColumnClick(ByVal ColumnHeader As LvwColumnHeader)

    Dim ii As Long

    With lvOS
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
'! Procedure   (Функция)   :   Sub lvOS_ItemDblClick
'! Description (Описание)  :   [Двойнок клик по элементу списка вызывает форму редактирования]
'! Parameters  (Переменные):   Item (LvwListItem)
'                              Button (Integer)
'!--------------------------------------------------------------------------------
Private Sub lvOS_ItemDblClick(ByVal item As LvwListItem, ByVal Button As Integer)
    TransferOSData
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReadOptions
'! Description (Описание)  :   [Читаем настройки программы и заполняем поля]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ReadOptions()
    ' загрузить список ОС
    LoadList_lvOS
    ' Остальные параметры
    chkUpdate.Value = mbUpdateCheck
    chkUpdateBeta.Value = mbUpdateCheckBeta
    chkSilentDll.Value = mbSilentDLL
    chkRemoveTemp.Value = mbDelTmpAfterClose
    chkDebug.Value = mbDebugStandart
    chkDebugTime2File.Value = mbDebugTime2File
    chkDebugLog2AppPath.Value = mbDebugLog2AppPath
    ucDebugLogPath.Path = strDebugLogPathTemp
    txtDebugLogName.Text = strDebugLogNameTemp
    txtDebugLogLevel.Text = lngDetailMode
    chkFormMaximaze.Value = mbStartMaximazed
    chkFormSizeSave.Value = mbSaveSizeOnExit
    chkTempPath.Value = mbTempPath
    ucTempPath.Path = strAlternativeTempPath
    chkHideOtherProcess.Value = mbHideOtherProcess

    ' Режим при старте
    LoadComboList
    LoadStartMode
    ' Параметры выделения при старте
    chkCheckAll.Value = Abs(mbCheckAllGroup)
    chkHideOther.Value = Abs(mbListOnlyGroup)
    
    'MainForm
    txtFormHeight.Value = lngMainFormHeight
    txtFormWidth.Value = lngMainFormWidth

    'Пути к программам
    If mbPatnAbs Then
        'Секция DPInst
        ucDPInst86Path.Path = strDPInstExePath86
        ucDPInst64Path.Path = strDPInstExePath64
        'Секция Arc
        ucArch86Path.Path = strArh7zExePath86
        ucArch64Path.Path = strArh7zExePath64
        ucArchPathSFX.Path = strArh7zSFXPATH
        ucArchPathSFXConfig.Path = strArh7zSFXConfigPath
        ucArchPathSFXConfigEn.Path = strArh7zSFXConfigPathEn
    Else
        'Секция DPInst
        ucDPInst86Path.Path = Replace$(strDPInstExePath86, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucDPInst64Path.Path = Replace$(strDPInstExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        'Секция Arc
        ucArch86Path.Path = Replace$(strArh7zExePath86, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArch64Path.Path = Replace$(strArh7zExePath64, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFX.Path = Replace$(strArh7zSFXPATH, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFXConfig.Path = Replace$(strArh7zSFXConfigPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        ucArchPathSFXConfigEn.Path = Replace$(strArh7zSFXConfigPathEn, strAppPathBackSL, vbNullString, , , vbTextCompare)
    End If

    ' Настройки DpInst
    chkLegacyMode.Value = mbDpInstLegacyMode
    chkPromptIfDriverIsNotBetter.Value = mbDpInstPromptIfDriverIsNotBetter
    chkForceIfDriverIsNotBetter.Value = mbDpInstForceIfDriverIsNotBetter
    chkSuppressAddRemovePrograms.Value = mbDpInstSuppressAddRemovePrograms
    chkSuppressWizard.Value = mbDpInstSuppressWizard
    chkQuietInstall.Value = mbDpInstQuietInstall
    chkScanHardware.Value = mbDpInstScanHardware
    ' Другие настройки
    txtCmdStringDPInst = CollectCmdString
    ' Загрузка списка скинов
    LoadSkinListCombo cmbImageMain, strPathImageMain
    cmbImageMain.Text = strImageMainName
    ' изменение активности элементов
    DebugCtlEnable CBool(chkDebug.Value)
    DebugCtlEnableLog2App Not CBool(chkDebugLog2AppPath.Value)
    TempCtlEnable CBool(chkTempPath.Value)
    UpdateCtlEnable CBool(chkUpdate.Value)
    ' Имя архива при старте
    SelectStartArchName
    txtArchNameShablon.Text = strArchNameCustom
    ' Инициализация параметров для изменения шрифта и цвета элементов
    InitializeObjectProperties
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveOptions
'! Description (Описание)  :   [Сохранение настроек в ини-файл]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    Dim miRezim          As Long
    Dim miArchName       As Long
    Dim cnt              As Long
    Dim lngOSCountNew    As Long
    Dim strSysIniTemp As String
    Dim strLogNameTemp   As String

    If mbIsDriveCDRoom And Not mbLoadIniTmpAfterRestart Then
        If strSysIni <> strWorkTempBackSL & strSettingIniFile Then
            MsgBox strMessages(38), vbInformation + vbApplicationModal, strProductName

            Exit Sub

        End If

    ElseIf mbIsDriveCDRoom And mbLoadIniTmpAfterRestart Then
        strSysIniTemp = strWinTemp & "Settings_" & strProjectName & "_TMP.ini"
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
    IniWriteStrPrivate "Main", "SilentDLL", chkSilentDll.Value, strSysIniTemp
    IniWriteStrPrivate "Main", "IconMainSkin", cmbImageMain.Text, strSysIniTemp
    IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", Abs(mbLoadIniTmpAfterRestart), strSysIniTemp
    IniWriteStrPrivate "Main", "ArchMode", cmbTypeBackUp.ListIndex, strSysIni
    IniWriteStrPrivate "Main", "DisableDEP", Abs(mbDisableDEP), strSysIniTemp
   
    ' Секция Debug
    IniWriteStrPrivate "Debug", "DebugEnable", chkDebug.Value, strSysIniTemp
    IniWriteStrPrivate "Debug", "DebugLogPath", ucDebugLogPath.Path, strSysIniTemp
    strLogNameTemp = strProjectName & "-LOG_%DATE%.txt"

    If LenB(txtDebugLogName.Text) Then
        If InStr(txtDebugLogName.Text, strDot) Then
            strLogNameTemp = txtDebugLogName.Text
        End If
    End If

    IniWriteStrPrivate "Debug", "DebugLogName", strLogNameTemp, strSysIniTemp
    IniWriteStrPrivate "Debug", "CleenHistory", 1, strSysIniTemp
    IniWriteStrPrivate "Debug", "Detailmode", txtDebugLogLevel.Text, strSysIniTemp
    IniWriteStrPrivate "Debug", "DebugLog2AppPath", chkDebugLog2AppPath.Value, strSysIniTemp
    IniWriteStrPrivate "Debug", "Time2File", Abs(mbDebugTime2File), strSysIniTemp

    'Секция Arc
    IniWriteStrPrivate "Arc", "PathExe", ucArch86Path.Path, strSysIniTemp
    IniWriteStrPrivate "Arc", "PathExe64", ucArch64Path.Path, strSysIniTemp
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

    'Секция DPInst
    IniWriteStrPrivate "DPInst", "PathExe", ucDPInst86Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PathExe64", ucDPInst64Path.Path, strSysIniTemp
    IniWriteStrPrivate "DPInst", "LegacyMode", chkLegacyMode.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", chkPromptIfDriverIsNotBetter.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", chkForceIfDriverIsNotBetter.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", chkSuppressAddRemovePrograms.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "SuppressWizard", chkSuppressWizard.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "QuietInstall", chkQuietInstall.Value, strSysIniTemp
    IniWriteStrPrivate "DPInst", "ScanHardware", chkScanHardware.Value, strSysIniTemp
    'Секция OS
    'Число ОС
    lngOSCountNew = lvOS.ListItems.count
    IniWriteStrPrivate "OS", "OSCount", lngOSCountNew, strSysIniTemp

    'Заполяем в цикле подсекции ОС
    For cnt = 1 To lngOSCountNew

        'Секция OS_N
        With lvOS.ListItems(cnt)
            IniWriteStrPrivate "OS_" & cnt, "Ver", .Text, strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "is64bit", .SubItems(1), strSysIniTemp
            IniWriteStrPrivate "OS_" & cnt, "drpFolder", .SubItems(2), strSysIniTemp
        End With

    Next
    'Секция MainForm
    IniWriteStrPrivate "MainForm", "Width", txtFormWidth.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "Height", txtFormHeight.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "StartMaximazed", chkFormMaximaze.Value, strSysIniTemp
    mbSaveSizeOnExit = CBool(chkFormSizeSave.Value)
    IniWriteStrPrivate "MainForm", "SaveSizeOnExit", chkFormSizeSave.Value, strSysIniTemp
    IniWriteStrPrivate "MainForm", "FontName", strFontMainForm_Name, strSysIniTemp
    IniWriteStrPrivate "MainForm", "FontSize", lngFontMainForm_Size, strSysIniTemp
    IniWriteStrPrivate "MainForm", "HighlightColor", CStr(glHighlightColor), strSysIniTemp
    'Секция Buttons
    IniWriteStrPrivate "Button", "FontName", strFontBtn_Name, strSysIniTemp
    IniWriteStrPrivate "Button", "FontSize", miFontBtn_Size, strSysIniTemp
    IniWriteStrPrivate "Button", "FontUnderline", Abs(mbFontBtn_Underline), strSysIniTemp
    IniWriteStrPrivate "Button", "FontStrikethru", Abs(mbFontBtn_Strikethru), strSysIniTemp
    IniWriteStrPrivate "Button", "FontItalic", Abs(mbFontBtn_Italic), strSysIniTemp
    IniWriteStrPrivate "Button", "FontBold", Abs(mbFontBtn_Bold), strSysIniTemp
    IniWriteStrPrivate "Button", "FontColor", CStr(cmdFutureButton.ForeColor), strSysIniTemp
    IniWriteStrPrivate "Button", "Style", cmbButtonStyle.ListIndex, strSysIniTemp
    IniWriteStrPrivate "Button", "StyleColor", cmbButtonStyleColor.ListIndex, strSysIniTemp
    IniWriteStrPrivate "Button", "BackColor", ctlStatusBtnBackColor.Value, strSysIniTemp
    
    ' Приводим Ini файл к читабельному виду
    NormIniFile strSysIniTemp
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectStartArchName
'! Description (Описание)  :   [Режим архива при старте]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SelectStartMode
'! Description (Описание)  :   [Режим при старте]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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
'! Procedure   (Функция)   :   Sub TempCtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub TempCtlEnable(ByVal mbEnable As Boolean)
    ucTempPath.Enabled = mbEnable
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub TransferOSData
'! Description (Описание)  :   [Передача параметров ОС из спика в форму редактирования]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub TransferOSData()

    Dim ii As Long

    With lvOS
        ii = .SelectedItem.Index

        If ii >= 0 Then

            frmOSEdit.txtOSVer.Text = .ListItems.item(ii).Text
            frmOSEdit.ucPathDRP.Path = .ListItems.item(ii).SubItems(2)
            frmOSEdit.chk64bit.Value = CBool(.ListItems.item(ii).SubItems(1))

            frmOSEdit.Show vbModal, Me
        End If
        
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optArchCustom_Click
'! Description (Описание)  :   [Выбор режима имени архива]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optArchCustom_Click()
    txtArchNameShablon.Enabled = optArchCustom.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optArchModelPC_Click
'! Description (Описание)  :   [Выбор режима имени архива]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optArchModelPC_Click()
    txtArchNameShablon.Enabled = optArchCustom.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub optArchNamePC_Click
'! Description (Описание)  :   [Выбор режима имени архива]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub optArchNamePC_Click()
    txtArchNameShablon.Enabled = optArchCustom.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchNameShablon_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtArchNameShablon_GotFocus()

    HighlightActiveControl Me, txtArchNameShablon, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchNameShablon_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtArchNameShablon_LostFocus()

    HighlightActiveControl Me, txtArchNameShablon, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtCmdStringDPInst_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtCmdStringDPInst_GotFocus()
    HighlightActiveControl Me, txtCmdStringDPInst, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtCmdStringDPInst_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtCmdStringDPInst_LostFocus()
    HighlightActiveControl Me, txtCmdStringDPInst, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtDebugLogName_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtDebugLogName_GotFocus()
    HighlightActiveControl Me, txtDebugLogName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtDebugLogName_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtDebugLogName_LostFocus()
    HighlightActiveControl Me, txtDebugLogName, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchMacrosDate_DblClick
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtArchMacrosDate_DblClick()

    txtArchMacrosDate.SelStart = 0
    txtArchMacrosDate.SelLength = Len(txtArchMacrosDate.Text)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchMacrosOSBit_DblClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtArchMacrosOSBit_DblClick()

    txtArchMacrosOSBIT.SelStart = 0
    txtArchMacrosOSBIT.SelLength = Len(txtArchMacrosOSBIT.Text)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchMacrosOSVer_DblClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtArchMacrosOSVer_DblClick()

    txtArchMacrosOSVER.SelStart = 0
    txtArchMacrosOSVER.SelLength = Len(txtArchMacrosOSVER.Text)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchMacrosPCModel_DblClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtArchMacrosPCModel_DblClick()

    txtArchMacrosPCModel.SelStart = 0
    txtArchMacrosPCModel.SelLength = Len(txtArchMacrosPCModel.Text)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtArchMacrosPCName_DblClick
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtArchMacrosPCName_DblClick()

    txtArchMacrosPCName.SelStart = 0
    txtArchMacrosPCName.SelLength = Len(txtArchMacrosPCName.Text)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFXConfigEn_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFXConfigEn_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPathSFXConfigEn_GotFocus()

    HighlightActiveControl Me, ucArchPathSFXConfigEn, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFXConfigEn_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPathSFXConfigEn_LostFocus()

    HighlightActiveControl Me, ucArchPathSFXConfigEn, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFXConfig_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFXConfig_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPathSFXConfig_GotFocus()

    HighlightActiveControl Me, ucArchPathSFXConfig, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFXConfig_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPathSFXConfig_LostFocus()

    HighlightActiveControl Me, ucArchPathSFXConfig, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFX_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFX_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPathSFX_GotFocus()

    HighlightActiveControl Me, ucArchPathSFX, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArchPathSFX_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArchPathSFX_LostFocus()

    HighlightActiveControl Me, ucArchPathSFX, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArch64Path_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArch64Path_Click()

    Dim strTempPath As String

    With ucArch64Path
        If .FileCount Then
            strTempPath = .FileName
    
            If InStr(1, strTempPath, strAppPathBackSL, vbTextCompare) Then
                strTempPath = Replace$(strTempPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
            End If
        End If
    
        If LenB(strTempPath) Then
            .Path = strTempPath
        End If
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArch64Path_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArch64Path_GotFocus()
    HighlightActiveControl Me, ucArch64Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArch64Path_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArch64Path_LostFocus()
    HighlightActiveControl Me, ucArch64Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArch86Path_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArch86Path_Click()

    Dim strTempPath As String

    With ucArch86Path
        If .FileCount Then
            strTempPath = .FileName
    
            If InStr(1, strTempPath, strAppPathBackSL, vbTextCompare) Then
                strTempPath = Replace$(strTempPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
            End If
        End If
    
        If LenB(strTempPath) Then
            .Path = strTempPath
        End If
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArch86Path_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArch86Path_GotFocus()
    HighlightActiveControl Me, ucArch86Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucArch86Path_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucArch86Path_LostFocus()
    HighlightActiveControl Me, ucArch86Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDebugLogPath_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_Click()

    Dim strTempPath As String

    With ucDebugLogPath
        strTempPath = .FileName

        If InStr(1, strTempPath, strAppPathBackSL, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        End If
    
        If LenB(strTempPath) Then
            .Path = strTempPath
        End If
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDebugLogPath_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_GotFocus()
    HighlightActiveControl Me, ucDebugLogPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDebugLogPath_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDebugLogPath_LostFocus()
    HighlightActiveControl Me, ucDebugLogPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst64Path_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_Click()

    Dim strTempPath As String

    With ucDPInst64Path
        If .FileCount Then
            strTempPath = .FileName
    
            If InStr(1, strTempPath, strAppPathBackSL, vbTextCompare) Then
                strTempPath = Replace$(strTempPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
            End If
        End If
    
        If LenB(strTempPath) Then
            .Path = strTempPath
        End If
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst64Path_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_GotFocus()
    HighlightActiveControl Me, ucDPInst64Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst64Path_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst64Path_LostFocus()
    HighlightActiveControl Me, ucDPInst64Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst86Path_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_Click()

    Dim strTempPath As String

    With ucDPInst86Path
        If .FileCount Then
            strTempPath = .FileName
    
            If InStr(1, strTempPath, strAppPathBackSL, vbTextCompare) Then
                strTempPath = Replace$(strTempPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
            End If
        End If
    
        If LenB(strTempPath) Then
            .Path = strTempPath
        End If
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst86Path_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_GotFocus()
    HighlightActiveControl Me, ucDPInst86Path, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucDPInst86Path_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucDPInst86Path_LostFocus()
    HighlightActiveControl Me, ucDPInst86Path, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucTempPath_Click
'! Description (Описание)  :   [выбор каталога или файла]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_Click()

    Dim strTempPath As String

    With ucTempPath
        strTempPath = .Path

        If InStr(1, strTempPath, strAppPathBackSL, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        End If
    
        If LenB(strTempPath) Then
            .Path = strTempPath
        End If
    End With
    
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucTempPath_GotFocus
'! Description (Описание)  :   [Элемент в фокусе]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_GotFocus()
    HighlightActiveControl Me, ucTempPath, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ucTempPath_LostFocus
'! Description (Описание)  :   [Элемент вне фокуса]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ucTempPath_LostFocus()
    HighlightActiveControl Me, ucTempPath, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub UpdateCtlEnable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub UpdateCtlEnable(ByVal mbEnable As Boolean)
    chkUpdateBeta.Enabled = mbEnable
End Sub

