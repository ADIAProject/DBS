VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "frmMain"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12480
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
   ScaleHeight     =   7185
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin prjDIADBS.ctlUcStatusBar ctlUcStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   15
      Top             =   6480
      Width           =   12480
      _ExtentX        =   22013
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
      Theme           =   2
   End
   Begin prjDIADBS.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      Top             =   5985
      Visible         =   0   'False
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   873
      Max             =   10000
   End
   Begin prjDIADBS.ctlJCFrames frPanel 
      Height          =   6120
      Left            =   0
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
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
         Left            =   120
         Top             =   75
         Width           =   5895
         _ExtentX        =   10398
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
         Caption         =   "��������� ������ ���������"
         TextBoxHeight   =   21
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin prjDIADBS.CheckBoxW chkHideOther 
            Height          =   400
            Left            =   75
            TabIndex        =   7
            Top             =   1560
            Width           =   3495
            _ExtentX        =   6165
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
            Height          =   255
            Left            =   75
            TabIndex        =   2
            Top             =   500
            Width           =   1600
            _ExtentX        =   2831
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
            Caption         =   "frmMain.frx":0072
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.OptionButtonW optGrp2 
            Height          =   255
            Left            =   1800
            TabIndex        =   3
            Top             =   500
            Width           =   1600
            _ExtentX        =   2831
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
            Caption         =   "frmMain.frx":00A4
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.OptionButtonW optGrp3 
            Height          =   255
            Left            =   75
            TabIndex        =   4
            Top             =   850
            Width           =   1600
            _ExtentX        =   2831
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
            Caption         =   "frmMain.frx":00CA
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.OptionButtonW optGrp4 
            Height          =   255
            Left            =   1800
            TabIndex        =   5
            Top             =   850
            Width           =   1600
            _ExtentX        =   2831
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
            Caption         =   "frmMain.frx":00F0
            Transparent     =   -1  'True
         End
         Begin prjDIADBS.ctlJCbutton cmdCheckAll 
            Height          =   510
            Left            =   3720
            TabIndex        =   8
            Top             =   500
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "�������� ��"
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   1
         End
         Begin prjDIADBS.ctlJCbutton cmdUnCheckAll 
            Height          =   510
            Left            =   3720
            TabIndex        =   9
            Top             =   1100
            Width           =   2055
            _ExtentX        =   3625
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
            Caption         =   "����� ���������"
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   1
         End
         Begin prjDIADBS.CheckBoxW chkCheckAll 
            Height          =   400
            Left            =   75
            TabIndex        =   6
            Top             =   1200
            Width           =   3495
            _ExtentX        =   6165
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
         Left            =   6120
         Top             =   75
         Width           =   6255
         _ExtentX        =   11033
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
         Caption         =   "�������� ��������� ����� ��������� ���������"
         TextBoxHeight   =   21
         HeaderStyle     =   1
         GradientHeaderStyle=   2
         Begin prjDIADBS.ComboBoxW cmbTypeBackUp 
            Height          =   330
            Left            =   1800
            TabIndex        =   10
            Top             =   500
            Width           =   4335
            _ExtentX        =   7646
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
            Text            =   "frmMain.frx":017E
            CueBanner       =   "frmMain.frx":019E
         End
         Begin prjDIADBS.ctlJCbutton cmdStartBackUp 
            Height          =   510
            Left            =   3960
            TabIndex        =   0
            Top             =   925
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
            Caption         =   "Start Backup"
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
            ColorScheme     =   1
         End
         Begin prjDIADBS.ctlJCbutton cmdBreak 
            Height          =   510
            Left            =   3960
            TabIndex        =   1
            Top             =   1500
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
            Enabled         =   0   'False
            BackColor       =   12244692
            Caption         =   "Break"
            PictureAlign    =   0
            PicturePushOnHover=   -1  'True
            PictureShadow   =   -1  'True
            CaptionEffects  =   0
            TooltipBackColor=   0
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
            Caption         =   "��� ������"
            Alignment       =   0
            Begin prjDIADBS.TextBoxW txtArchName 
               Height          =   350
               Left            =   120
               TabIndex        =   14
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
               Text            =   "frmMain.frx":01BE
               CueBanner       =   "frmMain.frx":01DE
            End
            Begin prjDIADBS.OptionButtonW optArchModelPC 
               Height          =   255
               Left            =   1800
               TabIndex        =   13
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
               Caption         =   "frmMain.frx":01FE
               Transparent     =   -1  'True
            End
            Begin prjDIADBS.OptionButtonW optArchNamePC 
               Height          =   255
               Left            =   1800
               TabIndex        =   12
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
               Caption         =   "frmMain.frx":0240
               Transparent     =   -1  'True
            End
            Begin prjDIADBS.OptionButtonW optArchCustom 
               Height          =   325
               Left            =   120
               TabIndex        =   11
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
               Caption         =   "frmMain.frx":027C
               Transparent     =   -1  'True
            End
         End
         Begin prjDIADBS.LabelW lblTypeBackUp 
            Height          =   405
            Left            =   75
            TabIndex        =   16
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
            Caption         =   "Type of backup:"
         End
      End
      Begin prjDIADBS.ctlJCFrames frPanelLV 
         Height          =   3690
         Left            =   120
         Top             =   2280
         Width           =   12255
         _ExtentX        =   21616
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
         Caption         =   "������ ��������� ��������� ���������"
         TextBoxHeight   =   21
         ThemeColor      =   3
         GradientHeaderStyle=   1
      End
   End
   Begin VB.Menu mnuReCollectHWID 
      Caption         =   "�������� ����������"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "���������"
   End
   Begin VB.Menu mnuMainAbout 
      Caption         =   "�������"
      Begin VB.Menu mnuLinks 
         Caption         =   "������"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "������� ���������"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "������� �� ������"
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHomePage 
         Caption         =   "�������� �������� ���������"
      End
      Begin VB.Menu mnuHomePageForum 
         Caption         =   "����� ���������"
      End
      Begin VB.Menu mnuOsZoneNet 
         Caption         =   "���������� ��������� �� OsZone.net"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUpd 
         Caption         =   "��������� ���������� ���������"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModulesVersion 
         Caption         =   "������..."
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDonate 
         Caption         =   "������������� ������..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "� ���������..."
      End
   End
   Begin VB.Menu mnuMainLang 
      Caption         =   "����"
      Begin VB.Menu mnuLangStart 
         Caption         =   "������������ ��������� ���� ��� ������� (������ ����������)"
      End
      Begin VB.Menu mnuSep17 
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

Public WithEvents lvDevices       As cListView
Attribute lvDevices.VB_VarHelpID = -1

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
Private arrSourceDisksFiles()     As String
Private arrSourceDisksNames()     As String

Private strFormName               As String
Private lngFrameTime              As Long
Private lngFrameCount             As Long
Private lngBorderWidthX           As Long
Private lngBorderWidthY           As Long


'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub BlockControl
'! Description (��������)  :   [����������(�������������) ��������� ��������� ����� ��� ������ ������� �������]
'! Parameters  (����������):   mbEnable (Boolean)
'!--------------------------------------------------------------------------------
Private Sub BlockControl(ByVal mbEnable As Boolean)

    'Filter
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
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ChangeFrmMainCaption
'! Description (��������)  :   [��������� Caption �����]
'! Parameters  (����������):   lngPercentage (Long)
'!--------------------------------------------------------------------------------
Private Sub ChangeFrmMainCaption(Optional ByVal lngPercentage As Long)

    Dim strProgressValue As String

    Select Case strPCLangCurrentID

        Case "0419"
            strFrmMainCaptionTemp = "Drivers Backup Solution"
            strFrmMainCaptionTempDate = " (���� ������: "

        Case Else
            strFrmMainCaptionTemp = "Drivers Backup Solution"
            strFrmMainCaptionTempDate = " (Date Build: "
    End Select

    If lngPercentage Mod 999 Then
        If ProgressBar1.Visible Then
            strProgressValue = (lngPercentage \ 10) & "% (" & ctlUcStatusBar1.PanelText(1) & ") - "
        End If
    End If

    If LenB(strThisBuildBy) = 0 Then
        Me.Caption = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & strFrmMainCaptionTempDate & " @" & App.CompanyName
    Else
        Me.Caption = strProgressValue & strFrmMainCaptionTemp & " v." & strProductVersion & strFrmMainCaptionTempDate & " " & strThisBuildBy
    End If

End Sub

Private Sub chkHideOther_Click()

    chkCheckAll.Enabled = CBool(chkHideOther.Value)

    If optGrp1.Value Then
        optGrp1_Click
    End If

    If optGrp2.Value Then
        optGrp2_Click
    End If

    If optGrp3.Value Then
        optGrp3_Click
    End If

    If optGrp4.Value Then
        optGrp4_Click
    End If
End Sub

Private Sub cmdBreak_Click()

    mbBreakUpdateDBAll = True
End Sub

Private Sub cmdCheckAll_Click()

    Dim i As Integer

    With lvDevices

        For i = 0 To .Count

            If Not .ItemChecked(i) Then
                .ItemChecked(i) = True
            End If

        Next
    End With

    FindCheckCountList
End Sub

'# do backup of selected drivers #
Private Sub cmdStartBackUp_Click()

    ' ���������� �����
    StartBackUp
End Sub

Private Sub cmdUnCheckAll_Click()

    Dim i As Integer

    With lvDevices

        For i = 0 To .Count

            If .ItemChecked(i) Then
                .ItemChecked(i) = False
            End If

        Next
    End With

    FindCheckCountList
End Sub

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

    DebugMode "***CollectDestPathFiles-Start", 2
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
    arrSourceDisksFiles = arr2_SDF
    arrSourceDisksNames = arr1_SDN
    DebugMode "***CollectDestPathFiles-Finish", 2
End Sub

' ��� ������ 7z
Private Function CollectDpName(ByVal strPcName As String) As String

    Dim strDpName       As String
    Dim strDPName_Part1 As String
    Dim strDPName_Part2 As String
    Dim strDPName_Part3 As String

    strDPName_Part1 = "_wnt" & Mid$(strOsCurrentVersion, 1, 1)

    If mbIsWin64 Then
        strDPName_Part2 = "_x64_"
    Else
        strDPName_Part2 = "_x32_"
    End If

    strDPName_Part3 = Replace$(CStr(Date), ".", "-")
    strDPName_Part3 = SafeDir(strDPName_Part3)
    strDpName = "DP_" & strPcName & strDPName_Part1 & strDPName_Part2 & strDPName_Part3
    strDpName = SafeDir(strDpName)
    CollectDpName = Replace$(strDpName, " ", "_")
End Function

Private Sub CopyFile2Dest(ByRef arrZ() As String, _
                          ByVal strDestination As String, _
                          ByVal strDestFolderSection As String, _
                          ByVal strInfFile As String, _
                          Optional ByVal mbSectCopyFiles As Boolean = False)

    Dim strFileName        As String
    Dim strFileName_x()    As String
    Dim strFileNameFrom    As String
    Dim strFileNameTo      As String
    Dim strDestPath4File   As String
    Dim D                  As Long
    Dim ext                As String
    Dim cDir               As String
    Dim customDir          As String
    Dim OldValue           As Long
    Dim strDestinationTemp As String
    Dim lngArrCount        As Long
    Dim lngUBoundZ         As Long
    Dim lngUBoundFileName  As Long

    lngUBoundZ = UBound(arrZ)

    For D = 0 To lngUBoundZ
        strFileName = arrZ(D)

        ' ���� ������ ��������, �� ����������
        If LenB(strFileName) > 0 Then
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

            ' ������� ��� ������ �� ����� �����
            strFileName = SafeFileName(strFileName)

            ' ���� ������ ��������, �� ����������
            If LenB(strFileName) > 0 Then

                ' ���� ������ �������� ".", ������ ��� ������ ��� ��� �����
                If InStr(1, strFileName, ".") Then

                    ' ���� ����� ���������� ����
                    Dim lngUBound As Long

                    lngUBound = UBound(arrSourceDisksFiles, 1)

                    For lngArrCount = 1 To lngUBound

                        If StrComp(arrSourceDisksFiles(lngArrCount, 1), strFileName, vbTextCompare) = 0 Then
                            strDestinationTemp = arrSourceDisksFiles(lngArrCount, 2)
                            strDestinationTemp = PathCollect4Dest(strDestinationTemp, strDestination)
                            Exit For
                        Else
                            strDestinationTemp = strDestination
                        End If

                    Next

                    ' ������� ������� ����������, ���� ��� ���
                    If PathExists(strDestinationTemp) = False Then
                        CreateNewDirectory strDestinationTemp
                    End If

                    ' ���������� ������ ���� ����������� �����
                    If LenB(strFileNameTo) > 0 Then
                        If mbSectCopyFiles Then
                            strDestPath4File = BackslashAdd2Path(strDestinationTemp) & strFileNameTo
                        Else
                            strDestPath4File = BackslashAdd2Path(strDestinationTemp) & strFileName
                        End If

                    Else
                        strDestPath4File = BackslashAdd2Path(strDestinationTemp) & strFileName
                    End If

                    ' ���������� �������, ��� ������ ������ ���� �� ��������� ����
                    customDir = ReadFromINI("DestinationDirs", strDestFolderSection, strInfFile, vbNullString)

                    '���� ������� �� ���������, �� ���������� ������� �� �������
                    If LenB(customDir) = 0 Then
                        customDir = ReadFromINI("DestinationDirs", "DefaultDestDir", strInfFile, vbNullString)
                    End If

                    '���� ��� ����� �� ���������, �� ����������
                    If LenB(customDir) > 0 Then
                        '# if it is #
                        cDir = WhereIsDir(customDir, strInfFile)

                        ' ���� x64, �� ������������� ���������� ��������������� ��� ����� system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64DisableWow64FsRedirection", "kernel32.dll") Then
                                Wow64DisableWow64FsRedirection OldValue
                            End If
                        End If

                        ' ����������� �����
                        strFileNameFrom = cDir & strFileName

                        If PathExists(strFileNameFrom) Then
                            If PathExists(strDestPath4File) = False Then
                                CopyFileTo cDir & strFileName, strDestPath4File
                                DebugMode "******Backup File: FROM=" & strFileNameFrom & " TO=" & strDestPath4File
                            End If
                        End If

                        ' ���� ��� �������� ��������, �� ���� �� ���� �����
                        If InStr(1, cDir, strSysDir86 & "spool\Drivers\w32x86", vbTextCompare) > 0 Then

                            '# search for correctly driver if has more tha one printer #
                            ' ���� ���� �� ���� ����� strSysDir & "\spool\Drivers\w32x86"
                            If PathExists(strDestPath4File) = False Then
                                strFileNameFrom = CStr(SearchFilesInRoot(cDir, strFileName, True, True))

                                If LenB(strFileNameFrom) > 0 Then
                                    CopyFileTo strFileNameFrom, strDestPath4File
                                End If
                            End If
                        End If

                        ' ���� x64, �� �������� ������� ��������������� ��� ����� system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64RevertWow64FsRedirection", "kernel32.dll") Then
                                'Wow64DisableWow64FsRedirection OldValue
                                Wow64RevertWow64FsRedirection OldValue
                            End If
                        End If
                    End If

                    ' �������������� ����� ������ �� ����������, ���� ���� ��� ��� �� ������
                    If PathExists(strDestPath4File) = False Then
                        '���������� �����
                        ext = ExtFromFileName(strFileName)

                        ' ���� x64, �� ������������� ���������� ��������������� ��� ����� system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64DisableWow64FsRedirection", "kernel32.dll") Then
                                Wow64DisableWow64FsRedirection OldValue
                            End If
                        End If

                        If ext = "hlp" Then
                            If PathExists(BackslashAdd2Path(strWinDirHelp) & strFileName) Then
                                CopyFileTo BackslashAdd2Path(strWinDirHelp) & strFileName, strDestPath4File
                            End If

                        ElseIf ext = "sys" Then

                            If PathExists(strSysDirDrivers & strFileName) Then
                                CopyFileTo strSysDirDrivers & strFileName, strDestPath4File
                            End If

                            If PathExists(strSysDirDrivers64 & strFileName) Then
                                CopyFileTo strSysDirDrivers64 & strFileName, strDestPath4File
                            End If

                        Else

                            If PathExists(strSysDir86 & strFileName) Then
                                CopyFileTo strSysDir86 & strFileName, strDestPath4File
                            End If

                            If PathExists(strSysDir64 & strFileName) Then
                                CopyFileTo strSysDir64 & strFileName, strDestPath4File
                            End If
                        End If

                        ' ���� x64, �� �������� ������� ��������������� ��� ����� system32
                        If mbIsWin64 Then
                            If APIFunctionPresent("Wow64RevertWow64FsRedirection", "kernel32.dll") Then
                                Wow64RevertWow64FsRedirection OldValue
                            End If
                        End If
                    End If
                End If
            End If
        End If

    Next
End Sub

'! -----------------------------------------------------------
'!  �������     :  CreateMenuLngIndex
'!  ����������  :  Name As String
'!  ��������    :
'! -----------------------------------------------------------
Private Sub CreateMenuLngIndex(ByVal strName As String)

    Dim i As Long

    On Error Resume Next

    If Not mnuLang(0).Visible Then
        '���� ���� ��� �� �������
        mnuLang(0).Visible = True
        mnuLang(0).Caption = strName
    Else
        Load mnuLang(mnuLang.Count)
        mnuLang(mnuLang.Count - 1).Visible = True

        For i = mnuLang.UBound To mnuLang.LBound Step -1

            If i = mnuLang.LBound Then
                mnuLang(0).Caption = strName
                Exit For
            End If

            mnuLang(i).Caption = mnuLang(i - 1).Caption
        Next
    End If

    On Error GoTo 0

End Sub

Private Function DefineFolderBackUp() As String

    Dim i                 As Long
    Dim strDestFolder     As String
    Dim strDestFolderTemp As String
    Dim str_x64           As String

    If mbBackFolderPredefine Then

        For i = 0 To UBound(arrOSList)
            str_x64 = arrOSList(i, 1)
            strDestFolderTemp = arrOSList(i, 2)

            If InStr(1, arrOSList(i, 0), strOsCurrentVersion) Then
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

    If LenB(strDestFolder) > 0 Then
        DefineFolderBackUp = strDestFolder
    Else
        DefineFolderBackUp = strAppPathBackSL & "drivers\"
    End If
End Function

Private Function DoZip(ByVal strPackFolder As String, ByVal strDpName As String) As Boolean

    Dim cmdString             As String
    Dim strDpName7z           As String
    Dim strDpNameExt          As String
    Dim strDpNamewoExt        As String
    Dim mbCreateSFX        As Boolean
    Dim strDPInstPath         As String
    Dim lngNumFilesFromFolder As Long

    ' �������� ���������� ����� ������ (exe ��� 7Z)
    strDpNameExt = ExtFromFileName(strDpName)
    strDpNamewoExt = FileName_woExt(strDpName)

    If StrComp(strDpNameExt, "exe", vbTextCompare) = 0 Then
        strDpName7z = strDpNamewoExt & ".7z"
        mbCreateSFX = True
    Else
        strDpName7z = strDpName
    End If

    ' ������� ������ ������ ���� ����
    If PathExists(strDpName7z) Then
        DebugMode "***DoZip: Clean previous drivers archive "
        DeleteFiles strDpName7z
    End If

    If mbCreateSFX Then
        If PathExists(strDpName) Then
            DebugMode "***DoZip: Clean previous drivers archive "
            DeleteFiles strDpName
        End If

        ' �������� ����� DPInst ��� �����������
        strDPInstPath = PathNameFromPath(strDPInstExePath)
        DebugMode "******CopyFiles DPINST : " & strDPInstPath
        ChangeStatusTextAndDebug "Copying files from DPInst folder: " & strDPInstPath
        lngNumFilesFromFolder = rgbCopyFiles(strDPInstPath, strPackFolder, ALL_FILES)
        DebugMode "******CopyFiles - count files: " & lngNumFilesFromFolder
    End If

    ' ������ ������ ��������
    '..\7za.exe a ..\out\%1 -mmt=off -m0=BCJ2 -m1=LZMA2:d%dict%m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 *.ini -ir!*.inf
    cmdString = Kavichki & strArh7zExePATH & Kavichki & " a " & Kavichki & strDpName7z & Kavichki & " " & strArh7zParam1
    ChangeStatusTextAndDebug strMessages(97) & " " & strDpName7z, "Compressing...: " & cmdString

    If RunAndWait(cmdString, strPackFolder, vbHide) = False Then
        MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
        DoZip = False
        ChangeStatusTextAndDebug strMessages(13) & " " & strDpName7z, "Error on run : " & cmdString
    Else

        ' ��������� ��������� �� ��� 100%? ���� ��� �� ��������
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusTextAndDebug strMessages(13) & strDpName7z
            MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
            DoZip = False
        End If

        DoZip = True
        ChangeStatusTextAndDebug "7z-archive (STEP 1) successfully done!!!"
    End If

    ' ������ ������ ��������
    '..\7za.exe a ..\out\%1 -mmt=off -m0=BCJ2 -m1=LZMA2:d%dict%m:fb273 -m2=LZMA2:d512k -m3=LZMA2:d512k -mb0:1 -mb0s1:2 -mb0s2:3 -xr!*.inf -x!*.ini
    cmdString = Kavichki & strArh7zExePATH & Kavichki & " a " & Kavichki & strDpName7z & Kavichki & " " & strArh7zParam2
    ChangeStatusTextAndDebug strMessages(97) & " " & strDpName7z, "Compressing...: " & cmdString

    If RunAndWait(cmdString, strPackFolder, vbHide) = False Then
        MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
        DoZip = False
        ChangeStatusTextAndDebug strMessages(13) & " " & strDpName7z, "Error on run : " & cmdString
    Else

        ' ��������� ��������� �� ��� 100%? ���� ��� �� ��������
        If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
            ChangeStatusTextAndDebug strMessages(13) & strDpName7z
            MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
            DoZip = False
        End If

        DoZip = True
        ChangeStatusTextAndDebug "7z-archive (STEP 2) successfully done!!!"
    End If

    If mbCreateSFX Then

        ' ������ ������ �������� SFX
        'copy /b "d:\aWork\myProg\DriversBackuper\Tools\Arc\sfx\7zSD.sfx" + "d:\aWork\myProg\DriversBackuper\Tools\Arc\sfx\config.txt" + "D:\aWork\myProg\DriversBackuper\drivers\2k_xp_2003\x64\DP_0300-B01951_wnt5_x32_03-03-2011.7z" "D:\aWork\myProg\DriversBackuper\drivers\2k_xp_2003\x64\DP_0300-B01951_wnt5_x32_03-03-2011.exe"
        Select Case strPCLangCurrentID

            Case "0419"
                cmdString = "cmd.exe /C copy /b " & Kavichki & strArh7zSFXPATH & Kavichki & " + " & Kavichki & strArh7zSFXConfigPath & Kavichki & " + " & Kavichki & strDpName7z & Kavichki & " " & Kavichki & strDpName & Kavichki

            Case Else
                cmdString = "cmd.exe /C copy /b " & Kavichki & strArh7zSFXPATH & Kavichki & " + " & Kavichki & strArh7zSFXConfigPathEn & Kavichki & " + " & Kavichki & strDpName7z & Kavichki & " " & Kavichki & strDpName & Kavichki
        End Select

        ChangeStatusTextAndDebug strMessages(97) & " " & strDpName, "Creating SFX...: " & cmdString

        If RunAndWait(cmdString, strWorkTemp, vbHide) = False Then
            MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
            DoZip = False
            ChangeStatusTextAndDebug strMessages(13) & " " & strDpName, "Error on run : " & cmdString
        Else

            If PathExists(strDpName) Then
                If PathExists(strDpName7z) Then
                    DebugMode "***DoZip: Clean temp drivers archive "
                    DeleteFiles strDpName7z
                End If

                DoZip = True
                ChangeStatusTextAndDebug "7z-archive (STEP 3) successfully done!!! SFX-archive created"
            Else
                MsgBox strMessages(13) & vbNewLine & vbNewLine & cmdString, vbInformation, strProductName
                DoZip = False
                ChangeStatusTextAndDebug strMessages(13) & " " & strDpName, "Error on run : " & cmdString
            End If
        End If
    End If
End Function

Private Function ExpandArchNamebyEnvironment(ByVal strArchName As String) As String

    Dim R               As String
    Dim strDPName_OSVer As String
    Dim strDPName_OSBit As String
    Dim strDPName_DATE  As String

    If InStr(1, strArchName, "%") Then
        ' ���������������� ������ �� %OSVer%
        strDPName_OSVer = "wnt" & Mid$(strOsCurrentVersion, 1, 1)

        ' ���������������� �������� �� %OSBit%
        If mbIsWin64 Then
            strDPName_OSBit = "x64"
        Else
            strDPName_OSBit = "x32"
        End If

        ' ���������������� ���� %DATE%
        strDPName_DATE = Replace$(CStr(Date), ".", "-")
        strDPName_DATE = SafeDir(strDPName_DATE)
        ' ������ �������� ����������
        R = strArchName
        R = Replace$(R, "%PCNAME%", strCompName)
        R = Replace$(R, "%PCMODEL%", Replace$(strCompModel, " ", "_"))
        R = Replace$(R, "%OSVer%", strDPName_OSVer)
        R = Replace$(R, "%OSBit%", strDPName_OSBit)
        R = Replace$(R, "%DATE%", strDPName_DATE)
        R = Trim$(R)
        ExpandArchNamebyEnvironment = R
    Else
        ExpandArchNamebyEnvironment = strArchName
    End If
End Function

Private Sub FindCheckCountList()

    Dim miCount As Integer

    miCount = lvDevices.CheckedCount
    cmdStartBackUp.Caption = LocaliseString(strPCLangCurrentPath, Me.Name, "cmdStartBackUp", "Start Backup")

    If miCount > 0 Then

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
End Sub

Private Function FindCopyCatFile(ByVal strInfFilePath As String, ByVal strDestination As String) As String

    Dim strCatFile         As String
    Dim strCatFile_ntx86   As String
    Dim strCatFile_ntamd64 As String
    Dim strCatFile_nt      As String
    Dim strCatFilePath     As String
    Dim strCatFileFromInf  As String
    Dim mbExitGoto      As Boolean

    '# ���� � ����� inf - catalog file (������� ������������)
    strCatFile = ReadFromINI("Version", "CatalogFile", strInfFilePath, vbNullString)
    strCatFile_nt = ReadFromINI("Version", "CatalogFile.nt", strInfFilePath, vbNullString)
    strCatFile_ntx86 = ReadFromINI("Version", "CatalogFile.ntx86", strInfFilePath, vbNullString)
    strCatFile_ntamd64 = ReadFromINI("Version", "CatalogFile.ntamd64", strInfFilePath, vbNullString)
    strCatFile = SafeFileName(strCatFile)

    If LenB(strCatFile) = 0 Then
        If LenB(strCatFile_ntx86) > 0 Then
            strCatFile = strCatFile_ntx86
        ElseIf LenB(strCatFile_ntamd64) > 0 Then
            strCatFile = strCatFile_ntamd64
        ElseIf LenB(strCatFile_nt) > 0 Then
            strCatFile = strCatFile_nt
        Else
            strCatFile = vbNullString
        End If
    End If

    strCatFileFromInf = FileName_woExt(FileNameFromPath(strInfFilePath)) & ".cat"
CopyCatAgain:

    '# if has catalog file #
    If LenB(strCatFile) > 0 Then

        ' ���� ���� cat ��� �� ���� ����� strSysDirCatRoot c ������ �� ���������� �� ����� inf
        If PathExists(BackslashAdd2Path(strDestination) & strCatFile) = False Then
            strCatFilePath = CStr(SearchFilesInRoot(strSysDirCatRoot, strCatFile, True, True))

            If LenB(strCatFilePath) > 0 Then
                CopyFileTo strCatFilePath, BackslashAdd2Path(strDestination) & strCatFile
                DebugMode "***CatalogFile find in: " & strCatFilePath
            End If
        End If

        ' ���� ���� cat ��� �� ���� ����� strSysDirCatRoot c ������ ����������� ����� inf
        If PathExists(BackslashAdd2Path(strDestination) & strCatFile) = False Then
            strCatFilePath = CStr(SearchFilesInRoot(strSysDirCatRoot, strCatFileFromInf, True, True))

            If LenB(strCatFilePath) > 0 Then
                CopyFileTo strCatFilePath, BackslashAdd2Path(strDestination) & strCatFile
                DebugMode "***CatalogFile find in: " & strCatFilePath
            End If
        End If

        ' ���� ���� cat ��� �� ���� ����� strSysDirDRVStore
        If PathExists(BackslashAdd2Path(strDestination) & strCatFile) = False Then
            strCatFilePath = CStr(SearchFilesInRoot(strSysDirDRVStore, strCatFile, True, True))

            If LenB(strCatFilePath) > 0 Then
                CopyFileTo strCatFilePath, BackslashAdd2Path(strDestination) & strCatFile
                DebugMode "***CatalogFile find in: " & strCatFilePath
            End If
        End If

        ' ���� ���� cat ��� ��� �� ������, �� ���� ��� �� ���� ����� windows
        If PathExists(BackslashAdd2Path(strDestination) & strCatFile) = False Then
            strCatFilePath = CStr(SearchFilesInRoot(strWinDir, strCatFile, True, True))

            If LenB(strCatFilePath) > 0 Then
                CopyFileTo strCatFilePath, BackslashAdd2Path(strDestination) & strCatFile
                DebugMode "***CatalogFile find in: " & strCatFilePath
            End If
        End If

        ' ���� ���� ������, �� ��� ����� �������� ������� ������� ��� ����������� �������������
        If PathExists(BackslashAdd2Path(strDestination) & strCatFile) Then
            FindCopyCatFile = strCatFile
        Else

            '���� �� ������ ����? �� �������� ����� ��� ��������� �����  strCatFile_ntx86 � strCatFile_ntamd64
            If LenB(strCatFile_ntx86) > 0 Then
                If LenB(strCatFile_ntamd64) > 0 Then
                    If Not mbExitGoto Then
                        mbExitGoto = True
                        strCatFile = strCatFile_ntamd64
                        GoTo CopyCatAgain
                    End If
                End If
            End If
        End If
    End If

    If PathExists(BackslashAdd2Path(strDestination) & strCatFile) = False Then
        DebugMode "***CatalogFile not find: " & strCatFile
    End If
End Function

Private Sub FontCharsetChange()

    ' ���������� �����
    With Me.Font
        .Name = strFontMainForm_Name
        .Size = lngFontMainForm_Size
        .Charset = lngFont_Charset
    End With

    SetBtnFontProperties cmdCheckAll
    SetBtnFontProperties cmdUnCheckAll
    SetBtnFontProperties cmdStartBackUp
    SetBtnFontProperties cmdBreak
    frGroup.Font.Charset = lngFont_Charset
    frBackUp.Font.Charset = lngFont_Charset
    frArchName.Font.Charset = lngFont_Charset
    frPanelLV.Font.Charset = lngFont_Charset
    ctlUcStatusBar1.Font.Charset = lngFont_Charset
End Sub

Private Sub Form_Activate()

    If mbFirstStart Then
        If mbStartMaximazed Then
            Me.WindowState = vbMaximized
        End If

        DoEvents

        ' �������� ���������� ��� ������
        If mbUpdateCheck Then
            ChangeStatusTextAndDebug strMessages(58)
            CheckUpd
            mbFirstStart = False
        Else
            ShowUpdateToolTip
        End If

        ChangeStatusTextAndDebug strMessages(1)
        mbFirstStart = False
    End If

    mbFirstStart = False
End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_KeyDown
'!  ����������  :  KeyCode As Integer, Shift As Integer
'!  ��������    :  ��������� ������� ������ ����������
'! -----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ' ����� �� ��������� �� "Escape"
    If KeyCode = vbKeyEscape Then
        If MsgBox(strMessages(34), vbQuestion + vbYesNo, strProductName) = vbYes Then
            Unload Me
        End If
    End If

    If KeyCode = vbKeyF5 Then
        ' ���� ��������� � �����
        mnuReCollectHWID_Click
    End If

    ' ������ ������ "Ctrl"
    If Shift = 2 Then

        Select Case KeyCode

            Case 65
                ' Ctrl+A (��������� ������������� ������� ��� ���������)
                cmdCheckAll_Click

            Case 90
                ' Ctrl+Z (��������� ������������� ������� ��� ���������)
                cmdUnCheckAll_Click

            Case 19
                ' CTRL+Break (���������� ��������� ���������)
                If cmdBreak.Visible Then
                    cmdBreak_Click
                End If

            Case 49
                ' CTRL+1 (������������ ����� ��������)
                optGrp1.Value = True
                
            Case 50
                ' CTRL+2 (������������ ����� ��������)
                optGrp2.Value = True
                
            Case 51
                ' CTRL+3 (������������ ����� ��������)
                optGrp3.Value = True
                
            Case 52
                ' CTRL+4 (������������ ����� ��������)
                optGrp4.Value = True
                
        End Select
    End If
End Sub

Private Sub Form_Load()

    Dim i  As Long
    Dim ii As Long

    DebugMode "MainForm Show"
    SetupVisualStyles Me

    With Me
        ' �������� ������ ����� � ����������
        ' Icon for Exe-file
        SetIcon .hWnd, "APPICON", True
        SetIcon .hWnd, "FRMMAIN", False
        ' ����� ��������� �����
        strFormName = .Name
        ChangeFrmMainCaption
        ' ������������� ����� �� ���� �����
        .Width = lngMainFormWidth
        .Height = lngMainFormHeight
        ' ���������� ����� �� ������
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    LoadIconImage
    ' ������������� ���� (��� 3D)
    Me.Line (0, 15)-(ScaleWidth, 15), vbWhite
    Me.Line (0, 0)-(ScaleWidth, 0), GetSysColor(COLOR_BTNSHADOW)

    lngBorderWidthY = VPadding(Me)
    lngBorderWidthX = HPadding(Me)

    ' ������� StatusBar
    ctlUcStatusBar1.AddPanel strProductName
    ctlUcStatusBar1.PanelAutoSize(1) = False
    PrintFileInDebugLog strSysIni
    ' �������� ���� ������
    mnuMainLang.Visible = mbMultiLanguage

    ' ����������z ���������� � �������� ���� ������
    If mbMultiLanguage Then
        DebugMode "CreateLangList: " & UBound(arrLanguage)

        For i = UBound(arrLanguage, 2) To 1 Step -1
            CreateMenuLngIndex CStr(arrLanguage(2, i))
        Next
        Localise strPCLangCurrentPath

        For ii = mnuLang.LBound To mnuLang.UBound
            mnuLang(ii).Checked = arrLanguage(1, ii + 1) = strPCLangCurrentPath
        Next
        mnuLangStart.Checked = Not mbAutoLanguage
    Else
        ' ���������� �����
        FontCharsetChange
    End If

    '���������� ������ ������ �������� ��������� �����
    LoadComboList
    ' ��������� ������ ��������� �� ������� - �������� �� ��������� �����
    frmProgress.Show vbModal, Me
    ' ���������� ListView �� ������ ���������� ����
    LoadList_Device False
    'pbProgressBar.Visible = False
    ' ��������� ��������� ��� ������
    chkCheckAll.Value = Abs(mbCheckAllGroup)
    chkHideOther.Value = Abs(mbListOnlyGroup)
    ' ����� ��� ������
    SelectStartMode
    ' ��� ������ ��� ������
    SelectStartArchName
    ' ������� ���-�� ����������
    FindCheckCountList

    '    If lngFrameTime < 0 Then lngFrameTime = 1
    '    If lngFrameCount < 1 Then lngFrameCount = 20
    If Me.WindowState <> vbMinimized Then
        AnimateForm Me, aLoad, eZoomOut, lngFrameTime, lngFrameCount
    End If

    DebugMode "FrmMainLoad-Finish" & vbNewLine & _
              "======================================================================="
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_QueryUnload
'! Description (��������)  :   [���������� �������� �����]
'! Parameters  (����������):   Cancel (Integer)
'                              UnloadMode (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' ��������� ��������� �� �������� ����������, ���� ��� �� ��������� ����� �� ���������, ����� ��������� �������
    If mbCheckUpdNotEnd Then
        Cancel = UnloadMode = vbFormControlMenu Or vbFormCode
        Exit Sub
    End If
    
    ' �������� ��������� ������ ���� ���� � ���� ����� ��������
    If mbDelTmpAfterClose Then
        ChangeStatusTextAndDebug strMessages(81), , , , strMessages(130)

        '������ ���� ������ �� ���������� ���������
        If Not mbRestartProgram Then
            'Me.Hide
            DelTemp
        End If
    End If
    
    ' ���������� ���������� ��� ������
    If mbSaveSizeOnExit Then
        FRMStateSave
    End If

    ' ��������� ���� ��� ������
    If Not mbIsDriveCDRoom Then
        If mnuLangStart.Checked Then
            IniWriteStrPrivate "Main", "StartLanguageID", strPCLangCurrentID, strSysIni
        End If

        IniWriteStrPrivate "Main", "AutoLanguage", CStr(Abs(Not mnuLangStart.Checked)), strSysIni
    End If

    SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP", False

    If mbLoadIniTmpAfterRestart Then
        SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP_PATH", "-"

        If StrComp(FileNameFromPath(strSysIni), "Settings_DBS_TMP.ini", vbTextCompare) = 0 Then
            DeleteFiles strSysIni
        End If
    End If

    If lngFrameTime < 0 Then lngFrameTime = 2
    If lngFrameCount < 1 Then lngFrameCount = 40
    If Me.WindowState <> vbMinimized Then
        AnimateForm Me, aUnload, eZoomOut, lngFrameTime, lngFrameCount
    End If

    ' ��������� �� ������ ����� � ������ ����������
    ' ������ ����������
    lvDevices.Destroy
    Set lvDevices = Nothing
    Set frmMain = Nothing

    ' ��������� �� ������ �����
    UnloadAllForms strFormName
    
    Unload Me
    Set frmMain = Nothing
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub UnloadAllForms
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   FormToIgnore (String = vbNullString)
'!--------------------------------------------------------------------------------
Public Sub UnloadAllForms(Optional FormToIgnore As String = vbNullString)

    Dim F As Form

    For Each F In Forms

        If Not F Is Nothing Then
            If StrComp(F.Name, FormToIgnore, vbTextCompare) <> 0 Then
                Unload F
                Set F = Nothing
            End If
        End If

    Next F

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Resize
'! Description (��������)  :   [��������� �������� ��������� ��� ��������� �������� �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Resize()

    With Me

        If .WindowState <> vbMinimized Then
            If strOsCurrentVersion >= "6.0" Then
                frGroup.Left = 100
                frBackUp.Left = frGroup.Left + frGroup.Width + 120
            Else
                frGroup.Left = 120
                frBackUp.Left = frGroup.Left + frGroup.Width + 220
            End If

            ctlUcStatusBar1.PanelWidth(1) = (Me.Width \ Screen.TwipsPerPixelX)
            ListViewResize
        End If
    End With
End Sub

'Private Sub Form_Terminate()
'
'    If Forms.Count = 0 Then
'        UnloadApp
'    End If
'End Sub

'! -----------------------------------------------------------
'!  �������     :  FRMStateSave
'!  ����������  :
'!  ��������    :  ������ ��������� ���� � ini-���
'! -----------------------------------------------------------
Private Sub FRMStateSave()

    Dim miHeight      As Long
    Dim miWidth       As Long
    Dim miWindowState As Long

    ' ���� ��������� �������, �� ��������� ����������
    miHeight = CLng(Me.Height)
    miWidth = vbNullString & CLng(Me.Width) & vbNullString

    If Me.WindowState = vbMaximized Then
        miWindowState = 1
    Else
        miWindowState = 0
    End If

    IniWriteStrPrivate "MainForm", "Height", CStr(miHeight), strSysIni
    IniWriteStrPrivate "MainForm", "Width", CStr(miWidth), strSysIni
    IniWriteStrPrivate "MainForm", "StartMaximazed", CStr(miWindowState), strSysIni
End Sub

Private Function ListingDirectory(ByVal strPath As String, ByVal mbRecursion As Boolean) As String

    Dim strFileList_x() As String
    Dim strFileList     As String
    Dim strFileListTemp As String
    Dim ii              As Long

    DebugMode "***ListingDirectory-Start: source=" & strPath, 2

    If LenB(strPath) > 0 Then
        strFileList_x = SearchFilesInRoot(strPath, ALL_FILES, mbRecursion, False, False)
        strFileList = vbNullString

        If UBound(strFileList_x, 2) >= 0 Then
            If LenB(strFileList_x(0, 0)) Then

                Dim lngLBound As Long
                Dim lngUBound As Long

                lngLBound = LBound(strFileList_x, 2)
                lngUBound = UBound(strFileList_x, 2)

                For ii = lngLBound To lngUBound
                    strFileListTemp = FileNameFromPath(strFileList_x(0, ii))

                    If LenB(strFileListTemp) Then
                        strFileList = AppendStr(strFileList, strFileListTemp, ";")
                    End If

                Next
            End If
        End If

    Else
        DebugMode "***ListingDirectory-Source Path not defined", 2
    End If

    ListingDirectory = strFileList
    DebugMode "***ListingDirectory-Finish", 2
End Function

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
        frPanel.Height = .Height - ctlUcStatusBar1.Height - lngBorderWidthY

        If strOsCurrentVersion >= "6.0" And .WindowState <> vbMaximized Then
            frPanel.Width = .Width - lngBorderWidthX
        Else
            frPanel.Width = .Width
        End If

        lngLVPanelTop = frGroup.Top + frGroup.Height + 80
        lngLVPanelLeft = frGroup.Left
        lngLVPanelHeight = frPanel.Height - lngLVPanelTop - 120
        lngLVPanelWidhtTemp = frBackUp.Left + frBackUp.Width - frGroup.Left

        If strOsCurrentVersion >= "6.0" And .WindowState <> vbMaximized Then
            lngLVPanelWidht = .Width - lngBorderWidthX - lngLVPanelLeft * 2.9
        ElseIf strOsCurrentVersion >= "6.0" And .WindowState = vbMaximized Then
            lngLVPanelWidht = .Width - lngBorderWidthX - lngLVPanelLeft * 2
        Else
            lngLVPanelWidht = .Width - lngBorderWidthX - lngLVPanelLeft * 2
        End If

        If lngLVPanelWidht < lngLVPanelWidhtTemp Then
            lngLVPanelWidht = lngLVPanelWidhtTemp
        End If

        With frPanelLV
            .Top = lngLVPanelTop
            .Left = lngLVPanelLeft
            .Height = lngLVPanelHeight
            .Width = lngLVPanelWidht
            lngLVTop = .TextBoxHeight + 1
            lngLVHeight = (.Height / Screen.TwipsPerPixelY) - lngLVTop - 2
            lngLVWidht = (.Width / Screen.TwipsPerPixelX) - 4
        End With

        If Not (lvDevices Is Nothing) Then
            lvDevices.Move 2, lngLVTop, lngLVWidht, lngLVHeight
            lvDevices.Refresh
        End If
    End With
End Sub

'���������� ������ ������ �������� ��������� �����
Private Sub LoadComboList()

    Dim strFormName As String

    strFormName = CStr(Me.Name)
    ' ������ ���������
    cmbListTypeBackupElement1 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbListTypeBackupElement1", "����������������� ����� � ����������")
    cmbListTypeBackupElement2 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbListTypeBackupElement2", "7z-����� � ����������")
    cmbListTypeBackupElement3 = LocaliseString(strPCLangCurrentPath, strFormName, "cmbListTypeBackupElement3", "7zSFX-����� c �������������� ����� DPInst")

    With cmbTypeBackUp
        .Clear
        .AddItem cmbListTypeBackupElement1, 0
        .AddItem cmbListTypeBackupElement2, 1
        .AddItem cmbListTypeBackupElement3, 2

        ' ����� ������������� ��� �������
        If miArchMode < 0 Or miArchMode > .ListCount - 1 Then
            .ListIndex = 0
        Else
            .ListIndex = miArchMode
        End If
    End With
End Sub

Private Sub LoadIconImage()

    DebugMode "LoadIconImage-Start"
    '--------------------- ��������� ������
    LoadIconImage2BtnJC cmdStartBackUp, "BTN_STARTBACKUP", strPathImageMainWork
    LoadIconImage2BtnJC cmdBreak, "BTN_BREAK", strPathImageMainWork
    LoadIconImage2BtnJC cmdCheckAll, "BTN_CHECKMARK", strPathImageMainWork
    LoadIconImage2BtnJC cmdUnCheckAll, "BTN_UNCHECKMARK", strPathImageMainWork
    LoadIconImage2FrameJC frBackUp, "FRAME_BACKUP", strPathImageMainWork
    LoadIconImage2FrameJC frGroup, "FRAME_GROUP", strPathImageMainWork
    LoadIconImage2FrameJC frPanelLV, "FRAME_LIS", strPathImageMainWork
    DebugMode "LoadIconImage-End"
End Sub

'! -----------------------------------------------------------
'!  �������     :  LoadList_Device
'!  ����������  :
'!  ��������    :  ���������� ������� ���c�� ���������
'! -----------------------------------------------------------
Private Sub LoadList_Device(Optional ByVal mbViewed As Boolean = True, Optional ByVal lngMode As Long = 0)

    Dim ii        As Integer
    Dim lngNumRow As Long

    DebugMode "LoadList_Device-Start"
    DebugMode "***LoadList_Device: Mode=" & lngMode

    If lvDevices Is Nothing Then
        Set lvDevices = New cListView

        With lvDevices
            .Create frPanelLV.hWnd, LVS_REPORT Or LVS_AUTOARRANGE Or LVS_SHOWSELALWAYS, 0, 120, 500, 300, , WS_EX_STATICEDGE

            If mbViewed Then
                .SetStyleEx LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES
            Else
                .SetStyleEx LVS_EX_FULLROWSELECT Or LVS_EX_GRIDLINES Or LVS_EX_CHECKBOXES Or LVS_EX_TWOCLICKACTIVATE
            End If

            .AddColumn 1, strTableHwidHeader1, 300
            .AddColumn 2, strTableHwidHeader2, 100
            .AddColumn 3, strTableHwidHeader3, 100
            .AddColumn 4, strTableHwidHeader4, 100
            .AddColumn 5, strTableHwidHeader5, 200
            .AddColumn 6, strTableHwidHeader6, 100
            .AddColumn 7, strTableHwidHeader7, 100
            .AddColumn 8, strTableHwidHeader8, 100
            .AddColumn 9, strTableHwidHeader9, 200
            .AddColumn 10, strTableHwidHeader10, 200
        End With
    End If

    For ii = LBound(arrHwidsLocal, 2) To UBound(arrHwidsLocal, 2)

        Select Case lngMode

            Case 0, 3

                With lvDevices
                    .AddItem arrHwidsLocal(0, ii), , ii
                    .ItemText(1, ii) = arrHwidsLocal(1, ii)
                    .ItemText(2, ii) = arrHwidsLocal(2, ii)
                    .ItemText(3, ii) = arrHwidsLocal(3, ii)
                    .ItemText(4, ii) = arrHwidsLocal(4, ii)
                    .ItemText(5, ii) = arrHwidsLocal(5, ii)
                    .ItemText(6, ii) = arrHwidsLocal(6, ii)
                    .ItemText(7, ii) = arrHwidsLocal(7, ii)
                    .ItemText(8, ii) = arrHwidsLocal(8, ii)
                    .ItemText(9, ii) = arrHwidsLocal(9, ii)
                End With

            Case 1

                If InStr(1, arrHwidsLocal(3, ii), "microsoft", vbTextCompare) Then

                    With lvDevices
                        .AddItem arrHwidsLocal(0, ii), lngNumRow
                        .ItemText(1, lngNumRow) = arrHwidsLocal(1, ii)
                        .ItemText(2, lngNumRow) = arrHwidsLocal(2, ii)
                        .ItemText(3, lngNumRow) = arrHwidsLocal(3, ii)
                        .ItemText(4, lngNumRow) = arrHwidsLocal(4, ii)
                        .ItemText(5, lngNumRow) = arrHwidsLocal(5, ii)
                        .ItemText(6, lngNumRow) = arrHwidsLocal(6, ii)
                        .ItemText(7, lngNumRow) = arrHwidsLocal(7, ii)
                        .ItemText(8, lngNumRow) = arrHwidsLocal(8, ii)
                        .ItemText(9, lngNumRow) = arrHwidsLocal(9, ii)
                    End With

                    lngNumRow = lngNumRow + 1
                End If

            Case 2

                If InStr(1, arrHwidsLocal(3, ii), "microsoft", vbTextCompare) = 0 Then

                    With lvDevices
                        .AddItem arrHwidsLocal(0, ii), lngNumRow
                        .ItemText(1, lngNumRow) = arrHwidsLocal(1, ii)
                        .ItemText(2, lngNumRow) = arrHwidsLocal(2, ii)
                        .ItemText(3, lngNumRow) = arrHwidsLocal(3, ii)
                        .ItemText(4, lngNumRow) = arrHwidsLocal(4, ii)
                        .ItemText(5, lngNumRow) = arrHwidsLocal(5, ii)
                        .ItemText(6, lngNumRow) = arrHwidsLocal(6, ii)
                        .ItemText(7, lngNumRow) = arrHwidsLocal(7, ii)
                        .ItemText(8, lngNumRow) = arrHwidsLocal(8, ii)
                        .ItemText(9, lngNumRow) = arrHwidsLocal(9, ii)
                    End With

                    lngNumRow = lngNumRow + 1
                End If
        End Select

    Next
    DebugMode "LoadList_Device-Finish"
    'strTableHwidHeader1 = "*������������ ����������*")
    'strTableHwidHeader2 = "*���� ��������*")
    'strTableHwidHeader3 = "*������ ��������*")
    'strTableHwidHeader4 = "*�������������*")
    'strTableHwidHeader5 = "*����� ��������*")
    'strTableHwidHeader6 = "*��� ������*")
    'strTableHwidHeader7 = "*Inf-����*")
    'strTableHwidHeader8 = "*������ Inf-�����*")
    'strTableHwidHeader9 = "*HWID*")
    'strTableHwidHeader10 ="-ID ������-")
    'strTableHwidHeader11 ="-ID ���������� ����������-")
End Sub

Private Sub Localise(ByVal StrPathFile As String)

    Dim strFormName As String

    strFormName = CStr(Me.Name)
    ' �������� �����
    'Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    ' ���������� �����
    FontCharsetChange
    frGroup.Caption = LocaliseString(StrPathFile, strFormName, "frGroup", frGroup.Caption)
    optGrp1.Caption = LocaliseString(StrPathFile, strFormName, "optGrp1", optGrp1.Caption)
    optGrp2.Caption = LocaliseString(StrPathFile, strFormName, "optGrp2", optGrp2.Caption)
    optGrp3.Caption = LocaliseString(StrPathFile, strFormName, "optGrp3", optGrp3.Caption)
    optGrp4.Caption = LocaliseString(StrPathFile, strFormName, "optGrp4", optGrp4.Caption)
    chkHideOther.Caption = LocaliseString(StrPathFile, strFormName, "chkHideOther", chkHideOther.Caption)
    cmdCheckAll.Caption = LocaliseString(StrPathFile, strFormName, "cmdCheckAll", cmdCheckAll.Caption)
    cmdUnCheckAll.Caption = LocaliseString(StrPathFile, strFormName, "cmdUnCheckAll", cmdUnCheckAll.Caption)
    chkCheckAll.Caption = LocaliseString(StrPathFile, strFormName, "chkCheckAll", chkCheckAll.Caption)
    cmdBreak.Caption = LocaliseString(StrPathFile, strFormName, "cmdBreak", cmdBreak.Caption)
    frBackUp.Caption = LocaliseString(StrPathFile, strFormName, "frBackUp", frBackUp.Caption)
    cmdStartBackUp.Caption = LocaliseString(StrPathFile, strFormName, "cmdStartBackUp", cmdStartBackUp.Caption)
    lblTypeBackUp.Caption = LocaliseString(StrPathFile, strFormName, "lblTypeBackUp", lblTypeBackUp.Caption)
    frPanelLV.Caption = LocaliseString(StrPathFile, strFormName, "frPanelLV", frPanelLV.Caption)
    strTableHwidHeader1 = LocaliseString(StrPathFile, strFormName, "TableHeader1", "*������������ ����������*")
    strTableHwidHeader2 = LocaliseString(StrPathFile, strFormName, "TableHeader2", "*���� ��������*")
    strTableHwidHeader3 = LocaliseString(StrPathFile, strFormName, "TableHeader3", "*������ ��������*")
    strTableHwidHeader4 = LocaliseString(StrPathFile, strFormName, "TableHeader4", "*�������������*")
    strTableHwidHeader5 = LocaliseString(StrPathFile, strFormName, "TableHeader5", "*����� ��������*")
    strTableHwidHeader6 = LocaliseString(StrPathFile, strFormName, "TableHeader6", "*��� ������*")
    strTableHwidHeader7 = LocaliseString(StrPathFile, strFormName, "TableHeader7", "*Inf-����*")
    strTableHwidHeader8 = LocaliseString(StrPathFile, strFormName, "TableHeader8", "*������ Inf-�����*")
    strTableHwidHeader9 = LocaliseString(StrPathFile, strFormName, "TableHeader9", "*HWID*")
    strTableHwidHeader10 = LocaliseString(StrPathFile, strFormName, "TableHeader10", "-ID ������-")
    strTableHwidHeader11 = LocaliseString(StrPathFile, strFormName, "TableHeader11", "-ID ���������� ����������-")
    ' ����
    mnuReCollectHWID.Caption = LocaliseString(StrPathFile, strFormName, "mnuReCollectHWID", mnuReCollectHWID.Caption)
    mnuOptions.Caption = LocaliseString(StrPathFile, strFormName, "mnuOptions", mnuOptions.Caption)
    mnuMainAbout.Caption = LocaliseString(StrPathFile, strFormName, "mnuMainAbout", mnuMainAbout.Caption)
    mnuLinks.Caption = LocaliseString(StrPathFile, strFormName, "mnuLinks", mnuLinks.Caption)
    mnuHistory.Caption = LocaliseString(StrPathFile, strFormName, "mnuHistory", mnuHistory.Caption)
    mnuHelp.Caption = LocaliseString(StrPathFile, strFormName, "mnuHelp", mnuHelp.Caption)
    mnuHomePage.Caption = LocaliseString(StrPathFile, strFormName, "mnuHomePage", mnuHomePage.Caption)
    mnuHomePageForum.Caption = LocaliseString(StrPathFile, strFormName, "mnuHomePageForum", mnuHomePageForum.Caption)
    mnuOsZoneNet.Caption = LocaliseString(StrPathFile, strFormName, "mnuOsZoneNet", mnuOsZoneNet.Caption)
    mnuCheckUpd.Caption = LocaliseString(StrPathFile, strFormName, "mnuCheckUpd", mnuCheckUpd.Caption)
    mnuDonate.Caption = LocaliseString(StrPathFile, strFormName, "mnuDonate", mnuDonate.Caption)
    'mnuLicence.Caption = LocaliseString(StrPathFile, strFormName, "mnuLicence", mnuLicence.Caption)
    mnuAbout.Caption = LocaliseString(StrPathFile, strFormName, "mnuAbout", mnuAbout.Caption)
    mnuModulesVersion.Caption = LocaliseString(StrPathFile, strFormName, "mnuModulesVersion", mnuModulesVersion.Caption)
    mnuMainLang.Caption = LocaliseString(StrPathFile, strFormName, "mnuMainLang", mnuMainLang.Caption)
    mnuLangStart.Caption = LocaliseString(StrPathFile, strFormName, "mnuLangStart", mnuLangStart.Caption)
    LoadComboList
    ChangeFrmMainCaption
    frArchName.Caption = LocaliseString(StrPathFile, strFormName, "frArchName", frArchName.Caption)
    optArchNamePC.Caption = LocaliseString(StrPathFile, strFormName, "optArchNamePC", optArchNamePC.Caption)
    optArchModelPC.Caption = LocaliseString(StrPathFile, strFormName, "optArchModelPC", optArchModelPC.Caption)
    optArchCustom.Caption = LocaliseString(StrPathFile, strFormName, "optArchCustom", optArchCustom.Caption)
    '��������� ����������� ���������
    LocaliseMessage strPCLangCurrentPath
End Sub

'
Private Sub lvDevices_ColumnClick(ByVal iColumn As Long)

    'toggle the sort order for use in the CompareXX routines
    sOrder = Not sOrder

    Select Case iColumn

        Case 0, 1, 3, 4, 5, 6, 8

            'Use sort routine to sort by text
            If sOrder Then
                lvDevices.Sort iColumn, stText, soAscending
            Else
                lvDevices.Sort iColumn, stText, soDescending
            End If

            '
        Case 2, 7

            'Use sort routine to sort by number
            If sOrder Then
                lvDevices.Sort iColumn, stNumber, soAscending
            Else
                lvDevices.Sort iColumn, stNumber, soDescending
            End If

            'Case 3:
            'Use sort routine to sort by number
            'If sOrder Then
            'Call lvDevices.Sort(iColumn, stNumber, soAscending)
            'Else
            'Call lvDevices.Sort(iColumn, stNumber, soDescending)
            'End If
    End Select

    '���������� - ����� ���������� ����������� ListView �� �������������
End Sub

Private Sub lvDevices_DblClick(ByVal iItem As Long, ByVal Button As MouseButtonConstants)

    Dim strOrigHwid As String

    If Button = vbLeftButton Then
        strOrigHwid = lvDevices.ItemText(6, iItem)
        OpenDeviceProp strOrigHwid
    End If
End Sub

Private Sub lvDevices_KeyUp(ByVal KeyCode As Long, ByVal Shift As Integer)

    If KeyCode = 32 Then
        FindCheckCountList
    End If
End Sub

Private Sub lvDevices_MouseUp(ByVal Button As MouseButtonConstants, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal Shift As Integer)

    FindCheckCountList
End Sub

'! -----------------------------------------------------------
'!  �������     :  mnuAbout_Click
'!  ����������  :
'!  ��������    :  ���� - � ���������
'! -----------------------------------------------------------
Private Sub mnuAbout_Click()

    frmAbout.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  �������     :  mnuCheckUpd_Click
'!  ����������  :
'!  ��������    :  ���� - ��������� ����������
'! -----------------------------------------------------------
Private Sub mnuCheckUpd_Click()

    CheckUpd False
End Sub

Private Sub mnuDonate_Click()

    frmDonate.Show vbModal, Me
End Sub

'! -----------------------------------------------------------
'!  �������     :  mnuHistory_Click
'!  ����������  :
'!  ��������    :  ���� - ������� ���������
'! -----------------------------------------------------------
Private Sub mnuHistory_Click()

    Dim cmdString       As String
    Dim strFilePathTemp As String

    strFilePathTemp = strAppPath & "\Tools\DocsDBS\" & strPCLangCurrentID & "\history.txt"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPath & "\Tools\DocsDBS\0409\history.txt"
    End If

    cmdString = Kavichki & strFilePathTemp & Kavichki
    RunUtilsShell cmdString, False
End Sub

Private Sub mnuHomePage_Click()

    RunUtilsShell Kavichki & strUrl_MainWWWSite & Kavichki, False
End Sub

Private Sub mnuHomePageForum_Click()

    RunUtilsShell Kavichki & strUrl_MainWWWForum & Kavichki, False
End Sub

Private Sub mnuLang_Click(Index As Integer)

    Dim i                      As Long
    Dim ii                     As Long
    Dim strPathLng             As String
    Dim strPCLangCurrentIDTemp As String
    Dim strPCLangCurrentID_x() As String

    i = Index + 1

    For ii = mnuLang.LBound To mnuLang.UBound
        mnuLang(ii).Checked = ii = Index
    Next
    strPathLng = arrLanguage(1, i)
    ChangeStatusTextAndDebug "Select language: " & arrLanguage(2, i)
    strPCLangCurrentPath = strPathLng
    strPCLangCurrentIDTemp = arrLanguage(3, i)
    lngFont_Charset = GetCharsetFromLng(CLng(arrLanguage(6, i)))

    If InStr(1, strPCLangCurrentIDTemp, ";") Then
        strPCLangCurrentID_x = Split(strPCLangCurrentIDTemp, ";")
        strPCLangCurrentID = strPCLangCurrentID_x(0)
    Else
        strPCLangCurrentID = strPCLangCurrentIDTemp
    End If

    ' ���������� �����������
    Localise strPCLangCurrentPath
    ' ����������� �������
    mnuReCollectHWID_Click
End Sub

Private Sub mnuLangStart_Click()

    mnuLangStart.Checked = Not mnuLangStart.Checked
End Sub

'! -----------------------------------------------------------
'!  �������     :  mnuLinks_Click
'!  ����������  :
'!  ��������    :  ���� - ������
'! -----------------------------------------------------------
Private Sub mnuLinks_Click()

    Dim cmdString       As String
    Dim strFilePathTemp As String

    strFilePathTemp = strAppPath & "\Tools\DocsDBS\" & strPCLangCurrentID & "\Links.html"

    If PathExists(strFilePathTemp) = False Then
        strFilePathTemp = strAppPath & "\Tools\DocsDBS\0409\Links.html"
    End If

    cmdString = Kavichki & strFilePathTemp & Kavichki
    RunUtilsShell cmdString, False
End Sub

'! -----------------------------------------------------------
'!  �������     :  mnuModulesVersion_Click
'!  ����������  :
'!  ��������    :  ���� - ������ �������
'! -----------------------------------------------------------
Private Sub mnuModulesVersion_Click()

    VerModules
End Sub

'! -----------------------------------------------------------
'!  �������     :  mnuOptions_Click
'!  ����������  :
'!  ��������    :  ���� - ���������
'! -----------------------------------------------------------
Private Sub mnuOptions_Click()

    Dim i As Long

    frmOptions.Show vbModal, Me

    If mbRestartProgram Then
        ' ��������� �� ������ ����� � ������ ����������
        ' ������ ����������
        'Set CFm_sbStatusBar = Nothing
        lvDevices.Destroy
        Set lvDevices = Nothing
        Set frmMain = Nothing

'        If Not mbIsDesignMode Then
'            Unhook
'        End If

        For i = Forms.Count - 1 To 1 Step -1

            If Forms(i).Name <> "frmMain" Then
                Unload Forms(i)
            End If

        Next
        Set frmMain = Nothing
        'FreeLibrary m_hMod
        ' �������������� �����
        ShellExecute Me.hWnd, "open", App.EXEName, vbNullString, strAppPath, SW_SHOWNORMAL
        End
    End If
End Sub

Private Sub mnuOsZoneNet_Click()

    RunUtilsShell Kavichki & "http://forum.oszone.net/thread-190814.html" & Kavichki, False
End Sub

Private Sub mnuReCollectHWID_Click()

    ReCollectHWID
    ' ����� ��� ������
    SelectStartMode
    FindCheckCountList
End Sub

Private Sub OpenDeviceProp(ByVal strHwid As String)

    Dim cmdString       As String
    Dim cmdStringParams As String
    Dim nRetShellEx     As Boolean

    cmdString = "rundll32.exe"
    cmdStringParams = "devmgr.dll,DeviceProperties_RunDLL /DeviceID " & strHwid
    DebugMode "cmdString: " & cmdString
    DebugMode "cmdStringParams: " & cmdStringParams
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL, cmdStringParams)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Sub optArchCustom_Click()

    Dim strTempString As String

    With txtArchName
        .Locked = False
        .Enabled = True
        strTempString = SafeDir(ExpandArchNamebyEnvironment(strArchNameCustom))

        If LenB(SafeDir(strTempString)) > 0 Then
            .Text = strTempString
        Else
            .Text = CollectDpName(strCompName)
        End If
    End With
End Sub

Private Sub optArchModelPC_Click()

    With txtArchName
        .Text = CollectDpName(strCompModel)
        .Locked = True
        .Enabled = False
    End With
End Sub

Private Sub optArchNamePC_Click()

    With txtArchName
        .Text = CollectDpName(strCompName)
        .Locked = True
        .Enabled = False
    End With
End Sub

Private Sub optGrp1_Click()

    Dim i As Integer

    If Not CBool(chkHideOther.Value) Then
        lvDevices.Clear
        lvDevices.Destroy
        Set lvDevices = Nothing
        
        LoadList_Device False
        cmdUnCheckAll_Click

        With lvDevices

            For i = 0 To .Count

                If InStr(1, .ItemText(3, i), "microsoft", vbTextCompare) Then
                    If Not .ItemChecked(i) Then
                        If chkCheckAll.Value Then
                            .ItemChecked(i) = True
                        End If
                    End If

                Else

                    If .ItemChecked(i) Then
                        .ItemChecked(i) = False
                    End If
                End If

            Next
        End With

    Else
        lvDevices.Clear
        lvDevices.Destroy
        Set lvDevices = Nothing
        LoadList_Device False, 1

        If chkCheckAll.Value And chkCheckAll.Enabled Then
            cmdCheckAll_Click
        End If
    End If

    FindCheckCountList
    ListViewResize
End Sub

Private Sub optGrp2_Click()

    Dim i As Integer

    If Not CBool(chkHideOther.Value) Then
        lvDevices.Clear
        lvDevices.Destroy
        Set lvDevices = Nothing
        LoadList_Device False
        cmdUnCheckAll_Click

        With lvDevices

            For i = 0 To .Count

                If InStr(1, .ItemText(3, i), "microsoft", vbTextCompare) = 0 Then
                    If Not .ItemChecked(i) Then
                        If chkCheckAll.Value Then
                            .ItemChecked(i) = True
                        End If
                    End If

                Else

                    If .ItemChecked(i) Then
                        .ItemChecked(i) = False
                        .ItemCut(i) = True
                    End If
                End If

            Next
        End With

    Else
        lvDevices.Clear
        lvDevices.Destroy
        Set lvDevices = Nothing
        LoadList_Device False, 2

        If chkCheckAll.Value And chkCheckAll.Enabled Then
            cmdCheckAll_Click
        End If
    End If

    FindCheckCountList
    ListViewResize
End Sub

Private Sub optGrp3_Click()

    If Not CBool(chkHideOther.Value) Then
        If chkCheckAll.Value Then
            cmdCheckAll_Click
        End If
    Else
        lvDevices.Clear
        lvDevices.Destroy
        Set lvDevices = Nothing
        LoadList_Device False, 3

        If chkCheckAll.Value Then
            If chkCheckAll.Enabled Then
                If chkCheckAll.Value Then
                    cmdCheckAll_Click
                End If
            End If
        End If
    End If

    FindCheckCountList
    ListViewResize
End Sub

Private Sub optGrp4_Click()

    Dim i As Integer

    If Not CBool(chkHideOther.Value) Then
        lvDevices.Clear
        lvDevices.Destroy
        Set lvDevices = Nothing
        LoadList_Device False
        cmdUnCheckAll_Click

        With lvDevices

            For i = 0 To .Count

                If LenB(.ItemText(7, i)) = 0 Then
                    If Not .ItemChecked(i) Then
                        If chkCheckAll.Value Then
                            .ItemChecked(i) = True
                        End If
                    End If

                Else

                    If .ItemChecked(i) Then
                        .ItemChecked(i) = False
                    End If
                End If

            Next
        End With

    Else
        lvDevices.Clear
        lvDevices.Destroy
        Set lvDevices = Nothing
        LoadList_Device False, 4
    End If

    FindCheckCountList
    ListViewResize
End Sub

Private Sub ReCollectHWID()

    ChangeStatusTextAndDebug strMessages(3)
    ' ������� ����view
    lvDevices.Clear
    lvDevices.Refresh
    DoEvents
    lvDevices.Destroy
    Set lvDevices = Nothing
    DoEvents
    
    ' �������� �������� ������
    ProgressBar1.SetTaskBarProgressState PrbTaskBarStateInProgress
    frmProgress.Show vbModal, Me
    
    ' � ������ ������������� ������ ���������
    LoadList_Device False
    ListViewResize
    
    ProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
    
    ChangeStatusTextAndDebug strMessages(5)
End Sub

Private Sub RunUtilsShell(ByVal strPathUtils As String, Optional ByVal mbCollectPath As Boolean = True)

    Dim nRetShellEx As Boolean
    Dim cmdString   As String

    If mbCollectPath Then
        cmdString = PathCollect(strPathUtils)
    Else
        cmdString = strPathUtils
    End If

    DebugMode "cmdString: " & cmdString
    nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)
    DebugMode "cmdString: " & nRetShellEx
End Sub

Private Function SearchSectInSect(ByRef arrZ() As String) As String()

    Dim strFileName      As String
    Dim D                As Long
    Dim strFileNameSect  As String
    Dim strFileName_x()  As String
    Dim strSectionList() As String
    Dim n                As Long
    Dim i                As Long
    Dim miMaxCountArr    As Long

    miMaxCountArr = 100
    ' ������������ ���-�� ��������� � �������
    ReDim strSectionList(miMaxCountArr) As String

    For D = 1 To UBound(arrZ, 1)
        strFileName = TrimNull(arrZ(D, 2))
        strFileNameSect = arrZ(D, 1)

        ' ����������� ��� ����� ";"
        If InStr(1, strFileName, ";", vbTextCompare) Then
            strFileName = Trim$(Left$(strFileName, InStrRev(strFileName, ";") - 1))
        End If

        If StrComp(strFileNameSect, "CopyFiles", vbTextCompare) = 0 Then
            strFileName_x = Split(strFileName, ",")

            For i = 0 To UBound(strFileName_x)

                ' ���� ������� � ������� ���������� ������ ��� ���������, �� ����������� ����������� �������
                If n = miMaxCountArr Then
                    miMaxCountArr = miMaxCountArr + miMaxCountArr
                    ReDim Preserve strSectionList(1, miMaxCountArr)
                End If

                strSectionList(n) = strFileName_x(i)
                n = n + 1
            Next
        End If

    Next

    ' ������������� ������ �� �������� ���-�� �������
    If n > 0 Then
        ReDim Preserve strSectionList(n - 1)
    Else
        ReDim Preserve strSectionList(0)
    End If

    SearchSectInSect = strSectionList
End Function

' ����� ��� ������
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

' ����� ��� ������
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

Private Sub StartBackUp()

    Dim destDir               As String
    Dim destDirDialog         As String
    Dim n                     As Long
    Dim i                     As Long
    Dim D                     As Long
    Dim dest                  As String
    Dim Z()                   As String
    Dim Z2()                  As String
    Dim Z3()                  As String
    Dim Z4()                  As String
    Dim Z5()                  As String
    Dim ZF1()                 As String
    Dim inf()                 As String
    Dim strInfFileName        As String
    Dim strInfFile2Path       As String
    Dim lngArrCount           As Long
    Dim lvCount               As Long
    Dim lvCountCheck          As Long
    Dim TimeScriptRun         As Long
    Dim TimeScriptFinish      As Long
    Dim AllTimeScriptRun      As String
    Dim miPbInterval          As Long
    Dim miPbNext              As Long
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

    DebugMode "cmdStartBackUp_Click-Start"
    TimeScriptRun = 0
    AllTimeScriptRun = vbNullString
    TimeScriptRun = GetTickCount

    '# ���� ���� ���������� ������
    If lvDevices.CheckedCount = 0 Then
        MsgBox strMessages(6), vbInformation + vbOKOnly, strProductName
    Else

        '# ������ �������� �����
        With New CommonDialog
            If mbIsDriveCDRoom Then
                .InitDir = strAppPathBackSL & "drivers\"
            Else
                .InitDir = DefineFolderBackUp
            End If
    
            .DialogTitle = strMessages(2)
            
            If .ShowFolder = True Then
                destDirDialog = .FileName
            End If
    
        End With
    
        If LenB(destDirDialog) = 0 Then
            '# if user cancel #
            Exit Sub
        End If

        DebugMode "StartBackUp: Destination=" & destDirDialog

        '��������� ���� ����� �������
        If mbBlockListOnBackup Then
            DebugMode "BlockListOnBackup: TRUE"
            EnableWindow lvDevices.hWnd, 0
        End If

        ' ��������� �������� �� ����� ��������
        DebugMode "BlockControl: TRUE"
        BlockControl True
        MousePointer = 11
        '# display hourglass cursor while read #
        DoEvents

        '��������� ���� �������� ���������� ������
        If LenB(Trim$(txtArchName)) > 0 Then
            destDir = BackslashAdd2Path(destDirDialog) & txtArchName
            'CollectDpName
        Else
            destDir = BackslashAdd2Path(destDirDialog) & CollectDpName(strCompName)
        End If

        DebugMode "***StartBackUp: Destination directory: " & destDir

        If PathExists(destDir) Then
            DebugMode "***StartBackUp: Clean destination directory: " & destDir
            ChangeStatusTextAndDebug strMessages(82)
            DelRecursiveFolder destDir
        End If

        lvCountCheck = lvDevices.CheckedCount
        ' ���������� ProgressBar
        With ProgressBar1
            .Value = 0
            .Visible = True
            .SetTaskBarProgressState PrbTaskBarStateInProgress
            .SetTaskBarProgressValue .Value, .Max
        End With
        
        miPbInterval = Round(10000 / lvCountCheck)
        miPbNext = 0
        '# loop al drivers in grid #
        n = -1
        numCat = 1
        lvCount = lvDevices.Count
        DebugMode "***StartBackUp: Count of drivers: " & lvCount
        DebugMode "***StartBackUp: Count of checked drivers: " & lvCountCheck

        For i = 0 To lvCount - 1
            mbBackUPedFiles = False

            '# ���� � ����� ���������� ������
            If lvDevices.ItemChecked(i) Then
                DebugMode "____________________________________________________________________"
                DebugMode "***StartBackUp: DRIVER in List �" & (i + 1)
                '��������� ������ ������
                strDriverDesc = SafeDir(lvDevices.ItemText(0, i))
                strClass = SafeDir(lvDevices.ItemText(5, i))
                strInfFileName = lvDevices.ItemText(6, i)
                'strInfFileName = "oem6.inf"
                strInfSection = lvDevices.ItemText(7, i)
                DebugMode "***StartBackUp: DRIVER=" & strDriverDesc & " Inf=" & strInfFileName

                ' ��������� �������
                If mbBreakUpdateDBAll Then
                    MsgBox strMessages(27) & vbNewLine & strDriverDesc, vbInformation, strProductName
                    DebugMode "***StartBackUp: BREAK by USER"
                    Exit For
                End If

                n = n + 1
                strStatusMsgTemp = strMessages(9) & " (" & n + 1 & " " & strMessages(108) & " " & lvCountCheck & "): " & strDriverDesc & ": "
                ChangeStatusTextAndDebug strStatusMsgTemp
                ReDim Preserve inf(n)
                '# ������� ���������� ��������
                dest = BackslashAdd2Path(destDir) & strClass & vbBackslash & strDriverDesc
                strInfFile2Path = BackslashAdd2Path(dest) & strInfFileName
                DebugMode "***StartBackUp: DestForDriver=" & dest

                ' ���� �������� inf-���� ����������, �� ����������, ���� ��� ����������
                If PathExists(strInfDir & strInfFileName) Then

                    ' ���� �������� ���, �� �������
                    If PathExists(dest) = False Then
                        CreateNewDirectory dest
                        numCat = 1
                    Else

                        ' � ���� ����, �� ������ �� ��� ������������ ����� �������, ������ ��� �����
                        If PathExists(strInfFile2Path) = False Then
                            dest = dest & "_" & numCat
                            CreateNewDirectory dest
                            numCat = numCat + 1
                        End If
                    End If

                    strInfFile2Path = BackslashAdd2Path(dest) & strInfFileName
                    '# �������� ���-���� � ������� ����������
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Copy Inf-File"
                    DebugMode strStatusMsgTemp & "Analizing '[SourceDisksFiles]'"
                    CopyFileTo strInfDir & strInfFileName, strInfFile2Path
                    'CopyFileTo "c:\oem6.inf", strInfFile2Path
                    DoEvents
                    '# �������� cat-���� � ������� ����������
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Search CatalogFile"
                    DebugMode strStatusMsgTemp & "Search CatalogFile"
                    strCatFileName4Inf = FindCopyCatFile(strInfFile2Path, dest)

                    ' ���� ���������� cat-����, �� ��������������� inf-���� � ��� cat-�����
                    If LenB(strCatFileName4Inf) > 0 Then
                        strInfFile2Path4Cat = PathNameFromPath(strInfFile2Path) & vbBackslash & FileName_woExt(strCatFileName4Inf) & ".inf"

                        If MoveFileTo(strInfFile2Path, strInfFile2Path4Cat) Then
                            strInfFile2Path = strInfFile2Path4Cat
                        End If
                    End If

                    DoEvents
                    ' ������������� ���� � �������� ��� ����� �� �������� c:\WINDOWS\system32\DRVSTORE\
                    DebugMode "***" & strStatusMsgTemp & "Analizing DRVSTORE"
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Analizing DriverStore folder"

                    If strOsCurrentVersion < "6.0" Then
                        If LenB(strCatFileName4Inf) > 0 And strOsCurrentVersion >= "5.1" Then
                            If PathExists(BackslashAdd2Path(dest) & strCatFileName4Inf) Then
                                ' ��������� ������ �� Hash
                                strDataSHA1 = CalcHashFile(BackslashAdd2Path(dest) & strCatFileName4Inf, CAPICOM_HASH_ALGORITHM_SHA1)

                                ZF1 = SearchFoldersInRoot(strSysDirDRVStore, "*" & "_" & UCase$(strDataSHA1) & "*")

                                Dim lngUBoundZF1 As Long

                                lngUBoundZF1 = UBound(ZF1, 2)

                                For D = 0 To lngUBoundZF1
                                    strFolderPath = ZF1(0, D)
                                    strFileNameInf = ZF1(1, D)

                                    If LenB(strFolderPath) > 0 Then
                                        If LenB(strFileNameInf) > 0 Then
                                            strFileNameInf = BackslashAdd2Path(strFolderPath) & strFileNameInf & ".inf"

                                            If PathExists(strFileNameInf) Then

                                                '��������� ������ �� Hash SHA1-�����
                                                mbCompare = CompareFilesByHashCAPICOM(strFileNameInf, strInfFile2Path)

                                                If mbCompare Then
                                                    ' ������� ���������� inf, ����� �� ���� ������
                                                    DeleteFiles strInfFile2Path
                                                    strInfFile2Path = strFileNameInf
                                                    ' �������� ���������� ������
                                                    DebugMode "******CopyFiles from DrvStore: " & strFolderPath
                                                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Copying files from DriverStore folder"
                                                    lngNumFilesFromFolder = rgbCopyFiles(strFolderPath, dest, ALL_FILES)
                                                    DebugMode "******CopyFiles - count files: " & lngNumFilesFromFolder
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

                        If LenB(strFileNameInf) > 0 Then
                            If PathExists(strFileNameInf) Then

                                '��������� ������ �� Hash SHA1-�����
                                mbCompare = CompareFilesByHashCAPICOM(strFileNameInf, strInfFile2Path)

                                If mbCompare Then
                                    ' ��������� ���� �������� � ����������
                                    strFolderPath = PathNameFromPath(strFileNameInf)
                                    ' ������� ���������� inf, ����� �� ���� ������
                                    DeleteFiles strInfFile2Path
                                    strInfFile2Path = strFileNameInf
                                    ' �������� ���������� DrvStore � ������� ����������
                                    DebugMode "******CopyFiles from DrvStore: " & strFolderPath
                                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Copying files from DriverStore folder"
                                    lngNumFilesFromFolder = rgbCopyFiles(strFolderPath, dest, ALL_FILES)
                                    DebugMode "******CopyFiles - count files: " & lngNumFilesFromFolder
                                End If
                            End If
                        End If
                    End If

                    ' ����������� ������ sourcediskfiles sourcedisknames  � ������ ������ ���� ������ � ����� ���� �� ���� ����������
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Collecting path of files information"
                    CollectDestPathFiles strInfFile2Path
                    '#  ������ INF - ��� SourceDisksFiles �� ������ ����� DefaultDestDir
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Analyzing '[SourceDisksFiles]'"
                    DebugMode "***" & strStatusMsgTemp & "Analizing '[SourceDisksFiles]'"
                    Z = LoadIniSectionKeys("SourceDisksFiles", strInfFile2Path)
                    CopyFile2Dest Z, dest, "DefaultDestDir", strInfFile2Path
                    DoEvents
                    '#  ������ INF - �� �������������� ������ DefaultDestDir
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Analyzing '[DestinationDirs]'"
                    DebugMode "***" & strStatusMsgTemp & "Analizing '[DestinationDirs]'"
                    Z2 = LoadIniSectionKeys("DestinationDirs", strInfFile2Path)

                    Dim lngUBoundZ2 As Long

                    lngUBoundZ2 = UBound(Z2)

                    For lngArrCount = 0 To lngUBoundZ2
                        strSectionName = Z2(lngArrCount)

                        If LenB(strSectionName) > 0 Then
                            If StrComp(strSectionName, "DefaultDestDir", vbTextCompare) <> 0 Then
                                Z = LoadIniSectionKeys(strSectionName, strInfFile2Path)
                                DebugMode "***" & strStatusMsgTemp & "Analizing section: " & strSectionName, 2
                                CopyFile2Dest Z, dest, strSectionName, strInfFile2Path, True
                            End If
                        End If

                    Next
                    DoEvents
                    ' �������������� ������ ������ �� �������� CopyFiles
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Analyzing CopyFiles '" & strInfSection & "'"
                    DebugMode "***" & strStatusMsgTemp & "Analizing section by CopyFiles: " & strInfSection
                    Z4 = GetSectionMass(strInfSection, strInfFile2Path, False)
                    Z5 = SearchSectInSect(Z4)

                    Dim lngUBoundZ5 As Long

                    lngUBoundZ5 = UBound(Z5)

                    For lngArrCount = 0 To lngUBoundZ5
                        strSectionName = Z5(lngArrCount)

                        If LenB(strSectionName) > 0 Then
                            DebugMode "***" & strStatusMsgTemp & "Analizing section: " & strSectionName, 2
                            Z = LoadIniSectionKeys(strSectionName, strInfFile2Path)
                            CopyFile2Dest Z, dest, "DefaultDestDir", strInfFile2Path, True
                        End If

                    Next
                    DoEvents
                    ' �������������� ������ ������ �� �������� CopyFiles ������ strInfSection.CoInstallers
                    Erase Z4
                    Erase Z5
                    ChangeStatusTextAndDebug strStatusMsgTemp & vbNewLine & "Analyzing CopyFiles '" & strInfSection & ".CoInstallers'"
                    DebugMode "***" & strStatusMsgTemp & "Analizing section CoInstallers: " & strInfSection & ".CoInstallers"
                    Z4 = GetSectionMass(strInfSection & ".CoInstallers", strInfFile2Path, False)
                    Z5 = SearchSectInSect(Z4)
                    lngUBoundZ5 = UBound(Z5)

                    For lngArrCount = 0 To lngUBoundZ5
                        strSectionName = Z5(lngArrCount)

                        If LenB(strSectionName) > 0 Then
                            DebugMode "***" & strStatusMsgTemp & "Analizing section: " & strSectionName, 2
                            Z = LoadIniSectionKeys(strSectionName, strInfFile2Path)
                            CopyFile2Dest Z, dest, "DefaultDestDir", strInfFile2Path, True
                        End If

                    Next
                    DoEvents
                    ' ���� ����� � ������ ������ ��������� �����
                    Z3 = LoadIniSectionKeys(strInfSection, strInfFile2Path, False)
                    CopyFile2Dest Z3, dest, "DefaultDestDir", strInfFile2Path
                Else
                    DebugMode "StartBackUp: Inf-File NotExist=" & strInfDir & strInfFileName
                End If

                '# show progress #
                miPbNext = miPbNext + miPbInterval

                If miPbNext > 10000 Then
                    miPbNext = 10000
                End If

                With ProgressBar1
                    .Value = miPbNext
                    .SetTaskBarProgressValue .Value, .Max
                End With
                
                mbBackUPedFiles = True
            End If

            ' ���� ���-�� ���� ����������, �� ������� � ���, ���� �������� �������
            If mbBackUPedFiles And mbDebugEnable Then
                DoEvents
                strFileList = ListingDirectory(dest, True)
                DebugMode "***Content directory after backup: " & strFileList
            End If

            ' ������ �������
            Erase Z
            Erase Z2
            Erase Z3
            Erase Z4
            Erase Z5
            Erase ZF1
        Next
        DebugMode "***BackUp all Checked drivers finished."
        DoEvents
        TimeScriptFinish = GetTickCount
        AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish)

        ' ���� ������� �������
        If mbBreakUpdateDBAll Then
            mbBreakUpdateDBAll = False
            ChangeStatusTextAndDebug strMessages(66) & " " & AllTimeScriptRun, , True
        Else

            '# type of backup #
            Select Case cmbTypeBackUp.ListIndex

                    '# create ZIP #
                Case 1
                    With ProgressBar1
                        .Value = 9000
                        .SetTaskBarProgressValue .Value, .Max
                    End With

                    ChangeStatusTextAndDebug "Zipping driver files..."
                    str7zFileArchivePath = BackslashAdd2Path(destDirDialog) & txtArchName & ".7z"
                    DebugMode "StartBackUp: Zip to File=" & str7zFileArchivePath
                    mbDoZip = DoZip(destDir, str7zFileArchivePath)
                    DoEvents

                    If mbDoZip Then
                        '# delete temp folder #
                        ChangeStatusTextAndDebug "Delete temporary files...Please wait"
                        DelFolderBackUp destDir
                    End If

                    MousePointer = 0
                    TimeScriptFinish = GetTickCount
                    AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish)
                    With ProgressBar1
                        .Value = 10000
                        .SetTaskBarProgressValue .Value, .Max
                    End With

                    If mbDoZip Then
                        ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun, , True
                        MsgBox strMessages(10) & vbNewLine & str7zFileArchivePath, vbInformation + vbOKOnly, strProductName
                    Else
                        ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun, , True
                        MsgBox strMessages(12), vbInformation + vbOKOnly, strProductName
                    End If

                    '# create ZIP-SFX with DPInst #
                Case 2
                    With ProgressBar1
                        .Value = 9000
                        .SetTaskBarProgressValue .Value, .Max
                    End With
                    ChangeStatusTextAndDebug "Zipping driver files..."
                    str7zFileArchivePath = BackslashAdd2Path(destDirDialog) & txtArchName & ".exe"
                    DebugMode "StartBackUp: Zip to File=" & str7zFileArchivePath
                    mbDoZip = DoZip(destDir, str7zFileArchivePath)
                    DoEvents

                    If mbDoZip Then
                        '# delete temp folder #
                        ChangeStatusTextAndDebug "Delete temporary files...Please wait"
                        DelFolderBackUp destDir
                    End If

                    '# display default cursor #
                    MousePointer = 0
                    TimeScriptFinish = GetTickCount
                    AllTimeScriptRun = CalculateTime(TimeScriptRun, TimeScriptFinish)

                    If mbDoZip Then
                        ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun, , True
                        MsgBox strMessages(10) & vbNewLine & str7zFileArchivePath, vbInformation + vbOKOnly, strProductName
                    Else
                        ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun, , True
                        MsgBox strMessages(12), vbInformation + vbOKOnly, strProductName
                    End If

                Case Else
                    With ProgressBar1
                        .Value = 10000
                        .SetTaskBarProgressValue .Value, .Max
                    End With
                    ChangeStatusTextAndDebug strMessages(67) & " " & AllTimeScriptRun, , True
                    MsgBox strMessages(10), vbInformation + vbOKOnly, strProductName
            End Select
        End If

        '# show info of end process #
        ProgressBar1.Visible = False
    End If

    MousePointer = 0
    ' ������������ �������� �� ����� ��������
    BlockControl False
    DebugMode "BlockControl: TRUE"

    '������������ ���� ����� ������
    If mbBlockListOnBackup Then
        EnableWindow lvDevices.hWnd, 1
        DebugMode "BlockListOnBackup: FALSE"
        lvDevices.Refresh
    End If

    ProgressBar1.SetTaskBarProgressState PrbTaskBarStateNone
        
    DebugMode "cmdStartBackUp_Click-Finish"
End Sub

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

    If InStr(1, strString, Kavichki) Then
        strString = Replace$(strString, Kavichki, vbNullString)
    End If

    StringCleaner = strString
End Function

Private Sub txtArchName_KeyPress(KeyAscii As Integer)

    Dim sTemplate As String

    sTemplate = "!@#$%^&*()_+=\/:;?><|[],"

    If InStr(1, sTemplate, Chr$(KeyAscii), vbTextCompare) Then
        KeyAscii = 0
    End If
End Sub

'! -----------------------------------------------------------
'!  �������     :  VerModules
'!  ����������  :
'!  ��������    :  ����������� ������ �������
'! -----------------------------------------------------------
Private Sub VerModules()

    MsgBox strMessages(35) & vbNewLine & "7za.exe (x86)" & vbTab & vbTab & objFSO.GetFileVersion(strArh7zExePATH) & vbNewLine & "7zSD.sfx (SFX-Module)" & vbTab & objFSO.GetFileVersion(strArh7zSFXPATH) & vbNewLine & "DPinst.exe (x86)" & vbTab & vbTab & objFSO.GetFileVersion(strDPInstExePath) & vbNewLine & "DPinst.exe (x64)" & vbTab & vbTab & objFSO.GetFileVersion(strDPInstExePath64), vbInformation, strProductName
End Sub

'[SourceDisksNames.x86]
'1 = %DiskId%,,,.\B_32846
'
'[SourceDisksNames.ia64]
'1 = %DiskID%,,,.\B_32846
'[SourceDisksFiles]
'ati2cqag.dll = 1
'ati2dvag.dll = 1
'[SourceDisksNames.x86]
'1 = %CD%,,,
'2 = %CD%,,,"drivers\dot4\Win2000"
'3 = %CD%,,,"drivers\dot4\WinxP"
'
'[SourceDisksNames]
'1 = %CD%,,,
'
'[SourceDisksFiles.x86]
'; Driver
'HPZius12.sys = 2
'; Co-Installer for w2k/XP, thunk for 9X
'HPZc3212.dll = 1
'HPZuci12.dll = 1
'Hppaufd0.sys = 3
'
'[SourceDisksFiles]
'; Driver
'HPZius12.sys = 1,Drivers\dot4\win98
'; Co-Installer for w2k/XP, thunk for 9X
'HPZc3212.dll = 1,Drivers\dot4\win98
'HPZuci12.dll = 1,Drivers\dot4\win98
'[SourceDisksNames]
'0 = %SRCDISK1%, "fjwia.cab", 0000-0000
'[SourceDisksFiles]
'fi4120.dll = 0
'[SourceDisksNames.x86]
'0=%DiskName%
'[SourceDisksNames.amd64]
'0=%DiskName%
'
'[SourceDisksFiles.x86]
'rimsptsk.sys=0,,
'snymsico.dll=0,,
'
'[SourceDisksFiles.amd64]
'rimspx64.sys=0,,
'snymsico.dll=0,,
'[SourceDisksNames]
'1 = %SrcDiskId%,,,
'
'[SourceDisksFiles.x86] ; files for x86
'sncduvc.sys = 1
'snp2uvc.sys = 1
'vsnp2uvc.dll = 1
'rsnp2uvc.dll = 1
'csnp2uvc.dll = 1
'PLFSet.dll = 1
'
'[SourceDisksFiles.amd64] ; files for AMD64
'sncduvc.sys = 1,.\x64,
'snp2uvc.sys = 1,.\x64,
'vsnp2uvc.dll = 1
'rsnp2uvc.dll = 1
'csnp2uvc.dll = 1,.\x64,
'vsnpvc64.dll = 1,.\x64,
'rsnpvc64.dll = 1,.\x64,
'PLFSet.dll = 1
'CheckIniSectionExists SekName, IniFileName
