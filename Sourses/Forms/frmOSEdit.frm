VERSION 5.00
Begin VB.Form frmOSEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������������� ������"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOSEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   8400
   StartUpPosition =   1  'CenterOwner
   Begin prjDIADBS.TextBoxW txtOSVer 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   5415
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
   Begin prjDIADBS.CheckBoxW chk64bit 
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
      _ExtentX        =   4683
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
      Caption         =   "frmOSEdit.frx":000C
      Transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlUcPickBox ucPathDRP 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   5415
      _ExtentX        =   10398
      _ExtentY        =   556
      UseAutoForeColor=   0   'False
      DefaultExt      =   ""
      Enabled         =   0   'False
      FileFlags       =   524288
      Filters         =   "Supported files|*.*|All Files (*.*)"
      UseDialogText   =   0   'False
      Locked          =   -1  'True
      QualifyPaths    =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   6480
      TabIndex        =   4
      Top             =   1320
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
      Caption         =   "��������� ��������� � �����"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   750
      Left            =   4560
      TabIndex        =   3
      Top             =   1320
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
      Caption         =   "����� ��� ����������"
      CaptionEffects  =   0
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
   End
   Begin prjDIADBS.LabelW lblPathDRP 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2655
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
      Caption         =   "���� �� �������� � �������� ���������"
   End
   Begin prjDIADBS.LabelW lblOSVer 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2655
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
      Caption         =   "������ ��"
   End
End
Attribute VB_Name = "frmOSEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName As String

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Property Get CaptionW
'! Description (��������)  :   [��������� Caption-�����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Property Get CaptionW() As String
    Dim lngLenStr As Long
    
    lngLenStr = DefWindowProc(Me.hWnd, WM_GETTEXTLENGTH, 0, ByVal 0)
    CaptionW = Space$(lngLenStr)
    DefWindowProc Me.hWnd, WM_GETTEXT, Len(CaptionW) + 1, ByVal StrPtr(CaptionW)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Property Let CaptionW
'! Description (��������)  :   [��������� Caption-�����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Property Let CaptionW(ByVal NewValue As String)
    DefWindowProc Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue & vbNullChar)
End Property

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdExit_Click
'! Description (��������)  :   [����� ��� ����������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub cmdOK_Click
'! Description (��������)  :   [��������� � �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    SaveOptions
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub FontCharsetChange
'! Description (��������)  :   [��������� ������ �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub FontCharsetChange()

    ' ���������� �����
    With Me.Font
        .Name = strFontOtherForm_Name
        .Size = lngFontOtherForm_Size
        .Charset = lngFont_Charset
    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Activate
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    txtOSVer_Change
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_KeyDown
'! Description (��������)  :   [��������� ������� ������ ����������]
'! Parameters  (����������):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Form_Load
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub Form_Load()
    SetupVisualStyles Me

    With Me
        strFormName = .Name
        SetIcon .hWnd, strFormName, False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    ' ������������� �������� ������
    LoadIconImage2Object cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2Object cmdExit, "BTN_EXIT", strPathImageMainWork

    ' ����������� ����������
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' ���������� �����
        FontCharsetChange
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub Localise
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal strPathFile As String)
    ' ���������� ����� ��������� (��������� ������ �� �� ��� ������� �� �������������� ������)
    FontCharsetChange
    ' �������� �����
    Me.CaptionW = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    ' ������
    lblOSVer.Caption = LocaliseString(strPathFile, strFormName, "lblOSVer", lblOSVer.Caption)
    lblPathDRP.Caption = LocaliseString(strPathFile, strFormName, "lblPathDRP", lblPathDRP.Caption)
    chk64bit.Caption = LocaliseString(strPathFile, strFormName, "chk64bit", chk64bit.Caption)
    '������
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
    ' ��������� �������� ������ ������ � ���������
    ucPathDRP.ToolTipTexts(ucFolder) = strMessages(152)
    ucPathDRP.DialogMsg(ucFolder) = strMessages(152)
     
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SaveOptions
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    Dim ii As Long

    If mbAddInList Then
        ii = lngLastIdOS + 1

        With frmOptions.lvOS.ListItems.Add(, , txtOSVer)
            .SubItems(2) = ucPathDRP.Path

            If chk64bit.Value Then
                .SubItems(1) = "1"
            Else
                .SubItems(1) = "1"
            End If
        End With

    Else

        With frmOptions.lvOS
            ii = .SelectedItem.Index
            .ListItems.item(ii).Text = txtOSVer
            .ListItems.item(ii).SubItems(2) = ucPathDRP.Path

            If chk64bit.Value Then
                .ListItems.item(ii).SubItems(1) = "1"
            Else
                .ListItems.item(ii).SubItems(1) = "0"
            End If
        End With


    End If

    lngLastIdOS = frmOptions.lvOS.ListItems.count
    frmOptions.lvOS.Refresh
    mbAddInList = False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtOSVer_Change
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtOSVer_Change()
    cmdOK.Enabled = LenB(Trim$(txtOSVer)) And LenB(Trim$(ucPathDRP.Path))
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtOSVer_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtOSVer_GotFocus()
    HighlightActiveControl Me, txtOSVer, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub txtOSVer_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub txtOSVer_LostFocus()
    HighlightActiveControl Me, txtOSVer, False
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucPathDRP_Click
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucPathDRP_Click()

    Dim strTempPath As String

    With ucPathDRP
        strTempPath = .Path

        If InStr(1, strTempPath, strAppPathBackSL, vbTextCompare) Then
            strTempPath = Replace$(strTempPath, strAppPathBackSL, vbNullString, , , vbTextCompare)
        End If
    
        If LenB(strTempPath) Then
            .Path = strTempPath
        End If
    End With
    
    txtOSVer_Change

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucPathDRP_GotFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucPathDRP_GotFocus()
    HighlightActiveControl Me, ucPathDRP, True
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ucPathDRP_LostFocus
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Private Sub ucPathDRP_LostFocus()
    HighlightActiveControl Me, ucPathDRP, False
End Sub
