VERSION 5.00
Begin VB.Form frmOSEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������������� ������"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
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
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjDBS.ctlCheckBoxTVH chk64bit 
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
      _ExtentX        =   4683
      _ExtentY        =   370
      Caption         =   "�������� ��� 64-������ ��"
      Transparent     =   -1  'True
   End
   Begin VB.TextBox txtOSVer 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
   Begin prjDBS.ctlLabelTVH lblPathDRP 
      Height          =   495
      Left            =   120
      Top             =   720
      Width           =   2655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "���� �� �������� � �������� ���������"
      WordWrap        =   -1  'True
      ShadowStyle     =   0
   End
   Begin prjDBS.ctlLabelTVH lblOSVer 
      Height          =   375
      Left            =   120
      Top             =   240
      Width           =   2655
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "������ ��"
      ShadowStyle     =   0
   End
   Begin prjDBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   6480
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1323
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "��������� ��������� � �����"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin prjDBS.ctlJCbutton cmdExit 
      Height          =   735
      Left            =   4560
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "����� ��� ����������"
      PictureAlign    =   0
      PicturePushOnHover=   -1  'True
      PictureShadow   =   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin prjDBS.ctlUcPickBox ucPathDRP 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   5415
      _ExtentX        =   10398
      _ExtentY        =   556
      Color           =   0
      DefaultExt      =   ""
      DialogType      =   1
      Enabled         =   0   'False
      Filters         =   "Supported files|*.*|All Files (*.*)"
      ToolTipText3    =   "Click Here to Locate File"
   End
End
Attribute VB_Name = "frmOSEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'! -----------------------------------------------------------
'!  �������     :  cmdExit_Click
'!  ����������  :
'!  ��������    :
'! -----------------------------------------------------------
Private Sub cmdExit_Click()

    Unload Me
End Sub

'! -----------------------------------------------------------
'!  �������     :  cmdOK_Click
'!  ����������  :
'!  ��������    :
'! -----------------------------------------------------------
Private Sub cmdOK_Click()

    SaveOptions
    Unload Me
End Sub

Private Sub FontCharsetChange()

    ' ���������� �����
    Me.Font.Name = strOtherForm_FontName
    Me.Font.Size = lngOtherForm_FontSize
    Me.Font.Charset = lngDialog_Charset
    SetButtonProperties , cmdExit, True
    SetButtonProperties , cmdOK, True
End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_KeyDown
'!  ����������  :  KeyCode As Integer, Shift As Integer
'!  ��������    :  ��������� ������� ������ ����������
'! -----------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

'! -----------------------------------------------------------
'!  �������     :  Form_Load
'!  ����������  :
'!  ��������    :
'! -----------------------------------------------------------
Private Sub Form_Load()

    ' ������������� �������� ������ � ������� �������� ������
    'SetSmallIcon Me.hWnd
    'LoadIconImage2Btn cmdPathDRP, "BTN_OPEN", strPathImageMainWork
    'cmdPathDRP.Caption = vbNullString
    
    Call SetIcon(Me.hwnd, "FRMOSEDIT", False)
    
    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork

    ' ����������z ����������
    If mboolLanguageChange Then
        Localise strPCLangCurrentPath
    Else
        ' ���������� �����
        FontCharsetChange
    End If
End Sub

Private Sub Form_Terminate()

    If Forms.Count = 0 Then
        UnloadApp
    End If
End Sub

Private Sub Localise(strPathFile As String)

    Dim strFormName As String

    strFormName = CStr(Me.Name)
    ' ���������� ����� ��������� (��������� ������ �� �� ��� ������� �� �������������� ������)
    FontCharsetChange
    ' �������� �����
    Me.Caption = LocaliseString(strPathFile, strFormName, strFormName, Me.Caption)
    ' ������
    lblOSVer.Caption = LocaliseString(strPathFile, strFormName, "lblOSVer", lblOSVer.Caption)
    'lblPathDRP.Caption = LocaliseString(strPathFile, strFormName, "lblPathDRP", lblPathDRP.Caption)
    chk64bit.Caption = LocaliseString(strPathFile, strFormName, "chk64bit", chk64bit.Caption)
    '������
    cmdOK.Caption = LocaliseString(strPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(strPathFile, strFormName, "cmdExit", cmdExit.Caption)
End Sub

'! -----------------------------------------------------------
'!  �������     :  SaveOptions
'!  ����������  :
'!  ��������    :
'! -----------------------------------------------------------
Private Sub SaveOptions()

    Dim i As Long

    If mboolAddInList Then
        i = LastIdOS + 1

        With frmOptions
            .lvOS.AddItem txtOSVer, , i - 1
            .lvOS.ItemText(2, i - 1) = ucPathDRP.Path

            If chk64bit.Checked Then
                .lvOS.ItemText(1, i - 1) = "1"
            Else
                .lvOS.ItemText(1, i - 1) = "0"
            End If
        End With

        'FRMOPTIONS
    Else

        With frmOptions
            i = .lvOS.SelectedItem
            .lvOS.ItemText(0, i) = txtOSVer
            .lvOS.ItemText(2, i) = ucPathDRP.Path

            If chk64bit.Checked Then
                .lvOS.ItemText(1, i) = "1"
            Else
                .lvOS.ItemText(1, i) = "0"
            End If
        End With

        'FRMOPTIONS
    End If

    LastIdOS = frmOptions.lvOS.Count
    frmOptions.lvOS.Refresh
    mboolAddInList = False
End Sub

'! -----------------------------------------------------------
'!  �������     :  ucPathDRP_Click
'!  ����������  :
'!  ��������    :  ����� �������� ��� �����
'! -----------------------------------------------------------
Private Sub ucPathDRP_Click()

    Dim strTempPath As String

    strTempPath = ucPathDRP.Path

    If InStr(1, strTempPath, strAppPath, vbTextCompare) > 0 Then
        strTempPath = Replace$(strTempPath, strAppPath, vbNullString, , , vbTextCompare)
    End If


    If LenB(strTempPath) > 0 Then
        ucPathDRP.Path = strTempPath
    End If
End Sub
