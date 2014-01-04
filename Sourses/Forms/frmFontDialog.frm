VERSION 5.00
Begin VB.Form frmFontDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Locate Font and Color ..."
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjDIADBS.TextBoxW txtFont 
      Height          =   495
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   4275
      _extentx        =   7541
      _extenty        =   873
      font            =   "frmFontDialog.frx":0000
      text            =   "frmFontDialog.frx":0028
      alignment       =   2
      cuebanner       =   "frmFontDialog.frx":007A
   End
   Begin prjDIADBS.OptionButtonW optControl 
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
      _extentx        =   1720
      _extenty        =   450
      font            =   "frmFontDialog.frx":009A
      value           =   0   'False
      caption         =   "frmFontDialog.frx":00C2
   End
   Begin prjDIADBS.SpinBox txtFontSize 
      Height          =   315
      Left            =   1860
      TabIndex        =   4
      Top             =   420
      Width           =   675
      _extentx        =   1191
      _extenty        =   556
      font            =   "frmFontDialog.frx":00EE
      min             =   6
      max             =   20
      value           =   6
      allowonlynumbers=   -1  'True
   End
   Begin prjDIADBS.ctlColorButton ctlFontColor 
      Height          =   330
      Left            =   1980
      TabIndex        =   3
      Top             =   780
      Width           =   525
      _extentx        =   926
      _extenty        =   582
      icon            =   "frmFontDialog.frx":0116
   End
   Begin prjDIADBS.CheckBoxW chkItalic 
      Height          =   315
      Left            =   2700
      TabIndex        =   2
      Top             =   675
      Width           =   1575
      _extentx        =   2778
      _extenty        =   556
      font            =   "frmFontDialog.frx":0272
      caption         =   "frmFontDialog.frx":029A
      transparent     =   -1  'True
   End
   Begin prjDIADBS.CheckBoxW chkBold 
      Height          =   255
      Left            =   2700
      TabIndex        =   1
      Top             =   420
      Width           =   1575
      _extentx        =   2778
      _extenty        =   450
      font            =   "frmFontDialog.frx":02C6
      caption         =   "frmFontDialog.frx":02EE
      transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlFontCombo ctlFontCombo 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4335
      _extentx        =   7646
      _extenty        =   556
      previewtext     =   "ctlFontCombo1"
      combofontsize   =   10
      buttonovercolor =   0
      font            =   "frmFontDialog.frx":0316
   End
   Begin prjDIADBS.CheckBoxW chkUnderline 
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Top             =   960
      Width           =   1575
      _extentx        =   2778
      _extenty        =   556
      font            =   "frmFontDialog.frx":033E
      caption         =   "frmFontDialog.frx":0366
      transparent     =   -1  'True
   End
   Begin prjDIADBS.ctlJCbutton cmdOK 
      Height          =   750
      Left            =   2280
      TabIndex        =   7
      Top             =   1860
      Width           =   2100
      _extentx        =   3704
      _extenty        =   1323
      font            =   "frmFontDialog.frx":0398
      buttonstyle     =   13
      backcolor       =   12244692
      caption         =   "Сохранить изменения и выйти"
      picturealign    =   0
      picturepushonhover=   -1  'True
      pictureshadow   =   -1  'True
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin prjDIADBS.ctlJCbutton cmdExit 
      Height          =   750
      Left            =   60
      TabIndex        =   8
      Top             =   1860
      Width           =   2100
      _extentx        =   3704
      _extenty        =   1323
      font            =   "frmFontDialog.frx":03C0
      buttonstyle     =   13
      backcolor       =   12244692
      caption         =   "Выход без сохранения"
      picturealign    =   0
      picturepushonhover=   -1  'True
      pictureshadow   =   -1  'True
      captioneffects  =   0
      tooltipbackcolor=   0
      colorscheme     =   3
   End
   Begin prjDIADBS.LabelW lblFontSize 
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   420
      Width           =   1815
      _extentx        =   3201
      _extenty        =   661
      font            =   "frmFontDialog.frx":03E8
      backstyle       =   0
      caption         =   "Размер шрифта"
   End
   Begin prjDIADBS.LabelW lblFontColor 
      Height          =   375
      Left            =   60
      TabIndex        =   11
      Top             =   840
      Width           =   1815
      _extentx        =   3201
      _extenty        =   661
      font            =   "frmFontDialog.frx":0410
      backstyle       =   0
      caption         =   "Цвет шрифта"
   End
End
Attribute VB_Name = "frmFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFormName As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkBold_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkBold_Click()
    ctlFontCombo.ComboFontBold = chkBold.Value = 1
    txtFont.Font.Bold = chkBold.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkItalic_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkItalic_Click()
    ctlFontCombo.ComboFontItalic = chkItalic.Value = 1
    txtFont.Font.Italic = chkItalic.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlFontColor_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ctlFontColor_Click()
    txtFont.ForeColor = ctlFontColor.BackColor
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub chkUnderline_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub chkUnderline_Click()
    txtFont.Font.Underline = chkUnderline.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlFontCombo_FontNotFound
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   FontName (String)
'!--------------------------------------------------------------------------------
Private Sub ctlFontCombo_FontNotFound(FontName As String)
    MsgBox "Cant find this font: " & FontName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ctlFontCombo_SelectedFontChanged
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewFontName (String)
'!--------------------------------------------------------------------------------
Private Sub ctlFontCombo_SelectedFontChanged(NewFontName As String)
    txtFont.Font.Name = NewFontName
    ctlFontCombo.ClearUsedList
    ctlFontCombo.AddToUsedList NewFontName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_Activate
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Form_Activate()
    ctlFontCombo.SelectedFont = txtFont.Font.Name
    txtFontSize.Value = txtFont.Font.Size
    ctlFontCombo.PreviewText = txtFont.Text
    ctlFontCombo.AddToUsedList txtFont.Font.Name
    chkBold.Value = txtFont.Font.Bold
    chkItalic.Value = txtFont.Font.Italic
    chkUnderline.Value = txtFont.Font.Underline
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Form_KeyDown
'! Description (Описание)  :   [обработка нажатий клавиш клавиатуры]
'! Parameters  (Переменные):   KeyCode (Integer)
'                              Shift (Integer)
'!--------------------------------------------------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

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
        SetIcon .hWnd, "frmFontDialog", False
        .Left = (lngRightWorkArea - lngLeftWorkArea) / 2 - .Width / 2
        .Top = (lngBottomWorkArea - lngTopWorkArea) / 2 - .Height / 2
    End With

    ' Устанавливаем картинки кнопок и убираем описание кнопок
    LoadIconImage2BtnJC cmdOK, "BTN_SAVE", strPathImageMainWork
    LoadIconImage2BtnJC cmdExit, "BTN_EXIT", strPathImageMainWork

    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    txtFontSize.Min = 6
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Localise
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Private Sub Localise(ByVal StrPathFile As String)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
    ' Лэйблы
    lblFontSize.Caption = LocaliseString(StrPathFile, strFormName, "lblFontSize", lblFontSize.Caption)
    lblFontColor.Caption = LocaliseString(StrPathFile, strFormName, "lblFontColor", lblFontColor.Caption)
    chkBold.Caption = LocaliseString(StrPathFile, strFormName, "chkBold", chkBold.Caption)
    chkItalic.Caption = LocaliseString(StrPathFile, strFormName, "chkItalic", chkItalic.Caption)
    chkUnderline.Caption = LocaliseString(StrPathFile, strFormName, "chkUnderline", chkUnderline.Caption)
    txtFont.Text = LocaliseString(StrPathFile, strFormName, "txtFont", txtFont.Text)
    'Кнопки
    cmdOK.Caption = LocaliseString(StrPathFile, strFormName, "cmdOK", cmdOK.Caption)
    cmdExit.Caption = LocaliseString(StrPathFile, strFormName, "cmdExit", cmdExit.Caption)
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

    SetBtnFontProperties cmdExit
    SetBtnFontProperties cmdOK
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtFont_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtFont_Change()
    ctlFontCombo.PreviewText = txtFont.Text
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdExit_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub cmdOK_Click
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub cmdOK_Click()
    SaveOptions
    Unload Me
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveOptions
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveOptions()

    With txtFont
            
        If optControl.Item(0).Value Then
            strFontBtn_Name = .Font.Name
            miFontBtn_Size = .Font.Size
            mbFontBtn_Underline = .Font.Underline
            mbFontBtn_Strikethru = .Font.Strikethrough
            mbFontBtn_Bold = .Font.Bold
            mbFontBtn_Italic = .Font.Italic
            lngFontBtn_Color = .ForeColor
            SetBtnFontProperties frmOptions.cmdFutureButton
            frmOptions.cmdFutureButton.ForeColor = .ForeColor
            'frmOptions.cmdFutureButton.Refresh
        End If

    End With

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtFontSize_Change
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtFontSize_Change()
    txtFont.Font.Size = txtFontSize.Value
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub txtFontSize_TextChange
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub txtFontSize_TextChange()
    txtFont.Font.Size = txtFontSize.Value
End Sub
