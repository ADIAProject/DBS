VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Сбор информации о драйверах. Пожалуйста подождите..."
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin prjDIADBS.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      Top             =   120
      Width           =   7170
      _ExtentX        =   12647
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

Private mbRunProgress As Boolean
'Private WithEvents TaskBar2 As cITaskBarList3

Private Sub FontCharsetChange()

    ' Выставляем шрифт
    Me.Font.Name = strFontOtherForm_Name
    Me.Font.Size = lngFontOtherForm_Size
    Me.Font.Charset = lngFont_Charset
End Sub

Private Sub Form_Activate()

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
    Sleep 300
    Unload Me
    Set frmProgress = Nothing
End Sub

Private Sub Form_Load()

    mbRunProgress = True
    'Set TaskBar2 = New cITaskBarList3
        
    ' Локализациz приложения
    If mbMultiLanguage Then
        Localise strPCLangCurrentPath
    Else
        ' Выставляем шрифт
        FontCharsetChange
    End If

    ProgressBar1.Height = Me.Height - VPadding(Me)

'    If strOsCurrentVersion > "5.2" Then
'        ProgressBar1 = CCBorderStyleSingle
'    Else
'        ProgressBar1.BorderStyle = CCBorderStyleNone
'    End If
End Sub

'Private Sub Form_Terminate()
'
'    'Set TaskBar2 = Nothing
'    If Forms.Count = 0 Then
'        UnloadApp
'    End If
'End Sub

Private Sub Localise(StrPathFile As String)

    Dim strFormName As String

    strFormName = CStr(Me.Name)
    ' Выставляем шрифт элементов (действует только на те для которых не поддерживается Юникод)
    FontCharsetChange
    ' Название формы
    Me.Caption = LocaliseString(StrPathFile, strFormName, strFormName, Me.Caption)
End Sub

Public Sub ChangeProgressBarStatus(ByRef lngProgressValue As Long, ByVal lngProgressValuePlus As Long)

    lngProgressValue = lngProgressValue + lngProgressValuePlus

    If lngProgressValue > 10000 Then
        lngProgressValue = 10000
        Sleep 50
    End If
    ProgressBar1.Value = lngProgressValue

    With ProgressBar1
        .Value = lngProgressValue
        .SetTaskBarProgressValue .Value, .Max
    End With
    
    DoEvents
End Sub
