VERSION 5.00
Begin VB.UserControl ctlOptionBoxTVH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ToolboxBitmap   =   "ctlOptionBoxTVH.ctx":0000
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   5
      Left            =   1560
      Picture         =   "ctlOptionBoxTVH.ctx":0312
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   75
      Index           =   4
      Left            =   1440
      Picture         =   "ctlOptionBoxTVH.ctx":03A4
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1200
      Picture         =   "ctlOptionBoxTVH.ctx":0436
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   960
      Picture         =   "ctlOptionBoxTVH.ctx":0680
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   720
      Picture         =   "ctlOptionBoxTVH.ctx":08CA
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   480
      Picture         =   "ctlOptionBoxTVH.ctx":0B14
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "ctlOptionBoxTVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Private Enum E_CheckStatus
    eNormal = 0
    eMoveOver = 1
    eClickDown = 2
    eDisabled = 3
End Enum

#If False Then

    Private eNormal, eMoveOver, eClickDown, eDisabled
#End If

Public Enum E_AlignmentOpt
    eobLeft = 0
    eobRight = 1
    eobTop = 2
    eobBottom = 3
End Enum

#If False Then

    Private eobLeft, eobRight, eobTop, eobBottom
#End If

Private Const D          As Integer = 13
Private Const D_Box_Text As Integer = 3
Private bFocus           As Boolean
Private MouseEvent       As E_MouseEvent
Private bPrevButton      As Long
Private m_Caption        As String
Private m_Forecolor      As OLE_COLOR
Private m_BackColor      As OLE_COLOR
Private m_Transparent    As Boolean
Private m_Alignment      As E_AlignmentOpt
Private m_Shadow         As Boolean
Private m_ShadowColor    As OLE_COLOR
Private m_Enabled        As Boolean
Private m_Checked        As Boolean

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseLeave(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Public Property Get Alignment() As E_AlignmentOpt

    Alignment = m_Alignment
End Property

Public Property Let Alignment(new_Alignment As E_AlignmentOpt)

    m_Alignment = new_Alignment
    PropertyChanged "Alignment"
    Refresh
End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor
End Property

Public Property Let BackColor(new_BackColor As OLE_COLOR)

    m_BackColor = new_BackColor
    PropertyChanged "Backcolor"
    Refresh
End Property

Public Property Get Caption() As String

    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    PropertyChanged "Caption"
    Refresh
End Property

Public Property Get Checked() As Boolean

    Checked = m_Checked
End Property

Public Property Let Checked(ByVal new_Checked As Boolean)

    m_Checked = new_Checked
    PropertyChanged "Checked"
    Refresh
End Property

Public Sub ClearChecks()

    Dim o As Control

    For Each o In UserControl.Parent.Controls

        If (TypeOf o Is ctlOptionBoxTVH) Then
            If o.Container.hwnd = UserControl.ContainerHwnd Then
                If o.hwnd <> UserControl.hwnd Then
                    If o.Checked Then
                        o.Checked = False
                    End If
                End If
            End If
        End If

    Next
End Sub

Private Sub DrawCaption()

    Dim s             As String
    Dim t             As RECT
    Dim Flag          As Long

    Const D_Edge_Text As Integer = 1

    Dim dong          As Long

    With UserControl

        Select Case m_Alignment

            Case ecbLeft
                t.Left = D + D_Box_Text
                t.Right = .ScaleWidth - D_Edge_Text
                t.Top = 0
                t.Bottom = .ScaleHeight
                Flag = DT_LEFT

            Case ecbRight
                t.Left = D_Edge_Text
                t.Right = .ScaleWidth - D - D_Box_Text
                t.Top = 0
                t.Bottom = .ScaleHeight
                Flag = DT_RIGHT

            Case ecbTop
                t.Left = D_Edge_Text
                t.Right = .ScaleWidth - D_Edge_Text
                t.Top = D
                t.Bottom = .ScaleHeight
                Flag = DT_CENTER

            Case ecbBottom
                t.Left = D_Edge_Text
                t.Right = .ScaleWidth - D_Edge_Text
                t.Top = 0
                t.Bottom = .ScaleHeight - D
                Flag = DT_CENTER
        End Select

        Set Pic(0).Font = .Font
        's = IIf(m_TiengViet, mUnicode.VNI_Unicode(m_Caption), m_Caption)
        s = m_Caption
        dong = DrawTextW(Pic(0).hDC, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK)
        t.Top = t.Top + (t.Bottom - t.Top - dong) \ 2

        If m_Shadow And m_Enabled Then
            Offset t, 1
            .ForeColor = m_ShadowColor
            DrawTextW hDC, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK
            Offset t, -1
        End If

        .ForeColor = IIf(m_Enabled, m_Forecolor, RGB(167, 166, 170))
        DrawTextW hDC, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK

        If bFocus Then

            With t
                .Left = .Left - 1
                .Top = .Top - 2
                .Bottom = .Top + dong + 3
            End With

            't
            If dong / TextHeight(Left$(s, 1)) = 1 Then
                't.Right = t.Left + TextWidthW(.hdc, s) + 2
            Else
                't.Right = t.Right + 1
            End If

            .ForeColor = 0
            DrawFocusRect hDC, t
        End If
    End With
End Sub

Private Sub DrawOptionBoxXp2005(C As E_CheckStatus, Optional iCheck As Byte = 0)

    Dim Y  As Long
    Dim X  As Long
    Dim ID As Byte

    With UserControl

        If iCheck <> 1 Then
            If iCheck <> 2 Then
                iCheck = IIf(m_Checked, 1, 2)
            End If
        End If

        Select Case m_Alignment

            Case eobLeft
                Y = (.ScaleHeight - D) \ 2

            Case eobRight
                X = .ScaleWidth - D
                Y = (.ScaleHeight - D) \ 2

            Case eobTop
                X = (.ScaleWidth - D) \ 2
                Y = 0

            Case eobBottom
                X = (.ScaleWidth - D) \ 2
                Y = .ScaleHeight - D
        End Select

        Select Case C

            Case eClickDown
                ID = 2

            Case eDisabled
                ID = 3

            Case eMoveOver
                ID = 1

            Case eNormal
                ID = 0
        End Select

        TransparentBlt .hDC, X, Y, Pic(0).ScaleWidth, Pic(0).ScaleWidth, Pic(ID).hDC, 0, 0, Pic(0).ScaleWidth, Pic(0).ScaleWidth, vbRed

        If iCheck = 1 Then
            TransparentBlt .hDC, X + Pic(0).ScaleWidth \ 2 - Pic(5).ScaleWidth \ 2, Y + Pic(0).ScaleWidth \ 2 - Pic(5).ScaleHeight \ 2, Pic(5).ScaleWidth, Pic(5).ScaleHeight, Pic(IIf(C = eDisabled, 5, 4)).hDC, 0, 0, Pic(5).ScaleWidth, Pic(5).ScaleHeight, vbRed
        End If
    End With
End Sub

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    Refresh
End Property

Public Property Get Font() As StdFont

    Set Font = UserControl.Font
End Property

Public Property Set Font(New_Font As StdFont)

    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR

    ForeColor = m_Forecolor
End Property

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)

    m_Forecolor = New_ForeColor
    PropertyChanged "Forecolor"
    Refresh
End Property

Public Property Get hwnd() As Long

    hwnd = UserControl.hwnd
End Property

Private Sub Offset(t As RECT, ByVal D As Integer)

    With t
        .Left = .Left + D
        .Right = .Right + D
        .Top = .Top + D
        .Bottom = .Bottom + D
    End With

    't
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------
Public Sub Refresh()

    Dim ST As E_CheckStatus

    With UserControl
        .Cls
        .BackColor = IIf(m_Transparent, TransColor, m_BackColor)
        Set .Picture = Nothing

        If MouseEvent = eMouseLeaving Then
            ST = eNormal
        ElseIf MouseEvent = eMouseLeavingClicking Or MouseEvent = eMouseMoving Then
            ST = eMoveOver
        ElseIf MouseEvent = eMouseMovingClicking Then
            ST = eClickDown
        Else
            Exit Sub
        End If

        .Enabled = m_Enabled
        DrawOptionBoxXp2005 IIf(m_Enabled, ST, eDisabled)
        DrawCaption

        If m_Transparent Then
            .MaskColor = TransColor
            .MaskPicture = .Image
            .BackStyle = 0
        Else
            .BackStyle = 1
        End If
    End With
End Sub

Public Property Get ShadowColor() As OLE_COLOR

    ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(new_ShadowColor As OLE_COLOR)

    m_ShadowColor = new_ShadowColor
    PropertyChanged "ShadowColor"
    Refresh
End Property

Public Property Get Shadow() As Boolean

    Shadow = m_Shadow
End Property

Public Property Let Shadow(ByVal New_Shadow As Boolean)

    m_Shadow = New_Shadow
    PropertyChanged "Shadow"
    Refresh
End Property

Public Property Get Transparent() As Boolean

    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal new_Transparent As Boolean)

    m_Transparent = new_Transparent
    PropertyChanged "Transparent"
    Refresh
End Property

Private Sub UserControl_Click()

    If bPrevButton = 1 Then
        ClearChecks
        m_Checked = True
        Refresh
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()

    If bPrevButton = 1 Then
        UserControl_MouseDown 1, 0, 1, 1
    End If
End Sub

Private Sub UserControl_ExitFocus()

    bFocus = False
    Refresh
End Sub

Private Sub UserControl_GotFocus()

    bFocus = True

    If MouseEvent <> eMouseMovingClicking Then
        Refresh
    End If
End Sub

Private Sub UserControl_Initialize()

    Font.Name = "Arial Unicode MS"
    m_Caption = "Nu1t cho5n"
    m_Forecolor = 0
    m_Shadow = True
    m_ShadowColor = vbWhite
    m_BackColor = &H8000000F
    m_Enabled = True
    Pic(0).ScaleWidth = Pic(0).ScaleWidth
    Refresh
End Sub

Private Sub UserControl_InitProperties()

    'Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
        UserControl_Click
    End If

    If KeyCode = vbKeyRight Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyLeft Then
        SendKeys "+{TAB}"
    End If

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next

    If Button = 1 Then
        bFocus = True
        MouseEvent = eMouseMovingClicking
        Refresh
    End If

    bPrevButton = Button
    RaiseEvent MouseDown(Button, Shift, X, Y)
    UserControl.Parent.SetFocus

    On Error GoTo 0

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim t2   As POINT

    Static t As POINT

    With UserControl

        If X < 0 Or Y < 0 Or X > .ScaleWidth Or Y > .ScaleHeight Then
re:

            If Button = 1 And MouseEvent = eMouseMovingClicking Then
                MouseEvent = eMouseLeavingClicking
                Refresh
            ElseIf MouseEvent <> eMouseLeavingClicking Then
                MouseEvent = eMouseLeaving
                Refresh
            End If

            If Button <> 1 Then
                ReleaseCapture
            End If

            RaiseEvent MouseLeave(Button, Shift, X, Y)
        Else
            GetCursorPos t2

            If WindowFromPoint(t2.X, t2.Y) <> .hwnd Then
                GoTo re
            Else
                SetCapture hwnd
            End If

            If Button = 1 And MouseEvent = eMouseLeavingClicking Then
                MouseEvent = eMouseMovingClicking
                Refresh
            ElseIf Button = 1 Then
                MouseEvent = eMouseMovingClicking
                Refresh
            Else

                If t2.X = t.X Then
                    If t2.Y = t.Y Then
                        RaiseEvent MouseMove(Button, Shift, X, Y)
                        Exit Sub
                    End If
                End If

                MouseEvent = eMouseMoving
                Refresh
                GetCursorPos t
            End If

            RaiseEvent MouseMove(Button, Shift, X, Y)
        End If
    End With
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        If MouseEvent = eMouseMovingClicking Then
            MouseEvent = eMouseMoving
        Else
            MouseEvent = eMouseLeaving
        End If

        Refresh
    End If

    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set UserControl.Font = .ReadProperty("Font", Parent.Font)
        m_Forecolor = .ReadProperty("Forecolor", 0)
        m_BackColor = .ReadProperty("BackColor", Parent.BackColor)
        m_ShadowColor = .ReadProperty("ShadowColor", vbWhite)
        m_Shadow = .ReadProperty("Shadow", True)
        m_Alignment = .ReadProperty("Alignment", ecbLeft)
        m_Transparent = .ReadProperty("Transparent", False)
        m_Caption = .ReadProperty("Caption", "Nu1t kie63m")
        m_Enabled = .ReadProperty("Enabled", True)
        m_Checked = .ReadProperty("Checked", False)
        Refresh
    End With
End Sub

Private Sub UserControl_Resize()

    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Caption", m_Caption, "Nu1t kie63m"
        .WriteProperty "Forecolor", m_Forecolor, 0
        .WriteProperty "BackColor", m_BackColor, Parent.BackColor
        .WriteProperty "ShadowColor", m_ShadowColor, vbWhite
        .WriteProperty "Shadow", m_Shadow, True
        .WriteProperty "Alignment", m_Alignment, ecbLeft
        .WriteProperty "Transparent", m_Transparent, False
        .WriteProperty "Font", UserControl.Font, Parent.Font
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "Checked", m_Checked, False
        Refresh
    End With
End Sub
