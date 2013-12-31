VERSION 5.00
Begin VB.UserControl ctlCheckBoxTVH 
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ToolboxBitmap   =   "ctlCheckBoxTVH.ctx":0000
   Begin VB.PictureBox TPic 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1935
      ScaleHeight     =   75
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   225
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "ctlCheckBoxTVH"
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
    eGotFocus = 1
    eMoveOver = 2
    eClickDown = 3
    eDisabled = 4
End Enum

#If False Then

    Private eNormal, eGotFocus, eMoveOver, eClickDown, eDisabled
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
Private m_Alignment      As E_AlignmentCheckBox
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

Public Property Get Alignment() As E_AlignmentCheckBox

    Alignment = m_Alignment
End Property

Public Property Let Alignment(new_Alignment As E_AlignmentCheckBox)

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

Private Sub DrawCaption()

    Dim s             As String
    Dim t             As RECT
    Dim Flag          As Long

    Const D_Edge_Text As Integer = 1

    '<:-
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

        Set TPic.Font = .Font
        s = m_Caption
        dong = DrawTextW(TPic.hDC, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK)
        t.Top = t.Top + (t.Bottom - t.Top - dong) \ 2

        If m_Shadow Then
            If m_Enabled Then
                Offset t, 1
                .ForeColor = m_ShadowColor
                DrawTextW hDC, StrPtr(s), Len(s), t, Flag Or DT_NOCLIP Or DT_WORDBREAK
                Offset t, -1
            End If
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
                t.Right = t.Left + TextWidthW(hDC, s) + 2
            Else
                t.Right = t.Right + 1
            End If

            .ForeColor = 0
            DrawFocusRect hDC, t
        End If
    End With
End Sub

Private Sub DrawCheckBoxXp2005(C As E_CheckStatus, Optional iCheck As Byte = 0)

    'iCheck =1 : Checked
    'iCheck =2 : UnChecked
    Dim Y      As Long
    Dim X      As Long
    Dim i      As Byte
    Dim ArrC() As Long

    With UserControl

        If iCheck <> 1 Then
            If iCheck <> 2 Then
                iCheck = IIf(m_Checked, 1, 2)
            End If
        End If

        Select Case m_Alignment

            Case ecbLeft
                Y = (.ScaleHeight - D) \ 2

            Case ecbRight
                X = .ScaleWidth - D
                Y = (.ScaleHeight - D) \ 2

            Case ecbTop
                X = (.ScaleWidth - D) \ 2
                Y = 0

            Case ecbBottom
                X = (.ScaleWidth - D) \ 2
                Y = .ScaleHeight - D
        End Select

        Select Case C

            Case eNormal
                PSet (X + 1, Y + 1), RGB(226, 226, 221)
                Line (X + 1, Y + 2)-(X + 3, Y), RGB(226, 226, 221)
                Line (X + 1, Y + 3)-(X + 4, Y), RGB(226, 226, 221)
                PSet (X + D - 2, Y + D - 2), RGB(255, 255, 255)
                Line (X + D - 3, Y + D - 2)-(X + D - 1, Y + D - 4), RGB(255, 255, 255)
                Line (X + D - 4, Y + D - 2)-(X + D - 1, Y + D - 5), RGB(255, 255, 255)
                GradientColor2 RGB(226, 226, 221), RGB(255, 255, 255), (D - 4) * 2 - 1, ArrC

                For i = 0 To D - 6
                    Line (X + 1, Y + i + 4)-(X + i + 5, Y), ArrC(i + 1)
                    Line (X + 1 + i, Y + D - 2)-(X + D - 1, Y + i), ArrC(D - 5 + i)
                Next

            Case eMoveOver
                GradientColor2 RGB(255, 240, 207), RGB(248, 179, 48), (D - 2) * 2 - 1, ArrC

                For i = 0 To D - 3
                    Line (X + 1, Y + i + 1)-(X + i + 2, Y), ArrC(i)
                    Line (X + 1 + i, Y + D - 2)-(X + D - 1, Y + i), ArrC(D - 3 + i)
                Next

                For i = 0 To 6
                    Line (X + (D - 7) \ 2, Y + (D - 7) \ 2 + i)-(X + (D - 7) \ 2 + 7, Y + (D - 7) \ 2 + i), RGB(247, 247, 245)
                Next

            Case eClickDown
                GradientColor2 RGB(176, 176, 167), RGB(241, 239, 223), (D - 2) * 2 - 1, ArrC

                For i = 0 To D - 3
                    Line (X + 1, Y + i + 1)-(X + i + 2, Y), ArrC(i)
                    Line (X + 1 + i, Y + D - 2)-(X + D - 1, Y + i), ArrC(D - 3 + i)
                Next

            Case eDisabled
                .ForeColor = RGB(198, 197, 201)
                .FillStyle = 0
                .FillColor = vbWhite
                Rectangle hDC, X, Y, X + D, Y + D
                .FillStyle = 1
        End Select

        .ForeColor = IIf(C = eDisabled, RGB(198, 197, 201), RGB(28, 81, 128))
        Rectangle hDC, X, Y, X + D, Y + D

        If iCheck = 1 Then
            Line (X + (D - 7) \ 2, Y + (D - 7) \ 2 + 2)-(X + (D - 7) \ 2, Y + (D - 7) \ 2 + 5), IIf(C = eDisabled, RGB(198, 197, 201), RGB(33, 161, 33))
            Line (X + (D - 7) \ 2 + 1, Y + (D - 7) \ 2 + 3)-(X + (D - 7) \ 2 + 1, Y + (D - 7) \ 2 + 6), IIf(C = eDisabled, RGB(198, 197, 201), RGB(33, 161, 33))

            For i = 0 To 4
                Line (X + (D - 7) \ 2 + 2 + i, Y + (D - 7) \ 2 + 4 - i)-(X + (D - 7) \ 2 + 2 + i, Y + (D - 7) \ 2 + 7 - i), IIf(C = eDisabled, RGB(198, 197, 201), RGB(33, 161, 33))
            Next
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

Private Sub GradientColor2(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Depth As Integer, Result() As Long)

    Dim VR As Single
    Dim VG As Single
    Dim VB As Single
    Dim r  As Integer
    Dim G  As Integer
    Dim B  As Integer
    Dim R2 As Integer
    Dim G2 As Integer
    Dim B2 As Integer
    Dim t  As Long

    If Not Depth < 1 Then
        If Depth = 1 Then
            ReDim Result(0) As Long
            Result(0) = Color1
            Exit Sub
        End If

        t = (Color1 And 255)
        r = t And 255
        t = Int(Color1 / 256)
        G = t And 255
        t = Int(Color1 / 65536)
        B = t And 255
        t = (Color2 And 255)
        R2 = t And 255
        t = Int(Color2 / 256)
        G2 = t And 255
        t = Int(Color2 / 65536)
        B2 = t And 255
        VR = Abs(r - R2) / (Depth - 1)
        VG = Abs(G - G2) / (Depth - 1)
        VB = Abs(B - B2) / (Depth - 1)

        If R2 < r Then
            VR = -VR
        End If

        If G2 < G Then
            VG = -VG
        End If

        If B2 < B Then
            VB = -VB
        End If

        ReDim Result(Depth - 1) As Long

        For t = 0 To Depth - 1
            R2 = r + VR * t
            G2 = G + VG * t
            B2 = B + VB * t
            Result(t) = RGB(R2, G2, B2)
        Next
    End If
End Sub

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
        DrawCheckBoxXp2005 IIf(m_Enabled, ST, eDisabled)
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

Private Function TextWidthW(lngHDc As Long, s As String) As Long

    Dim sz As Size

    GetTextExtentPoint32 lngHDc, StrPtr(s), Len(s), sz
    TextWidthW = sz.cx
End Function

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
        m_Checked = Not (m_Checked)
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
    m_Caption = "Nu1t kie63m"
    m_Forecolor = 0
    m_Shadow = True
    m_ShadowColor = vbWhite
    m_BackColor = &H8000000F
    m_Enabled = True
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

            If MouseEvent = eMouseMovingClicking Then
                If Button = 1 Then
                    MouseEvent = eMouseLeavingClicking
                    Refresh
                End If

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

            If MouseEvent = eMouseLeavingClicking Then
                If Button = 1 Then
                    MouseEvent = eMouseMovingClicking
                    Refresh
                End If

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
