VERSION 5.00
Begin VB.UserControl ctlLabelTVH 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
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
   ToolboxBitmap   =   "ctlLabelTVH.ctx":0000
End
Attribute VB_Name = "ctlLabelTVH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Enum E_ShadowStyle
    sLeftTop = 0
    sRightTop = 1
    sLeftBottom = 2
    sRightBottom = 3
End Enum

Enum E_Alignment
    aLeft = 0
    aRight = 1
    aCenter = 2
End Enum

Enum E_BorderStyle
    None = 0
    Flat = 1
    Outline = 2
    [3D] = 3
    Frame1 = 4
    Frame2 = 5
End Enum

Enum E_BackColorStyle
    aSingleColor = 0
    aGradientColor = 1
End Enum

Enum E_GradientBackColorStyle
    aLeftToRight = 0
    aRightToLeft = 1
    aTopToBottom = 2
    aBottomToTop = 3
    aLeftTopToRightBottom = 4
    aLeftBottomToRightTop = 5
    aCenterToLeftRight = 6
    aCenterToTopBottom = 7
    aCenterToLeftTopNRightBottom = 8
    aCenterToLeftBottomNRightTop = 9
End Enum

'--------------------Private
'_____Attribute
Private m_AutoSize               As Boolean
Private m_BackColor              As OLE_COLOR
Private m_Forecolor              As OLE_COLOR
Private m_BorderColor            As OLE_COLOR
Private m_BorderStyle            As E_BorderStyle
Private m_BorderSize             As Long
Private m_Caption                As String
Private m_Font                   As StdFont
Private m_WordWrap               As Boolean
Private m_Transparent            As Boolean
Private m_EdgeSpace              As Long
Private m_OutlineColor           As OLE_COLOR
Private m_Shadow                 As Boolean
Private m_ShadowDepth            As Integer
Private m_ShadowStyle            As E_ShadowStyle
Private m_ShadowColorStart       As OLE_COLOR
Private m_ShadowColorEnd         As OLE_COLOR
Private m_Alignment              As E_Alignment
Private m_LineCount              As Long
Private m_BackColorStyle         As E_BackColorStyle
Private m_GradientBackColorStyle As E_GradientBackColorStyle
Private m_GradientBackColorStart As OLE_COLOR
Private m_GradientBackColorEnd   As OLE_COLOR
'----------------------End Private
'--------------------------------------------------
'--------Events--------------------------------
'--------------------------------------------------
Event Change()
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseLeave(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get Alignment() As E_Alignment

    Alignment = m_Alignment
End Property

Public Property Let Alignment(new_Alignment As E_Alignment)

    m_Alignment = new_Alignment
    PropertyChanged "Alignment"
    Fresh
End Property

Public Property Get AutoSize() As Boolean

    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(new_AutoSize As Boolean)

    m_AutoSize = new_AutoSize
    PropertyChanged "AutoSize"
    Fresh
End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor
End Property

Public Property Let BackColor(new_BackColor As OLE_COLOR)

    m_BackColor = new_BackColor
    PropertyChanged "BackColor"
    Fresh
End Property

Public Property Get BackColorStyle() As E_BackColorStyle

    BackColorStyle = m_BackColorStyle
End Property

Public Property Let BackColorStyle(new_BackColorStyle As E_BackColorStyle)

    m_BackColorStyle = new_BackColorStyle
    PropertyChanged "BackColorStyle"
    Fresh
End Property

Public Property Get BorderColor() As OLE_COLOR

    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(new_BorderColor As OLE_COLOR)

    m_BorderColor = new_BorderColor
    PropertyChanged "BorderColor"
    Fresh
End Property

Public Property Get BorderSize() As Long

    BorderSize = m_BorderSize
End Property

Public Property Let BorderSize(new_BorderSize As Long)

    If new_BorderSize >= 0 Then m_BorderSize = new_BorderSize
    PropertyChanged "BorderSize"
    Fresh
End Property

Public Property Get BorderStyle() As E_BorderStyle

    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(new_BorderStyle As E_BorderStyle)

    m_BorderStyle = new_BorderStyle
    PropertyChanged "BorderStyle"
    Fresh
End Property

'--------------------------------------------------
'--------Properties---------------------------
'--------------------------------------------------
Public Property Get Caption() As String

    Caption = m_Caption
End Property

Public Property Let Caption(New_Caption As String)

    m_Caption = New_Caption
    PropertyChanged "Caption"
    Fresh
End Property

Private Sub Draw(s As String)

    Dim tR As RECT
    Dim i  As Integer

    With UserControl
        .BackColor = IIf(m_Transparent, TransColor, m_BackColor)
        Cls
        Set .Picture = Nothing
        tR.Left = BorderSize + EdgeSpace + IIf(m_Shadow And (m_ShadowStyle = sLeftTop Or m_ShadowStyle = sLeftBottom), m_ShadowDepth, 0)
        tR.Top = BorderSize + EdgeSpace + IIf(m_Shadow And (m_ShadowStyle = sLeftTop Or m_ShadowStyle = sRightTop), m_ShadowDepth, 0)
        tR.Bottom = .ScaleHeight - BorderSize - EdgeSpace
        tR.Right = .ScaleWidth - BorderSize - EdgeSpace

        If m_BackColorStyle = aGradientColor And m_Transparent = False Then DrawGradientBackColor m_GradientBackColorStyle
        If m_Shadow Then
            DrawShadow s, tR, m_WordWrap, m_ShadowColorStart, ShadowColorEnd, m_ShadowDepth, m_ShadowStyle
        End If

        .ForeColor = m_Forecolor
        m_LineCount = DrawTextFunc(.hDC, s, tR, m_WordWrap) / TextHeight(Left$(s, 1))

        Select Case m_BorderStyle

            Case 0

                'None
            Case 1
                'Flat
                .ForeColor = m_BorderColor

                For i = 1 To BorderSize
                    Rectangle .hDC, i - 1, i - 1, .ScaleWidth - i + 1, .ScaleHeight - i + 1
                Next

            Case 2
                'Outline
                .ForeColor = m_OutlineColor
                DrawOutline .hDC, s, tR, m_WordWrap, BorderSize
                .ForeColor = m_Forecolor
                DrawTextFunc .hDC, s, tR, m_WordWrap

            Case 3, 4, 5
                '3D,Frame1,Frame2
                .ForeColor = m_Forecolor
                m_BorderSize = IIf(m_BorderStyle = 5, 3, 2)
                PropertyChanged "BorderSize"

                If m_BorderStyle = 3 Then
                    Line (0, 0)-(.ScaleWidth - 1, 0), RGB(167, 166, 170)
                    Line (0, 0)-(0, .ScaleHeight - 1), RGB(167, 166, 170)
                    Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight), RGB(255, 255, 255)
                    Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), RGB(255, 255, 255)
                    Line (1, 1)-(.ScaleWidth - 2, 1), RGB(133, 135, 140)
                    Line (1, 1)-(1, .ScaleHeight - 2), RGB(133, 135, 140)
                    Line (.ScaleWidth - 2, 1)-(.ScaleWidth - 2, .ScaleHeight - 1), RGB(220, 223, 228)
                    Line (1, .ScaleHeight - 2)-(.ScaleWidth - 1, .ScaleHeight - 2), RGB(220, 223, 228)
                ElseIf m_BorderStyle = 4 Then
                    Line (0, 0)-(.ScaleWidth - 1, 0), RGB(167, 166, 170)
                    Line (0, 0)-(0, .ScaleHeight - 1), RGB(167, 166, 170)
                    Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight), RGB(255, 255, 255)
                    Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), RGB(255, 255, 255)
                    Line (1, 1)-(.ScaleWidth - 2, 1), RGB(255, 255, 255)
                    Line (1, 1)-(1, .ScaleHeight - 2), RGB(255, 255, 255)
                    Line (.ScaleWidth - 2, 1)-(.ScaleWidth - 2, .ScaleHeight - 1), RGB(167, 166, 170)
                    Line (1, .ScaleHeight - 2)-(.ScaleWidth - 1, .ScaleHeight - 2), RGB(167, 166, 170)
                Else
                    .ForeColor = RGB(255, 255, 255)
                    Rectangle .hDC, 0, 0, .ScaleWidth, .ScaleHeight
                    Rectangle .hDC, 2, 2, .ScaleWidth - 2, .ScaleHeight - 2
                    .ForeColor = RGB(167, 166, 170)
                    Rectangle .hDC, 1, 1, .ScaleWidth - 1, .ScaleHeight - 1
                End If
        End Select

        If m_Transparent Then
            .BackStyle = 0
            .MaskColor = TransColor
            .MaskPicture = .Image
        Else
            .BackStyle = 1
        End If
    End With
End Sub

Private Sub DrawGradientBackColor(Style As E_GradientBackColorStyle)

    Dim C()   As Long
    Dim Depth As Integer
    Dim i     As Integer

    With UserControl

        '(IIf(i < .ScaleHeight, 0, i - .ScaleHeight + 1), IIf(i < .ScaleHeight, i, .ScaleHeight - 1))-(IIf(i < .ScaleWidth, i + 1, .ScaleWidth), IIf(i < .ScaleWidth, -1, i - .ScaleWidth)), c(i)
        On Error GoTo er

        If Style = aLeftToRight Or Style = aRightToLeft Then
            Depth = .ScaleWidth
        ElseIf Style = aTopToBottom Or Style = aBottomToTop Then
            Depth = .ScaleHeight
        ElseIf Style = aLeftTopToRightBottom Or Style = aLeftBottomToRightTop Then
            Depth = 2 * .ScaleWidth - 1
        ElseIf Style = aCenterToLeftRight Then
            Depth = (.ScaleWidth + 1) \ 2
        ElseIf Style = aCenterToTopBottom Then
            Depth = (.ScaleHeight + 1) \ 2
        ElseIf Style = aCenterToLeftTopNRightBottom Or aCenterToLeftBottomNRightTop Then
            Depth = .ScaleWidth
        End If

        GradientColor2 m_GradientBackColorStart, m_GradientBackColorEnd, Depth, C

        Select Case Style

            Case aLeftToRight:

                For i = 0 To .ScaleWidth - 1
                    Line (i, 0)-(i, .ScaleHeight), C(i)
                Next

            Case aRightToLeft:

                For i = 0 To .ScaleWidth - 1
                    Line (i, 0)-(i, .ScaleHeight), C(.ScaleWidth - i - 1)
                Next

            Case aTopToBottom:

                For i = 0 To .ScaleHeight - 1
                    Line (0, i)-(.ScaleWidth, i), C(i)
                Next

            Case aBottomToTop:

                For i = 0 To .ScaleHeight - 1
                    Line (0, i)-(.ScaleWidth, i), C(.ScaleHeight - i - 1)
                Next

            Case aLeftTopToRightBottom:
                'For i = 0 To .ScaleWidth + .ScaleHeight - 2
                '    t.Line (IIf(i < .ScaleHeight, 0, i - .ScaleHeight + 1), IIf(i < .ScaleHeight, i, .ScaleHeight - 1))-(IIf(i < .ScaleWidth, i + 1, .ScaleWidth), IIf(i < .ScaleWidth, -1, i - .ScaleWidth)), c(i)
                'Next i
                PSet (0, 0), C(0)
                PSet (.ScaleWidth - 1, .ScaleHeight - 1), C(.ScaleWidth + .ScaleWidth - 2)

                For i = 1 To .ScaleWidth + .ScaleWidth - 3
                    Line (i - .ScaleWidth + 1, .ScaleHeight - 1)-(i + 1, -1), C(i)
                Next

            Case aLeftBottomToRightTop:
                PSet (0, .ScaleHeight - 1), C(0)
                PSet (.ScaleWidth - 1, 0), C(.ScaleWidth + .ScaleWidth - 2)

                For i = 1 To .ScaleWidth + .ScaleWidth - 3
                    Line (i - .ScaleWidth + 1, -1)-(i + 1, .ScaleHeight - 1), C(i)
                Next

            Case aCenterToLeftRight:

                For i = 0 To .ScaleWidth \ 2 + (.ScaleWidth Mod 2) - 1
                    Line (i, 0)-(i, .ScaleHeight), C(.ScaleWidth \ 2 + (.ScaleWidth Mod 2) - 1 - i)
                    Line (i + .ScaleWidth \ 2, 0)-(i + .ScaleWidth \ 2, .ScaleHeight), C(i)
                Next

            Case aCenterToTopBottom:

                For i = 0 To .ScaleHeight \ 2 + (.ScaleHeight Mod 2) - 1
                    Line (0, i)-(.ScaleWidth, i), C(.ScaleHeight \ 2 + (.ScaleHeight Mod 2) - 1 - i)
                    Line (0, i + .ScaleHeight \ 2)-(.ScaleWidth, i + .ScaleHeight \ 2), C(i)
                Next

            Case aCenterToLeftTopNRightBottom:

                For i = 0 To Depth
                    Line (i - .ScaleWidth + 1, .ScaleHeight - 1)-(i + 1, -1), C(Depth - i)
                    Line (i, .ScaleHeight - 1)-(i + .ScaleWidth, -1), C(i)
                Next

            Case aCenterToLeftBottomNRightTop

                For i = 0 To Depth
                    Line (i - .ScaleWidth + 1, 0)-(i + 1, .ScaleHeight), C(Depth - i)
                    Line (i, 0)-(i + .ScaleWidth, .ScaleHeight), C(i)
                Next
        End Select
    End With

er:
End Sub

Private Sub DrawOutline(hDC As Long, s As String, tRect As RECT, wwrap As Boolean, Size As Integer)

    Dim tx As Integer, ty As Integer
    Dim t  As RECT

    t = tRect

    For ty = tRect.Top - Size To tRect.Top + Size
        For tx = tRect.Left - Size To tRect.Left + Size
            t.Right = t.Right - (t.Left - tx)
            t.Bottom = t.Bottom - (t.Top - ty)
            t.Left = tx
            t.Top = ty
            DrawTextFunc hDC, s, t, wwrap
        Next
    Next
End Sub

Private Sub DrawShadow(s As String, _
                       tR As RECT, _
                       wwrap As Boolean, _
                       Color_S As Long, _
                       Color_E As Long, _
                       Depth As Integer, _
                       Style As E_ShadowStyle)

    Dim AColor() As Long
    Dim t        As RECT
    Dim i        As Integer, dx As Integer, dy As Integer

    t = tR
    GradientColor2 Color_E, Color_S, Depth + 1, AColor

    Select Case Style

        Case 0
            dx = -1
            dy = -1

            'LeftTop
        Case 1
            dx = 1
            dy = -1

            'RightTop
        Case 2
            dx = -1
            dy = 1

            'LeftBottom
        Case 3
            dx = 1
            dy = 1
            'RightBottom
    End Select

    For i = Depth To 1 Step -1
        UserControl.ForeColor = AColor(Depth - i)
        t.Left = tR.Left + i * dx
        t.Top = tR.Top + i * dy
        t.Right = tR.Right + i * dx
        t.Bottom = tR.Bottom + i * dy
        DrawTextFunc UserControl.hDC, s, t, wwrap
    Next
End Sub

Private Function DrawTextFunc(hDC As Long, s As String, tRect As RECT, wwrap As Boolean) As Long

    If LenB(Trim$(s)) = 0 Then
        Exit Function
    End If

    Dim flags As Long

    flags = DT_NOCLIP
    flags = flags Or IIf(wwrap, DT_WORDBREAK, 0)
    flags = flags Or IIf(m_Alignment = 0, 0, IIf(m_Alignment = 1, DT_RIGHT, DT_CENTER))
    DrawTextFunc = DrawTextW(hDC, StrPtr(s), Len(s), tRect, flags)
End Function

Public Property Get EdgeSpace() As Long

    EdgeSpace = m_EdgeSpace
End Property

Public Property Let EdgeSpace(new_EdgeSpace As Long)

    If new_EdgeSpace >= 0 Then m_EdgeSpace = new_EdgeSpace
    PropertyChanged "EdgeSpace"
    Fresh
End Property

Public Property Get Font() As StdFont

    Set Font = m_Font
End Property

Public Property Set Font(New_Font As StdFont)

    Set m_Font = New_Font
    PropertyChanged "Font"
    Set UserControl.Font = m_Font
    Fresh
End Property

Public Property Get ForeColor() As OLE_COLOR

    ForeColor = m_Forecolor
End Property

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)

    m_Forecolor = New_ForeColor
    PropertyChanged "ForeColor"
    Fresh
End Property

'--------------------------------------------------
'--------Function more---------------------
'--------------------------------------------------
Private Sub Fresh()

    Dim s As String

    With UserControl

        Dim t As RECT

        s = m_Caption

        If m_AutoSize Then

            Dim dong As Long

            Set .Font = m_Font

            If m_WordWrap Then
                t.Left = BorderSize + EdgeSpace
                t.Top = BorderSize + EdgeSpace
                t.Bottom = 100
                t.Right = .ScaleWidth - BorderSize - EdgeSpace - IIf(m_Shadow, m_ShadowDepth, 0)
                dong = DrawTextFunc(.hDC, s, t, True)
            Else
                .Width = (TextWidthW(.hDC, s) + 2 * BorderSize + 2 * EdgeSpace + IIf(m_Shadow, m_ShadowDepth, 0)) * 15
                dong = .TextHeight("A")
            End If

            .Height = (dong + 2 * BorderSize + 2 * EdgeSpace + IIf(m_Shadow, m_ShadowDepth, 0)) * 15
        End If

        Draw s
        RaiseEvent Change
    End With
End Sub

Public Property Get GradientBackColorEnd() As OLE_COLOR

    GradientBackColorEnd = m_GradientBackColorEnd
End Property

Public Property Let GradientBackColorEnd(new_GradientBackColorEnd As OLE_COLOR)

    m_GradientBackColorEnd = new_GradientBackColorEnd
    PropertyChanged "GradientBackColorEnd"
    Fresh
End Property

Public Property Get GradientBackColorStart() As OLE_COLOR

    GradientBackColorStart = m_GradientBackColorStart
End Property

Public Property Let GradientBackColorStart(new_GradientBackColorStart As OLE_COLOR)

    m_GradientBackColorStart = new_GradientBackColorStart
    PropertyChanged "GradientBackColorStart"
    Fresh
End Property

Public Property Get GradientBackColorStyle() As E_GradientBackColorStyle

    GradientBackColorStyle = m_GradientBackColorStyle
End Property

Public Property Let GradientBackColorStyle(new_GradientBackColorStyle As E_GradientBackColorStyle)

    m_GradientBackColorStyle = new_GradientBackColorStyle
    PropertyChanged "GradientBackColorStyle"
    Fresh
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

Public Property Get hDC() As Long

    hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long

    hwnd = UserControl.hwnd
End Property

Public Property Get LineCount() As Long

    LineCount = m_LineCount
End Property

Public Property Get OutlineColor() As OLE_COLOR

    OutlineColor = m_OutlineColor
End Property

Public Property Let OutlineColor(new_OutlineColor As OLE_COLOR)

    If new_OutlineColor >= 0 Then m_OutlineColor = new_OutlineColor
    PropertyChanged "OutlineColor"
    Fresh
End Property

Public Property Get ShadowColorEnd() As OLE_COLOR

    ShadowColorEnd = m_ShadowColorEnd
End Property

Public Property Let ShadowColorEnd(new_ShadowColorEnd As OLE_COLOR)

    m_ShadowColorEnd = new_ShadowColorEnd
    PropertyChanged "ShadowColorEnd"
    Fresh
End Property

Public Property Get ShadowColorStart() As OLE_COLOR

    ShadowColorStart = m_ShadowColorStart
End Property

Public Property Let ShadowColorStart(new_ShadowColorStart As OLE_COLOR)

    m_ShadowColorStart = new_ShadowColorStart
    PropertyChanged "ShadowColorStart"
    Fresh
End Property

Public Property Get ShadowDepth() As Integer

    ShadowDepth = m_ShadowDepth
End Property

Public Property Let ShadowDepth(new_ShadowDepth As Integer)

    If new_ShadowDepth >= 0 Then m_ShadowDepth = new_ShadowDepth
    PropertyChanged "ShadowDepth"
    Fresh
End Property

Public Property Get Shadow() As Boolean

    Shadow = m_Shadow
End Property

Public Property Let Shadow(New_Shadow As Boolean)

    m_Shadow = New_Shadow
    PropertyChanged "Shadow"
    Fresh
End Property

Public Property Get ShadowStyle() As E_ShadowStyle

    ShadowStyle = m_ShadowStyle
End Property

Public Property Let ShadowStyle(new_ShadowStyle As E_ShadowStyle)

    m_ShadowStyle = new_ShadowStyle
    PropertyChanged "ShadowStyle"
    Fresh
End Property

Private Function TextWidthW(lngHDc As Long, s As String) As Long

    Dim sz As Size

    GetTextExtentPoint32 lngHDc, StrPtr(s), Len(s), sz
    TextWidthW = sz.cx
End Function

Public Property Get Transparent() As Boolean

    Transparent = m_Transparent
End Property

Public Property Let Transparent(new_Transparent As Boolean)

    m_Transparent = new_Transparent
    PropertyChanged "Transparent"
    Fresh
End Property

Private Sub UserControl_Click()

    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()

    RaiseEvent DblClick
End Sub

'--------------------------------------------------
'--------UserControl-------------------------
'--------------------------------------------------
Private Sub UserControl_Initialize()

    m_Caption = "Nha4n"
    m_BorderSize = 1
    m_EdgeSpace = 1
    m_BackColor = &H8000000F
    m_BorderColor = 0
    m_Forecolor = 0
    m_OutlineColor = &HFFFFFF
    m_Shadow = True
    m_ShadowDepth = 1
    m_ShadowColorStart = 0
    m_ShadowColorEnd = vbWhite
    Set m_Font = UserControl.Font
    Fresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

    On Error Resume Next

    UserControl.Parent.SetFocus
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With UserControl

        If X < 0 Or Y < 0 Or X > .ScaleWidth Or Y > .ScaleHeight Then
re:
            ReleaseCapture
            RaiseEvent MouseLeave(Button, Shift, X, Y)
        Else

            Dim t2 As POINT

            GetCursorPos t2

            If WindowFromPoint(t2.X, t2.Y) <> .hwnd Then
                GoTo re
            Else
                SetCapture hwnd
            End If

            RaiseEvent MouseMove(Button, Shift, X, Y)
        End If
    End With
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set m_Font = .ReadProperty("Font", Parent.Font)
        Set UserControl.Font = m_Font
        m_Caption = .ReadProperty("Caption", "Nha4n")
        m_AutoSize = .ReadProperty("AutoSize", False)
        m_WordWrap = .ReadProperty("WordWrap", False)
        m_BackColor = .ReadProperty("BackColor", &H8000000F)
        m_Forecolor = .ReadProperty("ForeColor", 0)
        m_BorderColor = .ReadProperty("BorderColor", 0)
        m_BorderSize = .ReadProperty("BorderSize", 1)
        m_BorderStyle = .ReadProperty("BorderStyle", 0)
        m_Transparent = .ReadProperty("Transparent", True)
        m_EdgeSpace = .ReadProperty("EdgeSpace", 1)
        m_OutlineColor = .ReadProperty("OutlineColor", &HFFFFFF)
        m_Shadow = .ReadProperty("Shadow", False)
        m_ShadowDepth = .ReadProperty("ShadowDepth", 1)
        m_ShadowStyle = .ReadProperty("ShadowStyle", sLeftTop)
        m_ShadowColorStart = .ReadProperty("ShadowColorStart", m_Forecolor)
        m_ShadowColorEnd = .ReadProperty("ShadowColorEnd", vbWhite)
        m_Alignment = .ReadProperty("Alignment", 0)
        m_BackColorStyle = .ReadProperty("BackColorStyle", 0)
        m_GradientBackColorStyle = .ReadProperty("GradientBackColorStyle", 0)
        m_GradientBackColorStart = .ReadProperty("GradientBackColorStart", 0)
        m_GradientBackColorEnd = .ReadProperty("GradientBackColorEnd", vbWhite)
        Fresh
    End With
End Sub

Private Sub UserControl_Resize()

    Fresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Font", m_Font, Parent.Font)
        Call .WriteProperty("Caption", m_Caption, "Nha4n")
        Call .WriteProperty("AutoSize", m_AutoSize, False)
        Call .WriteProperty("WordWrap", m_WordWrap, False)
        Call .WriteProperty("BackColor", m_BackColor, &H8000000F)
        Call .WriteProperty("ForeColor", m_Forecolor, 0)
        Call .WriteProperty("BorderColor", m_BorderColor, 0)
        Call .WriteProperty("BorderSize", m_BorderSize, 1)
        Call .WriteProperty("BorderStyle", m_BorderStyle, 0)
        Call .WriteProperty("Transparent", m_Transparent, True)
        Call .WriteProperty("EdgeSpace", m_EdgeSpace, 1)
        Call .WriteProperty("OutlineColor", m_OutlineColor, vbWhite)
        Call .WriteProperty("Shadow", m_Shadow, False)
        Call .WriteProperty("ShadowDepth", m_ShadowDepth, 1)
        Call .WriteProperty("ShadowStyle", m_ShadowStyle, sRightBottom)
        Call .WriteProperty("ShadowColorStart", m_ShadowColorStart, m_Forecolor)
        Call .WriteProperty("ShadowColorEnd", m_ShadowColorEnd, vbWhite)
        Call .WriteProperty("Alignment", m_Alignment, 0)
        Call .WriteProperty("BackColorStyle", m_BackColorStyle, 0)
        Call .WriteProperty("GradientBackColorStyle", m_GradientBackColorStyle, 0)
        Call .WriteProperty("GradientBackColorStart", m_GradientBackColorStart, 0)
        Call .WriteProperty("GradientBackColorEnd", m_GradientBackColorEnd, vbWhite)
    End With
End Sub

Public Property Get WordWrap() As Boolean

    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(new_WordWrap As Boolean)

    m_WordWrap = new_WordWrap
    PropertyChanged "WordWrap"
    Fresh
End Property
