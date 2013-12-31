VERSION 5.00
Begin VB.UserControl ctlTextInteger 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   FillColor       =   &H8000000E&
   BeginProperty Font 
      Name            =   "Arial Unicode MS"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   510
   ScaleWidth      =   2205
   ToolboxBitmap   =   "ctlTextInteger.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   915
      Top             =   45
   End
   Begin VB.TextBox Te 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   510
      TabIndex        =   0
      Top             =   15
      Width           =   975
   End
End
Attribute VB_Name = "ctlTextInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================
' Copyright © 2003 Alexander Drovosekov (apexsun@narod.ru) '
' Visit apexsun.narod.ru                                   '
' ==========================================================
Public Event onChange(Text As String)

Private m_Range    As Long
Private m_Caption  As String
Private m_MinValue As Long
Private Rg         As Long

Public Property Get AlignText() As AlignmentConstants

    AlignText = Te.Alignment
End Property

Public Property Let AlignText(ByVal vNewValue As AlignmentConstants)

    Te.Alignment = vNewValue
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518

    Caption = m_Caption
End Property

Public Property Let Caption(ByRef vNewValue As String)

    m_Caption = vNewValue
    UserControl_Resize
    DrawRows 0, 0
End Property

Private Function DrawRows(ByVal Xcur As Single, ByRef State As Byte) As Byte

    Dim X     As Long
    Dim Y     As Long
    Dim X2    As Long
    Dim Y2    As Long
    Dim cFill As Long
    Dim cText As Long

    AutoRedraw = True
    Cls
    UserControl.CurrentY = 15
    Print m_Caption
    X = Te.Left - 15
    Y = 0
    X2 = Width - 30
    Y2 = Te.Height + 15
    UserControl.Line (X, Y)-(X2, Y2), Te.BackColor, BF
    UserControl.Line (X, Y)-(X2, Y2), &H80000010, B
    X = X + Te.Width + 15
    X = Te.Left + Te.Width + 30
    Y = 30
    X2 = Width - 60
    Y2 = Height - 45
    Xcur = Switch(Xcur > X + 125, 2, Xcur >= X, 1, Xcur < X, 0)

    If Xcur = 1 Then
        If State = 1 Then
            cFill = &H8000000E
            cText = &H80000012
        ElseIf State = 2 Then
            cFill = &H8000000D
            cText = &H8000000E
        End If

    Else
        cFill = &H8000000F
        cText = &H80000011
    End If

    UserControl.Line (X, Y)-(X + 120, Y2), cFill, BF
    UserControl.Line (X + 60, Y + 60)-(X + 60, Y + 75), cText
    UserControl.Line (X + 45, Y + 75)-(X + 85, Y + 75), cText
    UserControl.Line (X + 30, Y + 90)-(X + 100, Y + 90), cText

    If Xcur = 2 Then
        If State = 1 Then
            cFill = &H8000000E
            cText = &H80000012
        ElseIf State = 2 Then
            cFill = &H8000000D
            cText = &H8000000E
        End If

    Else
        cFill = &H8000000F
        cText = &H80000011
    End If

    UserControl.Line (X + 120, Y)-(X2, Y2), cFill, BF
    UserControl.Line (X + 180, Y2 - 60)-(X + 180, Y2 - 75), cText
    UserControl.Line (X + 165, Y2 - 75)-(X + 210, Y2 - 75), cText
    UserControl.Line (X + 150, Y2 - 90)-(X + 225, Y2 - 90), cText
    UserControl.Line (X, Y)-(X2, Y2), &H80000010, B
    UserControl.Line (X + 120, Y)-(X + 120, Y2), &H80000010
    AutoRedraw = False
    DrawRows = Xcur
End Function

Public Property Get MinValue() As Long

    MinValue = m_MinValue
End Property

Public Property Let MinValue(ByVal vNewValue As Long)

    m_MinValue = vNewValue
End Property

Public Property Get StepRange() As Long

    StepRange = m_Range
End Property

Public Property Let StepRange(ByVal vNewValue As Long)

    m_Range = vNewValue
End Property

Private Sub Te_Change()

    Static oText As String

    If Te.Enabled = False Then
        Exit Sub
    End If

    If IsNumeric(Te.Text) = False Then
        Te.Text = oText
    End If

    oText = Te.Text
    RaiseEvent onChange(oText)
    Te.Text = oText

    If CLng(Te.Text) < 0 Then
        Te.Text = 0
    End If

    If m_MinValue <> 0 Then
        If CLng(Te.Text) < m_MinValue And Te.Text >= 0 Then
            Te.BackColor = &HFF&
        Else
            Te.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub Te_GotFocus()

    Te.BackColor = &H80000005
    DrawRows 0, 0
End Sub

Private Sub Te_KeyDown(KeyCode As Integer, Shift As Integer)

    If Te.Text = vbNullString Or Not IsNumeric(Te.Text) Then
        Te.Text = 0
    End If

    If KeyCode = 38 Then
        Te.Text = Int(Te.Text + m_Range)
    ElseIf KeyCode = 40 Then
        Te.Text = Int(Te.Text - m_Range)
    ElseIf KeyCode = 33 Then
        Te.Text = Int(Te.Text + 10 * m_Range)
    ElseIf KeyCode = 34 Then
        Te.Text = Int(Te.Text - 10 * m_Range)
    End If
End Sub

Private Sub Te_KeyPress(KeyAscii As Integer)

    If KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Te_LostFocus()

    Te.BackColor = &H80000016
    DrawRows 0, 0
End Sub

Public Property Get Text() As String

    Text = Te.Text
End Property

Public Property Let Text(ByRef vNewValue As String)

    With Te

        If .Text <> vNewValue Then
            .Enabled = False
            .Text = vNewValue
            .Enabled = True
        End If
    End With
End Property

Private Sub Timer1_Timer()

    Te.Text = CInt(Te.Text + Val(Timer1.Tag) * m_Range)
    Timer1.Interval = 200 - Rg
    DrawRows Width - IIf(Val(Timer1.Tag) = 1, 260, 120), 2

    If Rg < 130 Then
        Rg = Rg + 3
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Te.Text = vbNullString Or Not IsNumeric(Te.Text) Then
        Te.Text = 0
    End If

    On Error Resume Next

    Y = DrawRows(X, 2)
    Y = Switch(Y = 2, -1, Y = 1, 1, Y = 0, 0)
    Timer1.Tag = Y
    Timer1.Interval = 500
    Rg = 0
    Timer1.Enabled = True
    Te.Text = CInt(Te.Text + Y * Button * m_Range)
    Call DrawRows(X, 2)

    On Error GoTo 0

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If GetCapture <> hwnd Then
        SetCapture hwnd
        DrawRows X, 1
    ElseIf X < Te.Left + Te.Width Or X > Width Or Y < 0 Or Y > Height Then
        ReleaseCapture
        DrawRows 0, 0
    Else
        DrawRows X, 1
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Timer1.Enabled = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_Range = .ReadProperty("Range", 1)
        m_Caption = .ReadProperty("Caption", "Name")
        Te.Text = .ReadProperty("Text", 0)
        Te.Alignment = .ReadProperty("TextAlign", 1)
        m_MinValue = .ReadProperty("MinValue", 0)
    End With

    UserControl_Resize
    DrawRows 0, 0
End Sub

Private Sub UserControl_Resize()

    Te.Left = TextWidth(m_Caption) + 75

    On Error Resume Next

    Te.Width = UserControl.Width - Te.Left - 330
    UserControl.Height = Te.Height + 30
    DrawRows 0, 0

    On Error GoTo 0

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Caption", m_Caption, "Name")
        Call .WriteProperty("Text", Te.Text, vbNullString)
        Call .WriteProperty("Range", m_Range, 1)
        Call .WriteProperty("TextAlign", Te.Alignment, 1)
        Call .WriteProperty("MinValue", m_MinValue, 0)
    End With
End Sub
