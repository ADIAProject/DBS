Attribute VB_Name = "mChooseFont"
Option Explicit

Public Const LF_FACESIZE As Long = 32       'Font Dialog

'���������, �������� � ������ ChooseFont
'����� ����������� �������������� ������
Public Type LOGFONT
    lfHeight                                 As Long
    lfWidth                                  As Long
    lfEscapement                             As Long
    lfOrientation                            As Long
    lfWeight                                 As Long
    lfItalic                                 As Byte
    lfUnderline                              As Byte
    lfStrikeOut                              As Byte
    lfCharSet                                As Byte
    lfOutPrecision                           As Byte
    lfClipPrecision                          As Byte
    lfQuality                                As Byte
    lfPitchAndFamily                         As Byte
    lfFaceName(LF_FACESIZE)                  As Byte
End Type

Public Type tLOGFONT
    lfHeight                                As Long
    lfWidth                                 As Long
    lfEscapement                            As Long
    lfOrientation                           As Long
    lfWeight                                As Long
    lfItalic                                As Byte
    lfUnderline                             As Byte
    lfStrikeOut                             As Byte
    lfCharSet                               As Byte
    lfOutPrecision                          As Byte
    lfClipPrecision                         As Byte
    lfQuality                               As Byte
    lfPitchAndFamily                        As Byte
    lfFaceName                              As String * 32
End Type

'��������� � ����������� � ������ ��� ������� ChooseFont � ��.
Public Type CHOOSEFONT
    lStructSize As Long
    hWndOwner As Long           '  caller's window handle
    hDC As Long                 '  printer DC/IC or NULL
    lpLogFont As Long           '  ptr. to a LOGFONT struct
    iPointSize As Long          '  10 * size in points of selected font
    flags As Long               '  enum. type flags
    rgbColors As Long           '  returned text color
    lCustData As Long           '  data passed to hook fn.
    lpfnHook As Long            '  ptr. to hook function
    lpTemplateName As String    '  custom template name
    hInstance As Long           '  instance handle of.EXE that
    lpszStyle As String         '  return the style field here
    nFontType As Integer        '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long            '  minimum pt size allowed &
    nSizeMax As Long            '  max pt size allowed if
End Type

Private Const CF_INITTOLOGFONTSTRUCT As Long = &H40
Private Const SCREEN_FONTTYPE        As Long = &H2000
Private Const BOLD_FONTTYPE          As Long = &H100
Private Const FW_BOLD                As Integer = 700
Private Const LOGPIXELSY             As Integer = 90

'�������� �������� ����������
'�������� ���������� �� ����������
'���������� ������ � �����
'�������� ������ ������ ������
'������ ���������� ����� �� ������ �� ��������� LOGFONT
'�������� ���� �� �������� ���������� �� ������ ��������� ����������
'�������� ��� ������ �� ������ ��������� ����������
'Public Enum CommonDialog_Flags
'��� ���������� ������ ������� ������� ��. ���� Flags.htm
Private Const cdlCFWYSIWYG           As Long = &H8000
Private Const cdlCFBoth              As Long = &H3
Private Const cdlCFEffects           As Long = &H100
Private Const cdlCFScreenFonts       As Long = &H1

'����� ���� ������ �� ��������� ����������
'����� ���� ��� �������� ����� �������� �������� ������.
Private hWndTargetFont               As Long

'����������, �������� ���������� �������� ������
Public strDialog_FontName            As String
Public miDialog_FontSize             As Integer
Public mboolDialog_Italic            As Boolean
Public mboolDialog_Underline         As Boolean
Public mboolDialog_Strikethru        As Boolean
Public mboolDialog_Bold              As Boolean
Public lngDialog_Color               As Long

Private lngDialog_Language           As Long

Public lngDialog_Charset             As Long

'����������, �������� ���������� �������� ������ TAB
Public strDialogTab_FontName         As String
Public miDialogTab_FontSize          As Integer
Public mboolDialogTab_Italic         As Boolean
Public mboolDialogTab_Underline      As Boolean
Public mboolDialogTab_Strikethru     As Boolean
Public mboolDialogTab_Bold           As Boolean
Public lngDialogTab_Color            As Long

' Character sets:
Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const HANGUL_CHARSET = 129
Public Const GB2312_CHARSET = 134
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255
Public Const JOHAB_CHARSET = 130
Public Const HEBREW_CHARSET = 177
Public Const ARABIC_CHARSET = 178
Public Const GREEK_CHARSET = 161
Public Const TURKISH_CHARSET = 162
Public Const VIETNAMESE_CHARSET = 163
Public Const THAI_CHARSET = 222
Public Const EASTEUROPE_CHARSET = 238
Public Const RUSSIAN_CHARSET = 204
Public Const MAC_CHARSET = 77
Public Const BALTIC_CHARSET = 186

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (p1 As Any, p2 As Any) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long

Public Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreateFontIndirectT Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As tLOGFONT) As Long

Private Declare Function GetTextFace _
                Lib "gdi32.dll" _
                Alias "GetTextFaceA" (ByVal hDC As Long, _
                                      ByVal nCount As Long, _
                                      ByVal lpFacename As String) As Long

Public Sub GetButtonProperties(ByVal ButtonName As ctlXpButton)

    With ButtonName
        'hWndTargetFont = .mhwnd
        '���������� ��������� ������� ������� ������� � ����������
        strDialog_FontName = .Font.Name
        miDialog_FontSize = .Font.Size
        mboolDialog_Underline = .Font.Underline
        mboolDialog_Strikethru = .Font.Strikethrough
        mboolDialog_Bold = .Font.Bold
        mboolDialog_Italic = .Font.Italic
        lngDialog_Color = .TextColor
        lngDialog_Language = .Font.Charset
    End With
End Sub

Public Sub GetButtonPropertiesJC(ByVal ButtonName As ctlJCbutton)

    With ButtonName
        hWndTargetFont = .hwnd
        '���������� ��������� ������� ������� ������� � ����������
        strDialog_FontName = .Font.Name
        miDialog_FontSize = .Font.Size
        mboolDialog_Underline = .Font.Underline
        mboolDialog_Strikethru = .Font.Strikethrough
        mboolDialog_Bold = .Font.Bold
        mboolDialog_Italic = .Font.Italic
        lngDialog_Color = .ForeColor
    End With
End Sub

Public Sub GetTabProperties(ButtonName As ctlXpButton)

    With ButtonName
        hWndTargetFont = .mhwnd
        '���������� ��������� ������� ������� ������� � ����������
        strDialogTab_FontName = .Font.Name
        miDialogTab_FontSize = .Font.Size
        mboolDialogTab_Underline = .Font.Underline
        mboolDialogTab_Strikethru = .Font.Strikethrough
        mboolDialogTab_Bold = .Font.Bold
        mboolDialogTab_Italic = .Font.Italic
        lngDialogTab_Color = .TextColor
    End With
End Sub

Public Sub SetButtonProperties(Optional ByVal ButtonName As ctlXpButton, _
                               Optional ByVal ButtonNameJC As ctlJCbutton, _
                               Optional ByVal IsJCButton As Boolean = False)

    Dim ctlObject As Object

    '���������� ��������� ������� ������� ������� � ����������
    If IsJCButton Then
        Set ctlObject = ButtonNameJC
    Else
        Set ctlObject = ButtonName
    End If

    With ctlObject
        .Font.Name = strDialog_FontName
        .Font.Size = miDialog_FontSize
        .Font.Underline = mboolDialog_Underline
        .Font.Strikethrough = mboolDialog_Strikethru
        .Font.Bold = mboolDialog_Bold
        .Font.Italic = mboolDialog_Italic
        .Font.Charset = lngDialog_Charset
        '.ForeColor = lngDialog_Color
    End With
End Sub

Public Sub ShowFont()

    Dim DialogFlags     As Long
    Dim CF              As CHOOSEFONT
    Dim LF              As LOGFONT
    Dim TempByteArray() As Byte
    Dim ByteArrayLimit  As Long
    Dim FontToUse       As Long
    Dim tbuf            As String * 80
    Dim X               As Long
    Dim uFlag           As Long
    Dim retvalue        As Long

    '��� ��������� �������� ���� ������� ������ ������
    '��������� � ����������� � ������
    '�������� ����� � ��� �����
    '��������� ����������� ��������� ������� � ���������� ����
    '�������� �������� ��. � ����� Flags.htm
    DialogFlags = cdlCFBoth Or cdlCFEffects
    uFlag = DialogFlags And (&H1 Or &H2 Or &H3 Or &H4 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H10000 Or &H20000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
    '�������������� ����� ������ �� ������� � ANSI
    TempByteArray = StrConv(strDialog_FontName & vbNullChar, vbFromUnicode)
    ByteArrayLimit = UBound(TempByteArray)

    '������ ����� ������ ��� ����� ������
    '���������� ��������� LogFont ���������� �������� �������������� ������
    With LF

        For X = 0 To ByteArrayLimit
            .lfFaceName(X) = TempByteArray(X)
            '��� ������
        Next X

        '������ ������, ��� GetDeviceCaps(hDC, LOGPIXELSY) - ���������� �������� �� ���� �� ������
        .lfHeight = miDialog_FontSize / 72 * GetDeviceCaps(GetDC(hWndTargetFont), LOGPIXELSY)
        '������
        .lfItalic = mboolDialog_Italic * -1
        '������������
        .lfUnderline = mboolDialog_Underline * -1
        '�������������
        .lfStrikeOut = mboolDialog_Strikethru * -1

        '����������
        If mboolDialog_Bold Then
            .lfWeight = FW_BOLD
        End If

        .lfCharSet = lngDialog_Language
    End With

    '���������� ��������� ChooseFont ����� �������������� � ������� ChooseFont
    With CF
        '����� ���������
        .lStructSize = Len(CF)
        '����� ���� ����
        .hWndOwner = hWndTargetFont
        '����� ��������� ���������� ����
        .hDC = GetDC(hWndTargetFont)
        '����� ��� �������� ��������� LOGFONT, ������� �������� ������ � �������������� ������
        .lpLogFont = lstrcpy(LF, LF)

        '����������
        If Not uFlag Then
            .flags = cdlCFScreenFonts
            .flags = uFlag
        Else
            .flags = uFlag Or cdlCFWYSIWYG
        End If

        '����� ������� ����������� ����
        .flags = .flags Or cdlCFEffects Or CF_INITTOLOGFONTSTRUCT
        '������� ����
        .rgbColors = lngDialog_Color
        .lCustData = 0
        .lpfnHook = 0
        .lpTemplateName = 0
        .hInstance = 0
        .lpszStyle = 0
        '��� ������
        .nFontType = SCREEN_FONTTYPE
        .nSizeMin = 0
        .nSizeMax = 0
        '������ ������
        .iPointSize = miDialog_FontSize * 10
    End With

    '����� ����������� ����
    retvalue = CHOOSEFONT(CF)

    If retvalue = 0 Then
        '<:-) :SUGGESTION: Empty 'If X Then' structure could be Replaced with 'If Not X Then' and 'Else' removed.
        '����� ���� ������� ��������
        'If mCancelError Then
        'Err.Raise (retvalue)
        'End If
    Else

        '����� ���� ������� ���������
        With LF
            mboolDialog_Italic = .lfItalic * -1
            mboolDialog_Underline = .lfUnderline * -1
            mboolDialog_Strikethru = .lfStrikeOut * -1
            lngDialog_Language = .lfCharSet
        End With

        With CF
            miDialog_FontSize = .iPointSize \ 10
            lngDialog_Color = .rgbColors
            mboolDialog_Bold = .nFontType And BOLD_FONTTYPE
        End With

        '�������� ����������� ������
        FontToUse = CreateFontIndirect(LF)

        If FontToUse = 0 Then
            Exit Sub
        End If

        '���� ����� �� ������ - ����� �� �������
        '������� ������ "�����" �� ����������(���� ���������� - PictureBox)
        SelectObject CF.hDC, FontToUse
        '�������� ��� ������ ���������� � �����
        retvalue = GetTextFace(CF.hDC, 79, tbuf)
        '��������� ��� ������ � ����������
        strDialog_FontName = Mid$(tbuf, 1, retvalue)
    End If

    '�������� ���� ������ ����������
    SetTextColor CF.hDC, lngDialog_Color
End Sub
