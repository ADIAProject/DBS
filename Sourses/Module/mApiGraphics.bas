Attribute VB_Name = "mApiGraphics"
Option Explicit

' --Formatting Text Consts
Public Const DT_LEFT       As Long = &H0
Public Const DT_CENTER     As Long = &H1
Public Const DT_RIGHT      As Long = &H2
Public Const DT_NOCLIP     As Long = &H100
Public Const DT_WORDBREAK  As Long = &H10
Public Const DT_CALCRECT   As Long = &H400
Public Const DT_RTLREADING As Long = &H20000 ' Right to left
Public Const DT_DRAWFLAG   As Long = DT_CENTER Or DT_WORDBREAK
Public Const DT_TOP        As Long = &H0
Public Const DT_BOTTOM     As Long = &H8
Public Const DT_VCENTER    As Long = &H4
Public Const DT_SINGLELINE As Long = &H20
Public Const DT_WORD_ELLIPSIS = &H40000
Public Const TransColor      As Long = &H8000000F

'   DrawEdge Message Constants
Public Const BDR_RAISEDOUTER As Long = &H1
Public Const BDR_SUNKENOUTER As Long = &H2
Public Const BDR_RAISEDINNER As Long = &H4
Public Const BDR_SUNKENINNER As Long = &H8
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const BF_LEFT      As Long = &H1
Public Const BF_TOP       As Long = &H2
Public Const BF_RIGHT     As Long = &H4
Public Const BF_BOTTOM    As Long = &H8
Public Const BF_RECT      As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_SUNKEN95 As Long = &HA
Public Const BDR_RAISED95 As Long = &H5

' §§§§§§§§§§§§§§§§§§§§§§§§§§ ImageList §§§§§§§§§§§§§§§§§§§§§§§§§§
Public Const SM_CXICON                    As Long = 11
Public Const SM_CYICON                    As Long = 12
Public Const SM_CYSMICON                  As Long = 50
Public Const SM_CXSMICON                  As Long = 49

' --System Hand Pointer
Public Const IDC_HAND     As Long = 32649

' --drawing Icon Constants
Public Const DI_NORMAL    As Long = &H3

Public Type Size
    cx                           As Long
    cy                           As Long
End Type

Public Type RECT
    Left                         As Long
    Top                          As Long
    Right                        As Long
    Bottom                       As Long
End Type

Public Type POINT
    X                            As Long
    Y                            As Long
End Type

Public Type RGB
    Red                                As Byte
    Green                              As Byte
    Blue                               As Byte
End Type

Public Type RGBTRIPLE
    rgbBlue                                 As Byte
    rgbGreen                                As Byte
    rgbRed                                  As Byte
End Type

'  RGB Colors structure
Public Type RGBColor
    r                                       As Single
    G                                       As Single
    B                                       As Single
End Type

Public Type RGBQUAD
    rgbBlue                                 As Byte
    rgbGreen                                As Byte
    rgbRed                                  As Byte
    rgbAlpha                                As Byte
End Type

Public Type ICONINFO
    fIcon                                   As Long
    xHotspot                                As Long
    yHotspot                                As Long
    hbmMask                                 As Long
    hbmColor                                As Long
End Type

'  for gradient painting and bitmap tiling
Public Type BITMAPINFOHEADER
    biSize                                  As Long
    biWidth                                 As Long
    biHeight                                As Long
    biPlanes                                As Integer
    biBitCount                              As Integer
    biCompression                           As Long
    biSizeImage                             As Long
    biXPelsPerMeter                         As Long
    biYPelsPerMeter                         As Long
    biClrUsed                               As Long
    biClrImportant                          As Long
End Type

'flicker free drawing
Public Type BITMAP
    bmType                                  As Long
    bmWidth                                 As Long
    bmHeight                                As Long
    bmWidthBytes                            As Long
    bmPlanes                                As Integer
    bmBitsPixel                             As Integer
    bmBits                                  As Long
End Type

Public Type BITMAPINFO
    bmiHeader                               As BITMAPINFOHEADER
    bmiColors                               As RGBTRIPLE
End Type

''Tooltip Window Types
Public Type TOOLINFO
    lSize                                   As Long
    lFlags                                  As Long
    lhWnd                                   As Long
    lID                                     As Long
    lpRect                                  As RECT
    hInstance                               As Long
    lpStr                                   As String
    lParam                                  As Long
End Type

'Tooltip Window Types [for UNICODE support]
Public Type TOOLINFOW
    lSize                                   As Long
    lFlags                                  As Long
    lhWnd                                   As Long
    lID                                     As Long
    lpRect                                  As RECT
    hInstance                               As Long
    lpStrW                                  As Long
    lParam                                  As Long
End Type

Public Type BITMAPINFO8
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT) As Long
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetCapture Lib "user32.dll" () As Long
Public Declare Function Rectangle _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal X1 As Long, _
                                ByVal Y1 As Long, _
                                ByVal X2 As Long, _
                                ByVal Y2 As Long) As Long

Public Declare Function DrawTextW _
               Lib "user32.dll" (ByVal hDC As Long, _
                                 ByVal lpStr As Long, _
                                 ByVal nCount As Long, _
                                 lpRect As RECT, _
                                 ByVal wFormat As Long) As Long

Public Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function GetTextExtentPoint32 _
               Lib "gdi32.dll" _
               Alias "GetTextExtentPoint32W" (ByVal hDC As Long, _
                                              ByVal lpsz As Long, _
                                              ByVal cbString As Long, _
                                              lpSize As Size) As Long

Public Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetTextColor Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetObject _
               Lib "gdi32.dll" _
               Alias "GetObjectA" (ByVal hObject As Long, _
                                   ByVal nCount As Long, _
                                   lpObject As Any) As Long

Public Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CopyRect Lib "user32.dll" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect _
               Lib "user32.dll" (lpRect As RECT, _
                                 ByVal X1 As Long, _
                                 ByVal Y1 As Long, _
                                 ByVal X2 As Long, _
                                 ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn _
               Lib "user32.dll" (ByVal hwnd As Long, _
                                 ByVal hRgn As Long, _
                                 ByVal bRedraw As Boolean) As Long

Public Declare Function LoadCursor _
               Lib "user32.dll" _
               Alias "LoadCursorA" (ByVal hInstance As Long, _
                                    ByVal lpCursorName As Long) As Long

Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function MoveToEx _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                lpPoint As POINT) As Long

Public Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreatePen _
               Lib "gdi32.dll" (ByVal nPenStyle As Long, _
                                ByVal nWidth As Long, _
                                ByVal crColor As Long) As Long

Public Declare Function RedrawWindow _
               Lib "user32" (ByVal hwnd As Long, _
                             lprcUpdate As RECT, _
                             ByVal hrgnUpdate As Long, _
                             ByVal fuRedraw As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function CreateWindowEx _
               Lib "user32" _
               Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                        ByVal lpClassName As String, _
                                        ByVal lpWindowName As String, _
                                        ByVal dwStyle As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal nWidth As Long, _
                                        ByVal nHeight As Long, _
                                        ByVal hWndParent As Long, _
                                        ByVal hMenu As Long, _
                                        ByVal hInstance As Long, _
                                        lpParam As Any) As Long

Public Declare Function SetWindowPos _
               Lib "user32" (ByVal hwnd As Long, _
                             ByVal hWndInsertAfter As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal cx As Long, _
                             ByVal cy As Long, _
                             ByVal wFlags As Long) As Long

Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function MoveWindow _
               Lib "user32" (ByVal hwnd As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal bRepaint As Long) As Long

Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINT) As Long
Public Declare Function DrawText _
               Lib "user32.dll" _
               Alias "DrawTextA" (ByVal hDC As Long, _
                                  ByVal lpStr As String, _
                                  ByVal nCount As Long, _
                                  lpRect As RECT, _
                                  ByVal wFormat As Long) As Long

Public Declare Function DrawIconEx _
               Lib "user32.dll" (ByVal hDC As Long, _
                                 ByVal xLeft As Long, _
                                 ByVal yTop As Long, _
                                 ByVal hIcon As Long, _
                                 ByVal cxWidth As Long, _
                                 ByVal cyWidth As Long, _
                                 ByVal istepIfAniCur As Long, _
                                 ByVal hbrFlickerFreeDraw As Long, _
                                 ByVal diFlags As Long) As Long

Public Declare Function BitBlt _
               Lib "gdi32.dll" (ByVal hDCDest As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal nWidth As Long, _
                                ByVal nHeight As Long, _
                                ByVal hdcSrc As Long, _
                                ByVal XSrc As Long, _
                                ByVal YSrc As Long, _
                                ByVal dwRop As Long) As Long

Public Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function SetPixel _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal crColor As Long) As Long

Public Declare Function CreateCompatibleBitmap _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal nWidth As Long, _
                                ByVal nHeight As Long) As Long

Public Declare Function CreateBitmap _
               Lib "gdi32.dll" (ByVal nWidth As Long, _
                                ByVal nHeight As Long, _
                                ByVal nPlanes As Long, _
                                ByVal nBitCount As Long, _
                                lpBits As Any) As Long

Public Declare Function DrawEdge _
               Lib "user32.dll" (ByVal hDC As Long, _
                                 qrc As RECT, _
                                 ByVal Edge As Long, _
                                 ByVal grfFlags As Long) As Long

Public Declare Function OleTranslateColor _
               Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                   ByVal HPALETTE As Long, _
                                   pccolorref As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Public Declare Function DrawThemeBackground _
               Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                  ByVal lhDC As Long, _
                                  ByVal iPartId As Long, _
                                  ByVal iStateId As Long, _
                                  pRect As RECT, _
                                  pClipRect As RECT) As Long

Public Declare Function GetThemeBackgroundRegion _
               Lib "uxtheme.dll" (ByVal hTheme As Long, _
                                  ByVal hDC As Long, _
                                  ByVal iPartId As Long, _
                                  ByVal iStateId As Long, _
                                  pRect As RECT, _
                                  pRegion As Long) As Long

Public Declare Function GetCurrentThemeName _
               Lib "uxtheme.dll" (ByVal pszThemeFileName As String, _
                                  ByVal dwMaxNameChars As Integer, _
                                  ByVal pszColorBuff As String, _
                                  ByVal cchMaxColorChars As Integer, _
                                  ByVal pszSizeBuff As String, _
                                  ByVal cchMaxSizeChars As Integer) As Long

Public Declare Function StretchBlt _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal nWidth As Long, _
                                ByVal nHeight As Long, _
                                ByVal hSrcDC As Long, _
                                ByVal XSrc As Long, _
                                ByVal YSrc As Long, _
                                ByVal nSrcWidth As Long, _
                                ByVal nSrcHeight As Long, _
                                ByVal dwRop As Long) As Long

Public Declare Function SetLayout Lib "gdi32.dll" (ByVal hDC As Long, ByVal dwLayout As Long) As Long
Public Declare Function TransparentBlt _
               Lib "msimg32.dll" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal nWidth As Long, _
                                  ByVal nHeight As Long, _
                                  ByVal hSrcDC As Long, _
                                  ByVal XSrc As Long, _
                                  ByVal YSrc As Long, _
                                  ByVal nSrcWidth As Long, _
                                  ByVal nSrcHeight As Long, _
                                  ByVal crTransparent As Long) As Boolean

Public Declare Function CreateDIBSection8 _
               Lib "gdi32.dll" _
               Alias "CreateDIBSection" (ByVal hDC As Long, _
                                         pBitmapInfo As BITMAPINFO8, _
                                         ByVal un As Long, _
                                         ByVal lplpVoid As Long, _
                                         ByVal Handle As Long, _
                                         ByVal dw As Long) As Long

Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function OleTranslateColorByRef _
               Lib "oleaut32.dll" _
               Alias "OleTranslateColor" (ByVal lOleColor As Long, _
                                          ByVal lHPalette As Long, _
                                          ByVal lColorRef As Long) As Long

Public Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetPixelV _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal crColor As Long) As Long

Public Declare Function CreateRectRgn _
               Lib "gdi32.dll" (ByVal X1 As Long, _
                                ByVal Y1 As Long, _
                                ByVal X2 As Long, _
                                ByVal Y2 As Long) As Long

Public Declare Function CombineRgn _
               Lib "gdi32.dll" (ByVal hDestRgn As Long, _
                                ByVal hSrcRgn1 As Long, _
                                ByVal hSrcRgn2 As Long, _
                                ByVal nCombineMode As Long) As Long

Public Declare Function RoundRect _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal Left As Long, _
                                ByVal Top As Long, _
                                ByVal Right As Long, _
                                ByVal Bottom As Long, _
                                ByVal EllipseWidth As Long, _
                                ByVal EllipseHeight As Long) As Long

Public Declare Function CreatePolygonRgn _
               Lib "gdi32.dll" (lpPoint As Any, _
                                ByVal nCount As Long, _
                                ByVal nPolyFillMode As Long) As Long

Public Declare Function CreateRoundRectRgn _
               Lib "gdi32.dll" (ByVal X1 As Long, _
                                ByVal Y1 As Long, _
                                ByVal X2 As Long, _
                                ByVal Y2 As Long, _
                                ByVal X3 As Long, _
                                ByVal Y3 As Long) As Long

Public Declare Function GetDIBits _
               Lib "gdi32.dll" (ByVal aHDC As Long, _
                                ByVal hBitmap As Long, _
                                ByVal nStartScan As Long, _
                                ByVal nNumScans As Long, _
                                lpBits As Any, _
                                lpBI As BITMAPINFO, _
                                ByVal wUsage As Long) As Long

Public Declare Function SetDIBitsToDevice _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal dx As Long, _
                                ByVal dy As Long, _
                                ByVal SrcX As Long, _
                                ByVal SrcY As Long, _
                                ByVal Scan As Long, _
                                ByVal NumScans As Long, _
                                Bits As Any, _
                                BitsInfo As BITMAPINFO, _
                                ByVal wUsage As Long) As Long

Public Declare Function StretchDIBits _
               Lib "gdi32.dll" (ByVal hDC As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal dx As Long, _
                                ByVal dy As Long, _
                                ByVal SrcX As Long, _
                                ByVal SrcY As Long, _
                                ByVal wSrcWidth As Long, _
                                ByVal wSrcHeight As Long, _
                                lpBits As Any, _
                                lpBitsInfo As Any, _
                                ByVal wUsage As Long, _
                                ByVal dwRop As Long) As Long

Public Declare Function GetNearestColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
