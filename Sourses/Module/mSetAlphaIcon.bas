Attribute VB_Name = "mSetAlphaIcon"
Option Explicit

Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Declare Function LoadImageAsString _
                Lib "user32" _
                Alias "LoadImageA" (ByVal hInst As Long, _
                                    ByVal lpsz As String, _
                                    ByVal uType As Long, _
                                    ByVal cxDesired As Long, _
                                    ByVal cyDesired As Long, _
                                    ByVal fuLoad As Long) As Long

Public Sub SetIcon(ByVal hwnd As Long, _
                   ByVal sIconResName As String, _
                   Optional ByVal bSetAsAppIcon As Boolean = True)

    Dim lhWndTop   As Long
    Dim lhWnd      As Long
    Dim cx         As Long
    Dim cy         As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long

    If (bSetAsAppIcon) Then
        ' Find VB's hidden parent window:
        lhWnd = hwnd
        lhWndTop = lhWnd

        Do While Not (lhWnd = 0)
            lhWnd = GetWindow(lhWnd, GW_OWNER)

            If Not (lhWnd = 0) Then
                lhWndTop = lhWnd
            End If
        Loop
    End If

    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)

    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    End If

    SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)

    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    End If

    SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall

End Sub
