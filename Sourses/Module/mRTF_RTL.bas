Attribute VB_Name = "mRTF_RTL"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Azkary
'Program Author   : Elsheshtawy, Ahmed Amin
'Home Page        : http://www.islamware.com
'Copyrights © 2007 Islamware. All rights reserved.
'==========================================================
'Permission to use, copy, modify, and distribute this software and its
'documentation for any purpose and without fee is hereby granted.
'==========================================================
Option Explicit

Public Const EM_SETPARAFORMAT = WM_USER + 71
Public Const EM_SETBIDIOPTIONS = WM_USER + 200
Public Const PFM_DIR = &H10000 'Direction mask bit
Public Const PFE_RTLPAR = &H1 'RTL paragraph style bit
Public Const BOM_DEFPARADIR = &H1 'Default direction mask
Public Const BOE_RTLDIR = &H1 'Default RTL para style

Type vbParaFormat
    cbSize As Long
    dwMask As Long
    wNumbering As Integer
    wEffects As Integer
    dxStrtIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    rgxTabs(31) As Long
End Type

Public vbPF As vbParaFormat
Public Declare Function SendPFMessage _
               Lib "user32.dll" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As vbParaFormat) As Long

Type vbBiDiOptions
    cbSize As Long
    wMask As Integer
    wEffects As Integer
End Type

Public vbBO As vbBiDiOptions
Public Declare Function SendBOMessage _
               Lib "user32.dll" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As vbBiDiOptions) As Long

Public Function SetParaDirection(hWndRTF As Long, Direction As Integer) As Boolean

    vbPF.cbSize = LenB(vbPF)
    'Size
    vbPF.dwMask = PFM_DIR
    'Attribute to set
    vbPF.wEffects = Direction
    'New direction
    SetParaDirection = SendPFMessage(hWndRTF, EM_SETPARAFORMAT, 0, vbPF)
End Function

'This code requires a rich text box control. Buttons named btnLTR and btnRTL are supported but not necassarily needed
'Option Explicit
'Private Sub btnLTR_Click()
'    SetParaDirection (Not PFE_RTLPAR)
'    RichTextBox1.SetFocus
'End Sub
'Private Sub btnRTL_Click()
'    SetParaDirection (PFE_RTLPAR)
'    RichTextBox1.SetFocus
'End Sub
'
'Private Sub Form_Load()
'    Dim Result As Long
'    SetParaDirection (PFE_RTLPAR)
'    vbBO.cbSize = LenB(vbBO) 'Size
'    vbBO.wMask = BOM_DEFPARADIR 'Attribute to set
'    vbBO.wEffects = BOE_RTLDIR 'Default direction
'    Result = SendBOMessage(RichTextBox1.hWnd, EM_SETBIDIOPTIONS, 0, vbBO)
'    RichTextBox1.RightMargin = 2000
'End Sub
'\\Code//
