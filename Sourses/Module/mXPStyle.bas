Attribute VB_Name = "mXPStyle"
Option Explicit

' Модуль для инициализации стиля XP+ в программах, требуется файл манифеста в ресурсах программы
Private Type tagInitCommonControlsEx
    lngSize                             As Long            '- Размер структуры
    lngICC                              As Long            '- Какие классы загружать
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Long

Private Const ICC_ANIMATE_CLASS      As Long = &H80     ' Load animate control class
Private Const ICC_BAR_CLASSES        As Long = &H4      ' Load toolbar, status bar, trackbar, tooltip control classes
Private Const ICC_COOL_CLASSES       As Long = &H400    ' Load rebar control class
Private Const ICC_DATE_CLASSES       As Long = &H100    ' Load date and time picker control class
Private Const ICC_HOTKEY_CLASS       As Long = &H40     ' Load hot key control class
Private Const ICC_INTERNET_CLASSES   As Long = &H800    ' Load IP address class
Private Const ICC_LINK_CLASS         As Long = &H8000&  ' Load a hyperlink control class. Must have trailing ampersand.
Private Const ICC_LISTVIEW_CLASSES   As Long = &H1      ' Load list-view and header control classes
Private Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000   ' Load a native font control class
Private Const ICC_PAGESCROLLER_CLASS As Long = &H1000   ' Load pager control class
Private Const ICC_PROGRESS_CLASS     As Long = &H20     ' Load progress bar control class
Private Const ICC_STANDARD_CLASSES   As Long = &H4000   ' Load user controls that include button, edit, static, listbox,
Private Const ICC_TREEVIEW_CLASSES   As Long = &H2      ' Load tree-view and tooltip control classes
Private Const ICC_TAB_CLASSES        As Long = &H8      ' Load tab and tooltip control classes
Private Const ICC_UPDOWN_CLASS       As Long = &H10     ' Load up-down control class
Private Const ICC_USEREX_CLASSES     As Long = &H200    ' Load ComboBoxEx class
Private Const ICC_WIN95_CLASSES      As Long = &HFF     ' Load animate control, header, hot key, list-view, progress bar,
Private Const ALL_FLAGS              As Long = ICC_ANIMATE_CLASS Or ICC_BAR_CLASSES Or ICC_COOL_CLASSES Or ICC_DATE_CLASSES Or ICC_HOTKEY_CLASS Or ICC_INTERNET_CLASSES Or ICC_LINK_CLASS Or ICC_LISTVIEW_CLASSES Or ICC_NATIVEFNTCTL_CLASS Or ICC_PAGESCROLLER_CLASS Or ICC_PROGRESS_CLASS Or ICC_STANDARD_CLASSES Or ICC_TREEVIEW_CLASSES Or ICC_TAB_CLASSES Or ICC_UPDOWN_CLASS Or ICC_USEREX_CLASSES Or ICC_WIN95_CLASSES

Public m_hMod                        As Long
                                    
'! -----------------------------------------------------------
'!  Функция     :  InitXPStyle
'!  Переменные  :
'!  Описание    :
'! -----------------------------------------------------------
Public Function InitXPStyle() As Boolean

    Dim iccex As tagInitCommonControlsEx

    On Error GoTo Use_Old_Version

    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ALL_FLAGS
    End With

    m_hMod = LoadLibrary("shell32.dll")

    ' VB will generate error 453 "Specified DLL function not found"
    ' if InitCommonControlsEx can't be located in the library. The
    ' error is trapped and the original InitCommonControls is called
    ' instead below.
    If InitCommonControlsEx(iccex) = 0 Then
        InitCommonControls
        InitXPStyle = False
    Else
        InitXPStyle = True
    End If

    On Error GoTo 0

    Exit Function
Use_Old_Version:
    InitCommonControls

    On Error GoTo 0

End Function
