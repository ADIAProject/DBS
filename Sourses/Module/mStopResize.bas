Attribute VB_Name = "mStopResize"
Option Explicit

'Эжели хоца, сделайте из этого модуля класс
'я попробовал, не получилось, а возиться не хочеться
Private Const WM_GETMINMAXINFO As Long = &H24

Private Type MINMAXINFO
    ptReserved                        As POINT
    ptMaxSize                         As POINT
    ptMaxPosition                     As POINT
    ptMinTrackSize                    As POINT
    ptMaxTrackSize                    As POINT
End Type

Private lpPrevWndProc As Long
Private gHW           As Long

Private Type Resize
    xMin                              As Single
    yMin                              As Single
    xMax                              As Single
    yMax                              As Single
End Type

Private rResize As Resize

Private Declare Function DefWindowProc _
                Lib "user32.dll" _
                Alias "DefWindowProcA" (ByVal hwnd As Long, _
                                        ByVal wMsg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long

Private Declare Sub CopyMemoryToMinMaxInfo _
                Lib "kernel32.dll" _
                Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, _
                                       ByVal hpvSource As Long, _
                                       ByVal cbCopy As Long)

Private Declare Sub CopyMemoryFromMinMaxInfo _
                Lib "kernel32.dll" _
                Alias "RtlMoveMemory" (ByVal hpvDest As Long, _
                                       hpvSource As MINMAXINFO, _
                                       ByVal cbCopy As Long)

Public Sub Hook(ByVal wHWND As Long, _
                Optional ByVal X_Min As Single = 0, _
                Optional ByVal Y_Min As Single = 0, _
                Optional ByVal X_Max As Single = 0, _
                Optional ByVal Y_Max As Single = 0)

    'Стартуем сабклассинг
    gHW = wHWND

    'запомним хэндл, чтобы воспользоваться им при остановке классинга
    With rResize
        .xMax = X_Max
        .yMax = Y_Max
        .xMin = X_Min
        .yMin = Y_Min
    End With

    'RRESIZE
    'rResize
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()

    'Останавливаем сабклассинг
    SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
End Sub

Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim MinMax As MINMAXINFO

    'Проверка, ресайзим ли мы окно
    If uMsg = WM_GETMINMAXINFO Then
        'Необходимо для заголовка child MDI окна (когда развернуто на весь экран)
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
        'получаем заданные по умолчанию параметры настройки Минимакса
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Определяем новый минимальный размер окна
        'Если не присвоить какое-либо значение в MinMax.ptMinTrackSize.x(y), то
        'При ресайзе это значение будет игнорироваться. Тоже самое и для максимальног значения
        If rResize.xMin <> 0 Then
            MinMax.ptMinTrackSize.X = rResize.xMin
        End If

        If rResize.yMin <> 0 Then
            MinMax.ptMinTrackSize.Y = rResize.yMin
        End If

        'Определяем новый максимальный размер окна
        If rResize.xMax <> 0 Then
            MinMax.ptMaxTrackSize.X = rResize.xMax
        End If

        If rResize.yMax <> 0 Then
            MinMax.ptMaxTrackSize.Y = rResize.yMax
        End If

        'Копируем нашу структуру обратно
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    End If
End Function
