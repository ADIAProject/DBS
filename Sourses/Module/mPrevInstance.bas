Attribute VB_Name = "mPrevInstance"
'��������� ����� ���������� ���� �� ���������
Option Explicit

Private Declare Function OpenIcon Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

'! -----------------------------------------------------------
'!  �������     :  ShowPrevInstance
'!  ����������  :
'!  ��������    :  ���������� ���������� ����� ���������, ���� ��������� �������� ������
'! -----------------------------------------------------------
Public Sub ShowPrevInstance()

    Dim OldTitle        As String
    Dim ll_WindowHandle As Long

    OldTitle = App.Title
    App.Title = "This App Will Be Closed"
    ll_WindowHandle = FindWindow("ThunderRT6Main", OldTitle)

    If ll_WindowHandle <> 0 Then
        ll_WindowHandle = GetWindow(ll_WindowHandle, GW_HWNDPREV)
        OpenIcon ll_WindowHandle
        SetForegroundWindow ll_WindowHandle
        End
    End If
End Sub
