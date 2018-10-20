Attribute VB_Name = "MouseLeave"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Public Const GWL_WNDPROC = (-4&)
Public Const WM_DESTROY = &H2
Public Const TME_LEAVE = &H2&
Public Const WM_MOUSELEAVE = &H2A3&
Public Const WM_KEYUP = &H101
Dim PrevWndProc&
Public Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Function SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_DESTROY Then MLTerminate (hWnd)
    If Msg = WM_MOUSELEAVE Then
        SubWndProc = CallWindowProc(PrevWndProc, hWnd, WM_KEYUP, -108, 0&)
    End If
    SubWndProc = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
End Function
Public Sub MLInit(hWnd As Long)
    PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub
Public Sub MLTerminate(hWnd As Long)
    Call SetWindowLong(hWnd, GWL_WNDPROC, PrevWndProc)
End Sub
Public Sub Reload(hWnd As Long)
    Dim ET As TRACKMOUSEEVENTTYPE
    ET.cbSize = Len(ET)
    ET.hwndTrack = hWnd
    ET.dwFlags = TME_LEAVE
    TrackMouseEvent ET
End Sub
