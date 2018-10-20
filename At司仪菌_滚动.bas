Attribute VB_Name = "At司仪菌_滚动"
'本模块代码来自《顾名思义》
'█████████████████████████████████████████████████████████████
'④加入一个模块并录入代码
Option Explicit
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_KEYUP = &H101
' -- 引用Win32Api –
'得到默认的窗口消息处理过程的地址需要的API
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'设置一个新的窗口消息处理过程的地址需要的API
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'给指定的窗口消息处理过程传递消息需要的API
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const GWL_WNDPROC = (-4&)
Dim PrevWndProc&
Private Const WM_DESTROY = &H2
Private Const WM_DRAWITEM = &H2B

'新的窗口消息处理过程，将被插入到默认处理过程之前
Private Function SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If Msg = WM_DESTROY Then Terminate (hWnd)


If Msg = WM_MOUSEWHEEL Then
If (wParam And &HFF000000) = &H0& Then
SubWndProc = CallWindowProc(PrevWndProc, hWnd, WM_KEYUP, &HFF00, 0&)
Else
SubWndProc = CallWindowProc(PrevWndProc, hWnd, WM_KEYUP, &HFF01, 0&)
End If
Exit Function
End If

'调用默认的窗口处理过程

SubWndProc = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
End Function
'子类化入口
Public Sub Init(hWnd As Long)
PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub
'子类化出口
Public Sub Terminate(hWnd As Long)
Call SetWindowLong(hWnd, GWL_WNDPROC, PrevWndProc)
End Sub

