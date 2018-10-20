VERSION 5.00
Begin VB.UserControl PVScrollBar 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   255
   ScaleHeight     =   3495
   ScaleWidth      =   255
   Begin P控件集.PUIMgr PM 
      Left            =   1440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Progress 
      BackColor       =   &H00FF7402&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   255
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "PVScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'██████↓定义存储属性的变量↓███████████████████████████████████████████████████████
Dim C_Color_Top As OLE_COLOR '滚动块颜色
Dim C_Color_Back As OLE_COLOR '滚动块背景颜色
Dim C_Is_Enabled As Boolean '是否有效
Dim C_Value As Single '值
Dim C_Size As Single '滚动块占总大小的多少
Dim C_Value_Wheel_Change As Single '鼠标滚轮滚动时改变的值
Dim C_Style_Border As Border '边框形式
Dim C_Color_Border As OLE_COLOR '边框颜色
'██████↓定义事件↓███████████████████████████████████████████████████████
Public Event Scroll(NValue As Single) '滚动事件
Public Event Change(NValue As Single) '值改变事件
Public Event Click() '单击事件
Public Event DblClick() '双击事件
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single) '鼠标按下事件,NValue为新值
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single) '鼠标触碰事件,NValue为新值
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single) '鼠标弹起事件,NValue为新值
'██████↓定义使用中所需的变量↓███████████████████████████████████████████████████████
Dim ClickedY As Single '鼠标按下的位置
Dim GoalY As Single '滚动块要到达的位置
Dim MouseDowned As Boolean '鼠标是否按下
'██████↓各种属性↓███████████████████████████████████████████████████████
Public Property Get Color_Top() As OLE_COLOR '滚动块颜色
    Color_Top = C_Color_Top
End Property

Public Property Let Color_Top(ByVal vNewValue As OLE_COLOR)
    C_Color_Top = vNewValue
    Progress.BackColor = C_Color_Top
    PropertyChanged "Color_Top"
End Property

Public Property Get Color_Back() As OLE_COLOR '滚动块背景颜色
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    UserControl.BackColor = vNewValue
    PropertyChanged "Color_Back"
End Property

Public Property Get Is_Enabled() As Boolean '是否有效
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    Progress.Visible = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Value() As Single '值
    Value = C_Value
End Property

Public Property Let Value(ByVal vNewValue As Single)
    If vNewValue > 1 Then '如果值大于1
        C_Value = 1
    ElseIf vNewValue < 0 Then '如果值小于0
        C_Value = 0
    Else
        C_Value = vNewValue '更新值
    End If
    PM.MoveSmly Progress, 0, (UserControl.Height - Progress.Height) * C_Value, 10 '平滑得移动滚动块到置顶位置
    RaiseEvent Change(C_Value) '触发值改变事件
    PropertyChanged "Value"
End Property

Public Property Get Size() As Single '滚动块占总大小的多少
    Size = C_Size
End Property

Public Property Let Size(ByVal vNewValue As Single)
    If vNewValue > 0.8 Then '如果值大于0.8
        C_Size = 0.8
    ElseIf (vNewValue * UserControl.Height) < 45 Then '如果滚动块小于3像素
        C_Size = 45 / UserControl.Height
    Else
        C_Size = vNewValue '更新值
    End If
    Progress.Height = UserControl.Height * C_Size '更新滚动块大小
    PropertyChanged "Size"
End Property

Public Property Get Value_Wheel_Change() As Single '鼠标滚轮滚动时改变的值
    Value_Wheel_Change = C_Value_Wheel_Change
End Property

Public Property Let Value_Wheel_Change(ByVal vNewValue As Single)
    If vNewValue > 0.5 Then '如果值大于0.5
        C_Value_Wheel_Change = 0.5
    ElseIf vNewValue < 0.01 Then '如果值小于0.01
        C_Value_Wheel_Change = 0.01
    Else
        C_Value_Wheel_Change = vNewValue '更新值
    End If
    PropertyChanged "Value_Wheel_Change"
End Property

Public Property Get Style_Border() As Border '边框形式
    Style_Border = C_Style_Border
End Property

Public Property Let Style_Border(ByVal vNewValue As Border)
    C_Style_Border = vNewValue
    PropertyChanged "Style_Border"
End Property

Public Property Get Color_Border() As OLE_COLOR '边框形式
    Color_Border = C_Color_Border
End Property

Public Property Let Color_Border(ByVal vNewValue As OLE_COLOR)
    C_Color_Border = vNewValue
    PropertyChanged "Color_Border"
End Property
'██████↓各种事件↓███████████████████████████████████████████████████████
Private Sub Progress_Click() '滚动块的单击事件
    If Is_Enabled = True Then RaiseEvent Click '触发单击事件
End Sub

Private Sub Progress_DblClick() '滚动块的双击事件
    If Is_Enabled = True Then RaiseEvent DblClick '触发双击事件
End Sub

Private Sub Progress_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = -108 Then Shape1.Visible = False
End Sub

Private Sub Progress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '滚动块的鼠标按下事件
    If Button = 1 Then ClickedY = Y '如果鼠标左键按下,记下鼠标按下的位置
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y, C_Value) '如果有效,触发鼠标按下事件
    If Button = 1 Then MouseDowned = True '如果鼠标左键按下,鼠标按下
End Sub

Private Sub Progress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Reload Progress.hWnd
    If MouseDowned Then
        If (Progress.Top - ClickedY + Y) < 0 Then
'            PM.MoveSmly Progress, 0, 0, 1
            GoalY = 0
        ElseIf (Progress.Top - ClickedY + Y) > (UserControl.Height - Progress.Height) Then
'            PM.MoveSmly Progress, 0, UserControl.Height - Progress.Height, 1
            GoalY = UserControl.Height - Progress.Height
        Else
'            PM.MoveSmly Progress, 0, Progress.Top - ClickedY + Y, 1
            GoalY = Progress.Top - ClickedY + Y
        End If
        Progress.Top = GoalY
    End If
    If MouseDowned Then
        RaiseEvent Scroll(GoalY / (UserControl.Height - Progress.Height))
        RaiseEvent Change(GoalY / (UserControl.Height - Progress.Height))
    End If
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y, C_Value)
    If Is_Enabled = True Then
        Shape1.Width = Progress.Width
        Shape1.Height = Progress.Height
        Select Case C_Style_Border
        Case 0
            '
        Case 1
            Shape1.BorderColor = RGB(Abs(255 - C_Color_Top Mod 256), Abs(255 - (C_Color_Top Mod 65536) \ 256), Abs(255 - C_Color_Top \ 65536))
            Shape1.Visible = True
        Case 2
            Shape1.BorderColor = C_Color_Border
            Shape1.Visible = True
        End Select
    End If
End Sub

Private Sub Progress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickedY = 0
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y, C_Value)
    MouseDowned = False
    C_Value = GoalY / (UserControl.Height - Progress.Height)
    Value = C_Value
End Sub

Private Sub UserControl_Initialize()
    C_Color_Top = &HFF7402
    C_Color_Back = &HF2AF00
    C_Is_Enabled = True
    C_Value = 0
    C_Size = 0.2
    C_Value_Wheel_Change = 0.05
    C_Style_Border = 0
    C_Color_Border = &H0&
    Init UserControl.hWnd
    MLInit Progress.hWnd
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = -256 Then
        If Me.Value > Value_Wheel_Change Then
            Me.Value = Me.Value - Value_Wheel_Change
        Else
            Me.Value = 0
        End If
    ElseIf KeyCode = -255 Then
        If Me.Value < (1 - Value_Wheel_Change) Then
            Me.Value = Me.Value + Value_Wheel_Change
        Else
            Me.Value = 1
        End If
    End If
    If Is_Enabled Then RaiseEvent Scroll(C_Value) '触发滚动事件
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y, C_Value)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y, C_Value)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y, C_Value)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Top = PropBag.ReadProperty("Color_Top", &HFF7402)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Value = PropBag.ReadProperty("Value", 0)
    C_Size = PropBag.ReadProperty("Size", 0.2)
    C_Value_Wheel_Change = PropBag.ReadProperty("Value_Wheel_Change", 0.05)
    C_Style_Border = PropBag.ReadProperty("Style_Border", 0)
    C_Color_Border = PropBag.ReadProperty("Color_Border", &H0&)
    Progress.BackColor = C_Color_Top
    UserControl.BackColor = Color_Back
    Progress.Visible = C_Is_Enabled
    Progress.Top = (UserControl.Height - Progress.Height) * C_Value
    Progress.Height = UserControl.Height * C_Size
End Sub

Private Sub UserControl_Resize()
    Progress.Top = (UserControl.Height - Progress.Height) * C_Value
    Progress.Height = UserControl.Height * C_Size
    Progress.Width = UserControl.Width
    Shape1.Width = Progress.Width
    Shape1.Height = Progress.Height
End Sub

Private Sub UserControl_Terminate()
    Terminate UserControl.hWnd
    MLTerminate Progress.hWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Top", C_Color_Top, &HFF7402)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Value", C_Value, 0)
    Call PropBag.WriteProperty("Size", C_Size, 0.2)
    Call PropBag.WriteProperty("Value_Wheel_Change", C_Value_Wheel_Change, 0.05)
    Call PropBag.WriteProperty("Style_Border", C_Style_Border, 0)
    Call PropBag.WriteProperty("Color_Border", C_Color_Border, &H0&)
End Sub



