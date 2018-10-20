VERSION 5.00
Begin VB.UserControl PHScrollBar 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   255
   ScaleWidth      =   3495
   Begin P控件集.PUIMgr PM 
      Left            =   1920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Progress 
      BackColor       =   &H00FF7402&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   0
      Width           =   855
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
End
Attribute VB_Name = "PHScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim C_Color_Top As OLE_COLOR
Dim C_Color_Back As OLE_COLOR
Dim C_Is_Enabled As Boolean
Dim C_Value As Single
Dim C_Size As Single
Dim C_Value_Wheel_Change As Single
Dim C_Style_Border As Border
Dim C_Color_Border As OLE_COLOR

Public Event Scroll(NValue As Single)
Public Event Change(NValue As Single)
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)

Dim ClickedX As Single
Dim GoalX As Single
Dim MouseDowned As Boolean

Public Property Get Color_Top() As OLE_COLOR
    Color_Top = C_Color_Top
End Property

Public Property Let Color_Top(ByVal vNewValue As OLE_COLOR)
    C_Color_Top = vNewValue
    Progress.BackColor = C_Color_Top
    PropertyChanged "Color_Top"
End Property

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    UserControl.BackColor = vNewValue
    PropertyChanged "Color_Back"
End Property

Public Property Get Is_Enabled() As Boolean
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    Progress.Visible = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Value() As Single
    Value = C_Value
End Property

Public Property Let Value(ByVal vNewValue As Single)
    If vNewValue > 1 Then
        C_Value = 1
    ElseIf vNewValue < 0 Then
        C_Value = 0
    Else
        C_Value = vNewValue
    End If
    PM.MoveSmly Progress, (UserControl.Width - Progress.Width) * C_Value, 0, 10
'    Progress.Left = (UserControl.Width - Progress.Width) * C_Value
    RaiseEvent Change(C_Value)
    PropertyChanged "Value"
End Property

Public Property Get Size() As Single
    Size = C_Size
End Property

Public Property Let Size(ByVal vNewValue As Single)
    If vNewValue > 0.8 Then
        C_Size = 0.8
    ElseIf (vNewValue * UserControl.Width) < 45 Then
        C_Size = 45 / UserControl.Width
    Else
        C_Size = vNewValue
    End If
    Progress.Width = UserControl.Width * C_Size
    PropertyChanged "Size"
End Property

Public Property Get Value_Wheel_Change() As Single
    Value_Wheel_Change = C_Value_Wheel_Change
End Property

Public Property Let Value_Wheel_Change(ByVal vNewValue As Single)
    If vNewValue > 0.5 Then
        C_Value_Wheel_Change = 0.5
    ElseIf vNewValue < 0.01 Then
        C_Value_Wheel_Change = 0.01
    Else
        C_Value_Wheel_Change = vNewValue
    End If
    PropertyChanged "Value_Wheel_Change"
End Property

Public Property Get Style_Border() As Border
    Style_Border = C_Style_Border
End Property

Public Property Let Style_Border(ByVal vNewValue As Border)
    C_Style_Border = vNewValue
    PropertyChanged "Style_Border"
End Property

Public Property Get Color_Border() As OLE_COLOR
    Color_Border = C_Color_Border
End Property

Public Property Let Color_Border(ByVal vNewValue As OLE_COLOR)
    C_Color_Border = vNewValue
    PropertyChanged "Color_Border"
End Property

Private Sub Progress_Click()
    If Is_Enabled = True Then RaiseEvent Click
End Sub

Private Sub Progress_DblClick()
    If Is_Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub Progress_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = -108 Then Shape1.Visible = False
End Sub

Private Sub Progress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then ClickedX = X
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y, C_Value)
    If Button = 1 Then MouseDowned = True
End Sub

Private Sub Progress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Reload Progress.hWnd
    If MouseDowned Then
        If (Progress.Left - ClickedX + X) < 0 Then
'            PM.MoveSmly Progress, 0, 0, 1
            GoalX = 0
        ElseIf (Progress.Left - ClickedX + X) > (UserControl.Width - Progress.Width) Then
'            PM.MoveSmly Progress, UserControl.Width - Progress.Width, 0, 1
            GoalX = UserControl.Width - Progress.Width
        Else
'            PM.MoveSmly Progress, Progress.Left - ClickedX + X, 0, 1
            GoalX = Progress.Left - ClickedX + X
        End If
        Progress.Left = GoalX
    End If
    If MouseDowned Then
        RaiseEvent Scroll(GoalX / (UserControl.Width - Progress.Width))
        RaiseEvent Change(GoalX / (UserControl.Width - Progress.Width))
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
    ClickedX = 0
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y, C_Value)
    MouseDowned = False
    C_Value = GoalX / (UserControl.Width - Progress.Width)
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
    Progress.Left = (UserControl.Width - Progress.Width) * C_Value
    Progress.Width = UserControl.Width * C_Size
End Sub

Private Sub UserControl_Resize()
    Progress.Left = (UserControl.Width - Progress.Width) * C_Value
    Progress.Width = UserControl.Width * C_Size
    Progress.Height = UserControl.Height
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

