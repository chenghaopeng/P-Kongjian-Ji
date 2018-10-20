VERSION 5.00
Begin VB.UserControl PSwitch 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   255
   ScaleWidth      =   3495
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
   Begin P¿Ø¼þ¼¯.PUIMgr PM 
      Left            =   1200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "PSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim C_Color_Top As OLE_COLOR
Dim C_Color_Back As OLE_COLOR
Dim C_Color_Back_True As OLE_COLOR
Dim C_Is_Enabled As Boolean
Dim C_Value As Boolean
Dim C_Style_Border As Border
Dim C_Color_Border As OLE_COLOR

Public Event ValueChange(NValue As Boolean)
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Boolean)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Boolean)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Boolean)

Dim ClickedX As Single
Dim MouseDowned As Boolean
Dim ChangeValue As Boolean

Private Sub Refresh()
    If Progress.Left >= (UserControl.Width - Progress.Width) / 2 Then
        C_Value = True
        PM.MoveSmly Progress, UserControl.Width - Progress.Width, 0, 10
    Else
        C_Value = False
        PM.MoveSmly Progress, 0, 0, 10
    End If
    If C_Value = True Then
        UserControl.BackColor = C_Color_Back_True
    Else
        UserControl.BackColor = C_Color_Back
    End If
End Sub

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
    Refresh
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_Back_True() As OLE_COLOR
    Color_Back_True = C_Color_Back_True
End Property

Public Property Let Color_Back_True(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_True = vNewValue
    Refresh
    PropertyChanged "Color_Back_True"
End Property

Public Property Get Is_Enabled() As Boolean
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    Progress.Visible = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Value() As Boolean
    Value = C_Value
End Property

Public Property Let Value(ByVal vNewValue As Boolean)
    C_Value = vNewValue
    Refresh
    PropertyChanged "Value"
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
            PM.MoveSmly Progress, 0, 0, 1
        ElseIf (Progress.Left - ClickedX + X) > (UserControl.Width - Progress.Width) Then
            PM.MoveSmly Progress, UserControl.Width - Progress.Width, 0, 1
        Else
            PM.MoveSmly Progress, Progress.Left - ClickedX + X, 0, 1
        End If
    End If
    If Progress.Left >= (UserControl.Width - Progress.Width) / 2 Then
        UserControl.BackColor = C_Color_Back_True
    Else
        UserControl.BackColor = C_Color_Back
    End If
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y, C_Value)
    If Is_Enabled = True Then
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
    Refresh
    If ClickedX <> 0 Then
        RaiseEvent ValueChange(C_Value)
    End If
    ClickedX = 0
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y, C_Value)
    MouseDowned = False
End Sub

Private Sub UserControl_Initialize()
    C_Color_Top = &HFF7402
    C_Color_Back = &HF2AF00
    C_Color_Back_True = &HF2AF00
    C_Is_Enabled = True
    C_Value = False
    C_Style_Border = 1
    C_Color_Border = &H0&
    Init UserControl.hWnd
    MLInit Progress.hWnd
    Refresh
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = -256 Then
        If Value = False Then Exit Sub
        Progress.Left = (UserControl.Width - Progress.Width) / 2 - 15
        'PM.MoveSmly Progress, 0, 0, 10
    ElseIf KeyCode = -255 Then
        If Value = True Then Exit Sub
        Progress.Left = (UserControl.Width - Progress.Width) / 2 + 15
        'PM.MoveSmly Progress, UserControl.Width - Progress.Width, 0, 10
    End If
    Refresh
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
    C_Color_Back_True = PropBag.ReadProperty("Color_Back_True", &HF2AF00)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Value = PropBag.ReadProperty("Value", False)
    C_Style_Border = PropBag.ReadProperty("Style_Border", 1)
    C_Color_Border = PropBag.ReadProperty("Color_Border", &H0&)
    Progress.BackColor = C_Color_Top
    UserControl.BackColor = Color_Back
    Progress.Visible = C_Is_Enabled
    Refresh
End Sub

Private Sub UserControl_Resize()
    Progress.Width = UserControl.Width / 2
    Progress.Height = UserControl.Height
    Shape1.Width = Progress.Width
    Shape1.Height = Progress.Height
    Refresh
End Sub

Private Sub UserControl_Terminate()
    Terminate UserControl.hWnd
    MLTerminate Progress.hWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Top", C_Color_Top, &HFF7402)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)
    Call PropBag.WriteProperty("Color_Back_True", C_Color_Back_True, &HF2AF00)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Value", C_Value, False)
    Call PropBag.WriteProperty("Style_Border", C_Style_Border, 1)
    Call PropBag.WriteProperty("Color_Border", C_Color_Border, &H0&)
End Sub


