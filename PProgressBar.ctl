VERSION 5.00
Begin VB.UserControl PProgressBar 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ScaleHeight     =   255
   ScaleWidth      =   3855
   Begin VB.Shape Progress 
      BackColor       =   &H00FF7402&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF7402&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "PProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim C_Color_Top As OLE_COLOR
Dim C_Color_Back As OLE_COLOR
Dim C_Is_Enabled As Boolean
Dim C_Value As Single

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)

Public Property Get Color_Top() As OLE_COLOR
    Color_Top = C_Color_Top
End Property

Public Property Let Color_Top(ByVal vNewValue As OLE_COLOR)
    C_Color_Top = vNewValue
    Progress.BackColor = C_Color_Top
    Progress.BorderColor = C_Color_Top
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
    Progress.Width = UserControl.Width * C_Value
    PropertyChanged "Value"
End Property

Private Sub UserControl_Click()
    If Is_Enabled = True Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If Is_Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    C_Color_Top = &HFF7402
    C_Color_Back = &HF2AF00
    C_Is_Enabled = True
    C_Value = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, x, y, x / UserControl.Width)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, x, y, x / UserControl.Width)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, x, y, x / UserControl.Width)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Top = PropBag.ReadProperty("Color_Top", &HFF7402)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Value = PropBag.ReadProperty("Value", 1)
    Progress.BackColor = C_Color_Top
    Progress.BorderColor = C_Color_Top
    UserControl.BackColor = Color_Back
    Progress.Width = UserControl.Width * C_Value
End Sub

Private Sub UserControl_Resize()
    Progress.Height = UserControl.Height
    Progress.Width = UserControl.Width * C_Value
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Top", C_Color_Top, &HFF7402)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Value", C_Value, 1)
End Sub
