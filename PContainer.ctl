VERSION 5.00
Begin VB.UserControl PContainer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F2AF00&
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   ControlContainer=   -1  'True
   FillColor       =   &H00B38200&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   ScaleHeight     =   1485
   ScaleWidth      =   3435
   Begin P控件集.PUIMgr PUI 
      Left            =   1560
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   120
   End
   Begin P控件集.PUIMgr PUI2 
      Left            =   2160
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "PContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim C_Color_Back As OLE_COLOR
Dim C_Color_Back_Down As OLE_COLOR
Dim C_Color_Circle As OLE_COLOR
Dim C_Color_Back_ChangeSpeed As Integer
Dim C_Size_Circle_ChangeSpeed_1 As Integer
Dim C_Size_Circle_ChangeSpeed_2 As Integer

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ClickedX As Single, ClickedY As Single, Downed As Boolean, NowR As Integer, MinR As Integer, i As Integer

Public Sub Mouse_Down(X As Single, Y As Single)
    UserControl_MouseDown 0, 0, X, Y
End Sub

Public Sub Mouse_Up()
    UserControl_MouseUp 0, 0, 0, 0
End Sub

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    Cls
    BackColor = C_Color_Back
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_Back_Down() As OLE_COLOR
    Color_Back_Down = C_Color_Back_Down
End Property

Public Property Let Color_Back_Down(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Down = vNewValue
    PropertyChanged "Color_Back_Down"
End Property

Public Property Get Color_Circle() As OLE_COLOR
    Color_Circle = C_Color_Circle
End Property

Public Property Let Color_Circle(ByVal vNewValue As OLE_COLOR)
    C_Color_Circle = vNewValue
    FillColor = C_Color_Circle
    PropertyChanged "Color_Circle"
End Property

Public Property Get Color_Back_ChangeSpeed() As Integer
    Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
End Property

Public Property Let Color_Back_ChangeSpeed(ByVal vNewValue As Integer)
    C_Color_Back_ChangeSpeed = vNewValue
    PropertyChanged "Color_Back_ChangeSpeed"
End Property

Public Property Get Size_Circle_ChangeSpeed_1() As Integer
    Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
End Property

Public Property Let Size_Circle_ChangeSpeed_1(ByVal vNewValue As Integer)
    C_Size_Circle_ChangeSpeed_1 = vNewValue
    PropertyChanged "Size_Circle_ChangeSpeed_1"
End Property

Public Property Get Size_Circle_ChangeSpeed_2() As Integer
    Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
End Property

Public Property Let Size_Circle_ChangeSpeed_2(ByVal vNewValue As Integer)
    C_Size_Circle_ChangeSpeed_2 = vNewValue
    PropertyChanged "Size_Circle_ChangeSpeed_2"
End Property

Private Sub PUI_ColorSmlyIng(nColor As Long)
    BackColor = nColor
    UserControl.Circle (ClickedX, ClickedY), NowR, C_Color_Circle
End Sub

Private Sub PUI2_ColorSmlyComplete()
    Cls
    RaiseEvent Click
End Sub

Private Sub PUI2_ColorSmlyIng(nColor As Long)
    BackColor = nColor
End Sub

Private Sub Timer1_Timer()
    If Downed Then
        For i = 1 To C_Size_Circle_ChangeSpeed_1
            NowR = NowR + 15
            UserControl.Circle (ClickedX, ClickedY), NowR, C_Color_Circle
        Next
    Else
        Timer1.Enabled = False
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    If MinR > NowR Then
        For i = 1 To C_Size_Circle_ChangeSpeed_2
            NowR = NowR + 15
            UserControl.Circle (ClickedX, ClickedY), NowR, C_Color_Circle
        Next
    Else
        Timer2.Enabled = False
        Cls
        PUI2.ColorSmly BackColor, C_Color_Back, C_Color_Back_ChangeSpeed, 1
    End If
End Sub

Private Sub UserControl_Initialize()
    C_Color_Back = &HF2AF00
    C_Color_Back_Down = &HE3A500
    C_Color_Circle = &HB38200
    C_Color_Back_ChangeSpeed = 2
    C_Size_Circle_ChangeSpeed_1 = 1
    C_Size_Circle_ChangeSpeed_2 = 6
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PUI.StopColorSmly
    PUI2.StopColorSmly
    Cls
    Timer1.Enabled = False
    Timer2.Enabled = False
    ClickedX = X
    ClickedY = Y
    NowR = 0
    MinR = Sqr(X * X + Y * Y) + 15
    If MinR < Sqr(X * X + (UserControl.Height - Y) * (UserControl.Height - Y)) Then MinR = Sqr(X * X + (UserControl.Height - Y) * (UserControl.Height - Y)) + 15
    If MinR < Sqr((UserControl.Width - X) * (UserControl.Width - X) + Y * Y) Then MinR = Sqr((UserControl.Width - X) * (UserControl.Width - X) + Y * Y) + 15
    If MinR < Sqr((UserControl.Width - X) * (UserControl.Width - X) + (UserControl.Height - Y) * (UserControl.Height - Y)) Then MinR = Sqr((UserControl.Width - X) * (UserControl.Width - X) + (UserControl.Height - Y) * (UserControl.Height - Y)) + 15
    PUI.ColorSmly BackColor, C_Color_Back_Down, C_Color_Back_ChangeSpeed, 1
    Downed = True
    Timer1.Enabled = True
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Downed = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)
    C_Color_Back_Down = PropBag.ReadProperty("Color_Back_Down", &HE3A500)
    C_Color_Circle = PropBag.ReadProperty("Color_Circle", &HB38200)
    C_Color_Back_ChangeSpeed = PropBag.ReadProperty("Color_Back_ChangeSpeed", 2)
    C_Size_Circle_ChangeSpeed_1 = PropBag.ReadProperty("Size_Circle_ChangeSpeed_1", 1)
    C_Size_Circle_ChangeSpeed_2 = PropBag.ReadProperty("Size_Circle_ChangeSpeed_2", 6)
    BackColor = C_Color_Back
    FillColor = C_Color_Circle
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)
    Call PropBag.WriteProperty("Color_Back_Down", C_Color_Back_Down, &HE3A500)
    Call PropBag.WriteProperty("Color_Circle", C_Color_Circle, &HB38200)
    Call PropBag.WriteProperty("Color_Back_ChangeSpeed", C_Color_Back_ChangeSpeed, 2)
    Call PropBag.WriteProperty("Size_Circle_ChangeSpeed_1", C_Size_Circle_ChangeSpeed_1, 1)
    Call PropBag.WriteProperty("Size_Circle_ChangeSpeed_2", C_Size_Circle_ChangeSpeed_2, 6)
End Sub
