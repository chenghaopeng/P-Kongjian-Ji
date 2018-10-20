VERSION 5.00
Begin VB.UserControl PButtonE 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin P控件集.PContainer PC 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1931
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PButtonE"
         BeginProperty Font 
            Name            =   "等线 Light"
            Size            =   14.25
            Charset         =   134
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "等线 Light"
            Size            =   14.25
            Charset         =   134
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
   End
End
Attribute VB_Name = "PButtonE"
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
Dim C_Text As String
Dim C_Font As Font
Dim C_Color_Text As OLE_COLOR

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Refresh()
    PC.Color_Back = C_Color_Back
    PC.Color_Back_Down = C_Color_Back_Down
    PC.Color_Circle = C_Color_Circle
    PC.Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
    PC.Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
    PC.Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
    Set L.Font = C_Font
    L = C_Text
    L.ForeColor = C_Color_Text
    L.Top = (PC.Height - L.Height) / 2
    L.Left = (PC.Width - L.Width) / 2
End Sub

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    Refresh
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_Back_Down() As OLE_COLOR
    Color_Back_Down = C_Color_Back_Down
End Property

Public Property Let Color_Back_Down(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Down = vNewValue
    Refresh
    PropertyChanged "Color_Back_Down"
End Property

Public Property Get Color_Circle() As OLE_COLOR
    Color_Circle = C_Color_Circle
End Property

Public Property Let Color_Circle(ByVal vNewValue As OLE_COLOR)
    C_Color_Circle = vNewValue
    Refresh
    PropertyChanged "Color_Circle"
End Property

Public Property Get Color_Back_ChangeSpeed() As Integer
    Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
End Property

Public Property Let Color_Back_ChangeSpeed(ByVal vNewValue As Integer)
    C_Color_Back_ChangeSpeed = vNewValue
    Refresh
    PropertyChanged "Color_Back_ChangeSpeed"
End Property

Public Property Get Size_Circle_ChangeSpeed_1() As Integer
    Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
End Property

Public Property Let Size_Circle_ChangeSpeed_1(ByVal vNewValue As Integer)
    C_Size_Circle_ChangeSpeed_1 = vNewValue
    Refresh
    PropertyChanged "Size_Circle_ChangeSpeed_1"
End Property

Public Property Get Size_Circle_ChangeSpeed_2() As Integer
    Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
End Property

Public Property Let Size_Circle_ChangeSpeed_2(ByVal vNewValue As Integer)
    C_Size_Circle_ChangeSpeed_2 = vNewValue
    Refresh
    PropertyChanged "Size_Circle_ChangeSpeed_2"
End Property

Public Property Get Text() As String
    Text = C_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    C_Text = vNewValue
    Refresh
    PropertyChanged "Text"
End Property

Public Property Get Font() As Font
    Set Font = C_Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set C_Font = vNewValue
    Refresh
    PropertyChanged "Font"
End Property

Public Property Get Color_Text() As OLE_COLOR
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    Refresh
    PropertyChanged "Color_Text"
End Property

Private Sub L_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PC.Mouse_Down X + L.Left, Y + L.Top
    RaiseEvent MouseDown(Button, Shift, X + L.Left, Y + L.Top)
End Sub

Private Sub L_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X + L.Left, Y + L.Top)
End Sub

Private Sub L_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PC.Mouse_Up
    RaiseEvent MouseUp(Button, Shift, X + L.Left, Y + L.Top)
End Sub

Private Sub PC_Click()
    RaiseEvent Click
End Sub

Private Sub PC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub PC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
    C_Color_Back = &HF2AF00
    C_Color_Back_Down = &HE3A500
    C_Color_Circle = &HB38200
    C_Color_Back_ChangeSpeed = 2
    C_Size_Circle_ChangeSpeed_1 = 1
    C_Size_Circle_ChangeSpeed_2 = 6
    C_Text = "PButtonE"
    Set C_Font = Label1.Font
    C_Color_Text = vbBlack
    Refresh
End Sub

Private Sub UserControl_Resize()
    PC.Height = UserControl.Height
    PC.Width = UserControl.Width
    L.Top = (PC.Height - L.Height) / 2
    L.Left = (PC.Width - L.Width) / 2
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)
    C_Color_Back_Down = PropBag.ReadProperty("Color_Back_Down", &HE3A500)
    C_Color_Circle = PropBag.ReadProperty("Color_Circle", &HB38200)
    C_Color_Back_ChangeSpeed = PropBag.ReadProperty("Color_Back_ChangeSpeed", 2)
    C_Size_Circle_ChangeSpeed_1 = PropBag.ReadProperty("Size_Circle_ChangeSpeed_1", 1)
    C_Size_Circle_ChangeSpeed_2 = PropBag.ReadProperty("Size_Circle_ChangeSpeed_2", 6)
    C_Text = PropBag.ReadProperty("Text", "PButtonE")
    Set C_Font = PropBag.ReadProperty("Font", Label1.Font)
    C_Color_Text = PropBag.ReadProperty("Color_Text", vbBlack)
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)
    Call PropBag.WriteProperty("Color_Back_Down", C_Color_Back_Down, &HE3A500)
    Call PropBag.WriteProperty("Color_Circle", C_Color_Circle, &HB38200)
    Call PropBag.WriteProperty("Color_Back_ChangeSpeed", C_Color_Back_ChangeSpeed, 2)
    Call PropBag.WriteProperty("Size_Circle_ChangeSpeed_1", C_Size_Circle_ChangeSpeed_1, 1)
    Call PropBag.WriteProperty("Size_Circle_ChangeSpeed_2", C_Size_Circle_ChangeSpeed_2, 6)
    Call PropBag.WriteProperty("Text", C_Text, "PButtonE")
    Call PropBag.WriteProperty("Font", C_Font, Label1.Font)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, vbBlack)
End Sub

