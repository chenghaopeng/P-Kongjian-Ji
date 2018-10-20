VERSION 5.00
Begin VB.UserControl PCheckBoxE 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ScaleHeight     =   1095
   ScaleWidth      =   2055
   Begin P控件集.PContainer PC 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1931
      Color_Back      =   16299591
      Color_Back_Down =   16297782
      Color_Circle    =   16296487
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
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PCheckBoxE"
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
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "PCheckBoxE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim C_Color_Back_1 As OLE_COLOR
Dim C_Color_Back_Down_1 As OLE_COLOR
Dim C_Color_Circle_1 As OLE_COLOR
Dim C_Color_Back_2 As OLE_COLOR
Dim C_Color_Back_Down_2 As OLE_COLOR
Dim C_Color_Circle_2 As OLE_COLOR
Dim C_Color_Back_ChangeSpeed As Integer
Dim C_Size_Circle_ChangeSpeed_1 As Integer
Dim C_Size_Circle_ChangeSpeed_2 As Integer
Dim C_Text As String
Dim C_Font As Font
Dim C_Value As Boolean
Dim C_Color_Text As OLE_COLOR

Public Event ValueChange(NewValue As Boolean)

Private Sub Refresh()
    If Not C_Value Then
        PC.Color_Back = C_Color_Back_1
        PC.Color_Back_Down = C_Color_Back_Down_1
        PC.Color_Circle = C_Color_Circle_1
    Else
        PC.Color_Back = C_Color_Back_2
        PC.Color_Back_Down = C_Color_Back_Down_2
        PC.Color_Circle = C_Color_Circle_2
    End If
    PC.Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
    PC.Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
    PC.Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
    Set L.Font = C_Font
    L = C_Text
    L.ForeColor = C_Color_Text
    L.Top = (PC.Height - L.Height) / 2
    L.Left = (PC.Width - L.Width) / 2
End Sub

Public Property Get Color_Back_1() As OLE_COLOR
    Color_Back_1 = C_Color_Back_1
End Property

Public Property Let Color_Back_1(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_1 = vNewValue
    Refresh
    PropertyChanged "Color_Back_1"
End Property

Public Property Get Color_Back_Down_1() As OLE_COLOR
    Color_Back_Down_1 = C_Color_Back_Down_1
End Property

Public Property Let Color_Back_Down_1(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Down_1 = vNewValue
    Refresh
    PropertyChanged "Color_Back_Down_1"
End Property

Public Property Get Color_Circle_1() As OLE_COLOR
    Color_Circle_1 = C_Color_Circle_1
End Property

Public Property Let Color_Circle_1(ByVal vNewValue As OLE_COLOR)
    C_Color_Circle_1 = vNewValue
    Refresh
    PropertyChanged "Color_Circle_1"
End Property

Public Property Get Color_Back_2() As OLE_COLOR
    Color_Back_2 = C_Color_Back_2
End Property

Public Property Let Color_Back_2(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_2 = vNewValue
    Refresh
    PropertyChanged "Color_Back_2"
End Property

Public Property Get Color_Back_Down_2() As OLE_COLOR
    Color_Back_Down_2 = C_Color_Back_Down_2
End Property

Public Property Let Color_Back_Down_2(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Down_2 = vNewValue
    Refresh
    PropertyChanged "Color_Back_Down_2"
End Property

Public Property Get Color_Circle_2() As OLE_COLOR
    Color_Circle_2 = C_Color_Circle_2
End Property

Public Property Let Color_Circle_2(ByVal vNewValue As OLE_COLOR)
    C_Color_Circle_2 = vNewValue
    Refresh
    PropertyChanged "Color_Circle_2"
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

Public Property Get Value() As Boolean
    Value = C_Value
End Property

Public Property Let Value(ByVal vNewValue As Boolean)
    C_Value = vNewValue
    PropertyChanged "Value"
    Refresh
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
End Sub

Private Sub L_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PC.Mouse_Up
End Sub

Private Sub PC_Click()
    C_Value = Not C_Value
    RaiseEvent ValueChange(C_Value)
    PropertyChanged "Value"
    Refresh
End Sub

Private Sub UserControl_Initialize()
    C_Color_Back_1 = &HF8B647
    C_Color_Back_Down_1 = &HF8AF36
    C_Color_Circle_1 = &HF8AA27
    C_Color_Back_2 = &HF2AF00
    C_Color_Back_Down_2 = &HE3A500
    C_Color_Circle_2 = &HB38200
    C_Color_Back_ChangeSpeed = 2
    C_Size_Circle_ChangeSpeed_1 = 1
    C_Size_Circle_ChangeSpeed_2 = 6
    C_Text = "PCheckBoxE"
    Set C_Font = Label1.Font
    C_Value = False
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
    C_Color_Back_1 = PropBag.ReadProperty("Color_Back_1", &HF8B647)
    C_Color_Back_Down_1 = PropBag.ReadProperty("Color_Back_Down_1", &HF8AF36)
    C_Color_Circle_1 = PropBag.ReadProperty("Color_Circle_1", &HF8AA27)
    C_Color_Back_2 = PropBag.ReadProperty("Color_Back_2", &HF2AF00)
    C_Color_Back_Down_2 = PropBag.ReadProperty("Color_Back_Down_2", &HE3A500)
    C_Color_Circle_2 = PropBag.ReadProperty("Color_Circle_2", &HB38200)
    C_Color_Back_ChangeSpeed = PropBag.ReadProperty("Color_Back_ChangeSpeed", 2)
    C_Size_Circle_ChangeSpeed_1 = PropBag.ReadProperty("Size_Circle_ChangeSpeed_1", 1)
    C_Size_Circle_ChangeSpeed_2 = PropBag.ReadProperty("Size_Circle_ChangeSpeed_2", 6)
    C_Text = PropBag.ReadProperty("Text", "PCheckBoxE")
    Set C_Font = PropBag.ReadProperty("Font", Label1.Font)
    C_Value = PropBag.ReadProperty("Value", False)
    C_Color_Text = PropBag.ReadProperty("Color_Text", vbBlack)
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back_1", C_Color_Back_1, &HF8B647)
    Call PropBag.WriteProperty("Color_Back_Down_1", C_Color_Back_Down_1, &HF8AF36)
    Call PropBag.WriteProperty("Color_Circle_1", C_Color_Circle_1, &HF8AA27)
    Call PropBag.WriteProperty("Color_Back_2", C_Color_Back_2, &HF2AF00)
    Call PropBag.WriteProperty("Color_Back_Down_2", C_Color_Back_Down_2, &HE3A500)
    Call PropBag.WriteProperty("Color_Circle_2", C_Color_Circle_2, &HB38200)
    Call PropBag.WriteProperty("Color_Back_ChangeSpeed", C_Color_Back_ChangeSpeed, 2)
    Call PropBag.WriteProperty("Size_Circle_ChangeSpeed_1", C_Size_Circle_ChangeSpeed_1, 1)
    Call PropBag.WriteProperty("Size_Circle_ChangeSpeed_2", C_Size_Circle_ChangeSpeed_2, 6)
    Call PropBag.WriteProperty("Text", C_Text, "PCheckBoxE")
    Call PropBag.WriteProperty("Font", C_Font, Label1.Font)
    Call PropBag.WriteProperty("Value", C_Value, False)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, vbBlack)
End Sub
