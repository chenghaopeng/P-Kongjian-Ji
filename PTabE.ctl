VERSION 5.00
Begin VB.UserControl PTabE 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1425
   ScaleWidth      =   4800
   Begin P控件集.PUIMgr PUI 
      Left            =   3000
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin P控件集.PButtonE PBE1 
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      Text            =   "<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "等线 Light"
         Size            =   10.5
         Charset         =   134
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin P控件集.PButtonE PBE2 
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      Text            =   ">"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "等线 Light"
         Size            =   10.5
         Charset         =   134
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox P 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   2
      Top             =   0
      Width           =   2535
      Begin P控件集.PCheckBoxE PO 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Text            =   "PTabE"
      End
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
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "PTabE"
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
Dim C_Color_Text As OLE_COLOR
Dim C_ExtendWidth As Integer
Dim C_ScrollSpeed As Single

Public Event IndexChange(NewIndex As Integer, LastIndex As Integer)

Public Sub SetIndex(ByVal Index As Integer)
    If Index > PO.UBound Then Exit Sub
    PO_ValueChange Index, True
End Sub

Private Sub Refresh()
    PBE1.Color_Back = C_Color_Back_2
    PBE1.Color_Back_Down = C_Color_Back_Down_2
    PBE1.Color_Circle = C_Color_Circle_2
    PBE1.Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
    PBE1.Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
    PBE1.Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
    PBE1.Color_Text = C_Color_Text
    PBE2.Color_Back = C_Color_Back_2
    PBE2.Color_Back_Down = C_Color_Back_Down_2
    PBE2.Color_Circle = C_Color_Circle_2
    PBE2.Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
    PBE2.Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
    PBE2.Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
    PBE2.Color_Text = C_Color_Text
    Dim i As Integer, s() As String
    s = Split(C_Text, "|")
    For i = 1 To PO.UBound
        Unload PO(i)
    Next
    Set L.Font = C_Font
    P.Left = 0
    L = s(0)
    PO(0).Width = L.Width + C_ExtendWidth
    PO(0).Text = s(0)
    PO(0).Color_Back_1 = C_Color_Back_1
    PO(0).Color_Back_Down_1 = C_Color_Back_Down_1
    PO(0).Color_Circle_1 = C_Color_Circle_1
    PO(0).Color_Back_2 = C_Color_Back_2
    PO(0).Color_Back_Down_2 = C_Color_Back_Down_2
    PO(0).Color_Circle_2 = C_Color_Circle_2
    PO(0).Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
    PO(0).Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
    PO(0).Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
    Set PO(0).Font = C_Font
    PO(0).Color_Text = C_Color_Text
    P.Width = PO(0).Width
    PO(0).Value = True
    On Error Resume Next
    For i = 1 To UBound(s)
        Load PO(i)
        L = s(i)
        PO(i).Width = L.Width + C_ExtendWidth
        PO(i).Text = s(i)
        PO(i).Color_Back_1 = C_Color_Back_1
        PO(i).Color_Back_Down_1 = C_Color_Back_Down_1
        PO(i).Color_Circle_1 = C_Color_Circle_1
        PO(i).Color_Back_2 = C_Color_Back_2
        PO(i).Color_Back_Down_2 = C_Color_Back_Down_2
        PO(i).Color_Circle_2 = C_Color_Circle_2
        PO(i).Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
        PO(i).Size_Circle_ChangeSpeed_1 = C_Size_Circle_ChangeSpeed_1
        PO(i).Size_Circle_ChangeSpeed_2 = C_Size_Circle_ChangeSpeed_2
        Set PO(i).Font = C_Font
        PO(i).Color_Text = C_Color_Text
        Set PO(i).Container = P
        PO(i).Left = PO(i - 1).Left + PO(i - 1).Width
        PO(i).Value = False
        PO(i).Visible = True
        P.Width = PO(i).Left + PO(i).Width
    Next
    UserControl_Resize
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

Public Property Get Color_Text() As OLE_COLOR
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    Refresh
    PropertyChanged "Color_Text"
End Property

Public Property Get ExtendWidth() As Integer
    ExtendWidth = C_ExtendWidth
End Property

Public Property Let ExtendWidth(ByVal vNewValue As Integer)
    C_ExtendWidth = vNewValue
    Refresh
    PropertyChanged "ExtendWidth"
End Property

Public Property Get ScrollSpeed() As Single
    ScrollSpeed = C_ScrollSpeed
End Property

Public Property Let ScrollSpeed(ByVal vNewValue As Single)
    C_ScrollSpeed = vNewValue
    PropertyChanged "ScrollSpeed"
End Property

Private Sub PBE1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If P.Left + (P.Width - UserControl.Width + 510) * C_ScrollSpeed < 0 Then
        PUI.MoveSmly P, P.Left + (P.Width - UserControl.Width + 510) * C_ScrollSpeed, P.Top, 1
    Else
        PUI.MoveSmly P, 0, P.Top, 1
    End If
End Sub

Private Sub PBE2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If P.Left - (P.Width - UserControl.Width + 510) * C_ScrollSpeed > -(P.Width - UserControl.Width + 510) Then
        PUI.MoveSmly P, P.Left - (P.Width - UserControl.Width + 510) * C_ScrollSpeed, P.Top, 1
    Else
        PUI.MoveSmly P, -(P.Width - UserControl.Width + 510), P.Top, 1
    End If
End Sub

Private Sub PO_ValueChange(Index As Integer, NewValue As Boolean)
    If NewValue = False Then
        PO(Index).Value = True
    End If
    Dim i As Integer, j As Integer
    j = -1
    For i = 0 To PO.UBound
        If Index <> i And PO(i).Value Then
            PO(i).Value = False
            j = i
            Exit For
        End If
    Next
    RaiseEvent IndexChange(Index, j)
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
    C_Text = "PTabE"
    Set C_Font = Label1.Font
    C_Color_Text = vbBlack
    C_ExtendWidth = 240
    C_ScrollSpeed = 0.2
    Refresh
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
    C_Text = PropBag.ReadProperty("Text", "PTabE")
    Set C_Font = PropBag.ReadProperty("Font", Label1.Font)
    C_Color_Text = PropBag.ReadProperty("Color_Text", vbBlack)
    C_ExtendWidth = PropBag.ReadProperty("ExtendWidth", 240)
    C_ScrollSpeed = PropBag.ReadProperty("ScrollSpeed", 0.2)
    Refresh
End Sub

Private Sub UserControl_Resize()
    P.Left = 0
    P.Height = UserControl.Height
    PBE1.Height = UserControl.Height
    PBE2.Height = UserControl.Height
    PBE1.Left = UserControl.Width - 510
    PBE2.Left = UserControl.Width - 255
    Dim i As Integer, Total As Long
    For i = 0 To PO.UBound
        PO(i).Height = UserControl.Height
    Next
    Total = PO(PO.UBound).Width + PO(PO.UBound).Left
    If Total <= UserControl.Width Then
        PBE1.Visible = False
        PBE2.Visible = False
    Else
        PBE1.Visible = True
        PBE2.Visible = True
    End If
    P.Width = Total
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
    Call PropBag.WriteProperty("Text", C_Text, "PTabE")
    Call PropBag.WriteProperty("Font", C_Font, Label1.Font)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, vbBlack)
    Call PropBag.WriteProperty("ExtendWidth", C_ExtendWidth, 240)
    Call PropBag.WriteProperty("ScrollSpeed", C_ScrollSpeed, 0.2)
End Sub

