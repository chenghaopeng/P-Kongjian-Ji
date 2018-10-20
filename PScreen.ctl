VERSION 5.00
Begin VB.UserControl PScreen 
   BackColor       =   &H00000000&
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   810
   ScaleHeight     =   720
   ScaleWidth      =   810
   Begin VB.PictureBox p 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   240
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   15
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label FontTmp 
      BeginProperty Font 
         Name            =   "µ»œﬂ Light"
         Size            =   11.25
         Charset         =   134
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape s 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   30
      Index           =   0
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   30
   End
End
Attribute VB_Name = "PScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum Shapes: Rectangle = 1: Round = 2: End Enum
Dim C_Color_Back As OLE_COLOR
Dim C_Color_Text As OLE_COLOR
Dim C_Color_Text_Back As OLE_COLOR
Dim C_Text As String
Dim C_Size As Integer
Dim C_Font As Font
Dim C_Distance As Integer
Dim C_Style_Shape As Shapes

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    PropertyChanged "Color_Back"
    Refresh
End Property

Public Property Get Color_Text() As OLE_COLOR
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    PropertyChanged "Color_Text"
    Refresh
End Property

Public Property Get Color_Text_Back() As OLE_COLOR
    Color_Text_Back = C_Color_Text_Back
End Property

Public Property Let Color_Text_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Text_Back = vNewValue
    PropertyChanged "Color_Text_Back"
    Refresh
End Property

Public Property Get Text() As String
    Text = C_Text
End Property
Public Property Let Text(ByVal vNewValue As String)
    C_Text = vNewValue
    PropertyChanged "Text"
    Refresh
End Property

Public Property Get Size() As Integer
    Size = C_Size
End Property

Public Property Let Size(ByVal vNewValue As Integer)
    If vNewValue < 15 Then vNewValue = 15
    If vNewValue > 300 Then vNewValue = 300
    C_Size = vNewValue
    PropertyChanged "Size"
    Refresh
End Property

Public Property Get Font() As Font
    Set Font = C_Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set C_Font = vNewValue
    Refresh
    PropertyChanged "Font"
End Property

Public Property Get Distance() As Integer
    Distance = C_Distance
End Property

Public Property Let Distance(ByVal vNewValue As Integer)
    If vNewValue < 0 Then vNewValue = 0
    If vNewValue > 150 Then vNewValue = 150
    C_Distance = vNewValue
    PropertyChanged "Distance"
    Refresh
End Property

Public Property Get Style_Shape() As Shapes
    Style_Shape = C_Style_Shape
End Property

Public Property Let Style_Shape(ByVal vNewValue As Shapes)
    C_Style_Shape = vNewValue
    PropertyChanged "Style_Shape"
    Refresh
End Property

Private Sub Refresh()
    UserControl.BackColor = C_Color_Back
    P.Cls
    P.Width = UserControl.Width
    P.Height = UserControl.Height
    Set P.Font = C_Font
    P.Print C_Text
    Dim i As Integer, j As Integer, c As Integer
    For i = 1 To s.UBound
        Unload s(i)
    Next
    c = 0
    For i = 1 To (UserControl.Width - C_Size) \ (C_Size + C_Distance) + 1
        For j = 1 To (UserControl.Height - C_Size) \ (C_Size + C_Distance) + 1
            c = c + 1
            Load s(c)
            If C_Style_Shape = 1 Then
                s(c).Shape = 0
            Else
                s(c).Shape = 3
            End If
            s(c).Width = C_Size + 15
            s(c).Height = C_Size + 15
            s(c).Left = (i - 1) * (C_Size + C_Distance)
            s(c).Top = (j - 1) * (C_Size + C_Distance)
            s(c).BackColor = GetPoint(i - 1, j - 1)
            s(c).Visible = True
        Next
    Next
End Sub

Private Function GetPoint(ByVal X As Integer, ByVal Y As Integer) As Long
    Dim Color As Long
    Color = P.Point(X * 15, Y * 15)
    P.PSet (X * 15, Y * 15), &H0&
    If Color = &H0& Then
        GetPoint = C_Color_Text_Back
    Else
        GetPoint = C_Color_Text
    End If
End Function

Private Sub UserControl_Initialize()
    C_Color_Back = &H0&
    C_Color_Text = &HC0C0C0
    C_Color_Text_Back = &H808080
    C_Text = ""
    C_Size = 60
    C_Distance = 15
    C_Style_Shape = 1
    Set C_Font = FontTmp.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &H0&)
    C_Color_Text = PropBag.ReadProperty("Color_Text", &HC0C0C0)
    C_Color_Text_Back = PropBag.ReadProperty("Color_Text_Back", &H808080)
    C_Text = PropBag.ReadProperty("Text", "")
    C_Size = PropBag.ReadProperty("Size", 60)
    C_Distance = PropBag.ReadProperty("Distance", 15)
    C_Style_Shape = PropBag.ReadProperty("Style_Shape", 1)
    Set C_Font = PropBag.ReadProperty("Font", FontTmp.Font)
    Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &H0&)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &HC0C0C0)
    Call PropBag.WriteProperty("Color_Text_Back", C_Color_Text_Back, &H808080)
    Call PropBag.WriteProperty("Text", C_Text, "")
    Call PropBag.WriteProperty("Size", C_Size, 60)
    Call PropBag.WriteProperty("Distance", C_Distance, 15)
    Call PropBag.WriteProperty("Style_Shape", C_Style_Shape, 1)
    Call PropBag.WriteProperty("Font", C_Font, FontTmp.Font)
End Sub

