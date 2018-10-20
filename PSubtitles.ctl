VERSION 5.00
Begin VB.UserControl PSubtitles 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin P控件集.PUIMgr PUI 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   135
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1080
      Top             =   600
   End
   Begin VB.Label FontTmp 
      BeginProperty Font 
         Name            =   "等线 Light"
         Size            =   9.75
         Charset         =   134
         Weight          =   300
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "PSubtitles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim C_TextsAndLinks As String
Dim C_Color_Text As OLE_COLOR
Dim C_Color_Text_End As OLE_COLOR
Dim C_Color_Back As OLE_COLOR
Dim C_Font As Font
Dim C_Is_Enabled As Boolean
Dim C_Is_Back_Transparent As Boolean
Dim C_Interval As Integer
Dim C_Text_Align As Integer
Dim C_Is_Random As Boolean
Dim C_Color_Back_ChangeSpeed As Integer

Dim NowIndex As Integer
Dim s() As String

Private Sub Refresh()
    Timer1.Enabled = False
    UserControl.Cls
    With Label1
        Set .Font = C_Font
        .ForeColor = C_Color_Text
    End With
    NowIndex = 0
    s = Split(C_TextsAndLinks, "|")
    Label1.Caption = Left(s(0), InStr(s(0), ",") - 1)
    Label1.Tag = Right(s(0), Len(s(0)) - InStr(s(0), ","))
    UserControl.BackColor = C_Color_Back
    If C_Is_Back_Transparent Then If BeAbleToBeBackTransparent Then UserControl.PaintPicture UserControl.Extender.Container.Image, 0, 0, UserControl.Width, UserControl.Height, UserControl.Extender.Left, UserControl.Extender.Top, UserControl.Width, UserControl.Height
    SetTextPos
    Timer1.Interval = C_Interval
    Timer1.Enabled = True
End Sub

Private Sub SetTextPos()
    Picture1.Width = Label1.Width
    Picture1.Height = Label1.Height
    Select Case C_Text_Align
        Case 1
            Picture1.Left = 0
            Picture1.Top = 0
        Case 2
            Picture1.Left = (UserControl.Width - Label1.Width) / 2
            Picture1.Top = 0
        Case 3
            Picture1.Left = UserControl.Width - Label1.Width
            Picture1.Top = 0
        Case 4
            Picture1.Left = 0
            Picture1.Top = (UserControl.Height - Label1.Height) / 2
        Case 5
            Picture1.Left = (UserControl.Width - Label1.Width) / 2
            Picture1.Top = (UserControl.Height - Label1.Height) / 2
        Case 6
            Picture1.Left = UserControl.Width - Label1.Width
            Picture1.Top = (UserControl.Height - Label1.Height) / 2
        Case 7
            Picture1.Left = 0
            Picture1.Top = UserControl.Height - Label1.Height
        Case 8
            Picture1.Left = (UserControl.Width - Label1.Width) / 2
            Picture1.Top = UserControl.Height - Label1.Height
        Case 9
            Picture1.Left = UserControl.Width - Label1.Width
            Picture1.Top = UserControl.Height - Label1.Height
    End Select
    Picture1.PaintPicture UserControl.Image, 0, 0, , , Picture1.Left, Picture1.Top
End Sub

Private Function BeAbleToBeBackTransparent() As Boolean
    On Error GoTo Err
    Dim pppp As StdPicture
    Set pppp = UserControl.Extender.Container.Image
    BeAbleToBeBackTransparent = True
    Exit Function
Err:
    BeAbleToBeBackTransparent = False
End Function

Public Property Get TextsAndLinks() As String
    TextsAndLinks = C_TextsAndLinks
End Property

Public Property Let TextsAndLinks(ByVal vNewValue As String)
    If vNewValue = "" Or InStr(vNewValue, ",") <= 0 Then vNewValue = "EXAMPLE,longdows.cn"
    C_TextsAndLinks = vNewValue
    Refresh
    PropertyChanged "TextsAndLinks"
End Property

Public Property Get Color_Text() As OLE_COLOR
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    Refresh
    PropertyChanged "Color_Text"
End Property

Public Property Get Color_Text_End() As OLE_COLOR
    Color_Text_End = C_Color_Text_End
End Property

Public Property Let Color_Text_End(ByVal vNewValue As OLE_COLOR)
    C_Color_Text_End = vNewValue
    PropertyChanged "Color_Text_End"
End Property

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    Refresh
    PropertyChanged "Color_Back"
End Property

Public Property Get Font() As Font
    Set Font = C_Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set C_Font = vNewValue
    Refresh
    PropertyChanged "Font"
End Property

Public Property Get Is_Enabled() As Boolean
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Is_Back_Transparent() As Boolean
    Is_Back_Transparent = C_Is_Back_Transparent
End Property

Public Property Let Is_Back_Transparent(ByVal vNewValue As Boolean)
    C_Is_Back_Transparent = vNewValue
    Refresh
    PropertyChanged "Is_Back_Transparent"
End Property

Public Property Get Interval() As Integer
    Interval = C_Interval
End Property

Public Property Let Interval(ByVal vNewValue As Integer)
    C_Interval = vNewValue
    Refresh
    PropertyChanged "Interval"
End Property

Public Property Get Text_Align() As Integer
    Text_Align = C_Text_Align
End Property

Public Property Let Text_Align(ByVal vNewValue As Integer)
    If vNewValue < 1 Then vNewValue = 1
    If vNewValue > 9 Then vNewValue = 9
    C_Text_Align = vNewValue
    SetTextPos
    PropertyChanged "Text_Align"
End Property

Public Property Get Is_Random() As Boolean
    Is_Random = C_Is_Random
End Property

Public Property Let Is_Random(ByVal vNewValue As Boolean)
    C_Is_Random = vNewValue
    PropertyChanged "Is_Random"
End Property

Public Property Get Color_Back_ChangeSpeed() As Integer
    Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
End Property

Public Property Let Color_Back_ChangeSpeed(ByVal vNewValue As Integer)
    C_Color_Back_ChangeSpeed = vNewValue
    If C_Color_Back_ChangeSpeed < 1 Then C_Color_Back_ChangeSpeed = 1
    If C_Color_Back_ChangeSpeed > 30 Then C_Color_Back_ChangeSpeed = 30
    PropertyChanged "Color_Back_ChangeSpeed"
End Property

Private Sub Label1_Click()
    If Not C_Is_Enabled Then Exit Sub
    Shell "explorer http:\\" & Label1.Tag
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Reload Picture1.hWnd
    PUI.ColorSmly Label1.ForeColor, C_Color_Text_End, C_Color_Back_ChangeSpeed, 1
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = -108 Then
        PUI.ColorSmly Label1.ForeColor, C_Color_Text, C_Color_Back_ChangeSpeed, 1
    End If
End Sub

Private Sub PUI_ColorSmlyIng(nColor As Long)
    Label1.ForeColor = nColor
End Sub

Private Sub Timer1_Timer()
    Dim T As Integer
    If C_Is_Random And UBound(s) > 3 Then
        Randomize
        T = Int(Rnd() * (UBound(s) + 1)) - 1
        Do Until T >= 0 And T <= UBound(s) And T <> NowIndex
            Randomize
            T = Int(Rnd() * (UBound(s) + 1)) - 1
        Loop
    Else
        If NowIndex = UBound(s) Then
            T = 0
        Else
            T = NowIndex + 1
        End If
    End If
    NowIndex = T
    Label1.Caption = Left(s(T), InStr(s(T), ",") - 1)
    Label1.Tag = Right(s(T), Len(s(T)) - InStr(s(T), ","))
    SetTextPos
    If C_Is_Back_Transparent Then If BeAbleToBeBackTransparent Then UserControl.PaintPicture UserControl.Extender.Container.Image, 0, 0, UserControl.Width, UserControl.Height, UserControl.Extender.Left, UserControl.Extender.Top, UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_Initialize()
    C_TextsAndLinks = "EXAMPLE,longdows.cn"
    C_Color_Text = &HFFFFFF
    C_Color_Text_End = &HFF7402
    C_Color_Back = &HF2AF00
    Set C_Font = FontTmp.Font
    C_Is_Enabled = True
    C_Is_Back_Transparent = True
    C_Interval = 3000
    C_Text_Align = 4
    C_Is_Random = True
    C_Color_Back_ChangeSpeed = 20
    Refresh
    
    MLInit Picture1.hWnd
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_TextsAndLinks = PropBag.ReadProperty("TextsAndLinks", "EXAMPLE,longdows.cn")
    C_Color_Text = PropBag.ReadProperty("Color_Text", &HFFFFFF)
    C_Color_Text_End = PropBag.ReadProperty("Color_Text_End", &HFF7402)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)
    Set C_Font = PropBag.ReadProperty("Font", FontTmp.Font)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Is_Back_Transparent = PropBag.ReadProperty("Is_Back_Transparent", True)
    C_Interval = PropBag.ReadProperty("Interval", 3000)
    C_Text_Align = PropBag.ReadProperty("Text_Align", 4)
    C_Is_Random = PropBag.ReadProperty("Is_Random", True)
    C_Color_Back_ChangeSpeed = PropBag.ReadProperty("Color_Back_ChangeSpeed", 20)
    Refresh
End Sub

Private Sub UserControl_Resize()
    SetTextPos
End Sub

Private Sub UserControl_Terminate()
    MLTerminate Picture1.hWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TextsAndLinks", C_TextsAndLinks, "EXAMPLE,longdows.cn")
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &HFFFFFF)
    Call PropBag.WriteProperty("Color_Text_End", C_Color_Text_End, &HFF7402)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)
    Call PropBag.WriteProperty("Font", C_Font, FontTmp.Font)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Is_Back_Transparent", C_Is_Back_Transparent, True)
    Call PropBag.WriteProperty("Interval", C_Interval, 3000)
    Call PropBag.WriteProperty("Text_Align", C_Text_Align, 4)
    Call PropBag.WriteProperty("Is_Random", C_Is_Random, True)
    Call PropBag.WriteProperty("Color_Back_ChangeSpeed", C_Color_Back_ChangeSpeed, 20)
End Sub
