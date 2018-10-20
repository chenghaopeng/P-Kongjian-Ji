VERSION 5.00
Begin VB.UserControl PWin8Form 
   BackColor       =   &H00F4AD6D&
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ControlContainer=   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   4695
   Begin P控件集.PUIMgr PM 
      Left            =   1680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin P控件集.PUIMgr PUI 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin P控件集.PButton PBBig 
      Height          =   315
      Left            =   3450
      TabIndex        =   4
      ToolTipText     =   "最大化"
      Top             =   15
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      Color_Back      =   16035181
      Color_Back_Down =   10051645
      Color_Begin     =   16035181
      Color_End       =   11756854
      Color_Text      =   2631720
      Text            =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Can_Text_Move   =   0   'False
      Text_Deviate_X  =   2
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   120
      MousePointer    =   1  'Arrow
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   120
      Width           =   240
      Begin VB.Image imaIcon 
         Height          =   240
         Left            =   0
         MousePointer    =   1  'Arrow
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
   End
   Begin P控件集.PButton PBClose 
      Height          =   315
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "关闭"
      Top             =   15
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      Color_Back      =   5263559
      Color_Back_Down =   5197766
      Color_Begin     =   5263559
      Color_End       =   4408288
      Color_Text      =   16777215
      Text            =   "×"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   12.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Can_Text_Move   =   0   'False
      Text_Deviate_Y  =   -2
   End
   Begin P控件集.PButton PBSmall 
      Height          =   315
      Left            =   3060
      TabIndex        =   5
      ToolTipText     =   "最小化"
      Top             =   15
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   556
      Color_Back      =   16035181
      Color_Back_Down =   10051645
      Color_Begin     =   16035181
      Color_End       =   11756854
      Color_Text      =   2631720
      Text            =   "—"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Can_Text_Move   =   0   'False
      Text_Deviate_X  =   1
      Text_Deviate_Y  =   2
   End
   Begin P控件集.PUIMgr PUI2 
      Left            =   1080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      MousePointer    =   1  'Arrow
      ScaleHeight     =   2895
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Image vIcon 
      Height          =   480
      Left            =   2640
      Picture         =   "PWin8Form.ctx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2400
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label labTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PWin8Form"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00282839&
      Height          =   300
      Left            =   480
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   90
      Width           =   2460
   End
   Begin VB.Shape SmallBorder 
      BorderColor     =   &H00CB8549&
      Height          =   2925
      Left            =   105
      Top             =   465
      Width           =   4485
   End
   Begin VB.Shape BigBorder 
      BorderColor     =   &H00CB8549&
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "PWin8Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Const SPI_GETWORKAREA = 48
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim C_Icon As Picture
Dim C_Picture As Picture
Dim C_Caption As String
Dim C_Color_Border As OLE_COLOR
Dim C_Color_Frame As OLE_COLOR
Dim C_Color_Back As OLE_COLOR
Dim C_Is_Stretch As Boolean
Dim C_Can_Move_Smoothly As Boolean
Dim C_Is_Enabled As Boolean
Dim C_Has_MinButton As Boolean
Dim C_Has_MaxButton As Boolean
Dim C_Has_CloseButton As Boolean
Dim C_Has_Icon As Boolean
Dim C_Is_Resizable As Boolean

Dim ClickedX As Single
Dim ClickedY As Single
Dim ChangingSize As Integer
Dim LastY As Single

Dim CanResize As Boolean
Dim IsBig As Boolean

Dim yWidth As Integer, yHeight As Integer, yLeft As Integer, yTop As Integer

Private Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

Public Property Get Icon() As Picture
    Set Icon = C_Icon
End Property

Public Property Set Icon(ByVal vNewValue As Picture)
    Set C_Icon = vNewValue
    Set imaIcon.Picture = C_Icon
    PropertyChanged "Icon"
End Property

Public Property Get Picture() As Picture
    Set Picture = C_Picture
End Property

Public Property Set Picture(ByVal vNewValue As Picture)
    Set C_Picture = vNewValue
    picContainer.Cls
    If (C_Is_Stretch) And (Not (C_Picture Is Nothing)) Then
        Set Image1.Picture = C_Picture
        picContainer.PaintPicture Image1.Picture, 0, 0, picContainer.Width, picContainer.Height, 0, 0, Image1.Width, Image1.Height
    Else
        Set picContainer.Picture = C_Picture
    End If
    PropertyChanged "Picture"
End Property

Public Property Get Caption() As String
    Caption = C_Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    C_Caption = vNewValue
    labTitle = C_Caption
    PropertyChanged "Caption"
End Property

Public Property Get Color_Border() As OLE_COLOR
    Color_Border = C_Color_Border
End Property

Public Property Let Color_Border(ByVal vNewValue As OLE_COLOR)
    C_Color_Border = vNewValue
    BigBorder.BorderColor = C_Color_Border
    SmallBorder.BorderColor = C_Color_Border
    PropertyChanged "Color_Border"
End Property

Public Property Get Color_Frame() As OLE_COLOR
    Color_Frame = C_Color_Frame
End Property

Public Property Let Color_Frame(ByVal vNewValue As OLE_COLOR)
    C_Color_Frame = vNewValue
    UserControl.BackColor = C_Color_Frame
    picIcon.BackColor = C_Color_Frame
    PBSmall.Color_Back = C_Color_Frame
    PBBig.Color_Back = C_Color_Frame
    PropertyChanged "Color_Frame"
End Property

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    picContainer.BackColor = C_Color_Back
    PropertyChanged "Color_Back"
End Property

Public Property Get Is_Stretch() As Boolean
    Is_Stretch = C_Is_Stretch
End Property

Public Property Let Is_Stretch(ByVal vNewValue As Boolean)
    C_Is_Stretch = vNewValue
    picContainer.Cls
    If (C_Is_Stretch) And (Not (C_Picture Is Nothing)) Then
        Set Image1.Picture = C_Picture
        picContainer.PaintPicture Image1.Picture, 0, 0, picContainer.Width, picContainer.Height, 0, 0, Image1.Width, Image1.Height
    Else
        Set picContainer.Picture = C_Picture
    End If
    PropertyChanged "Is_Stretch"
End Property

Public Property Get Can_Move_Smoothly() As Boolean
    Can_Move_Smoothly = C_Can_Move_Smoothly
End Property

Public Property Let Can_Move_Smoothly(ByVal vNewValue As Boolean)
    C_Can_Move_Smoothly = vNewValue
    PropertyChanged "Can_Move_Smoothly"
End Property

Public Property Get Is_Enabled() As Boolean
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Has_MinButton() As Boolean
    Has_MinButton = C_Has_MinButton
End Property

Public Property Let Has_MinButton(ByVal vNewValue As Boolean)
    C_Has_MinButton = vNewValue
    UserControl_Resize
    PropertyChanged "Has_MinButton"
End Property

Public Property Get Has_MaxButton() As Boolean
    Has_MaxButton = C_Has_MaxButton
End Property

Public Property Let Has_MaxButton(ByVal vNewValue As Boolean)
    C_Has_MaxButton = vNewValue
    UserControl_Resize
    PropertyChanged "Has_MaxButton"
End Property

Public Property Get Has_CloseButton() As Boolean
    Has_CloseButton = C_Has_CloseButton
End Property

Public Property Let Has_CloseButton(ByVal vNewValue As Boolean)
    C_Has_CloseButton = vNewValue
    UserControl_Resize
    PropertyChanged "Has_CloseButton"
End Property

Public Property Get Has_Icon() As Boolean
    Has_Icon = C_Has_Icon
End Property

Public Property Let Has_Icon(ByVal vNewValue As Boolean)
    C_Has_Icon = vNewValue
    UserControl_Resize
    PropertyChanged "Has_Icon"
End Property

Public Property Get Is_Resizable() As Boolean
    Is_Resizable = C_Is_Resizable
End Property

Public Property Let Is_Resizable(ByVal vNewValue As Boolean)
    C_Is_Resizable = vNewValue
    PropertyChanged "Is_Resizable"
End Property

Private Sub imaIcon_DblClick()
    Unload UserControl.Extender.Container
End Sub

Private Sub labTitle_DblClick()
    PBBig_Click
End Sub

Private Sub PBBig_Click()
    If C_Has_MaxButton = False Then Exit Sub
    If Not IsBig Then
        yWidth = UserControl.Extender.Container.Width
        yHeight = UserControl.Extender.Container.Height
        yLeft = UserControl.Extender.Container.Left
        yTop = UserControl.Extender.Container.Top
        If C_Can_Move_Smoothly Then
            PUI.MoveSmly UserControl.Extender.Container, 0, 0, 1
            PUI.SizeSmly UserControl.Extender.Container, Screen.Width, Screen.Height - GetTaskbarHeight, 1
        Else
            UserControl.Extender.Container.Left = 0
            UserControl.Extender.Container.Top = 0
            UserControl.Extender.Container.Width = Screen.Width
            UserControl.Extender.Container.Height = Screen.Height - GetTaskbarHeight
        End If
        UserControl.Width = Screen.Width
        UserControl.Height = Screen.Height - GetTaskbarHeight
        CanResize = C_Is_Resizable
        C_Is_Resizable = False
        IsBig = True
        PBBig.Text = "2"
    Else
        If C_Can_Move_Smoothly Then
            PUI.MoveSmly UserControl.Extender.Container, yLeft, yTop, 1
            PUI.SizeSmly UserControl.Extender.Container, yWidth, yHeight, 1
        Else
            UserControl.Extender.Container.Width = yWidth
            UserControl.Extender.Container.Height = yHeight
            UserControl.Extender.Container.Left = yLeft
            UserControl.Extender.Container.Top = yTop
        End If
        UserControl.Width = yWidth
        UserControl.Height = yHeight
        C_Is_Resizable = CanResize
        IsBig = False
        PBBig.Text = "1"
    End If
    UserControl_Resize
End Sub

Private Sub PBClose_Click()
    Unload UserControl.Extender.Container
End Sub

Private Sub PBSmall_Click()
    UserControl.Extender.Container.WindowState = 1
End Sub

Private Sub UserControl_DblClick()
    If LastY < 480 Then PBBig_Click
End Sub

Private Sub UserControl_Initialize()
    ClickedX = -1
    ClickedY = -1
    ChangingSize = 0
    Set C_Icon = vIcon.Picture
    Set C_Picture = Nothing
    C_Caption = "PWin8Form"
    C_Color_Border = &HCB8549
    C_Color_Frame = &HF4AD6D
    C_Color_Back = &HFFFFFF
    C_Is_Stretch = False
    C_Can_Move_Smoothly = False
    C_Is_Enabled = True
    C_Has_MinButton = True
    C_Has_MaxButton = True
    C_Has_CloseButton = True
    C_Has_Icon = True
    C_Is_Resizable = True
    picContainer.BackColor = C_Color_Back
    UserControl.BackColor = C_Color_Frame
    picIcon.BackColor = C_Color_Frame
    PBSmall.Color_Back = C_Color_Frame
    PBBig.Color_Back = C_Color_Frame
    BigBorder.BorderColor = C_Color_Border
    SmallBorder.BorderColor = C_Color_Border
    labTitle = C_Caption
    picContainer.Cls
    If (C_Is_Stretch) And (Not (C_Picture Is Nothing)) Then
        Set Image1.Picture = C_Picture
        picContainer.PaintPicture Image1.Picture, 0, 0, picContainer.Width, picContainer.Height, 0, 0, Image1.Width, Image1.Height
    Else
        Set picContainer.Picture = C_Picture
    End If
    Set imaIcon.Picture = C_Icon
    UserControl_Resize
End Sub

Private Sub labTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (Not IsBig) Then
        ClickedX = X
        ClickedY = Y
    End If
End Sub

Private Sub labTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (ClickedX <> -1) And (ClickedY <> -1) Then
        If C_Can_Move_Smoothly Then
            PUI.MoveSmly UserControl.Extender.Container, UserControl.Extender.Container.Left + X - ClickedX, UserControl.Extender.Container.Top + Y - ClickedY, 1
        Else
            UserControl.Extender.Container.Left = UserControl.Extender.Container.Left + X - ClickedX
            UserControl.Extender.Container.Top = UserControl.Extender.Container.Top + Y - ClickedY
        End If
    End If
End Sub

Private Sub labTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickedX = -1
    ClickedY = -1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (Not IsBig) Then
        If UserControl.Width - X <= 120 And Y <= 120 Then
            ChangingSize = 2
        ElseIf UserControl.Width - X <= 120 And Y >= UserControl.Height - 120 Then
            ChangingSize = 4
        ElseIf X <= 120 And Y >= UserControl.Height - 120 Then
            ChangingSize = 6
        ElseIf X <= 120 And Y <= 120 Then
            ChangingSize = 8
        ElseIf X <= UserControl.Width And Y <= 120 Then
            ChangingSize = 1
        ElseIf UserControl.Width - X <= 120 And Y <= UserControl.Height Then
            ChangingSize = 3
        ElseIf X <= UserControl.Width And Y >= UserControl.Height - 120 Then
            ChangingSize = 5
        ElseIf X <= 120 And Y <= UserControl.Height Then
            ChangingSize = 7
        End If
        If C_Is_Resizable Then
            ClickedX = X
            ClickedY = Y
        Else
            ChangingSize = 0
            If Y < 480 Then
                ClickedX = X
                ClickedY = Y
            End If
        End If
    End If
    LastY = Y
'    If Not IsBig Then
'        nWidth = UserControl.Width
'        nHeight = UserControl.Height
'    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbNormal
    If C_Is_Resizable Then
        If (X <= 120 And Y <= 120) Or (UserControl.Width - X <= 120 And Y >= UserControl.Height - 120) Then
            MousePointer = vbSizeNWSE
        ElseIf (X <= 120 And Y >= UserControl.Height - 120) Or (UserControl.Width - X <= 120 And Y <= 120) Then
            MousePointer = vbSizeNESW
        ElseIf (X <= 120 And Y <= UserControl.Height) Or (UserControl.Width - X <= 120 And Y <= UserControl.Height) Then
            MousePointer = vbSizeWE
        ElseIf (X <= UserControl.Width And Y <= 120) Or (X <= UserControl.Width And Y >= UserControl.Height - 120) Then
            MousePointer = vbSizeNS
        Else
            MousePointer = vbNormal
        End If
    End If
    If (ClickedX <> -1) And (ClickedY <> -1) Then
        Dim w As Integer, H As Integer, L As Integer, T As Integer
        w = UserControl.Width
        H = UserControl.Height
        L = UserControl.Extender.Container.Left
        T = UserControl.Extender.Container.Top
        Select Case ChangingSize
        Case 0
            L = L + X - ClickedX
            T = T + Y - ClickedY
            If C_Can_Move_Smoothly Then
                PUI.MoveSmly UserControl.Extender.Container, L, T, 1
            Else
                UserControl.Extender.Container.Left = L
                UserControl.Extender.Container.Top = T
            End If
'            If (t + y <= 15) And (Not IsBig) Then
'                UserControl.Extender.Container.Left = l
'                UserControl.Extender.Container.Top = t
'                PBBig_Click
'            ElseIf (t + y > 15) And (IsBig) Then
'                PBBig_Click
'            End If
        Case 1
            H = H - Y + ClickedY
            T = T + Y - ClickedY
        Case 2
            w = w + X - ClickedX
            ClickedX = X
            H = H - Y + ClickedY
            T = T + Y - ClickedY
        Case 3
            w = w + X - ClickedX
            ClickedX = X
        Case 4
            w = w + X - ClickedX
            ClickedX = X
            H = H + Y - ClickedY
            ClickedY = Y
        Case 5
            H = H + Y - ClickedY
            ClickedY = Y
        Case 6
            L = L + X - ClickedX
            w = w - X + ClickedX
            H = H + Y - ClickedY
            ClickedY = Y
        Case 7
            L = L + X - ClickedX
            w = w - X + ClickedX
        Case 8
            L = L + X - ClickedX
            w = w - X + ClickedX
            H = H - Y + ClickedY
            T = T + Y - ClickedY
        End Select
        If w < 2355 Then w = 2355
        If H < 720 Then H = 720
        If C_Is_Resizable Then
            UserControl.Width = w
            UserControl.Height = H
            UserControl.Extender.Container.Width = w
            UserControl.Extender.Container.Height = H
            UserControl.Extender.Container.Left = L
            UserControl.Extender.Container.Top = T
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickedX = -1
    ClickedY = -1
    ChangingSize = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set C_Icon = PropBag.ReadProperty("Icon", Nothing)
    Set C_Picture = PropBag.ReadProperty("Picture", Nothing)
    C_Caption = PropBag.ReadProperty("Caption", "PWin8Form")
    C_Color_Border = PropBag.ReadProperty("Color_Border", &HCB8549)
    C_Color_Frame = PropBag.ReadProperty("Color_Frame", &HF4AD6D)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HFFFFFF)
    C_Is_Stretch = PropBag.ReadProperty("Is_Stretch", False)
    C_Can_Move_Smoothly = PropBag.ReadProperty("Can_Move_Smoothly", False)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Has_MinButton = PropBag.ReadProperty("Has_MinButton", True)
    C_Has_MaxButton = PropBag.ReadProperty("Has_MaxButton", True)
    C_Has_CloseButton = PropBag.ReadProperty("Has_CloseButton", True)
    C_Has_Icon = PropBag.ReadProperty("Has_Icon", True)
    C_Is_Resizable = PropBag.ReadProperty("Is_Resizable", True)
    picContainer.BackColor = C_Color_Back
    UserControl.BackColor = C_Color_Frame
    picIcon.BackColor = C_Color_Frame
    PBSmall.Color_Back = C_Color_Frame
    PBBig.Color_Back = C_Color_Frame
    BigBorder.BorderColor = C_Color_Border
    SmallBorder.BorderColor = C_Color_Border
    labTitle = C_Caption
    picContainer.Cls
    If (C_Is_Stretch) And (Not (C_Picture Is Nothing)) Then
        Set Image1.Picture = C_Picture
        picContainer.PaintPicture Image1.Picture, 0, 0, picContainer.Width, picContainer.Height, 0, 0, Image1.Width, Image1.Height
    Else
        Set picContainer.Picture = C_Picture
    End If
    Set imaIcon.Picture = C_Icon
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 2355 Then UserControl.Width = 2355
    If UserControl.Height < 720 Then UserControl.Height = 720
    BigBorder.Width = UserControl.Width
    BigBorder.Height = UserControl.Height
    SmallBorder.Width = UserControl.Width - 210
    SmallBorder.Height = UserControl.Height - 570
    If Not IsBig Then
        picContainer.Left = 120
        picContainer.Width = UserControl.Width - 240
        picContainer.Height = UserControl.Height - 600
        SmallBorder.Visible = True
    Else
        picContainer.Left = 0
        picContainer.Width = UserControl.Width
        picContainer.Height = UserControl.Height - 480
        SmallBorder.Visible = False
    End If
    If Not IsBig Then
        PBClose.Left = UserControl.Width - 120 - 735
    Else
        PBClose.Left = UserControl.Width - 735
    End If
    PBBig.Left = PBClose.Left - PBBig.Width
    PBSmall.Left = PBBig.Left - PBSmall.Width
    If Not IsBig Then
        PBClose.Top = 15
        PBBig.Top = 15
        PBSmall.Top = 15
        BigBorder.Visible = True
    Else
        PBClose.Top = 0
        PBBig.Top = 0
        PBSmall.Top = 0
        BigBorder.Visible = False
    End If
    picIcon.Visible = C_Has_Icon
    If C_Has_MinButton And C_Has_MaxButton And C_Has_CloseButton Then
        PBSmall.Color_Text = &H282828
        PBSmall.Is_Enabled = True
        PBSmall.Visible = True
        PBBig.Color_Text = &H282828
        PBBig.Is_Enabled = True
        PBBig.Visible = True
        PBClose.Color_Text = &HFFFFFF
        PBClose.Is_Enabled = True
        PBClose.Visible = True
    ElseIf C_Has_MinButton And C_Has_MaxButton Then
        PBSmall.Color_Text = &H282828
        PBSmall.Is_Enabled = True
        PBSmall.Visible = True
        PBBig.Color_Text = &H282828
        PBBig.Is_Enabled = True
        PBBig.Visible = True
        PBClose.Color_Text = RGB(69, 110, 145)
        PBClose.Is_Enabled = False
        PBClose.Visible = True
    ElseIf C_Has_MinButton And C_Has_CloseButton Then
        PBSmall.Color_Text = &H282828
        PBSmall.Is_Enabled = True
        PBSmall.Visible = True
        PBBig.Color_Text = RGB(69, 110, 145)
        PBBig.Is_Enabled = False
        PBBig.Visible = True
        PBClose.Color_Text = &HFFFFFF
        PBClose.Is_Enabled = True
        PBClose.Visible = True
    ElseIf C_Has_MaxButton And C_Has_CloseButton Then
        PBSmall.Color_Text = RGB(69, 110, 145)
        PBSmall.Is_Enabled = False
        PBSmall.Visible = True
        PBBig.Color_Text = &H282828
        PBBig.Is_Enabled = True
        PBBig.Visible = True
        PBClose.Color_Text = &HFFFFFF
        PBClose.Is_Enabled = True
        PBClose.Visible = True
    ElseIf C_Has_MinButton Then
        PBSmall.Color_Text = &H282828
        PBSmall.Is_Enabled = True
        PBSmall.Visible = True
        PBBig.Color_Text = RGB(69, 110, 145)
        PBBig.Is_Enabled = False
        PBBig.Visible = True
        PBClose.Color_Text = RGB(69, 110, 145)
        PBClose.Is_Enabled = False
        PBClose.Visible = True
    ElseIf C_Has_MaxButton Then
        PBSmall.Color_Text = RGB(69, 110, 145)
        PBSmall.Is_Enabled = False
        PBSmall.Visible = True
        PBBig.Color_Text = &H282828
        PBBig.Is_Enabled = True
        PBBig.Visible = True
        PBClose.Color_Text = RGB(69, 110, 145)
        PBClose.Is_Enabled = False
        PBClose.Visible = True
    ElseIf C_Has_CloseButton Then
        PBSmall.Color_Text = RGB(69, 110, 145)
        PBSmall.Is_Enabled = False
        PBSmall.Visible = False
        PBBig.Color_Text = RGB(69, 110, 145)
        PBBig.Is_Enabled = False
        PBBig.Visible = False
        PBClose.Color_Text = &HFFFFFF
        PBClose.Is_Enabled = True
        PBClose.Visible = True
    Else
        PBSmall.Color_Text = RGB(69, 110, 145)
        PBSmall.Is_Enabled = False
        PBSmall.Visible = False
        PBBig.Color_Text = RGB(69, 110, 145)
        PBBig.Is_Enabled = False
        PBBig.Visible = False
        PBClose.Color_Text = RGB(69, 110, 145)
        PBClose.Is_Enabled = False
        PBClose.Visible = False
    End If
    If C_Has_Icon Then
        labTitle.Left = 480
    Else
        labTitle.Left = 120
    End If
    If (C_Has_CloseButton = True) And (Not C_Has_MinButton) And (Not C_Has_MaxButton) Then
        labTitle.Width = PBClose.Left - 120 - labTitle.Left
    ElseIf (Not C_Has_CloseButton) And (Not C_Has_MinButton) And (Not C_Has_MaxButton) Then
        labTitle.Width = UserControl.Width - 120 - labTitle.Left
    Else
        labTitle.Width = PBSmall.Left - 120 - labTitle.Left
    End If
    If (C_Is_Stretch) And (Not (C_Picture Is Nothing)) Then
        picContainer.Cls
        Set Image1.Picture = C_Picture
        picContainer.PaintPicture Image1.Picture, 0, 0, picContainer.Width, picContainer.Height, 0, 0, Image1.Width, Image1.Height
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Icon", C_Icon, Nothing)
    Call PropBag.WriteProperty("Picture", C_Picture, Nothing)
    Call PropBag.WriteProperty("Caption", C_Caption, "PWin8Form")
    Call PropBag.WriteProperty("Color_Border", C_Color_Border, &HCB8549)
    Call PropBag.WriteProperty("Color_Frame", C_Color_Frame, &HF4AD6D)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HFFFFFF)
    Call PropBag.WriteProperty("Is_Stretch", C_Is_Stretch, False)
    Call PropBag.WriteProperty("Can_Move_Smoothly", C_Can_Move_Smoothly, False)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Has_MinButton", C_Has_MinButton, True)
    Call PropBag.WriteProperty("Has_MaxButton", C_Has_MaxButton, True)
    Call PropBag.WriteProperty("Has_CloseButton", C_Has_CloseButton, True)
    Call PropBag.WriteProperty("Has_Icon", C_Has_Icon, True)
    Call PropBag.WriteProperty("Is_Resizable", C_Is_Resizable, True)
End Sub
