VERSION 5.00
Begin VB.UserControl PPictureBox 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   ScaleHeight     =   5430
   ScaleWidth      =   6390
   Begin P控件集.PUIMgr PUIMgr1 
      Left            =   3240
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin P控件集.PVScrollBar PV 
      Height          =   5175
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   9128
      Color_Top       =   4210752
      Color_Back      =   8421504
   End
   Begin P控件集.PHScrollBar PH 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      Color_Top       =   4210752
      Color_Back      =   8421504
   End
   Begin VB.PictureBox Pic1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.PictureBox Pic2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   3
         Top             =   0
         Width           =   135
      End
   End
End
Attribute VB_Name = "PPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim C_Color_Top As OLE_COLOR
Dim C_Color_Back As OLE_COLOR
Dim C_Picture As Picture
Dim C_Value_V As Single
Dim C_Value_H As Single
Dim C_Is_Enabled As Boolean

Dim ClickedX As Single
Dim ClickedY As Single

Public Event Scroll()
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub ReSetState()
    PV.Value = 0
    PH.Value = 0
    Pic2.Left = 0
    Pic2.Top = 0
    ClickedX = 0
    ClickedY = 0
    If Pic2.Width > Pic1.Width Then
        PH.Is_Enabled = True
    Else
        PH.Is_Enabled = False
    End If
    If Pic2.Height > Pic1.Height Then
        PV.Is_Enabled = True
    Else
        PV.Is_Enabled = False
    End If
End Sub

Public Property Get Is_Enabled() As Boolean
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Picture() As Picture
    Set Picture = C_Picture
End Property

Public Property Set Picture(ByVal vNewValue As Picture)
    Set C_Picture = vNewValue
    Dim w As Long, H As Long
    w = Pic2.Width
    H = Pic2.Height
    Set Pic2.Picture = C_Picture
    If (Pic2.Width < w) Or (Pic2.Height < H) Then ReSetState
    PropertyChanged "Picture"
End Property

Public Property Get Value_V() As Single
    Value_V = C_Value_V
End Property

Public Property Let Value_V(ByVal vNewValue As Single)
    If vNewValue > 1 Then
        C_Value_V = 1
    ElseIf vNewValue < 0 Then
        C_Value_V = 0
    Else
        C_Value_V = vNewValue
    End If
    PV.Value = C_Value_V
    PV_Scroll C_Value_V
    PropertyChanged "Value_V"
End Property

Public Property Get Value_H() As Single
    Value_H = C_Value_H
End Property

Public Property Let Value_H(ByVal vNewValue As Single)
    If vNewValue > 1 Then
        C_Value_H = 1
    ElseIf vNewValue < 0 Then
        C_Value_H = 0
    Else
        C_Value_H = vNewValue
    End If
    PH.Value = C_Value_H
    PH_Scroll C_Value_H
    PropertyChanged "Value_H"
End Property

Public Property Get Color_Top() As OLE_COLOR
    Color_Top = C_Color_Top
End Property

Public Property Let Color_Top(ByVal vNewValue As OLE_COLOR)
    C_Color_Top = vNewValue
    PV.Color_Top = C_Color_Top
    PH.Color_Top = C_Color_Top
    PropertyChanged "Color_Top"
End Property

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    PV.Color_Back = C_Color_Back
    PH.Color_Back = C_Color_Back
    PropertyChanged "Color_Back"
End Property

Private Sub PV_Click()
    If Is_Enabled = True Then RaiseEvent Click
End Sub

Private Sub PV_DblClick()
    If Is_Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub PV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub PV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub PV_Scroll(NValue As Single)
    PUIMgr1.MoveSmly Pic2, Pic2.Left, -(Pic2.Height - Pic1.Height) * NValue, 1
    If Is_Enabled = True Then RaiseEvent Scroll
End Sub

Private Sub Pic1_Click()
    If Is_Enabled = True Then RaiseEvent Click
End Sub

Private Sub Pic1_DblClick()
    If Is_Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Pic2_Click()
    If Is_Enabled = True Then RaiseEvent Click
End Sub

Private Sub Pic2_DblClick()
    If Is_Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub Pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        If Pic2.Width > Pic1.Width Then ClickedX = X
        If Pic2.Height > Pic1.Height Then ClickedY = Y
    End If
End Sub

Private Sub Pic2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (ClickedX <> 0) And (ClickedY <> 0) Then
        Dim l As Single, T As Single
        If Pic2.Left - ClickedX + X < -(Pic2.Width - Pic1.Width) Then
            l = -(Pic2.Width - Pic1.Width)
        ElseIf (Pic2.Left - ClickedX + X) > 0 Then
            l = 0
        Else
            l = Pic2.Left - ClickedX + X
        End If
        If Pic2.Top - ClickedY + Y < -(Pic2.Height - Pic1.Height) Then
            T = -(Pic2.Height - Pic1.Height)
        ElseIf (Pic2.Top - ClickedY + Y) > 0 Then
            T = 0
        Else
            T = Pic2.Top - ClickedY + Y
        End If
        C_Value_H = (-l) / (Pic2.Width - Pic1.Width)
        PH.Value = C_Value_H
        C_Value_V = (-T) / (Pic2.Height - Pic1.Height)
        PV.Value = C_Value_V
        PUIMgr1.MoveSmly Pic2, l, T, 1
    End If
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Pic2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickedX = 0
    ClickedY = 0
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub PH_Click()
    If Is_Enabled = True Then RaiseEvent Click
End Sub

Private Sub PH_DblClick()
    If Is_Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub PH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub PH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub PH_Scroll(NValue As Single)
    PUIMgr1.MoveSmly Pic2, -(Pic2.Width - Pic1.Width) * NValue, Pic2.Top, 1
    If Is_Enabled = True Then RaiseEvent Scroll
End Sub

Private Sub UserControl_Click()
    If Is_Enabled = True Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If Is_Enabled = True Then RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Set C_Picture = Nothing
    C_Color_Top = &H404040
    C_Color_Back = &H808080
    C_Is_Enabled = True
    C_Value_V = 0
    C_Value_H = 0
    Set Pic2.Picture = C_Picture
    PV.Color_Top = C_Color_Top
    PH.Color_Top = C_Color_Top
    PV.Color_Back = C_Color_Back
    PH.Color_Back = C_Color_Back
    PV.Value = C_Value_V
    PH.Value = C_Value_H
    PH_Scroll C_Value_V
    PV_Scroll C_Value_H
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set C_Picture = PropBag.ReadProperty("Picture", Nothing)
    C_Color_Top = PropBag.ReadProperty("Color_Top", &H404040)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &H808080)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Value_V = PropBag.ReadProperty("Value_V", 0)
    C_Value_H = PropBag.ReadProperty("Value_H", 0)
    Set Pic2.Picture = C_Picture
    PV.Color_Top = C_Color_Top
    PH.Color_Top = C_Color_Top
    PV.Color_Back = C_Color_Back
    PH.Color_Back = C_Color_Back
    ReSetState
    PV.Value = C_Value_V
    PH.Value = C_Value_H
    PH_Scroll C_Value_V
    PV_Scroll C_Value_H
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 1200 Then UserControl.Width = 1200
    If UserControl.Height < 1200 Then UserControl.Height = 1200
    Pic1.Width = UserControl.Width - 255
    Pic1.Height = UserControl.Height - 255
    PV.Height = UserControl.Height - 255
    PV.Left = UserControl.Width - 255
    PH.Width = UserControl.Width - 255
    PH.Top = UserControl.Height - 255
    ReSetState
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", C_Picture, Nothing)
    Call PropBag.WriteProperty("Color_Top", C_Color_Top, &H404040)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &H808080)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Value_V", C_Value_V, 0)
    Call PropBag.WriteProperty("Value_H", C_Value_H, 0)
End Sub
