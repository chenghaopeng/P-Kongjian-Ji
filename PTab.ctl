VERSION 5.00
Begin VB.UserControl PTab 
   BackColor       =   &H00E1A400&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox P2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   600
      ScaleHeight     =   30
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin P控件集.PUIMgr PM 
      Left            =   360
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox p 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label FontTmp 
      BeginProperty Font 
         Name            =   "等线 Light"
         Size            =   11.25
         Charset         =   134
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label l 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "PTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim C_Color_Back As OLE_COLOR
Dim C_Color_Text As OLE_COLOR
Dim C_Picture As Picture
Dim C_Font As Font
Dim C_Is_Enabled As Boolean
Dim C_Distance_Transverse As Integer
Dim C_Distance_Vertical As Integer
Dim C_Color_Selected As OLE_COLOR
Dim C_Color_Selector As OLE_COLOR
Dim C_Color_Selector_Moved As OLE_COLOR
Dim C_Height_Selector As Integer
Dim C_Texts As String
Dim C_Is_AutoDisplay As Boolean
Dim C_Is_AutoUndisplay As Boolean

Private Type PControls
    PControl As Object
    Text As String
    HasControl As Boolean
End Type

Dim Items() As PControls
Dim Total As Integer
Dim Checked As Integer

Dim GoalLeft As Integer
Dim GoalWidth As Integer

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ItemSelected(NewIndex As Integer, LastIndex As Integer)

Private Sub Reset()
    Checked = -1
    Dim s() As String
    s = Split(C_Texts, "|")
    ReDim Items(UBound(s))
    Dim i As Integer
    For i = 0 To UBound(s)
        Items(i).Text = IIf(s(i) <> "", s(i), "?")
        Items(i).HasControl = False
    Next
    Total = UBound(s)
    Refresh
End Sub

Public Sub Refresh()
    Dim i As Integer
    For i = 1 To l.UBound
        Unload l(i)
    Next
    P.BackColor = C_Color_Selector
    P.Height = C_Height_Selector
    l(0).Left = 120
    l(0).Top = 120
    UserControl.BackColor = C_Color_Back
    l(0).ForeColor = C_Color_Text
    Set UserControl.Picture = C_Picture
    Set l(0).Font = C_Font
    If Total >= 0 Then
        P.Visible = True
        l(0).Visible = True
        l(0) = Items(0).Text
        Checked = 0
        If Not Items(0).PControl Is Nothing Then Items(0).PControl.Visible = True
        l(0).ForeColor = C_Color_Selected
        RaiseEvent ItemSelected(0, -1)
        L_MouseDown 0, 1, 0, 0, 0
        If Total = 0 Then Exit Sub
    End If
    For i = 1 To Total
        Load l(i)
        If Not Items(i).PControl Is Nothing Then Items(i).PControl.Visible = False
        l(i).Top = l(i - 1).Top
        l(i).Left = l(i - 1).Left + l(i - 1).Width + C_Distance_Transverse
        l(i) = Items(i).Text
        If l(i).Left > UserControl.Width - l(i).Width Then
            l(i).Top = l(i - 1).Top + l(i).Height + C_Distance_Vertical + P.Height
            l(i).Left = 120
        End If
        l(i).ForeColor = C_Color_Text
        l(i).Visible = True
    Next
    l(0).ForeColor = C_Color_Selected
    Checked = 0
    L_MouseDown 0, 1, 0, 0, 0
End Sub

Public Sub BeRelated(ByRef MyControl As Object)
    Dim i As Integer
    For i = 0 To Total
        If Items(i).HasControl = False Then
            Set Items(i).PControl = MyControl
            Items(i).HasControl = True
            Refresh
            Exit For
        End If
    Next
End Sub

Public Sub BeUnRelated(Index As Integer)
    If (Index < 0) Or (Index > UBound(Items)) Then Exit Sub
    If Not Items(Index).PControl Is Nothing Then Items(Index).PControl.Visible = False
    Set Items(Index).PControl = Nothing
    Items(Index).HasControl = False
    Refresh
End Sub

Public Sub Clear()
    Reset
End Sub

Public Sub GotoIndex(Index As Integer)
    L_MouseDown Index, 1, 0, 0, 0
End Sub

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

Public Property Get Picture() As Picture
    Set Picture = C_Picture
End Property

Public Property Set Picture(ByVal vNewValue As Picture)
    Set C_Picture = vNewValue
    PropertyChanged "Picture"
    Refresh
End Property

Public Property Get Is_Enabled() As Boolean
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Font() As Font
    Set Font = C_Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set C_Font = vNewValue
    Refresh
    PropertyChanged "Font"
End Property

Public Property Get Distance_Transverse() As Integer
    Distance_Transverse = C_Distance_Transverse
End Property

Public Property Let Distance_Transverse(ByVal vNewValue As Integer)
    If vNewValue < 0 Then vNewValue = 0
    C_Distance_Transverse = vNewValue
    PropertyChanged "Distance_Transverse"
    Refresh
End Property

Public Property Get Distance_Vertical() As Integer
    Distance_Vertical = C_Distance_Vertical
End Property

Public Property Let Distance_Vertical(ByVal vNewValue As Integer)
    If vNewValue < 0 Then vNewValue = 0
    C_Distance_Vertical = vNewValue
    PropertyChanged "Distance_Vertical"
    Refresh
End Property

Public Property Get Color_Selected() As OLE_COLOR
    Color_Selected = C_Color_Selected
End Property

Public Property Let Color_Selected(ByVal vNewValue As OLE_COLOR)
    C_Color_Selected = vNewValue
    PropertyChanged "Color_Selected"
    Refresh
End Property

Public Property Get Color_Selector() As OLE_COLOR
    Color_Selector = C_Color_Selector
End Property

Public Property Let Color_Selector(ByVal vNewValue As OLE_COLOR)
    C_Color_Selector = vNewValue
    PropertyChanged "Color_Selector"
    Refresh
End Property

Public Property Get Color_Selector_Moved() As OLE_COLOR
    Color_Selector_Moved = C_Color_Selector_Moved
End Property

Public Property Let Color_Selector_Moved(ByVal vNewValue As OLE_COLOR)
    C_Color_Selector_Moved = vNewValue
    PropertyChanged "Color_Selector_Moved"
    Refresh
End Property

Public Property Get Texts() As String
    Texts = C_Texts
End Property

Public Property Let Texts(ByVal vNewValue As String)
    If vNewValue = "" Then vNewValue = "PTab"
    C_Texts = vNewValue
    Reset
    PropertyChanged "Texts"
End Property

Public Property Get Height_Selector() As Integer
    Height_Selector = C_Height_Selector
End Property

Public Property Let Height_Selector(ByVal vNewValue As Integer)
    C_Height_Selector = vNewValue
    PropertyChanged "Height_Selector"
    Refresh
End Property

Public Property Get Is_AutoDisplay() As Boolean
    Is_AutoDisplay = C_Is_AutoDisplay
End Property

Public Property Let Is_AutoDisplay(ByVal vNewValue As Boolean)
    C_Is_AutoDisplay = vNewValue
    PropertyChanged "Is_AutoDisplay"
End Property

Public Property Get Is_AutoUndisplay() As Boolean
    Is_AutoUndisplay = C_Is_AutoUndisplay
End Property

Public Property Let Is_AutoUndisplay(ByVal vNewValue As Boolean)
    C_Is_AutoUndisplay = vNewValue
    PropertyChanged "Is_AutoUndisplay"
End Property

Private Sub L_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Dim xt As Boolean
    Dim last As Integer
    last = Checked
    If Index = Checked Then xt = True
    If (Not Items(Checked).PControl Is Nothing) And C_Is_AutoUndisplay Then Items(Checked).PControl.Visible = False
    l(Checked).ForeColor = C_Color_Text
    Checked = Index
    If (Not Items(Checked).PControl Is Nothing) And C_Is_AutoDisplay Then Items(Checked).PControl.Visible = True
    l(Checked).ForeColor = C_Color_Selected
    PM.MoveSmly P, l(Index).Left - 60, l(Index).Top + l(Index).Height + 45, 1, 5
    PM.SizeSmly P, l(Index).Width + 120, P.Height, 1, 5
    If Not xt Then RaiseEvent ItemSelected(Index, last)
    P2.Visible = False
End Sub

Private Sub L_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = Checked Then Exit Sub
    P2.BackColor = C_Color_Selector_Moved
    P2.Left = l(Index).Left - 30
    P2.Top = l(Index).Top + l(Index).Height
    P2.Width = l(Index).Width + 60
    P2.Visible = True
    Reload UserControl.hWnd
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    
    C_Color_Back = &HE1A400
    C_Color_Text = &H0&
    Set C_Picture = Nothing
    Set C_Font = FontTmp.Font
    C_Is_Enabled = True
    C_Distance_Transverse = 210
    C_Distance_Vertical = 75
    C_Color_Selected = &HFFFFFF
    C_Color_Selector = &HFFFFFF
    C_Color_Selector_Moved = &H404040
    C_Texts = "PTab"
    C_Height_Selector = 30
    C_Is_AutoDisplay = True
    C_Is_AutoUndisplay = True
    
    Reset

    MLInit UserControl.hWnd
    
    Init UserControl.hWnd
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = -108 Then P2.Visible = False
    If KeyCode = -256 Then
        If Checked > 0 Then L_MouseDown (Checked - 1), 1, 0, 0, 0
    ElseIf KeyCode = -255 Then
        If Checked <> l.UBound Then L_MouseDown (Checked + 1), 1, 0, 0, 0
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    P2.Visible = False
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HE1A400)
    C_Color_Text = PropBag.ReadProperty("Color_Text", &H0&)
    Set C_Picture = PropBag.ReadProperty("Picture", Nothing)
    Set C_Font = PropBag.ReadProperty("Font", FontTmp.Font)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Distance_Transverse = PropBag.ReadProperty("Distance_Transverse", 210)
    C_Distance_Vertical = PropBag.ReadProperty("Distance_Vertical", 75)
    C_Color_Selected = PropBag.ReadProperty("Color_Selected", &HFFFFFF)
    C_Color_Selector = PropBag.ReadProperty("Color_Selector", &HFFFFFF)
    C_Color_Selector_Moved = PropBag.ReadProperty("Color_Selector_Moved", &H404040)
    C_Texts = PropBag.ReadProperty("Texts", "PTab")
    C_Height_Selector = PropBag.ReadProperty("Height_Selector", 30)
    C_Is_AutoDisplay = PropBag.ReadProperty("Is_AutoDisplay", True)
    C_Is_AutoUndisplay = PropBag.ReadProperty("Is_AutoUndisplay", True)
    Reset
End Sub

Private Sub UserControl_Terminate()
    MLTerminate UserControl.hWnd
    Terminate UserControl.hWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HE1A400)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &H0&)
    Call PropBag.WriteProperty("Picture", C_Picture, Nothing)
    Call PropBag.WriteProperty("Font", C_Font, FontTmp.Font)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Distance_Transverse", C_Distance_Transverse, 210)
    Call PropBag.WriteProperty("Distance_Vertical", C_Distance_Vertical, 75)
    Call PropBag.WriteProperty("Color_Selected", C_Color_Selected, &HFFFFFF)
    Call PropBag.WriteProperty("Color_Selector", C_Color_Selector, &HFFFFFF)
    Call PropBag.WriteProperty("Color_Selector_Moved", C_Color_Selector_Moved, &H404040)
    Call PropBag.WriteProperty("Texts", C_Texts, "PTab")
    Call PropBag.WriteProperty("Height_Selector", C_Height_Selector, 30)
    Call PropBag.WriteProperty("Is_AutoDisplay", C_Is_AutoDisplay, True)
    Call PropBag.WriteProperty("Is_AutoUndisplay", C_Is_AutoUndisplay, True)
End Sub
