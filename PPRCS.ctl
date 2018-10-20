VERSION 5.00
Begin VB.UserControl PPRCS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   FillStyle       =   0  'Solid
   ScaleHeight     =   3495
   ScaleWidth      =   4695
   Begin P控件集.PMaths PMaths1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin P控件集.PListBox l 
      Height          =   975
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "PPRCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim C_Resolution As Single
Dim C_Color_Back As OLE_COLOR
Dim C_Color_Top As OLE_COLOR
Dim C_Grid As Boolean
Dim C_Color_Grid As OLE_COLOR

Dim w As Integer
Dim H As Integer

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Property Get Color_Top() As OLE_COLOR
    Color_Top = C_Color_Top
End Property

Public Property Let Color_Top(ByVal vNewValue As OLE_COLOR)
    C_Color_Top = vNewValue
    UserControl.ForeColor = C_Color_Top
    ReDraw
    PropertyChanged "Color_Top"
End Property

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    UserControl.BackColor = C_Color_Back
    ReDraw
    PropertyChanged "Color_Back"
End Property

Public Property Get Resolution() As Single
    Resolution = C_Resolution
    ReDraw
End Property

Public Property Let Resolution(ByVal vNewValue As Single)
    If vNewValue > 100 Then
        C_Resolution = 100
    ElseIf vNewValue < -100 Then
        C_Resolution = -100
    Else
        C_Resolution = vNewValue
    End If
    PropertyChanged "Resolution"
    ReDraw
End Property

Public Property Get Grid() As Boolean
    Grid = C_Grid
End Property

Public Property Let Grid(ByVal vNewValue As Boolean)
    C_Grid = vNewValue
    PropertyChanged "Grid"
    ReDraw
End Property

Public Property Get Color_Grid() As OLE_COLOR
    Color_Grid = C_Color_Grid
End Property

Public Property Let Color_Grid(ByVal vNewValue As OLE_COLOR)
    C_Color_Grid = vNewValue
    ReDraw
    PropertyChanged "Color_Grid"
End Property

Public Sub ReDraw()
    w = UserControl.Width
    H = UserControl.Height
    UserControl.Cls
    Dim x As Long, y As Long, i  As Long
    UserControl.Line (w / 2, 0)-(w / 2, H)
    UserControl.Line (0, H / 2)-(w, H / 2)
    PrintText "0", 0, w / 2 + 75, H / 2 - 210
    For i = 1 To 5
        If C_Grid Then
            ForeColor = C_Color_Grid
            Line (w / 2 / 6 * i, 0)-(w / 2 / 6 * i, H)
            Line (w / 2 + w / 2 / 6 * i, 0)-(w / 2 + w / 2 / 6 * i, H)
            Line (0, H / 2 / 6 * i)-(w, H / 2 / 6 * i)
            Line (0, H / 2 + H / 2 / 6 * i)-(w, H / 2 + H / 2 / 6 * i)
        End If
        
        ForeColor = C_Color_Top
        Line (w / 2 / 6 * i, H / 2 - 60)-(w / 2 / 6 * i, H / 2)
        PrintText (i - 6) * C_Resolution, 6, w / 2 / 6 * i, H / 2 - 60
        Line (w / 2 + w / 2 / 6 * i, H / 2 - 60)-(w / 2 + w / 2 / 6 * i, H / 2)
        PrintText i * C_Resolution, 6, w / 2 + w / 2 / 6 * i - 75, H / 2 - 60
        Line (w / 2 - 60, H / 2 / 6 * i)-(w / 2, H / 2 / 6 * i)
        PrintText (6 - i) * C_Resolution, 7, w / 2 - 180, H / 2 / 6 * i
        Line (w / 2 - 60, H / 2 + H / 2 / 6 * i)-(w / 2, H / 2 + H / 2 / 6 * i)
        PrintText -i * C_Resolution, 7, w / 2 - 75, H / 2 + H / 2 / 6 * i
    Next
    Dim s() As String, q() As String
    For i = 0 To l.ListCount - 1
        s = Split(l.List(i), ":")
        q = Split(s(1), ",")
        Select Case s(0)
        Case "Fun"
            DrawFunction q(0)
        Case "Line"
            Line2Points Val(q(0)), Val(q(1)), Val(q(2)), Val(q(3))
        Case "LineX"
            DrawLineX Val(q(0))
        Case "LineY"
            DrawLineY Val(q(0))
        Case "Point"
            DrawPoint Val(q(0)), Val(q(1)), q(2)
        End Select
    Next
End Sub

Public Sub DrawFunction(strFun As String)
    If l.ItemIsExists("Fun:" & strFun) = False Then l.AddItem "Fun:" & strFun
    Dim f As String, s As String, x As Single, y As Single, lx As Single, ly As Single, i As Integer
    lx = -15
    ly = -15
    s = FormatFunction(strFun)
    For i = 15 To w Step 15
        f = s
        x = ((6 - i / (w / 12)) * C_Resolution)
        f = Replace(f, "x", x)
        y = PMaths1.VBCodetoNum(f)
        If (lx <> -15) And (ly <> -15) Then LP x, y, lx, ly
        lx = x
        ly = y
        'DoEvents
    Next
End Sub

Public Sub Line2Points(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
    If l.ItemIsExists("Line:" & x1 & "," & y1 & "," & x2 & "," & y2) = False Then l.AddItem "Line:" & x1 & "," & y1 & "," & x2 & "," & y2
    Line ((6 * C_Resolution + x1) / C_Resolution * (w / 12), (6 * C_Resolution - y1) / C_Resolution * (H / 12))-((6 * C_Resolution + x2) / C_Resolution * (w / 12), (6 * C_Resolution - y2) / C_Resolution * (H / 12))
End Sub

Public Sub DrawLineX(x As Single)
    If l.ItemIsExists("LineX:" & x) = False Then l.AddItem "LineX:" & x
    Dim t As Single
    t = (6 * C_Resolution - x) / C_Resolution * (H / 12)
    Line (t, 0)-(t, H)
End Sub

Public Sub DrawLineY(y As Single)
    If l.ItemIsExists("LineY:" & y) = False Then l.AddItem "LineY:" & y
    Dim t As Single
    t = (6 * C_Resolution - y) / C_Resolution * (w / 12)
    Line (0, t)-(w, t)
End Sub

Public Sub DrawPoint(x As Single, y As Single, strText As String)
    If l.ItemIsExists("Point:" & x & "," & y & "," & strText) = False Then l.AddItem "Point:" & x & "," & y & "," & strText
    UserControl.Circle ((6 * C_Resolution + x) / C_Resolution * (w / 12), (6 * C_Resolution - y) / C_Resolution * (H / 12)), 30
    PrintText strText, 3, (6 * C_Resolution + x) / C_Resolution * (w / 12), (6 * C_Resolution - y) / C_Resolution * (H / 12)
End Sub

Public Sub LoadTxt(strPath As String)
    l.ReadFile strPath, True
    ReDraw
End Sub

Public Sub SaveTxt(strPath As String)
    l.SaveAllItems strPath, True
End Sub

Public Sub SavePic(strPath As String)
    SavePicture UserControl.Image, strPath
End Sub

Public Sub Clear()
    l.Clear
    ReDraw
End Sub

Private Function FormatFunction(ByVal strFun As String) As String
    Dim i As Integer, s As String, f() As String
    f = Split(strFun, "x")
    If Left(strFun, 1) <> "x" Then
        s = ""
        For i = 0 To UBound(f) - 1
            s = s & f(i)
            If Right(s, 1) = "/" Then
                s = s & "x"
            ElseIf (Right(s, 1) >= "0") And (Right(s, 1) <= "9") Then
                   s = s & "*(x)"
            ElseIf Right(s, 1) = ")" Then
                   s = s & "*(x)"
            Else
                s = s & "(x)"
            End If
        Next
    Else
        s = "(x)"
        For i = 1 To UBound(f) - 1
            s = s & f(i)
            If Right(s, 1) = "/" Then
                s = s & "x"
            ElseIf (Right(s, 1) >= "0") And (Right(s, 1) <= "9") Then
                   s = s & "*(x)"
            ElseIf Right(s, 1) = ")" Then
                   s = s & "*(x)"
            Else
                s = s & "(x)"
            End If
        Next
    End If
    FormatFunction = s & f(UBound(f))
End Function

Private Sub LP(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
    On Error Resume Next
    Line ((6 * C_Resolution + x1) / C_Resolution * (w / 12), (6 * C_Resolution - y1) / C_Resolution * (H / 12))-((6 * C_Resolution + x2) / C_Resolution * (w / 12), (6 * C_Resolution - y2) / C_Resolution * (H / 12))
End Sub

Private Sub PrintText(txt As Variant, lx As Integer, x As Integer, y As Integer)
    Label1 = txt
    UserControl.CurrentX = x
    UserControl.CurrentY = y
    Select Case lx
    Case 1 '右下角
        UserControl.CurrentX = x
        UserControl.CurrentY = y
    Case 2 '左下角
        UserControl.CurrentX = x - Label1.Width
        UserControl.CurrentY = y
    Case 3 '左上角
        UserControl.CurrentX = x - Label1.Width
        UserControl.CurrentY = y - Label1.Height
    Case 4 '右上角
        UserControl.CurrentX = x
        UserControl.CurrentY = y - Label1.Height
    Case 5 '正下居中
        UserControl.CurrentX = x - Label1.Width / 2
        UserControl.CurrentY = y
    Case 6 '正上居中
        UserControl.CurrentX = x - Label1.Width / 2
        UserControl.CurrentY = y - Label1.Height
    Case 7 '左中间
        UserControl.CurrentX = x - Label1.Width
        UserControl.CurrentY = y - Label1.Height / 2
    Case 8 '右中间
        UserControl.CurrentX = x
        UserControl.CurrentY = y - Label1.Height / 2
    End Select
    Print txt
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    C_Resolution = 1
    C_Color_Top = &H80000012
    C_Color_Back = &HFFFFFF
    C_Grid = True
    C_Color_Grid = &HE0E0E0
    UserControl.ForeColor = C_Color_Top
    UserControl.BackColor = C_Color_Back
    ReDraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Resolution = PropBag.ReadProperty("Resolution", 1)
    C_Color_Top = PropBag.ReadProperty("Color_Top", &H80000012)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HFFFFFF)
    C_Grid = PropBag.ReadProperty("Grid", True)
    C_Color_Grid = PropBag.ReadProperty("Color_Grid", &HE0E0E0)
    UserControl.ForeColor = C_Color_Top
    UserControl.BackColor = C_Color_Back
    ReDraw
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 2535 Then UserControl.Width = 2535
    If UserControl.Height < 2535 Then UserControl.Height = 2535
    ReDraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Resolution", C_Resolution, 1)
    Call PropBag.WriteProperty("Color_Top", C_Color_Top, &H80000012)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HFFFFFF)
    Call PropBag.WriteProperty("Grid", C_Grid, True)
    Call PropBag.WriteProperty("Color_Grid", C_Color_Grid, &HE0E0E0)
End Sub
