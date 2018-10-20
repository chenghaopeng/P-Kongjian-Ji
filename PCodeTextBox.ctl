VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl PCodeTextBox 
   BackColor       =   &H00808080&
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ScaleHeight     =   3495
   ScaleWidth      =   4815
   Begin VB.PictureBox Copying 
      BackColor       =   &H00FFFFCC&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   3375
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin P控件集.PProgressBar CopyingPP 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   53
      End
      Begin VB.Label CopyingL 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在粘贴，请稍候.."
         BeginProperty Font 
            Name            =   "等线 Light"
            Size            =   15.75
            Charset         =   134
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2670
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "等线 Light"
         Size            =   10.5
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "等线 Light"
         Size            =   10.5
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin RichTextLib.RichTextBox VBRule1 
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"PCodeTextBox.ctx":0000
   End
   Begin VB.PictureBox GDT 
      BackColor       =   &H00FF7402&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   4680
      ScaleHeight     =   2175
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   600
   End
   Begin P控件集.PUIMgr PUI 
      Left            =   1200
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin RichTextLib.RichTextBox TextTmp 
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"PCodeTextBox.ctx":0276
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "等线 Light"
         Size            =   10.5
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin P控件集.PUIMgr PUI2 
      Left            =   1680
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin RichTextLib.RichTextBox VBRule2 
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"PCodeTextBox.ctx":0309
   End
   Begin RichTextLib.RichTextBox VBRule3 
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"PCodeTextBox.ctx":03B5
   End
   Begin VB.PictureBox dmrq 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin RichTextLib.RichTextBox R 
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"PCodeTextBox.ctx":0452
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "等线 Light"
            Size            =   10.5
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   360
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Label LN 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "等线 Light"
            Size            =   9
            Charset         =   134
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   0
         Width           =   45
      End
   End
End
Attribute VB_Name = "PCodeTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim GDTMouseDown As Boolean, ClickedY As Single

Dim C_Color_Back As OLE_COLOR
Dim C_Color_Back_Editing As OLE_COLOR
Dim C_Color_Back_Moved As OLE_COLOR
Dim C_Color_Back_Number As OLE_COLOR
Dim C_Color_Fore_Number As OLE_COLOR
Dim C_Color_Text_Number As OLE_COLOR
Dim C_Color_Text_Common As OLE_COLOR
Dim C_Rule_1 As String
Dim C_Rule_2 As String
Dim C_Rule_3 As String
Dim C_Font_Text_Common As Font
Dim C_Font_Number As Font
Dim C_Line_Height As Integer

Private Type Rule1
    Key As String
    Yanse As OLE_COLOR
    Jiacu As Boolean
    Xieti As Boolean
End Type
Dim R1() As Rule1

Private Type Rule2
    Key1 As String
    Key2 As String
    Yanse As OLE_COLOR
    Jiacu As Boolean
    Xieti As Boolean
End Type
Dim R2() As Rule2

Private Type Rule3
    Key1 As String
    Key2 As String
    Yanse As OLE_COLOR
    Jiacu As Boolean
    Xieti As Boolean
End Type
Dim R3 As Rule3

Private Sub Refresh()
    Dim i As Integer
    Dim s() As String, T() As String, tt() As String
    s = Split(C_Rule_1, "|")
    ReDim R1(UBound(s))
    For i = 0 To UBound(s)
        T = Split(s(i), ":")
        R1(i).Key = "," & T(0) & ","
        tt = Split(T(1), ",")
        R1(i).Yanse = Val(tt(0))
        R1(i).Jiacu = Val(tt(1))
        R1(i).Xieti = Val(tt(2))
    Next
    s = Split(C_Rule_2, "|")
    ReDim R2(UBound(s))
    For i = 0 To UBound(s)
        T = Split(s(i), ":")
        R2(i).Key1 = Left(T(0), InStr(T(0), ",") - 1)
        R2(i).Key2 = Right(T(0), Len(T(0)) - InStr(T(0), ","))
        tt = Split(T(1), ",")
        R2(i).Yanse = Val(tt(0))
        R2(i).Jiacu = Val(tt(1))
        R2(i).Xieti = Val(tt(2))
    Next
    If C_Rule_3 = "" Then
        R3.Key1 = "none"
    Else
        T = Split(C_Rule_3, ":")
        R3.Key1 = Left(T(0), InStr(T(0), ",") - 1)
        R3.Key2 = Right(T(0), Len(T(0)) - InStr(T(0), ","))
        tt = Split(T(1), ",")
        R3.Yanse = Val(tt(0))
        R3.Jiacu = Val(tt(1))
        R3.Xieti = Val(tt(2))
    End If
    
    Set Text1.Font = C_Font_Number
    Set Text2.Font = C_Font_Text_Common
    UserControl.BackColor = C_Color_Back
    dmrq.BackColor = C_Color_Back_Number
    
    For i = 0 To R.UBound
        R(i).BackColor = C_Color_Back
        R(i).Tag = ""
        Set R(i).Font = C_Font_Text_Common
        R(i).Height = C_Line_Height
        R(i).Top = i * C_Line_Height
    Next
    ToVB6Code 0
    For i = 1 To R.UBound
        If R(i).Tag = "" Then
            ToVB6Code i
        End If
    Next
    For i = 0 To LN.UBound
        LN(i).ForeColor = C_Color_Fore_Number
    Next
    
    
    RefreshLineNumber 0
    'R(0).SetFocus
    R(0).BackColor = C_Color_Back_Editing
End Sub

Private Function Log10(ByVal a As Double) As Integer
    Log10 = Int(Log(a) / Log(10))
End Function

Private Function Rule2_Find(ByVal a As String) As Integer
    Rule2_Find = -1
    Dim i As Integer
    For i = 0 To UBound(R2)
        If R2(i).Key1 = a Then
            Rule2_Find = i
            Exit Function
        End If
    Next
End Function

Private Sub ToVB6Code(ByVal Index As Integer)
    If R(Index).Text = "" Then Exit Sub
    Set TextTmp.Font = Text2.Font
    TextTmp.Text = R(Index).Text
    TextTmp.SelStart = 0
    TextTmp.SelLength = Len(TextTmp.Text)
    TextTmp.SelColor = C_Color_Text_Common
    Dim hp As Integer, T As String, L As Integer, P As Long, a As String, s As Integer, i As Integer, b As Integer
    hp = R(Index).SelStart: T = LCase(TextTmp.Text): L = Len(T): P = 0
    Do Until P >= L
        P = P + 1
        a = Mid(T, P, 1)
        s = P
        TextTmp.SelStart = s - 1
        If a >= "a" And a <= "z" Then
            For i = s + 1 To L
                a = Mid(T, i, 1)
                If (a < "a" Or a > "z") And (a < "0" Or a > "9") And a <> "_" Then
                    P = i - 1
                    Exit For
                End If
                If i = L Then P = L + 1
            Next
            a = Mid(T, s, P - s + 1)
            For i = 0 To UBound(R1)
                If InStr(R1(i).Key, "," & a & ",") <> 0 Then
                    TextTmp.SelLength = Len(a)
                    TextTmp.SelColor = R1(i).Yanse
                    TextTmp.SelBold = R1(i).Jiacu
                    TextTmp.SelItalic = R1(i).Xieti
                    Exit For
                End If
            Next
        ElseIf a >= "0" And a <= "9" Then
            If s <> 1 Then If Mid(T, s - 1, 1) = "." Then s = s - 1
            For i = s + 1 To L
                a = Mid(T, i, 1)
                If (a < "0" Or a > "9") And (a <> ".") Then
                    P = i - 1
                    Exit For
                End If
                If i = L Then P = L + 1
            Next
            TextTmp.SelLength = P - s + 1
            TextTmp.SelColor = C_Color_Text_Number
            TextTmp.SelBold = Text1.FontBold
            TextTmp.SelItalic = Text1.FontItalic
        ElseIf Rule2_Find(a) > -1 Then
            b = Rule2_Find(a)
            If R2(b).Key2 = "all" Then
                P = L + 1
            Else
                P = P + 1
                a = Mid(T, P, 1)
                Do Until a = R2(b).Key2
                    P = P + 1
                    a = Mid(T, P, 1)
                    If P > L Then Exit Do
                Loop
            End If
            TextTmp.SelLength = P - s + 1
            TextTmp.SelColor = R2(b).Yanse
            TextTmp.SelBold = R2(b).Jiacu
            TextTmp.SelItalic = R2(b).Xieti
            If P > L Then Exit Do
        ElseIf a = R3.Key1 Then
            P = L + 1
            TextTmp.SelLength = P - s
            TextTmp.SelColor = R3.Yanse
            TextTmp.SelBold = R3.Jiacu
            TextTmp.SelItalic = R3.Xieti
            R(Index).Tag = "1"
            For i = Index + 1 To R.UBound
                R(i).Tag = "1"
                R(i).SelStart = 0
                R(i).SelLength = Len(R(i).Text)
                R(i).SelColor = R3.Yanse
                R(i).SelBold = R3.Jiacu
                R(i).SelItalic = R3.Xieti
                If InStr(R(i), R3.Key2) <> 0 Then
                    Exit For
                End If
            Next
            Exit Do
        End If
    Loop
    R(Index).TextRTF = TextTmp.TextRTF
    R(Index).SelStart = hp
End Sub

Private Sub RefreshLineNumber(Optional ByVal Index As Integer = -1)
    Dim i As Integer
    LN(LN.UBound).AutoSize = False
    LN(LN.UBound).AutoSize = True
    LN(LN.UBound) = "   " & LN.UBound + 1 & " "
    LN(LN.UBound).Height = R(R.UBound).Height
    LN(LN.UBound).Left = 0
    R(R.UBound).Left = LN(LN.UBound).Width + 15
    R(R.UBound).Width = dmrq.Width - R(R.UBound).Left
    For i = R.UBound - 1 To 0 Step -1
        LN(i) = "   " & i + 1 & " "
        LN(i).Height = R(i).Height
        LN(i).Width = LN(LN.UBound).Width
        R(i).Left = LN(LN.UBound).Width + 15
        LN(i).Left = 0
        R(i).Width = R(R.UBound).Width
    Next
    dmrq.Height = R(R.UBound).Height * (R.UBound + 1)
    Line1.x1 = LN(LN.UBound).Width
    Line1.y1 = 0
    Line1.x2 = LN(LN.UBound).Width
    Line1.y2 = dmrq.Height
    RefreshContainerPos Index
End Sub

Private Sub RefreshContainerPos(Optional ByVal Index As Integer = -1)
    If Index = -1 Then Exit Sub
    If (R(0).Height * (R.UBound + 1)) <= UserControl.Height Then
        ChangeScrollBarValue 0
        GDT.Visible = False
        PUI.MoveSmly dmrq, dmrq.Left, 0, 1
        Exit Sub
    End If
    If R(Index).Top < -dmrq.Top Then
        PUI.MoveSmly dmrq, dmrq.Left, -R(Index).Top, 1
        PUI2.MoveSmly GDT, GDT.Left, R(Index).Top / (dmrq.Height - UserControl.Height) * (UserControl.Height - GDT.Height), 1
    ElseIf R(Index).Top > UserControl.Height - dmrq.Top - R(Index).Height Then
        PUI.MoveSmly dmrq, dmrq.Left, UserControl.Height - R(Index).Top - R(Index).Height, 1
        PUI2.MoveSmly GDT, GDT.Left, -(UserControl.Height - R(Index).Top - R(Index).Height) / (dmrq.Height - UserControl.Height) * (UserControl.Height - GDT.Height), 1
    End If
    If dmrq.Height + dmrq.Top < UserControl.Height And dmrq.Height > UserControl.Height Then
        PUI.MoveSmly dmrq, dmrq.Left, UserControl.Height - dmrq.Height, 1
    End If
End Sub

Private Sub ChangeScrollBarValue(ByVal NewValue As Single)
    If Not GDT.Visible Then Exit Sub
    GDT.Top = (UserControl.Height - GDT.Height) * NewValue
    dmrq.Top = -NewValue * (dmrq.Height - UserControl.Height)
End Sub

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    PropertyChanged "Color_Back"
    Refresh
End Property

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back_Editing(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Editing = vNewValue
    PropertyChanged "Color_Back_Editing"
    Refresh
End Property

Public Property Get Color_Back_Editing() As OLE_COLOR
    Color_Back_Editing = C_Color_Back_Editing
End Property

Public Property Let Color_Back_Moved(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Moved = vNewValue
    PropertyChanged "Color_Back_Moved"
    Refresh
End Property

Public Property Get Color_Back_Moved() As OLE_COLOR
    Color_Back_Moved = C_Color_Back_Moved
End Property

Public Property Let Color_Back_Number(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Number = vNewValue
    PropertyChanged "Color_Back_Number"
    Refresh
End Property

Public Property Get Color_Back_Number() As OLE_COLOR
    Color_Back_Number = C_Color_Back_Number
End Property

Public Property Let Color_Fore_Number(ByVal vNewValue As OLE_COLOR)
    C_Color_Fore_Number = vNewValue
    PropertyChanged "Color_Fore_Number"
    Refresh
End Property

Public Property Get Color_Fore_Number() As OLE_COLOR
    Color_Fore_Number = C_Color_Fore_Number
End Property

Public Property Let Color_Text_Number(ByVal vNewValue As OLE_COLOR)
    C_Color_Text_Number = vNewValue
    PropertyChanged "Color_Text_Number"
    Refresh
End Property

Public Property Get Color_Text_Number() As OLE_COLOR
    Color_Text_Number = C_Color_Text_Number
End Property

Public Property Let Color_Text_Common(ByVal vNewValue As OLE_COLOR)
    C_Color_Text_Common = vNewValue
    PropertyChanged "Color_Text_Common"
    Refresh
End Property

Public Property Get Color_Text_Common() As OLE_COLOR
    Color_Text_Common = C_Color_Text_Common
End Property

Public Property Let Rule_1(ByVal vNewValue As String)
    C_Rule_1 = vNewValue
    PropertyChanged "Rule_1"
    Refresh
End Property

Public Property Get Rule_1() As String
    Rule_1 = C_Rule_1
End Property

Public Property Let Rule_2(ByVal vNewValue As String)
    C_Rule_2 = vNewValue
    PropertyChanged "Rule_2"
    Refresh
End Property

Public Property Get Rule_2() As String
    Rule_2 = C_Rule_2
End Property

Public Property Let Rule_3(ByVal vNewValue As String)
    C_Rule_3 = vNewValue
    PropertyChanged "Rule_3"
    Refresh
End Property

Public Property Get Rule_3() As String
    Rule_3 = C_Rule_3
End Property

Public Property Set Font_Text_Common(ByVal vNewValue As Font)
    Set C_Font_Text_Common = vNewValue
    PropertyChanged "Font_Text_Common"
    Refresh
End Property

Public Property Get Font_Text_Common() As Font
    Set Font_Text_Common = C_Font_Text_Common
End Property

Public Property Set Font_Number(ByVal vNewValue As Font)
    Set C_Font_Number = vNewValue
    PropertyChanged "Font_Number"
    Refresh
End Property

Public Property Get Font_Number() As Font
    Set Font_Number = C_Font_Number
End Property

Public Property Let Line_Height(ByVal vNewValue As Integer)
    C_Line_Height = vNewValue
    PropertyChanged "Line_Height"
    Refresh
End Property

Public Property Get Line_Height() As Integer
    Line_Height = C_Line_Height
End Property

Private Sub dmrq_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not GDT.Visible Then
        If (R(0).Height * (R.UBound + 1)) > UserControl.Height Then
            GDT.Visible = True
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub GDT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GDTMouseDown = True
    ClickedY = Y
End Sub

Private Sub GDT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not GDTMouseDown Then Exit Sub
    Dim T As Single
    T = GDT.Top + Y - ClickedY
    If T < 0 Then T = 0
    If T > UserControl.Height - GDT.Height Then T = UserControl.Height - GDT.Height
    PUI.MoveSmly dmrq, dmrq.Left, T / (GDT.Height - UserControl.Height) * (dmrq.Height - UserControl.Height), 1
    PUI2.MoveSmly GDT, GDT.Left, T, 1
End Sub

Private Sub GDT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GDTMouseDown = False
End Sub

Private Sub R_Change(Index As Integer)
    If R(Index).Tag = "1" Then
        If InStr(R(Index).Text, R3.Key1) = 0 Then
            If Index = 0 Or R(Index - 1).Tag = "" Then
                R(Index).Tag = ""
                ToVB6Code Index
                Dim i As Integer
                For i = Index + 1 To R.UBound
                    If R(i).Tag = "" Then Exit For
                    If InStr(R(Index).Text, R3.Key1) = 0 Then
                        R(i).Tag = ""
                        ToVB6Code i
                    End If
                Next
            ElseIf Index <> R.UBound Then
                If R(Index + 1).Tag = "" And InStr(R(Index).Text, R3.Key2) = 0 Then
                    For i = Index + 1 To R.UBound
                        R(i).Tag = "1"
                        R(i).SelStart = 0
                        R(i).SelLength = Len(R(i).Text)
                        R(i).SelColor = R3.Yanse
                        R(i).SelBold = R3.Jiacu
                        R(i).SelItalic = R3.Xieti
                        If InStr(R(i), R3.Key2) <> 0 Then
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Else
        ToVB6Code Index
    End If
End Sub

Private Sub R_GotFocus(Index As Integer)
    R(Index).BackColor = C_Color_Back_Editing
End Sub

Private Sub R_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = R(Index).SelStart
    If KeyCode = 13 Then
        Dim T As String
        T = Mid(R(Index).Text, R(Index).SelStart + 1, Len(R(Index).Text) - R(Index).SelStart)
        If R(Index).SelStart <> 0 Then
            R(Index).Text = Left(R(Index).Text, R(Index).SelStart)
        Else
            R(Index).Text = ""
        End If
        Load R(R.UBound + 1)
        Load LN(LN.UBound + 1)
        With R(R.UBound)
            .Top = R(R.UBound - 1).Top + .Height
            dmrq.Height = .Height * (R.UBound + 1)
            .Text = ""
            .BackColor = C_Color_Back
        End With
        With LN(LN.UBound)
            .Visible = True
            .Top = R(LN.UBound - 1).Top + .Height
        End With
        RefreshLineNumber Index + 1
        For i = R.UBound To Index + 2 Step -1
            R(i).TextRTF = R(i - 1).TextRTF
        Next
        R(R.UBound).Visible = True
        R(Index + 1).Text = T
        R(Index + 1).SetFocus
    ElseIf KeyCode = 8 And R(Index).SelStart = 0 Then
        If Index <> 0 Then
            Dim P As Integer
            P = Len(R(Index - 1).Text)
            R(Index - 1).Text = R(Index - 1).Text & R(Index).Text
            For i = Index To R.UBound - 1
                R(i).TextRTF = R(i + 1).TextRTF
            Next
            R(Index - 1).SetFocus
            R(Index - 1).SelStart = P
            Unload R(R.UBound)
            Unload LN(LN.UBound)
            RefreshLineNumber Index - 1
        End If
    ElseIf KeyCode = 38 Then
        If Index <> 0 Then
            If i <= Len(R(Index - 1).Text) Then
                R(Index - 1).SetFocus
                R(Index - 1).SelStart = i
            Else
                R(Index - 1).SetFocus
                R(Index - 1).SelStart = Len(R(Index - 1).Text)
            End If
            RefreshContainerPos Index - 1
        End If
    ElseIf KeyCode = 40 Then
        If Index <> R.UBound Then
            If i <= Len(R(Index + 1).Text) Then
                R(Index + 1).SetFocus
                R(Index + 1).SelStart = i
            Else
                R(Index + 1).SetFocus
                R(Index + 1).SelStart = Len(R(Index - 1).Text)
            End If
            RefreshContainerPos Index + 1
        End If
    ElseIf KeyCode = 37 And i = 0 Then
        If Index <> 0 Then
            R(Index - 1).SetFocus
            R(Index - 1).SelStart = Len(R(Index - 1).Text)
            RefreshContainerPos Index - 1
        End If
    ElseIf KeyCode = 39 And i = Len(R(Index).Text) Then
        If Index <> R.UBound Then
            R(Index + 1).SetFocus
            R(Index + 1).SelStart = 0
            RefreshContainerPos Index + 1
        End If
    End If
End Sub

Private Sub R_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        Dim strText As String, i As Integer, j As Integer, T As String, a As Integer, b As Integer, c As Integer, d As Integer
        strText = Clipboard.GetText
        Dim s() As String
        s = Split(strText, vbCrLf)
        Copying.Height = UserControl.Height
        Copying.Width = UserControl.Width
        CopyingL.Left = (UserControl.Width - CopyingL.Width) \ 2
        CopyingL.Top = (UserControl.Height - CopyingL.Height - 30) \ 2
        CopyingL = "正在加载新代码行"
        CopyingPP.Width = UserControl.Width
        CopyingPP.Top = UserControl.Height - 30
        Copying.Visible = True
        DoEvents
        Dim Total As Integer
        Total = UBound(s)

        T = Mid(R(Index).Text, R(Index).SelStart + 1, Len(R(Index).Text) - R(Index).SelStart)
        If R(Index).SelStart <> 0 Then
            R(Index).Text = Left(R(Index).Text, R(Index).SelStart)
        Else
            R(Index).Text = ""
        End If
        
        a = Index + 1
        b = R.UBound
        c = Total + Total + b - a + 1
        d = 0
        For i = 1 To Total
            Load R(R.UBound + 1)
            Load LN(LN.UBound + 1)
            With R(R.UBound)
                .Visible = True
                .Top = R(R.UBound - 1).Top + .Height
                dmrq.Height = .Height * (R.UBound + 1)
                .Text = ""
                .BackColor = C_Color_Back
            End With
            With LN(LN.UBound)
                .Visible = True
                .Top = R(LN.UBound - 1).Top + .Height
            End With
            d = d + 1
            CopyingPP.Value = d / c
            DoEvents
        Next
        CopyingL = "正在移动新代码行"
        For i = b To a Step -1
            R(R.UBound - (b - i)).TextRTF = R(i).TextRTF
            If R.UBound - (b - i) <> i Then R(i).Text = ""
            d = d + 1
            CopyingPP.Value = d / c
            DoEvents
        Next
        CopyingL = "正在显示新代码"
        For i = a To a + Total - 1
            R(i).Text = s(i - a + 1)
            d = d + 1
            CopyingPP.Value = d / c
            DoEvents
        Next
        CopyingL = "正在处理"
        R(a + Total - 1).SetFocus
        R(a + Total - 1).Text = R(a + Total - 1).Text & T
        R(a + Total - 1).SelStart = Len(R(a + Total - 1).Text) - Len(T)
        RefreshLineNumber R.UBound - 1
        Copying.Visible = False
    End If
End Sub

Private Sub R_LostFocus(Index As Integer)
    R(Index).SelStart = 0
    R(Index).BackColor = C_Color_Back
End Sub

Private Sub R_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not GDT.Visible Then
        If (R(0).Height * (R.UBound + 1)) > UserControl.Height Then
            GDT.Visible = True
            Timer1.Enabled = True
        End If
    End If
    Dim i As Integer
    For i = 0 To R.UBound
        If R(i).BackColor = C_Color_Back_Moved Then R(i).BackColor = C_Color_Back
    Next
    If R(Index).BackColor <> C_Color_Back_Editing Then R(Index).BackColor = C_Color_Back_Moved
End Sub

Private Sub Timer1_Timer()
    If GDTMouseDown Then Exit Sub
    Dim hwn As Long
    hwn = GetPointhWnd
    If hwn = UserControl.hWnd Then Exit Sub
    If hwn = dmrq.hWnd Then Exit Sub
    If hwn = GDT.hWnd Then Exit Sub
    Dim i As Integer
    For i = 0 To R.UBound
        If hwn = R(i).hWnd Then Exit Sub
    Next
    GDT.Visible = False
    Timer1.Enabled = False
End Sub

Private Sub UserControl_Initialize()
    C_Color_Back = &HFFFFFF
    C_Color_Back_Editing = &HFFFFCC
    C_Color_Back_Moved = &HFFFFDF
    C_Color_Back_Number = &HF0F0F0
    C_Color_Fore_Number = &H0
    C_Color_Text_Number = &HFF00FF
    C_Color_Text_Common = &H0
    C_Rule_1 = VBRule1.Text
    C_Rule_2 = VBRule2.Text
    C_Rule_3 = VBRule3.Text
    Set C_Font_Text_Common = TextTmp.Font
    Set C_Font_Number = TextTmp.Font
    C_Line_Height = 255
    
    Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not GDT.Visible Then
        If (R(0).Height * (R.UBound + 1)) > UserControl.Height Then
            GDT.Visible = True
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HFFFFFF)
    C_Color_Back_Editing = PropBag.ReadProperty("Color_Back_Editing", &HFFFFCC)
    C_Color_Back_Moved = PropBag.ReadProperty("Color_Back_Moved", &HFFFFDF)
    C_Color_Back_Number = PropBag.ReadProperty("Color_Back_Number", &HF0F0F0)
    C_Color_Fore_Number = PropBag.ReadProperty("Color_Fore_Number", &H0)
    C_Color_Text_Number = PropBag.ReadProperty("Color_Text_Number", &HFF00FF)
    C_Color_Text_Common = PropBag.ReadProperty("Color_Text_Common", &H0)
    C_Rule_1 = PropBag.ReadProperty("Rule_1", VBRule1.Text)
    C_Rule_2 = PropBag.ReadProperty("Rule_2", VBRule2.Text)
    C_Rule_3 = PropBag.ReadProperty("Rule_3", VBRule3.Text)
    Set C_Font_Text_Common = PropBag.ReadProperty("Font_Text_Common", TextTmp.Font)
    Set C_Font_Number = PropBag.ReadProperty("Font_Number", TextTmp.Font)
    C_Line_Height = PropBag.ReadProperty("Line_Height", 255)
    
    Refresh
End Sub

Private Sub UserControl_Resize()
    dmrq.Top = 0
    dmrq.Width = UserControl.Width
    dmrq.Height = (R.UBound + 1) * R(0).Height
    GDT.Left = UserControl.Width - GDT.Width
    RefreshLineNumber
    GDT.Height = UserControl.Height / 5
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HFFFFFF)
    Call PropBag.WriteProperty("Color_Back_Editing", C_Color_Back_Editing, &HFFFFCC)
    Call PropBag.WriteProperty("Color_Back_Moved", C_Color_Back_Moved, &HFFFFDF)
    Call PropBag.WriteProperty("Color_Back_Number", C_Color_Back_Number, &HF0F0F0)
    Call PropBag.WriteProperty("Color_Fore_Number", C_Color_Fore_Number, &H0)
    Call PropBag.WriteProperty("Color_Text_Number", C_Color_Text_Number, &HFF00FF)
    Call PropBag.WriteProperty("Color_Text_Common", C_Color_Text_Common, &H0)
    Call PropBag.WriteProperty("Rule_1", C_Rule_1, VBRule1.Text)
    Call PropBag.WriteProperty("Rule_2", C_Rule_2, VBRule2.Text)
    Call PropBag.WriteProperty("Rule_3", C_Rule_3, VBRule3.Text)
    Call PropBag.WriteProperty("Font_Text_Common", C_Font_Text_Common, TextTmp.Font)
    Call PropBag.WriteProperty("Font_Number", C_Font_Number, TextTmp.Font)
    Call PropBag.WriteProperty("Line_Height", C_Line_Height, 255)
End Sub
