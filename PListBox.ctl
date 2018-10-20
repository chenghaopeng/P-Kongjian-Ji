VERSION 5.00
Begin VB.UserControl PListBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   ScaleHeight     =   4215
   ScaleWidth      =   3615
   Begin P¿Ø¼þ¼¯.PVScrollBar H 
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2143
   End
   Begin VB.Label FontTmp2 
      BeginProperty Font 
         Name            =   "µÈÏß Light"
         Size            =   12
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
   Begin VB.Label FontTmp 
      BeginProperty Font 
         Name            =   "µÈÏß Light"
         Size            =   11.25
         Charset         =   134
         Weight          =   300
         Underline       =   0   'False
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
   Begin VB.Label L 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   675
   End
End
Attribute VB_Name = "PListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim C_Color_Back As OLE_COLOR
Dim C_Color_Text As OLE_COLOR
Dim C_Color_Top_ScrollBar As OLE_COLOR
Dim C_Color_Back_ScrollBar As OLE_COLOR
Dim C_Picture As Picture
'Dim C_Font_Name As String
'Dim C_Font_Size As Integer
'Dim C_Font_Bold As Boolean
'Dim C_Font_Italic As Boolean
'Dim C_Font_Underline As Boolean
Dim C_Font As Font
Dim C_Font_Selected As Font
Dim C_Is_Enabled As Boolean
Dim C_Distance_Item As Integer
Dim C_Height_Item As Integer
'Dim C_Font_Size_Selected As Integer
Dim C_Color_Top_Selected As OLE_COLOR
Dim C_Color_Back_Selected As OLE_COLOR
Dim C_Color_Text_Moved As OLE_COLOR
Dim C_Color_Back_Moved As OLE_COLOR
Dim C_Style_Number As Num

Dim Items() As Variant
Dim Total As Long
Dim Checked As Long
Dim TopItem As Long

Public Event ListIndexChanged(Index As Long)
Public Event ListClick(Index As Long)
Public Event ListDblClick(Index As Long)
Public Event ListMouseDown(Index As Long)
Public Event ListMouseMove(Index As Long)
Public Event ListMouseUp(Index As Long)
Public Event Scroll(Value As Single)

Public Enum Num
    None = 0
    Arabic = 1
    Chinese = 2
    Round = 3
End Enum

Public Sub AddItem(ByVal Item As Variant, Optional ByVal Index As Long = -1)
    If Total < (UBound(Items) + 1) Then
        Total = Total + 1
        If (Index = -1) Or (Index > (Total - 1)) Then
            Items(Total - 1) = Item
        Else
            Dim i As Long, T As Variant
            For i = Total - 1 To Index + 1 Step -1
                Items(i) = Items(i - 1)
            Next
            Items(Index) = Item
            Refresh
        End If
        If (Total - 1) < Int((UserControl.Height - C_Height_Item) / (C_Height_Item + C_Distance_Item)) + 1 Then Refresh
        If (Total - 1) > Int((UserControl.Height - C_Height_Item) / (C_Height_Item + C_Distance_Item)) + 1 Then
            H.Is_Enabled = True
            H.Value = TopItem / (Total - 1)
        End If
    End If
End Sub

Public Sub Clear()
    ReDim Items(99999)
    Total = 0
    Refresh
End Sub

Public Sub RemoveItem(ByVal Index As Long)
    If Index > Total - 1 Then Exit Sub
    Dim i As Long
    For i = Index + 1 To Total - 1
        Items(i - 1) = Items(i)
    Next
    Items(Total - 1) = ""
    Total = Total - 1
    Refresh
End Sub

Public Function List(ByVal Index As Long) As Variant
    If (Index > Total - 1) Or (Index < 0) Then Exit Function
    List = Items(Index)
End Function

Public Function ListCount() As Long
    ListCount = Total
End Function

Public Function ListIndex() As Long
    ListIndex = Checked
End Function

Public Sub SetIndex(ByVal Index As Long)
    If Index = -1 Then
        If (Checked >= TopItem) And (Checked <= TopItem + l.UBound) Then
            l(Checked - TopItem).BackStyle = 0
            Set l(Checked - TopItem).Font = C_Font
'            l(Checked - TopItem).FontSize = C_Font_Size
            l(Checked - TopItem).ForeColor = C_Color_Text
        End If
        Checked = -1
        RaiseEvent ListIndexChanged(Checked)
    ElseIf (Index >= TopItem) And (Index <= TopItem + l.UBound) Then
        If (Checked >= TopItem) And (Checked <= TopItem + l.UBound) Then
            l(Checked - TopItem).BackStyle = 0
            Set l(Checked - TopItem).Font = C_Font
'            l(Checked - TopItem).FontSize = C_Font_Size
            l(Checked - TopItem).ForeColor = C_Color_Text
        End If
        Checked = Index
        l(Checked - TopItem).BackStyle = 1
        l(Checked - TopItem).BackColor = C_Color_Back_Selected
        l(Checked - TopItem).ForeColor = C_Color_Top_Selected
        Set l(Checked - TopItem).Font = C_Font_Selected
'        l(Checked - TopItem).FontSize = C_Font_Size_Selected
        RaiseEvent ListIndexChanged(Checked)
    ElseIf (Index >= 0) And (Index <= Total - 1) Then
        Checked = Index
        If Index <= Total - l.UBound - 1 Then
            H.Value = Index / (Total - 1)
            H_Scroll Index / (Total - 1)
        Else
            H.Value = 1
            H_Scroll 1
        End If
    End If
End Sub

Public Function Text() As Variant
    If Checked = -1 Then
        Text = ""
    Else
        Text = List(Checked)
    End If
End Function

Public Sub ChangeText(ByVal Index As Long, ByVal Item As Variant)
    If (Index > Total - 1) Or (Index < 0) Then Exit Sub
    Items(Index) = Item
    If (Index >= TopItem) And (Index <= TopItem + l.UBound) Then
        l(Checked - TopItem) = FormatNumber(Index + 1) & Item
    End If
End Sub

Public Sub ExchangeText(ByVal Index1 As Long, ByVal Index2 As Long)
    If (Index1 > Total - 1) Or (Index2 > Total - 1) Then Exit Sub
    Dim T As Variant
    T = Items(Index1)
    Items(Index1) = Items(Index2)
    Items(Index2) = T
    Refresh
End Sub

Public Sub MoveItem(ByVal Index As Long, ByVal Goal As Long)
    If (Index1 > Total - 1) Or (Index2 > Total - 1) Then Exit Sub
    Dim T As Variant
    T = Items(Index)
    RemoveItem Index
    AddItem T, Goal
End Sub

Public Sub Refresh()
    Checked = -1
    H.Value = 0
    TopItem = 0
    Dim i As Integer, s
    For i = 1 To l.UBound
        Unload l(i)
    Next
    s = Int((UserControl.Height - C_Height_Item) / (C_Height_Item + C_Distance_Item))
    If s > Total - 1 Then
        H.Is_Enabled = False
        s = Total - 1
    Else
        H.Is_Enabled = True
    End If
    l(0).Left = 0
    l(0).Top = 0
    l(0).Width = UserControl.Width - H.Width
    l(0).Height = C_Height_Item
    l(0).BackStyle = 0
'    l(0).FontName = C_Font_Name
'    l(0).FontSize = C_Font_Size
'    l(0).FontBold = C_Font_Bold
'    l(0).FontItalic = C_Font_Italic
'    l(0).FontUnderline = C_Font_Underline
    Set l(0).Font = C_Font
    l(0).ForeColor = C_Color_Text
    l(0).Visible = True
    If s > -1 Then
        l(0) = FormatNumber(1) & Items(0)
        For i = 1 To s
            Load l(i)
            l(i).Left = 0
            l(i).Top = i * (C_Height_Item + C_Distance_Item)
            l(i) = FormatNumber(i + 1) & Items(i)
            l(i).Visible = True
        Next
    Else
        l(0) = ""
        l(0).Visible = False
    End If
End Sub

Public Sub BackTransparent()
    Dim i As Long
    For i = 0 To UserControl.ParentControls.Count - 1
        If UserControl.ParentControls.Item(i).Name = UserControl.Extender.Name Then
            UserControl.PaintPicture UserControl.Extender.Container.Image, 0, 0, UserControl.Width, UserControl.Height, UserControl.ParentControls.Item(i).Left, UserControl.ParentControls.Item(i).Top, UserControl.Width, UserControl.Height
            Exit Sub
        End If
    Next
End Sub

Public Sub BackReduction()
    UserControl.Cls
    Set UserControl.Picture = C_Picture
End Sub

Public Function ItemIsExists(ByVal Item As Variant, Optional ByVal Index As Long = 0) As Integer
    ItemIsExists = False
    If (Index > (Total - 1)) Or (Index < 0) Then Exit Function
    For i = Index To Total - 1
        If Items(i) = Item Then
            ItemIsExists = i
            Exit Function
        End If
    Next
End Function

Public Sub SaveAllItems(ByVal strPath As String, Optional ByVal Encryption As Boolean)
    If strPath = "" Then Exit Sub
    Dim i As Long
    Open strPath For Output As #1
        For i = 0 To Total - 1
            If Encryption Then
                Print #1, Encrypt(Items(i))
            Else
                Print #1, Items(i)
            End If
        Next
    Close #1
End Sub

Public Sub ReadFile(ByVal strPath As String, Optional ByVal Encryption As Boolean)
    If strPath = "" Then Exit Sub
    Dim i As Long, T As String
    Open strPath For Input As #1
        Do Until EOF(1)
            Input #1, T
            If Encryption Then
                AddItem Declassified(T)
            Else
                AddItem T
            End If
        Loop
    Close #1
End Sub

Private Function FormatNumber(Number As Long) As String
    If C_Style_Number = 0 Then
        FormatNumber = ""
    ElseIf C_Style_Number = 1 Then
        FormatNumber = Number & ". "
    ElseIf C_Style_Number = 2 Then
        FormatNumber = Number & ". "
        FormatNumber = Replace(FormatNumber, "0", "Áã")
        FormatNumber = Replace(FormatNumber, "1", "Ò»")
        FormatNumber = Replace(FormatNumber, "2", "¶þ")
        FormatNumber = Replace(FormatNumber, "3", "Èý")
        FormatNumber = Replace(FormatNumber, "4", "ËÄ")
        FormatNumber = Replace(FormatNumber, "5", "Îå")
        FormatNumber = Replace(FormatNumber, "6", "Áù")
        FormatNumber = Replace(FormatNumber, "7", "Æß")
        FormatNumber = Replace(FormatNumber, "8", "°Ë")
        FormatNumber = Replace(FormatNumber, "9", "¾Å")
    ElseIf C_Style_Number = 3 Then
        FormatNumber = Number & ". "
        FormatNumber = Replace(FormatNumber, "0", "©–")
        FormatNumber = Replace(FormatNumber, "1", "¢Ù")
        FormatNumber = Replace(FormatNumber, "2", "¢Ú")
        FormatNumber = Replace(FormatNumber, "3", "¢Û")
        FormatNumber = Replace(FormatNumber, "4", "¢Ü")
        FormatNumber = Replace(FormatNumber, "5", "¢Ý")
        FormatNumber = Replace(FormatNumber, "6", "¢Þ")
        FormatNumber = Replace(FormatNumber, "7", "¢ß")
        FormatNumber = Replace(FormatNumber, "8", "¢à")
        FormatNumber = Replace(FormatNumber, "9", "¢á")
    End If
End Function

Public Property Get Color_Back() As OLE_COLOR
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    UserControl.BackColor = C_Color_Back
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_Text() As OLE_COLOR
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    Refresh
    PropertyChanged "Color_Text"
End Property

Public Property Get Color_Top_ScrollBar() As OLE_COLOR
    Color_Top_ScrollBar = C_Color_Top_ScrollBar
End Property

Public Property Let Color_Top_ScrollBar(ByVal vNewValue As OLE_COLOR)
    C_Color_Top_ScrollBar = vNewValue
    H.Color_Top = C_Color_Top_ScrollBar
    PropertyChanged "Color_Top_ScrollBar"
End Property

Public Property Get Color_Back_ScrollBar() As OLE_COLOR
    Color_Back_ScrollBar = C_Color_Back_ScrollBar
End Property

Public Property Let Color_Back_ScrollBar(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_ScrollBar = vNewValue
    H.Color_Back = C_Color_Back_ScrollBar
    PropertyChanged "Color_Back_ScrollBar"
End Property

Public Property Get Picture() As Picture
    Set Picture = C_Picture
End Property

Public Property Set Picture(ByVal vNewValue As Picture)
    Set C_Picture = vNewValue
    Set UserControl.Picture = C_Picture
    PropertyChanged "Picture"
End Property

Public Property Get Is_Enabled() As Boolean
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    PropertyChanged "Is_Enabled"
End Property

'Public Property Get Font_Name() As String
'    Font_Name = C_Font_Name
'End Property
'
'Public Property Let Font_Name(ByVal vNewValue As String)
'    C_Font_Name = vNewValue
'    Refresh
'    PropertyChanged "Font_Name"
'End Property
'
'Public Property Get Font_Size() As Integer
'    Font_Size = C_Font_Size
'End Property
'
'Public Property Let Font_Size(ByVal vNewValue As Integer)
'    If vNewValue <= 0 Then vNewValue = 1
'    C_Font_Size = vNewValue
'    Refresh
'    PropertyChanged "Font_Size"
'End Property
'
'Public Property Get Font_Bold() As Boolean
'    Font_Bold = C_Font_Bold
'End Property
'
'Public Property Let Font_Bold(ByVal vNewValue As Boolean)
'    C_Font_Bold = vNewValue
'    Refresh
'    PropertyChanged "Font_Bold"
'End Property
'
'Public Property Get Font_Italic() As Boolean
'    Font_Italic = C_Font_Italic
'End Property
'
'Public Property Let Font_Italic(ByVal vNewValue As Boolean)
'    C_Font_Italic = vNewValue
'    Refresh
'    PropertyChanged "Font_Italic"
'End Property
'
'Public Property Get Font_Underline() As Boolean
'    Font_Underline = C_Font_Underline
'End Property
'
'Public Property Let Font_Underline(ByVal vNewValue As Boolean)
'    C_Font_Underline = vNewValue
'    Refresh
'    PropertyChanged "Font_Underline"
'End Property

Public Property Get Font() As Font
    Set Font = C_Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set C_Font = vNewValue
    Refresh
    PropertyChanged "Font"
End Property

Public Property Get Font_Selected() As Font
    Set Font_Selected = C_Font_Selected
End Property

Public Property Set Font_Selected(ByVal vNewValue As Font)
    Set C_Font_Selected = vNewValue
    Refresh
    PropertyChanged "Font_Selected"
End Property

Public Property Get Distance_Item() As Integer
    Distance_Item = C_Distance_Item
End Property

Public Property Let Distance_Item(ByVal vNewValue As Integer)
    If vNewValue < 0 Then vNewValue = 0
    C_Distance_Item = vNewValue
    Refresh
    PropertyChanged "Distance_Item"
End Property

Public Property Get Height_Item() As Integer
    Height_Item = C_Height_Item
End Property

Public Property Let Height_Item(ByVal vNewValue As Integer)
    If vNewValue < 0 Then vNewValue = 0
    C_Height_Item = vNewValue
    Refresh
    PropertyChanged "Height_Item"
End Property

'Public Property Get Font_Size_Selected() As Integer
'    Font_Size_Selected = C_Font_Size_Selected
'End Property
'
'Public Property Let Font_Size_Selected(ByVal vNewValue As Integer)
'    If vNewValue <= 0 Then vNewValue = 1
'    C_Font_Size_Selected = vNewValue
'    Refresh
'    PropertyChanged "Font_Size_Selected"
'End Property

Public Property Get Color_Top_Selected() As OLE_COLOR
    Color_Top_Selected = C_Color_Top_Selected
End Property

Public Property Let Color_Top_Selected(ByVal vNewValue As OLE_COLOR)
    C_Color_Top_Selected = vNewValue
    Refresh
    PropertyChanged "Color_Top_Selected"
End Property

Public Property Get Color_Back_Selected() As OLE_COLOR
    Color_Back_Selected = C_Color_Back_Selected
End Property

Public Property Let Color_Back_Selected(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Selected = vNewValue
    Refresh
    PropertyChanged "Color_Back_Selected"
End Property

Public Property Get Color_Text_Moved() As OLE_COLOR
    Color_Text_Moved = C_Color_Text_Moved
End Property

Public Property Let Color_Text_Moved(ByVal vNewValue As OLE_COLOR)
    C_Color_Text_Moved = vNewValue
    Refresh
    PropertyChanged "Color_Text_Moved"
End Property

Public Property Get Color_Back_Moved() As OLE_COLOR
    Color_Back_Moved = C_Color_Back_Moved
End Property

Public Property Let Color_Back_Moved(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Moved = vNewValue
    Refresh
    PropertyChanged "Color_Back_Moved"
End Property

Public Property Get Style_Number() As Num
    Style_Number = C_Style_Number
End Property

Public Property Let Style_Number(ByVal vNewValue As Num)
    C_Style_Number = vNewValue
    Refresh
    PropertyChanged "Style_Number"
End Property

Private Sub H_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Single)
    Reload UserControl.hWnd
End Sub

Private Sub H_Scroll(NValue As Single)
    Dim i As Integer
    For i = 0 To l.UBound
        l(i) = FormatNumber(Int(NValue * (Total - 1 - l.UBound)) + i + 1) & Items(Int(NValue * (Total - 1 - l.UBound)) + i)
    Next
    TopItem = Int(NValue * (Total - 1 - l.UBound))
    For i = 0 To l.UBound
        If l(i).BackStyle <> 0 Then
            l(i).BackStyle = 0
            l(i).ForeColor = C_Color_Text
'            l(i).FontSize = C_Font_Size
            Set l(i).Font = C_Font
        End If
    Next
    If (Checked >= TopItem) And (Checked - 1 < TopItem + l.UBound) Then
        l(Checked - TopItem).BackStyle = 1
        l(Checked - TopItem).BackColor = C_Color_Back_Selected
        l(Checked - TopItem).ForeColor = C_Color_Top_Selected
'        l(Checked - TopItem).FontSize = C_Font_Size_Selected
        Set l(Checked - TopItem).Font = C_Font_Selected
    End If
    If C_Is_Enabled Then RaiseEvent Scroll(NValue)
End Sub

Private Sub L_Click(Index As Integer)
    RaiseEvent ListClick(Checked)
End Sub

Private Sub L_DblClick(Index As Integer)
    RaiseEvent ListDblClick(Checked)
End Sub

Private Sub L_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If C_Is_Enabled = False Then Exit Sub
    If Button <> 1 Then Exit Sub
    For i = 0 To l.UBound
        If l(i).BackStyle <> 0 Then
            l(i).BackStyle = 0
            l(i).ForeColor = C_Color_Text
'            l(i).FontSize = C_Font_Size
            Set l(i).Font = C_Font
        End If
    Next
    Checked = Index + TopItem
    l(Index).BackStyle = 1
    l(Index).BackColor = C_Color_Back_Selected
    l(Index).ForeColor = C_Color_Top_Selected
'    l(Index).FontSize = C_Font_Size_Selected
    Set l(Index).Font = C_Font_Selected
    RaiseEvent ListIndexChanged(Checked)
    RaiseEvent ListMouseDown(Checked)
End Sub

Private Sub L_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Reload UserControl.hWnd
    If C_Is_Enabled = False Then Exit Sub
    If (l(Index).BackStyle = 1) And (Index <> Checked - TopItem) Then Exit Sub
    For i = 0 To l.UBound
        If l(i).BackStyle <> 0 Then
            l(i).BackStyle = 0
            l(i).ForeColor = C_Color_Text
'            l(i).FontSize = C_Font_Size
            Set l(i).Font = C_Font
        End If
    Next
    l(Index).BackStyle = 1
    l(Index).BackColor = C_Color_Back_Moved
    l(Index).ForeColor = C_Color_Text_Moved
    If (Checked >= TopItem) And (Checked - 1 < TopItem + l.UBound) Then
        l(Checked - TopItem).BackStyle = 1
        l(Checked - TopItem).BackColor = C_Color_Back_Selected
        l(Checked - TopItem).ForeColor = C_Color_Top_Selected
'        l(Checked - TopItem).FontSize = C_Font_Size_Selected
        Set l(Checked - TopItem).Font = C_Font_Selected
    End If
    RaiseEvent ListMouseMove(Checked)
End Sub

Private Sub L_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If C_Is_Enabled Then RaiseEvent ListMouseUp(Checked)
End Sub

Private Sub UserControl_Initialize()
    ReDim Items(99999)
    
    C_Color_Back = &HFFFFFF
    C_Color_Text = &H0&
    C_Color_Top_ScrollBar = &HFF7402
    C_Color_Back_ScrollBar = &HF2AF00
    Set C_Picture = Nothing
'    C_Font_Name = "Î¢ÈíÑÅºÚ"
'    C_Font_Size = 11
'    C_Font_Bold = False
'    C_Font_Italic = False
'    C_Font_Underline = False
    Set C_Font = FontTmp.Font
    Set C_Font = FontTmp2.Font
    C_Is_Enabled = True
    C_Distance_Item = 0
    C_Height_Item = 300
'    C_Font_Size_Selected = 12
    C_Color_Top_Selected = &HFFFFFF
    C_Color_Back_Selected = &HFF7402
    C_Color_Text_Moved = &HFFFFFF
    C_Color_Back_Moved = &HF2AF00
    UserControl.BackColor = C_Color_Back
    UserControl.ForeColor = C_Color_Text
    Set UserControl.Picture = C_Picture
    H.Color_Top = C_Color_Top_ScrollBar
    H.Color_Back = C_Color_Back_ScrollBar
'    UserControl.FontName = C_Font_Name
'    UserControl.FontSize = C_Font_Size
'    UserControl.FontBold = C_Font_Bold
'    UserControl.FontItalic = C_Font_Italic
'    UserControl.FontUnderline = C_Font_Underline
    Set UserControl.Font = C_Font
    Style_Number = 1
    
    Refresh
    
    MLInit UserControl.hWnd
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    For i = 0 To l.UBound
        If l(i).BackStyle <> 0 Then
            l(i).BackStyle = 0
            l(i).ForeColor = C_Color_Text
'            l(i).FontSize = C_Font_Size
            Set l(i).Font = C_Font
        End If
    Next
    If (Checked >= TopItem) And (Checked - 1 < TopItem + l.UBound) Then
        l(Checked - TopItem).BackStyle = 1
        l(Checked - TopItem).BackColor = C_Color_Back_Selected
        l(Checked - TopItem).ForeColor = C_Color_Top_Selected
'        l(Checked - TopItem).FontSize = C_Font_Size_Selected
        Set l(Checked - TopItem).Font = C_Font_Selected
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Reload UserControl.hWnd
    If C_Is_Enabled = False Then Exit Sub
    For i = 0 To l.UBound
        If l(i).BackStyle <> 0 Then
            l(i).BackStyle = 0
            l(i).ForeColor = C_Color_Text
'            l(i).FontSize = C_Font_Size
            Set l(i).Font = C_Font
        End If
    Next
    If (Checked >= TopItem) And (Checked - 1 < TopItem + l.UBound) Then
        l(Checked - TopItem).BackStyle = 1
        l(Checked - TopItem).BackColor = C_Color_Back_Selected
        l(Checked - TopItem).ForeColor = C_Color_Top_Selected
'        l(Checked - TopItem).FontSize = C_Font_Size_Selected
        Set l(Checked - TopItem).Font = C_Font_Selected
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HFFFFFF)
    C_Color_Text = PropBag.ReadProperty("Color_Text", &H0&)
    C_Color_Top_ScrollBar = PropBag.ReadProperty("Color_Top_ScrollBar", &HFF7402)
    C_Color_Back_ScrollBar = PropBag.ReadProperty("Color_Back_ScrollBar", &HF2AF00)
    Set C_Picture = PropBag.ReadProperty("Picture", Nothing)
'    C_Font_Name = PropBag.ReadProperty("Font_Name", "Î¢ÈíÑÅºÚ")
'    C_Font_Size = PropBag.ReadProperty("Font_Size", 11)
'    C_Font_Bold = PropBag.ReadProperty("Font_Bold", False)
'    C_Font_Italic = PropBag.ReadProperty("Font_Italic", False)
'    C_Font_Underline = PropBag.ReadProperty("Font_Underline", False)
    Set C_Font = PropBag.ReadProperty("Font", FontTmp.Font)
    Set C_Font_Selected = PropBag.ReadProperty("Font_Selected", FontTmp2.Font)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Distance_Item = PropBag.ReadProperty("Distance_Item", 0)
    C_Height_Item = PropBag.ReadProperty("Height_Item", 300)
'    C_Font_Size_Selected = PropBag.ReadProperty("Font_Size_Selected", 12)
    C_Color_Top_Selected = PropBag.ReadProperty("Color_Top_Selected", &HFFFFFF)
    C_Color_Back_Selected = PropBag.ReadProperty("Color_Back_Selected", &HFF7402)
    C_Color_Text_Moved = PropBag.ReadProperty("Color_Text_Moved", &HFFFFFF)
    C_Color_Back_Moved = PropBag.ReadProperty("Color_Back_Moved", &HF2AF00)
    C_Style_Number = PropBag.ReadProperty("Style_Number", 1)
    UserControl.BackColor = C_Color_Back
    UserControl.ForeColor = C_Color_Text
    Set UserControl.Picture = C_Picture
    H.Color_Top = C_Color_Top_ScrollBar
    H.Color_Back = C_Color_Back_ScrollBar
'    UserControl.FontName = C_Font_Name
'    UserControl.FontSize = C_Font_Size
'    UserControl.FontBold = C_Font_Bold
'    UserControl.FontItalic = C_Font_Italic
'    UserControl.FontUnderline = C_Font_Underline
    UserControl.Font = C_Font
    
    Refresh
End Sub

Private Sub UserControl_Resize()
    H.Left = UserControl.Width - H.Width
    H.Height = UserControl.Height
    Refresh
End Sub

Private Sub UserControl_Terminate()
    MLTerminate UserControl.hWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HFFFFFF)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &H0&)
    Call PropBag.WriteProperty("Color_Top_ScrollBar", C_Color_Top_ScrollBar, &HFF7402)
    Call PropBag.WriteProperty("Color_Back_ScrollBar", C_Color_Back_ScrollBar, &HF2AF00)
    Call PropBag.WriteProperty("Picture", C_Picture, Nothing)
'    Call PropBag.WriteProperty("Font_Name", C_Font_Name, "Î¢ÈíÑÅºÚ")
'    Call PropBag.WriteProperty("Font_Size", C_Font_Size, 11)
'    Call PropBag.WriteProperty("Font_Bold", C_Font_Bold, False)
'    Call PropBag.WriteProperty("Font_Italic", C_Font_Italic, False)
'    Call PropBag.WriteProperty("Font_Underline", C_Font_Underline, False)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Font", C_Font, FontTmp.Font)
    Call PropBag.WriteProperty("Font_Selected", C_Font_Selected, FontTmp2.Font)
    Call PropBag.WriteProperty("Distance_Item", C_Distance_Item, 0)
    Call PropBag.WriteProperty("Height_Item", C_Height_Item, 300)
'    Call PropBag.WriteProperty("Font_Size_Selected", C_Font_Size_Selected, 12)
    Call PropBag.WriteProperty("Color_Top_Selected", C_Color_Top_Selected, &HFFFFFF)
    Call PropBag.WriteProperty("Color_Back_Selected", C_Color_Back_Selected, &HFF7402)
    Call PropBag.WriteProperty("Color_Text_Moved", C_Color_Text_Moved, &HFFFFFF)
    Call PropBag.WriteProperty("Color_Back_Moved", C_Color_Back_Moved, &HF2AF00)
    Call PropBag.WriteProperty("Style_Number", C_Style_Number, 1)
End Sub
