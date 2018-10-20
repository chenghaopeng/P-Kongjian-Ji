VERSION 5.00
Begin VB.UserControl PCheckBox 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ScaleHeight     =   375
   ScaleWidth      =   1455
   Begin P控件集.PButton B2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "等线"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin P控件集.PButton B1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Text            =   "×"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "等线 Light"
         Size            =   11.25
         Charset         =   134
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "PCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'██████↓定义存储属性的变量↓███████████████████████████████████████████████████████
Dim C_Color_Back As OLE_COLOR '背景颜色
Dim C_Color_End As OLE_COLOR '渐变结束的颜色
Dim C_Color_Text As OLE_COLOR '按钮文本的颜色
Dim C_Text As String '文本
Dim C_Font As Font '字体
Dim C_Is_Enabled As Boolean '是否可用
Dim C_Value As Boolean '值
'██████↓定义事件↓███████████████████████████████████████████████████████
Public Event ValueChange(NValue As Boolean) '值改变事件
Public Event Click() '单击事件
Public Event DblClick() '双击事件
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Boolean) '鼠标按下事件,NValue为新值
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Boolean) '鼠标触碰事件,NValue为新值
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, NValue As Boolean) '鼠标弹起事件,NValue为新值
'██████↓各种属性↓███████████████████████████████████████████████████████
Public Property Get Value() As Boolean '值
    Value = C_Value
End Property

Public Property Let Value(ByVal vNewValue As Boolean)
    C_Value = vNewValue
    If C_Value = True Then
        B1.Text = "√"
    Else
        B1.Text = "×"
    End If
    PropertyChanged "Value"
End Property

Public Property Get Is_Enabled() As Boolean '是否有效
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Font() As Font '字体
    Set Font = C_Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set C_Font = vNewValue
    Set B1.Font = C_Font
    Set B2.Font = C_Font
    PropertyChanged "Font"
End Property

Public Property Get Text() As String '文本
    Text = C_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    C_Text = vNewValue
    B2.Text = vNewValue
    PropertyChanged "Text"
End Property

Public Property Get Color_Back() As OLE_COLOR '背景颜色
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    B1.Color_Back = vNewValue
    B2.Color_Back = vNewValue
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_End() As OLE_COLOR '渐变结束的颜色
    Color_End = C_Color_End
End Property

Public Property Let Color_End(ByVal vNewValue As OLE_COLOR)
    C_Color_End = vNewValue
    B1.Color_End = vNewValue
    B2.Color_End = vNewValue
    PropertyChanged "Color_End"
End Property

Public Property Get Color_Text() As OLE_COLOR '文本颜色
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    B1.Color_Text = vNewValue
    B2.Color_Text = vNewValue
    PropertyChanged "Color_Text"
End Property
'██████↓各种事件↓███████████████████████████████████████████████████████
Private Sub B1_Click() 'B1的单击事件
    Value = Not (Value) '值取相反
    If C_Is_Enabled = True Then RaiseEvent Click '如果有效,触发单击事件
    If C_Is_Enabled = True Then RaiseEvent ValueChange(C_Value) '如果有效,触发值改变事件
End Sub

Private Sub B1_DblClick() 'B1的双击事件
    If C_Is_Enabled = True Then RaiseEvent DblClick '如果有效,触发双击事件
End Sub

Private Sub B1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B1的鼠标按下事件
    If C_Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y, C_Value) '如果有效,触发鼠标按下事件
End Sub

Private Sub B1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B1的鼠标触碰事件
    If C_Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y, C_Value) '如果有效,触发鼠标触碰事件
End Sub

Private Sub B1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B1的鼠标弹起事件
    If C_Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y, C_Value) '如果有效,触发鼠标弹起事件
End Sub

Private Sub B2_Click() 'B2的单击事件
    B1_Click '调用B1的单击事件
End Sub

Private Sub B2_DblClick() 'B2的双击事件
    B1_DblClick '调用B1的双击事件
End Sub

Private Sub B2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B2的鼠标按下事件
    B1_MouseDown Button, Shift, X, Y '调用B1的鼠标按下事件
End Sub

Private Sub B2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B2的鼠标触碰事件
    B1_MouseMove Button, Shift, X, Y '调用B1的鼠标触碰事件
End Sub

Private Sub B2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B2的鼠标弹起事件
    B1_MouseUp Button, Shift, X, Y '调用B1的鼠标弹起事件
End Sub

Private Sub UserControl_Initialize() '控件的加载事件
    C_Color_Back = &HF2AF00 '定义每种属性的初始值
    C_Color_End = &HFF7402
    C_Color_Text = &H0&
    C_Text = "PCheckBox"
    Set C_Font = FontTmp.Font
    C_Is_Enabled = True
    C_Value = False
    B1.Color_Text = C_Color_Text '配置每种属性
    B2.Color_Text = C_Color_Text
    B1.Color_End = C_Color_End
    B2.Color_End = C_Color_End
    B1.Color_Back = C_Color_Back
    B2.Color_Back = C_Color_Back
    B2.Text = C_Text
    Set B1.Font = C_Font
    Set B2.Font = C_Font
    If C_Value = True Then
        B1.Text = "√"
    Else
        B1.Text = "×"
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag) '控件的读取属性事件
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00) '读取各种属性和初始值
    C_Color_End = PropBag.ReadProperty("Color_End", &HFF7402)
    C_Color_Text = PropBag.ReadProperty("Color_Text", &H0&)
    C_Text = PropBag.ReadProperty("Text", "PButton")
    Set C_Font = PropBag.ReadProperty("Font", FontTmp.Font)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Value = PropBag.ReadProperty("Value", False)
    B1.Color_Text = C_Color_Text '配置各种属性
    B2.Color_Text = C_Color_Text
    B1.Color_End = C_Color_End
    B2.Color_End = C_Color_End
    B1.Color_Back = C_Color_Back
    B2.Color_Back = C_Color_Back
    B2.Text = C_Text
    Set B1.Font = C_Font
    Set B2.Font = C_Font
    If C_Value = True Then
        B1.Text = "√"
    Else
        B1.Text = "×"
    End If
End Sub

Private Sub UserControl_Resize() '控件的大小改变事件
    If UserControl.Width < UserControl.Height Then UserControl.Width = UserControl.Height '改变每个控件的大小和位置
    B1.Width = UserControl.Height
    B1.Height = UserControl.Height
    B2.Height = UserControl.Height
    B2.Width = UserControl.Width - B1.Width
    B2.Left = UserControl.Width - B2.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag) '控件的写属性事件
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00) '写入各种属性和初始值
    Call PropBag.WriteProperty("Color_End", C_Color_End, &HFF7402)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &H0&)
    Call PropBag.WriteProperty("Text", C_Text, "PButton")
    Call PropBag.WriteProperty("Font", C_Font, FontTmp.Font)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Value", C_Value, False)
End Sub
'↑↑↑↑↑↑↑The End↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

