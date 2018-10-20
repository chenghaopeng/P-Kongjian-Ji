#P控件集现已弃坑。

---
##简介
P控件集是一套基于VB、应用于VB的UI系统。
现最新版本为7，之前已有版本1、版本2、版本3、版本4、版本5。
现已包含25个控件，基本可以满足日常的程序界面设计。

##P控件集版本7说明
本程序需要VB运行库支持，请自行下载安装。
本程序将访问网络，请确认网络是否可用。
本程序引用了“等线 Light”字体，如果系统中没有该字体，会使程序中所有字体均显示为“宋体”，影响外观，但不影响使用。如有需要，请自行下载安装该字体。

##P控件集简史
2014年9月6日，P控件集开始制作
2014年9月7日，于百度VB吧发布版本1预告
—http://tieba.baidu.com/p/3281539319
2014年10月2日，发布P控件集版本1
—http://tieba.baidu.com/p/3327697860
2014年11月15日，发布P控件集版本2
—http://tieba.baidu.com/p/3411960638
2015年1月3日，发布P控件集版本3
—http://tieba.baidu.com/p/3505620343
2015年2月23日，于百度VB吧发布版本4预告
—http://tieba.baidu.com/p/3599612882
2015年3月8日，发布P控件集版本4
—http://tieba.baidu.com/p/3623734163
2015年4月17日，于百度VB吧发布版本5预告
—http://tieba.baidu.com/p/3707173054
2015年5月24日，发布P控件集版本5
—http://tieba.baidu.com/p/3782929039
2016年3月5日，于百度VB吧发布版本7预告
—http://tieba.baidu.com/p/4392381562
2016年3月27日，发布P控件集版本7
—http://tieba.baidu.com/p/4439538240
		
---
##P控件集各控件的使用方法
控件名：PButton （按钮）
属性：
.Color_Back（长整型）
..返回/设置默认状态时显示的颜色
.Color_Back_Down（长整型）
..返回/设置鼠标按下时显示的颜色
.Color_Begin（长整型）
..同Color_Back
.Color_End（长整型）
..返回/设置获得鼠标焦点后渐变颜色的终值
.Color_Text（长整型）
..返回/设置文本的颜色
.Color_Text_MouseMoved（长整型）
..返回/设置鼠标触碰时文本的颜色
.Text（变体）
..返回/设置显示于按钮之上的文本
.Font（字体类型）
..返回/设置显示于按钮之上的文本之字体
.Is_Enabled（布尔型）
..返回/设置一个布尔值，决定控件是否接受用户事件
.Style_Border（自定义型）
..返回/设置按钮的边框形式
...无=1，与背景的对比色=2，用户自定义=3
.Color_Border（长整型）
..返回/设置边框的颜色
...当且仅当Style_Border=3时有效
.Can_Text_Move（布尔型）
..返回/设置一个布尔值，决定在鼠标按下时文本会否向右下角位移
.Color_Back_ChangeSpeed（整型）
..返回/设置背景颜色在渐变时的变化速度，值越小动画越缓慢越细腻
.Text_Deviate_X（整型）
..返回/设置文本在水平方向上的偏移
.Text_Deviate_Y（整型）
..返回/设置文本在竖直方向上的偏移
.Color_Back_TransparentDegree（整型）
..返回/设置按钮背景颜色的透明度
.Is_Text_Transparent（布尔型）
..返回/设置一个布尔值，决定文本是否随着背景同时透明
事件：
.Click()
..鼠标单击事件
.DblClick()
..鼠标双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标按下事件
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标触碰事件
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标弹起事件
方法：
.Refresh
..刷新/重绘按钮

---
控件名：PCheckBox （选择框）
属性：
.Color_Back
..返回/设置多选框的背景颜色
.Color_End （长整型）
..返回/设置多选框的渐变颜色的终值
.Color_Text （长整型）
..返回/设置多选框的文本颜色
.Text （变体）
..返回/设置显示于多选框上的文字
.Font（字体类型）
..返回/设置显示于按钮之上的文本之字体
.Is_Enabled （布尔型）
..返回/设置一个布尔值，决定多选框是否接受用户事件
事件：
.ValueChange(NValue （布尔型）)
..值改变事件
...NValue（布尔型）：返回多选框是否被选中
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, NValue （布尔型）)
..鼠标按下事件
...NValue（布尔型）：返回多选框是否被选中
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, NValue （布尔型）)
..鼠标触碰事件
...NValue（布尔型）：返回多选框是否被选中
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, NValue （布尔型）)
..鼠标弹起事件
...NValue（布尔型）：返回多选框是否被选中
方法：无

---
控件名：PSwitch （开关）
属性：
.Color_Top （长整型）
..返回/设置滑块颜色
.Color_Back （长整型）
..返回/设置背景颜色
.Color_Back_True （长整型）
..返回/设置选定时背景颜色
.Is_Enabled （布尔型）
..返回/设置是否接受用户事件
.Value （布尔型）
..返回/设置是否选定
.Style_Border （自定义型）
..返回/设置边框形式
...无=1，与背景的对比色=2，用户自定义=3
.Color_Border （长整型）
..返回/设置边框颜色
事件：
.ValueChange(NValue As Boolean)
..值改变事件
...NValue（布尔型）：返回当前滚动条的值
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Boolean)
..鼠标按下事件
...NValue（布尔型）：返回当前滚动条的值
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Boolean)
..鼠标触碰事件
...NValue（布尔型）：返回当前滚动条的值
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Boolean)
..鼠标弹起事件
...NValue（布尔型）：返回当前滚动条的值
方法：无

---
控件名：PProgressBar （进度条）
属性：
.Color_Top （长整型）
..返回/设置顶层颜色
.Color_Back （长整型）
..返回/设置背景颜色
.Is_Enabled （布尔型）
..返回/设置是否接受用户事件
.Value （单精度型）
..返回/设置进度条的值
事件：
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标按下事件
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标触碰事件
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标弹起事件
方法：无

---
控件名：PVScrollBar （竖滚动条）
属性：
.Color_Top （长整型）
..返回/设置滚动块的颜色
.Color_Back （长整型）
..返回/设置滚动条背景的颜色
.Is_Enabled （布尔型）
..返回/设置一个布尔值，决定滚动条是否接受用户事件
.Value （单精度型）
..返回/设置当前滚动条的值
.Size （单精度型）
..返回/设置滚动块占滚动条总宽度的百分比
.Value_Wheel_Change （单精度型）
..返回/设置鼠标滚轮每滚动一次Value变化的值
.Style_Border（自定义型）
..返回/设置边框形式
...无=1，与背景的对比色=2，用户自定义=3
.Color_Border（长整型）
..返回/设置边框的颜色
...当且仅当Style_Border=3时有效
事件：
.Scroll(NValue As Single)
..滚动事件
...NValue（单精度型）：返回当前滚动条的值
.Change(NValue As Single)
..值改变事件
...NValue（单精度型）：返回当前滚动条的值
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
..鼠标按下事件
...NValue（单精度型）：返回当前滚动条的值
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
..鼠标触碰事件
...NValue（单精度型）：返回当前滚动条的值
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
..鼠标弹起事件
...NValue（单精度型）：返回当前滚动条的值
方法：无

---
名称：PHScrollBar （横滚动条）
属性：
.Color_Top （长整型）
..返回/设置滚动块的颜色
.Color_Back （长整型）
..返回/设置滚动条背景的颜色
.Is_Enabled （布尔型）
..返回/设置一个布尔值，决定滚动条是否接受用户事件
.Value （单精度型）
..返回/设置当前滚动条的值
.Size （单精度型）
..返回/设置滚动块占滚动条总宽度的百分比
.Value_Wheel_Change （单精度型）
..返回/设置鼠标滚轮每滚动一次Value变化的值
.Style_Border（自定义型）
..返回/设置边框形式
...无=1，与背景的对比色=2，用户自定义=3
.Color_Border（长整型）
..返回/设置边框的颜色
...当且仅当Style_Border=3时有效
事件：
.Scroll(NValue As Single)
..滚动事件
...NValue（单精度型）：返回当前滚动条的值
.Change(NValue As Single)
..值改变事件
...NValue（单精度型）：返回当前滚动条的值
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
..鼠标按下事件
...NValue（单精度型）：返回当前滚动条的值
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
..鼠标触碰事件
...NValue（单精度型）：返回当前滚动条的值
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, NValue As Single)
..鼠标弹起事件
...NValue（单精度型）：返回当前滚动条的值
方法：无

---
控件名：PPictureBox （图片框）
属性：
.Color_Top （长整型）
..返回/设置滚动条滑块的颜色
.Color_Back （长整型）
..返回/设置滚动条的背景颜色
.Picture （图片型）
..返回/设置显示的图片
.Value_V （单精度型）
..返回/设置横滚动条的值
.Value_H （单精度型）
..返回/设置竖滚动条的值
.Is_Enabled （布尔型）
..返回/设置是否接受用户事件
事件：
.Scroll()
..滚动事件
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标按下事件
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标触碰事件
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标弹起事件
方法：无

---
控件名：PPRCS （坐标系）
属性：
.Resolution （单精度型）
..返回/设置坐标系的分辨率
.Color_Back （长整型）
..返回/设置背景颜色
.Color_Top （长整型）
..返回/设置线条颜色
.Grid （布尔型）
..返回/设置是否显示网格
.Color_Grid （长整型）
..返回/设置网格颜色
事件：
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标按下事件
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标触碰事件
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标弹起事件
方法：
.ReDraw()
..重绘坐标系
.DrawFunction(strFun As String)
..画y关于x的函数
...strFun（字符串）：必需，y关于x的合法函数表达式
.Line2Points(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
..连接两个点
...x1，y1，x2，y2（单精度型）：必需，欲连接的两点坐标
.DrawLineX(X As Single)
..过点（x，0）作y轴平行线
...x（单精度型）：必需，x的值
.DrawLineY(Y As Single)
..过点（0，y）作x轴平行线
...y（单精度型）：必需，y的值
.DrawPoint(X As Single, Y As Single, strText As String)
..描出（X,Y）并标上strText
...X，Y（单精度型）：必需，坐标
...strText（字符串）：必需，要标上的文字
.LoadTxt(strPath As String)
..加载文本
...strPath（字符串）：必需，路径
.SaveTxt(strPath As String)
..保存文本
...strPath（字符串）：必需，路径
.SavePic(strPath As String)
..保存为图片
...strPath（字符串）：必需，路径
.Clear()
..清空

---
控件名：PWinsock （Winsock封装）
属性：无
事件：
.ListenBegin()
..侦听开始事件
.ConnectBegin()
..连接开始事件
.ConnectFail()
..连接失败事件
.ConnectSucceed()
..连接成功事件
.ConnectClose()
..连接关闭事件
.DataSendError()
..数据发送失败事件
.DataSendSucceed()
..数据发送成功事件
.DataArrived(strTouwenjian As String, Shuju As Variant)
..数据接收事件
.DocumentSendError()
..文件接收失败事件
.DocumentSendSucceed()
..文件接收成功事件
.DocumentSending(Progress As Long, Total As Long)
..文件正在接收事件
.DocumentQuest()
..文件请求事件
.DocumentArrived()
..文件接收事件
.DocumentArriveError()
..文件接收失败事件
.DocumentComplete()
..文件接收完成事件
方法：
.ConnectionClose()
..关闭连接
.Connect(strIP As String)
..建立连接
...strIP（字符串）：必需，目标IP
.Listen()
..侦听
.ConnectIsOK() As Boolean
..获取连接是否完成
...返回成功1，失败0
.SendData(strTouwenjian As String, Shuju As Variant) As Boolean
..发送数据
...strTouwenjian（字符串）：必需，头文件
...Shuju（变体）：必需，数据
....返回成功1，失败0
.SendDocument(strPath As String) As Boolean
..发送文件
...strPath（字符串）：必需，文件路径
....返回成功1，失败0
.AcceptDocument(strPath As String) As Boolean
..同意接收文件
...strPath（字符串）：必需，保存路径
....返回成功1，失败0
.RefuseDocument() As Boolean
..拒绝接收文件
.RemoteHost()，RemoteHostIP()，RemotePort()
..获取对方主机，IP，端口
.LocalHostName()，LocalIP()，LocalPort()
..获取本地主机，IP，端口

---
名称：PListBox （列表框）
属性：
.Color_Back （长整型）
..返回/设置列表框的背景颜色
.Color_Text （长整型）
..返回/设置列表框的文本颜色
.Color_Top_ScrollBar （长整型）
..返回/设置列表框中滚动条的滚动块颜色
.Color_Back_ScrollBar （长整型）
..返回/设置列表框中滚动条的背景颜色
.Picture （图片型）
..返回/设置显示于列表框中的图片
.Font（字体类型）
..返回/设置显示于列表框中文本的字体
.Is_Enabled （布尔型）
..返回/设置列表框是否接受用户事件
.Distance_Item （整型）
..返回/设置列表框中每一项之间的间隔
.Height_Item （整型）
..返回/设置列表框中每一项的高度
.Font_Size_Selected （整型）
..返回/设置列表框中选中项的文字大小
.Color_Top_Selected （长整型）
..返回/设置列表框中选中项的字体颜色
.Color_Back_Selected （长整型）
..返回/设置列表框中选中项的背景颜色
.Color_Text_Moved （长整型）
..返回/设置列表框中鼠标触碰项的字体颜色
.Color_Back_Moved （长整型）
..返回/设置列表框中鼠标触碰项的背景颜色
.Style_Number （自定义型）
..返回/设置列表框中每一项前的标号类型
...无=0；阿拉伯数字=1；汉字数字=2；圆圈数字=3；
事件：
.ListIndexChanged(Index As Long)
..列表框选中项改变事件
...Index（长整型）：被选中的项
.ListClick(Index As Long)
..列表框单击事件
...Index（长整型）：被选中的项
.ListDblClick(Index As Long)
.ListClick(Index As Long)
..列表框双击事件
...Index（长整型）：被选中的项
.ListMouseDown(Index As Long)
..列表框鼠标按下事件
...Index（长整型）：被选中的项
.ListMouseMove(Index As Long)
..列表框鼠标触碰事件
...Index（长整型）：被选中的项
.ListMouseUp(Index As Long)
..列表框鼠标弹起事件
...Index（长整型）：被选中的项
.Scroll(Value As Single)
..列表框中滚动条的滚动事件
...Value（单精度型）：滚动条的值
方法： 
.AddItem(ByVal Item As Variant, Optional ByVal Index As Long = -1)
..添加项
...Item（变体）：必需，项的内容
...Index（长整型）：可选，项添加的位置，默认-1即最后
.Clear()
..清空列表
.RemoveItem(ByVal Index As Long)
..移除项
...Index（长整型）：必需，项的索引
.List(ByVal Index As Long) As Variant
..获取某项的内容
...Index（长整型）：必需，项的索引
....返回项的内容（变体）
.ListCount() As Long
..获取项数
...返回项数（长整型）
.ListIndex() As Long
..获取当前选中项的索引
...返回选中项索引（长整型）
.SetIndex(ByVal Index As Long)
..设置选中项的索引
...Index（长整型）：必需，欲选中项
.Text() As Variant
..获取选中项的文本
...返回选中项的文本（变体）
.ChangeText(ByVal Index As Long, ByVal Item As Variant)
..更改文本
...Index（长整型）：必需，欲更改文本的项的索引
...Item（变体）：必需，欲更改为的文本
.ExchangeText(ByVal Index1 As Long, ByVal Index2 As Long)
..交换文本
...Index1（长整型）：必需，欲交换的项1
...Index2（长整型）：必需，欲交换的项2
.MoveItem(ByVal Index As Long, ByVal Goal As Long)
..移动项
...Index（长整型）：必需，欲移动的项
...Goal（长整型）：必需，欲移动至的索引
....注：此操作不会覆盖原第Goal项文本
.Refresh()
..刷新
.BackTransparent()
..使列表框的背景透明
.BackReduction()
..使列表框的背景恢复
.ItemIsExists(ByVal Item As Variant, Optional ByVal Index As Long = 0) As Boolean
..判断某项是否存在
...Item（变体）：必需，欲查找的项的文本
...Index（长整型）：可选，表示查找开始的项，默认从头开始
....存在则TRUE，不存在则返回FALSE（布尔型）
.SaveAllItems(ByVal strPath As String, Optional ByVal Encryption As Boolean)
..保存所有项
...strPath（字符串）：必需，文件存储的路径
...Encryption（布尔型）：可选，决定文本存储时是否加密
.ReadFile(ByVal strPath As String, Optional ByVal Encryption As Boolean)
..读取文件
...strPath（字符串）：必需，文件路径
...Encryption（布尔型）：可选，决定文本读取时是否解密

---
控件名：PMaths （数学控件）
属性：无
事件：无
方法：
.Add(strShu1 As String, strShu2 As String) As String
..高精度加法
...strShu1（字符串）：必需，加数1
...strShu2（字符串）：必需，加数2
....返回和（字符串）
.Subtract(strShu1 As String, strShu2 As String) As String
..高精度减法
...strShu1（字符串）：必需，被减数
...strShu2（字符串）：必需，减数
....返回差（字符串）
.Multiply(strShu1 As String, strShu2 As String) As String
..高精度乘法
...strShu1（字符串）：必需，乘数1
...strShu2（字符串）：必需，乘数2
....返回积（字符串）
.Division(strShu1 As String, strShu2 As String) As String
..高精度除法
...strShu1（字符串）：必需，被除数
...strShu2（字符串）：必需，除数
....返回商的整数部分（字符串）
.VBCodetoNum(Code As String) As Single
..求VB表达式的值
...Code（字符串）：必需，VB表达式
....返回表达式的值（单精度型）

---
控件名：PScreen （像素屏）
属性：
.Color_Back （长整型）
..返回/设置背景颜色
.Color_Text （长整型）
..返回/设置字体颜色
.Color_Text_Back （长整型）
..返回/设置字体背景颜色
.Text （字符串）
..返回/设置文本
.Size （整型）
..返回/设置像素点大小
.Font （字体类型）
..返回/设置字体
.Distance （整型）
..返回/设置像素点间距离
.Style_Shape （自定义型）
..返回/设置像素点形状
...矩形=1，圆形=2
事件：无
方法：无

---
控件名：PTab （选项卡）
属性：
.Color_Back （长整型）
..返回/设置背景颜色
.Color_Text （长整型）
..返回/设置文本颜色
.Picture （图片类型）
..返回/设置背景图片
.Font （字体类型）
..返回/设置字体
.Is_Enabled （布尔型）
..返回/设置是否接受用户事件
.Distance_Transverse （整型）
..返回/设置横向间距
.Distance_Vertical （整型）
..返回/设置纵向间距
.Color_Selected （长整型）
..返回/设置被选择的文本颜色
.Color_Selector （长整型）
..返回/设置滑块颜色
.Color_Selector_Moved （长整型）
..返回/设置滑块被鼠标指向时的颜色
.Height_Selector （整型）
..返回/设置滑块的高度
.Texts （字符串）
..返回/设置文本
.Is_AutoDisplay （布尔型）
..返回/设置是否自动显示控件
.Is_AutoUndisplay （布尔型）
..返回/设置是否自动隐藏控件
事件：
.Click()
..单击事件
.DblClick()
..双击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标按下事件
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标触碰事件
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标弹起事件
.ItemSelected(NewIndex As Integer, LastIndex As Integer)
..选择事件
...NewIndex（整型）：返回新选项
...LastIndex（整型）：返回上次选项
方法：无

---
控件名：PWeather （天气控件）
属性：无
事件：无
方法：
.GetWethInfo_Today(Optional ByVal strCityName As String = "") As String
..获取今日天气
...strCityName（字符串）：可选，城市名
.GetWethInfo_Pred(Optional ByVal strCityName As String = "") As String
..获取昨日天气
...strCityName（字符串）：可选，城市名
.GetWethInfo_Succ(Optional ByVal Days As Integer = 1, Optional ByVal strCityName As String = "") As String
..获取今后天气
...Days（整型）：可选，今后的哪一天，1<Days<4
...strCityName（字符串）：可选，城市名

---
控件名：PWin8Form （仿Win8窗体）
属性：
.Icon （图片类型）
..返回/设置图标
.Picture （图片类型）
..返回/设置背景图片
.Caption （字符串）
..返回/设置标题
.Color_Border （长整型）
..返回/设置边线颜色
.Color_Frame （长整型）
..返回/设置边框颜色
.Color_Back （长整型）
..返回/设置背景颜色
.Is_Stretch （布尔型）
..返回/设置图片是否拉伸
.Can_Move_Smoothly （布尔型）
..返回/设置是否平滑移动
.Is_Enabled （布尔型）
..返回/设置是否接受用户事件
.Has_MinButton （布尔型）
..返回/设置是否具有最小化按钮
.Has_MaxButton （布尔型）
..返回/设置是否具有最大化按钮
.Has_CloseButton （布尔型）
..返回/设置是否具有关闭按钮
.Has_Icon （布尔型）
..返回/设置是否具有图标
.Is_Resizable （布尔型）
..返回/设置是否可以调整大小
事件：无
方法：无

---
控件名：PUIMgr（UI管家）
属性：无
事件：
.MoveSmlyComplete(Control As Object)
..平滑移动结束事件
...Control（对象）：返回对象
.SizeSmlyComplete(Control As Object)
..平滑改变大小结束事件
...Control（对象）：返回对象
.ColorSmlyIng(nColor As Long)
..正在颜色渐变事件
...nColor（长整型）：返回当前颜色
.ColorSmlyComplete()
..颜色渐变结束事件
方法：
.MoveSmly(ByRef Control As Object, ByVal nLeft As Long, ByVal nTop As Long, ByVal Delay As Integer, Optional ByVal Speed As Integer = 10)
..平滑移动控件
...Control（对象）：必需，对象
...nLeft，nTop（长整型）：必需，目标位置
...Delay（整型）：必需，延迟
...Speed（长整型）：可选，平滑度
.StopMoveSmly()
..停止平滑移动控件
.SizeSmly(ByRef Control As Object, ByVal nWidth As Long, ByVal nHeight As Long, ByVal Delay As Integer, Optional ByVal Speed As Integer = 10)
..平滑改变控件大小
...Control（对象）：必需，对象
...nWidth，nHeight（长整型）：必需，目标大小
...Delay（整型）：必需，延迟
...Speed（长整型）：可选，平滑度
.StopSizeSmly()
..停止平滑改变控件大小
.ColorSmly(ByVal CurrentColor As Long, ByVal GoalColor As Long, ByVal CGSPD As Integer, ByVal Delay As Integer)
..颜色渐变
...CurrentColor，GoalColor（长整型）：必需，当前/目标颜色
...CGSPD（整型）：必需，渐变速度
...Delay（整型）：必需，延迟
.ControlTransparent(ByRef Container As Object, ByRef Control As Object, ByVal Transparency As Integer)
..控件透明
...Container（对象）：必需，容器
...Control（对象）：必需，要透明的对象
...Transparency（整型）：必需，透明度，0-255

---
控件名：PUIMgrPlus （UI管家加强）
属性：无
事件：
.MoveSmlyComplete(Control As Object)
..平滑移动结束事件
...Control（对象）：返回对象
.ColorSmlyComplete(Index As Integer)
..颜色渐变结束事件
...Index（整型）：返回索引
.ColorSmlyIng(Index As Integer, nColor As Long)
..正在颜色渐变事件
...Index（整型）：返回索引
...nColor（长整型）：返回当前颜色
方法：
.MoveSmly(ByRef Control As Object, ByVal nLeft As Long, ByVal nTop As Long, ByVal Delay As Integer, ByVal Index As Integer, Optional ByVal Speed As Integer = 10)
..平滑移动控件
...Control（对象）：必需，对象
...nLeft，nTop（长整型）：必需，目标位置
...Delay（整型）：必需，延迟
...Index（整型）：必需，索引
...Speed（长整型）：可选，平滑度
.ColorSmly(ByVal CurrentColor As Long, ByVal GoalColor As Long, ByVal CGSPD As Integer, ByVal Delay As Integer)
..颜色渐变
...CurrentColor，GoalColor（长整型）：必需，当前/目标颜色
...CGSPD（整型）：必需，渐变速度
...Delay（整型）：必需，延迟
...Index（整型）：必需，索引

---
控件名：PNet （网络应用）
属性：无
事件：无
方法：
.GetHtmlCodeByXMLHTTP(ByVal strUrl As String) As String
..通过XMLHTTP获取网页源码
...strUrl（字符串）：必需，网页地址
....返回网页源码（字符串）
.GetHtmlCodeByInet(ByVal strUrl As String) As String
..通过Inet获取网页源码
...strUrl（字符串）：必需，网页地址
....返回网页源码（字符串）
.GetHtmlCodeByWebbrowser(ByVal strUrl As String) As String
..通过Webbrowser获取网页源码
...strUrl（字符串）：必需，网页地址
....返回网页源码（字符串）
.GetCurrentIP() As String
..获取当前外网IP
...返回当前外网IP（字符串）
.GetCurrentIPLoaction() As String
..获取当前外网IP所在地
...返回当前外网IP所在地（字符串）
.GetCurrentIPOperator() As String
..获取当前外网IP提供商
...返回当前外网IP提供商（字符串）
DownloadFile(strUrl As String, strSavePath As String) As Boolean
.下载文件
..strUrl（字符串）：必需，下载地址
..strSavePath（字符串）：必需，保存地址
...返回下载文件成功1，失败0（布尔值）

---
控件名：PSubtitles （字幕）
属性：
.TextsAndLinks （字符串）
..返回/设置文本和链接
.Color_Text （长整型）
..返回/设置字体颜色
.Color_Text_End （长整型）
..返回/设置字体渐变结束颜色
.Color_Back （长整型）
..返回/设置背景颜色
.Font （字体类型）
..返回/设置字体
.Is_Enabled （布尔型）
..返回/设置是否接受用户事件
.Is_Back_Transparent （布尔型）
..返回/设置背景是否透明
.Interval （整型）
..返回/设置字幕切换间隔
.Text_Align （整型）
..返回/设置文本对齐方式
.Is_Random （布尔型）
..返回/设置字幕切换是否随机
.Color_Back_ChangeSpeed （整型）
..返回/设置文本颜色渐变速度
事件：无
方法：无

---
控件名：PUpdate （更新控件）
属性：无
事件：无
方法：
.CheckUpdate()
..涉及服务器安全，保密

---
控件名：PCodeTextBox （代码框）
属性：
.Color_Back （长整型）
..返回/设置背景颜色
.Color_Back_Editing （长整型）
..返回/设置当前编辑行的背景颜色
.Color_Back_Moved （长整型）
..返回/设置当前鼠标指向行的背景颜色
.Color_Back_Number （长整型）
..返回/设置行号的背景颜色
.Color_Fore_Number （长整型）
..返回/设置行号的颜色
.Color_Text_Number （长整型）
..返回/设置代码中数字的颜色
.Color_Text_Common （长整型）
..返回/设置代码中普通文本的颜色
.Rule_1 （字符串）
..返回/设置代码高亮规则1
...保留字高亮规则
.Rule_2 （字符串）
..返回/设置代码高亮规则2
...符号高亮规则
.Rule_3 （字符串）
..返回/设置代码高亮规则3
...跨行符号高亮规则
.Font_Text_Common （字体类型）
..返回/设置代码中文本的字体
.Font_Number （字体类型）
..返回/设置代码中数字的字体
.Line_Height （整型）
..返回/设置每行的高度
事件：无
方法：无

---
控件名：PContainer （容器）
属性：
.C_Color_Back （长整型）
..返回/设置背景颜色
.Color_Back_Down （长整型）
..返回/设置鼠标按下时的背景颜色
.Color_Circle （长整型）
..返回/设置圆形的填充颜色
.Color_Back_ChangeSpeed （整型）
..返回/设置背景颜色的渐变速度
.Size_Circle_ChangeSpeed_1 （整型）
..返回/设置鼠标按下时圆形的大小变化速度
.Size_Circle_ChangeSpeed_2 （整型）
..返回/设置鼠标弹起后圆形的大小变化速度
事件：
.Click()
..鼠标单击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标按下事件
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标触碰事件
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标弹起事件
方法：无

---
控件名：PCodeTextBoxE （简选择框）
属性：
.Color_Back_1 （长整型）
..返回/设置未选定的背景颜色
.Color_Back_Down_1 （长整型）
..返回/设置未选定按下时背景颜色
.Color_Circle_1 （长整型）
..返回/设置未选定圆形的颜色
.Color_Back_2 （长整型）
..返回/设置选定的背景颜色
.Color_Back_Down_2 （长整型）
..返回/设置选定按下时背景颜色
.Color_Circle_2 （长整型）
..返回/设置选定圆形的颜色
.Color_Back_ChangeSpeed （整型）
..返回/设置背景颜色渐变速度
.Size_Circle_ChangeSpeed_1 （整型）
..返回/设置按下时圆形大小的变化速度
.Size_Circle_ChangeSpeed_2 （整型）
..返回/设置弹起时圆形大小的变化速度
.Text （字符串）
..返回/设置文本
.Font （字体类型）
..返回/设置字体
.Value （布尔型）
..返回/设置是否选定
.Color_Text （长整型）
..返回/设置文本颜色
事件：
.ValueChange(NewValue As Boolean)
..值改变事件
...NewValue（布尔型）：返回是否选定
方法：无

---
控件名：PButtonE （简按钮）
属性：
.C_Color_Back （长整型）
..返回/设置按钮的背景颜色
.Color_Back_Down （长整型）
..返回/设置按钮按下时的背景颜色
.Color_Circle （长整型）
..返回/设置按钮上圆形的填充颜色
.Color_Back_ChangeSpeed （整型）
..返回/设置按钮背景颜色的渐变速度
.Size_Circle_ChangeSpeed_1 （整型）
..返回/设置按钮按下时圆形的大小变化速度
.Size_Circle_ChangeSpeed_2 （整型）
..返回/设置按钮弹起后圆形的大小变化速度
.Text （字符串）
..返回/设置按钮上的文本
.Font （字体类型）
..返回/设置按钮上文本的字体
.Color_Text （长整型）
..返回/设置按钮上文本的颜色
事件：
.Click()
..鼠标单击事件
.MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标按下事件
.MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标触碰事件
.MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
..鼠标弹起事件
方法：无

---
控件名：PTabE （简选项卡）
属性：
.Color_Back_1 （长整型）
..返回/设置未选定的背景颜色
.Color_Back_Down_1 （长整型）
..返回/设置未选定按下时背景颜色
.Color_Circle_1 （长整型）
..返回/设置未选定圆形的颜色
.Color_Back_2 （长整型）
..返回/设置选定的背景颜色
.Color_Back_Down_2 （长整型）
..返回/设置选定按下时背景颜色
.Color_Circle_2 （长整型）
..返回/设置选定圆形的颜色
.Color_Back_ChangeSpeed （整型）
..返回/设置背景颜色渐变速度
.Size_Circle_ChangeSpeed_1 （整型）
..返回/设置按下时圆形大小的变化速度
.Size_Circle_ChangeSpeed_2 （整型）
..返回/设置弹起时圆形大小的变化速度
.Text （字符串）
..返回/设置文本
.Font （字体类型）
..返回/设置字体
.Color_Text （长整型）
..返回/设置文本颜色
.ExtendWidth （整型）
..返回/设置每个选项的拓展宽度
.ScrollSpeed （单精度型）
..返回/设置滑块的滚动速度
事件：
.IndexChange(NewIndex As Integer, LastIndex As Integer)
..选项改变事件
...NewIndex（整型）：返回新选项
...LastIndex（整型）：返回上次选项
方法：
.SetIndex(ByVal Index As Integer)
..改变选项
...Index（整型）：必需，要切换到的索引