VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl PWinsock 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   480
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Wj 
      Left            =   360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sj 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "PWinsock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event ListenBegin()
Public Event ConnectBegin()
Public Event ConnectFail()
Public Event ConnectSucceed()
Public Event ConnectClose()
Public Event DataSendError()
Public Event DataSendSucceed()
Public Event DataArrived(strTouwenjian As String, Shuju As Variant)
Public Event DocumentSendError()
Public Event DocumentSendSucceed()
Public Event DocumentSending(Progress As Long, Total As Long)
Public Event DocumentQuest()
Public Event DocumentArrived()
Public Event DocumentArriveError()
Public Event DocumentComplete()

Dim Connected As Boolean
Dim AcceptingOrSendingDocument As Boolean
Dim DocumentPath As String

Public Sub ConnectionClose()
    Sj.Close
    Wj.Close
    Connected = False
End Sub

Public Sub Connect(strIP As String)
    RaiseEvent ConnectBegin
    Connected = False
    Sj.Close
    Wj.Close
    Sj.Connect strIP, 5000
    Wj.Connect strIP, 5001
    Timer1.Enabled = True
End Sub

Public Sub Listen()
    RaiseEvent ListenBegin
    Connected = False
    Sj.Close
    Wj.Close
    Sj.LocalPort = 5000
    Wj.LocalPort = 5001
    Sj.Bind
    Wj.Bind
    Sj.Listen
    Wj.Listen
    Timer1.Enabled = True
End Sub

Public Function ConnectIsOK() As Boolean
    ConnectIsOK = Connected
End Function

Public Function SendData(strTouwenjian As String, Shuju As Variant) As Boolean
    On Error GoTo Err
    
    If (strTouwenjian = "sjswj") Or (strTouwenjian = "ejswj") Or (strTouwenjian = "acceptd") Or (strTouwenjian = "refuse") Then GoTo Err
    
    Dim Bag As New PropertyBag
    Dim bytData() As Byte
    Bag.WriteProperty "lx", strTouwenjian
    Bag.WriteProperty "sj", Shuju
    bytData = Bag.Contents
    Sj.SendData bytData
    SendData = True
    RaiseEvent DataSendSucceed
Err:
    SendData = False
    RaiseEvent DataSendError
End Function

Public Function SendDocument(strPath As String) As Boolean
    On Error GoTo Err
    
    If strPath = "" Then GoTo Err
    If Dir(strPath, vbHidden + vbReadOnly + vbSystem) = "" Then GoTo Err
    DocumentPath = strPath
    Dim Bag As New PropertyBag
    Dim bytData() As Byte
    Bag.WriteProperty "lx", "sjswj"
    Bag.WriteProperty "sj", ""
    bytData = Bag.Contents
    Sj.SendData bytData
    Timer2.Enabled = True
    SendDocument = True
    Exit Function
Err:
    SendDocument = False
End Function

Public Function AcceptDocument(strPath As String) As Boolean
    On Error GoTo Err
    
    Dim Bag As New PropertyBag
    Dim bytData() As Byte
    Bag.WriteProperty "lx", "acceptd"
    Bag.WriteProperty "sj", ""
    bytData = Bag.Contents
    Sj.SendData bytData
    DocumentPath = strPath
    AcceptDocument = True
    Exit Function
Err:
    AcceptDocument = False
End Function

Public Function RefuseDocument() As Boolean
    On Error GoTo Err
    
    Dim Bag As New PropertyBag
    Dim bytData() As Byte
    Bag.WriteProperty "lx", "refuse"
    Bag.WriteProperty "sj", ""
    bytData = Bag.Contents
    Sj.SendData bytData
    DocumentPath = strPath
    RefuseDocument = True
    AcceptingOrSendingDocument = False
    Exit Function
Err:
    RefuseDocument = False
    AcceptingOrSendingDocument = False
End Function

Public Function RemoteHost()
    If Connected Then RemoteHost = Sj.RemoteHost
End Function

Public Function RemoteHostIP()
    If Connected Then RemoteHostIP = Sj.RemoteHostIP
End Function

Public Function RemotePort()
    If Connected Then RemotePort = Sj.RemotePort
End Function

Public Function LocalHostName()
    If Connected Then LocalHostName = Sj.LocalHostName
End Function

Public Function LocalIP()
    If Connected Then LocalIP = Sj.LocalIP
End Function

Public Function LocalPort()
    If Connected Then LocalPort = Sj.LocalPort
End Function

Private Sub Sj_Close()
    RaiseEvent ConnectClose
    Connected = False
    Sj.Close
    Wj.Close
End Sub

Private Sub Sj_ConnectionRequest(ByVal requestID As Long)
    Sj.Close
    Sj.Accept requestID
End Sub

Private Sub Sj_DataArrival(ByVal bytesTotal As Long)
    Dim ByteRecv() As Byte
    Dim StringRecv As String, strTouwenjian As String, Shuju As Variant
    Dim Bag As New PropertyBag
    Dim MemBuf() As Byte
    ReDim ByteRecv(bytesTotal - 1)
    Sj.GetData ByteRecv
    StringRecv = ByteRecv
    TotalData = TotalData & StringRecv
    MemBuf = TotalData
    Bag.Contents = MemBuf
    strTouwenjian = Bag.ReadProperty("lx")
    Shuju = Bag.ReadProperty("sj")

    If strTouwenjian = "sjswj" Then
        DocumentPath = Shuju
        AcceptingOrSendingDocument = True
        RaiseEvent DocumentQuest
    ElseIf strTouwenjian = "ejswj" Then
        AcceptingOrSendingDocument = False
        RaiseEvent DocumentComplete
    ElseIf strTouwenjian = "acceptd" Then
        AcceptingOrSendingDocument = True
        StartSend
    ElseIf strTouwenjian = "refuse" Then
        AcceptingOrSendingDocument = False
    Else
        RaiseEvent DataArrived(strTouwenjian, Shuju)
    End If
End Sub

Private Sub Sj_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent ConnectFail
    Connected = False
    Sj.Close
    Wj.Close
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If (Sj.State = 7) And (Wj.State = 7) Then
        RaiseEvent ConnectSucceed
        Connected = True
        Timer1.Enabled = False
    End If
End Sub

Private Sub StartSend()
    On Error GoTo Err
    
    If AcceptingOrSendingDocument Then
    
        Dim BytDate() As Byte
        Dim FileName As String
        Dim lngFile As Long
        Dim i As Long
    
        FileName = DocumentPath '取得文件名及路径
        lngFile = FileLen(FileName) \ 1024 '取得文件长度
    
        For i = 0 To lngFile
            ReDim myFile(1023) As Byte '初始化数组
            Open FileName For Binary As #1 '打开文件
                Get #1, i * 1024 + 1, myFile         '将文件写入数组
            Close #1 '关闭文件
            Wj.SendData myFile  '发送
            DoEvents
            RaiseEvent DocumentSending(i, lngFile)
        Next i
        
        AcceptingOrSendingDocument = False
        Dim Bag As New PropertyBag
        Dim bytData() As Byte
        Bag.WriteProperty "lx", "ejswj"
        Bag.WriteProperty "sj", ""
        bytData = Bag.Contents
        Sj.SendData bytData
        RaiseEvent DocumentSendSucceed
        
    End If
    
    Exit Sub
Err:
    RaiseEvent DocumentSendError
    AcceptingOrSendingDocument = False
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 480
End Sub

Private Sub Wj_Close()
    RaiseEvent ConnectClose
    Connected = False
    Sj.Close
    Wj.Close
End Sub

Private Sub Wj_ConnectionRequest(ByVal requestID As Long)
    Wj.Close
    Wj.Accept requestID
End Sub

Private Sub Wj_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo Err
    
    If AcceptingOrSendingDocument = False Then GoTo Err
    If DocumentPath = "" Then GoTo Err
    
    Dim myFile() As Byte
    Dim myLong As Single
    Dim myPath As String
    myPath = DocumentPath
    Open myPath For Binary As #1 '新建文件
        ReDim myFile(bytesTotal) '此处也可以是(0 To bytesTotal-1)
        Wj.GetData myFile
        myLong = FileLen(myPath)
        Put #1, myLong + 1, myFile '将收到的数据写入新文件中
    Close #1 '关闭
    RaiseEvent DocumentArrived
    
Err:
    RaiseEvent DocumentArriveError
End Sub

Private Sub Wj_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent ConnectFail
    Connected = False
    Sj.Close
    Wj.Close
    Timer1.Enabled = False
End Sub
