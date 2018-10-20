VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.UserControl PNet 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   30
      ExtentX         =   53
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "PNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetHtmlCodeByXMLHTTP(ByVal strUrl As String) As String
    Dim xh As Object
    Set xh = CreateObject("Microsoft.XMLHTTP")
    xh.Open "get", strUrl, True
    xh.send
    While xh.ReadyState <> 4
        DoEvents
    Wend
    GetHtmlCodeByXMLHTTP = BytesToBstr(xh.ResponseBody)
    Set xh = Nothing
End Function

Public Function GetHtmlCodeByInet(ByVal strUrl As String) As String
    GetHtmlCodeByInet = Inet1.OpenURL(strUrl)
End Function

Public Function GetHtmlCodeByWebbrowser(ByVal strUrl As String) As String
    Web.Navigate strUrl
    DoEvents
    Dim doc As Object
    Dim i As Object
    Set doc = Web.Document
    For Each i In doc.All
        GetHtmlCodeByWebbrowser = GetHtmlCodeByWebbrowser & Chr(13) & i.innerHtml
    Next
End Function

Public Function GetCurrentIP() As String
    'On Error Resume Next
    Dim c As String
    c = GetHtmlCodeByXMLHTTP("http://www.ip138.com/ips138.asp")
    Dim F1 As String, F2 As String
    F1 = "您的IP地址是：["
    F2 = "] 来自："
    Dim L1 As Long, L2 As Long
    L1 = InStr(c, F1) + 9
    L2 = InStr(c, F2)
    GetCurrentIP = Mid(c, L1, L2 - L1)
End Function

Public Function GetCurrentIPLoaction() As String
    'On Error Resume Next
    Dim c As String
    c = GetHtmlCodeByXMLHTTP("http://www.ip138.com/ips138.asp")
    Dim F1 As String
    F1 = "] 来自："
    Dim L1 As Long, L2 As Long
    L1 = InStr(c, F1) + 5
    L2 = L1
    Do Until Mid(c, L2, 1) = " "
        L2 = L2 + 1
    Loop
    GetCurrentIPLoaction = Mid(c, L1, L2 - L1)
End Function

Public Function GetCurrentIPOperator() As String
    'On Error Resume Next
    Dim c As String
    c = GetHtmlCodeByXMLHTTP("http://www.ip138.com/ips138.asp")
    Dim F1 As String
    F1 = "] 来自："
    Dim L1 As Long, L2 As Long
    L1 = InStr(c, F1) + 5
    Do Until Mid(c, L1, 1) = " "
        L1 = L1 + 1
    Loop
    L1 = L1 + 1
    L2 = L1
    Do Until Mid(c, L2, 1) = "<"
        L2 = L2 + 1
    Loop
    GetCurrentIPOperator = Mid(c, L1, L2 - L1)
End Function

Public Function DownloadFile(strUrl As String, strSavePath As String) As Boolean
        DownloadFile = URLDownloadToFile(0&, strUrl, strSavePath, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
End Function



Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 480
End Sub



'以下代码摘自http://www.newxing.com/Tech/Program/VisualBasic/XMLHttp_392.html
Private Function BytesToBstr(Bytes)
    Dim Unicode As String
    If IsUTF8(Bytes) Then
        Unicode = "UTF-8"
    Else
        Unicode = "GB2312"
    End If
    Dim objstream As Object
    Set objstream = CreateObject("ADODB.Stream")
    With objstream
        .Type = 1
        .mode = 3
        .Open
        .Write Bytes
        .Position = 0
        .Type = 2
        .Charset = Unicode
        BytesToBstr = .ReadText
       .Close
    End With
End Function

Private Function IsUTF8(Bytes) As Boolean
        Dim i As Long, AscN As Long, Length As Long
        Length = UBound(Bytes) + 1
        If Length < 3 Then
            IsUTF8 = False
            Exit Function
        ElseIf Bytes(0) = &HEF And Bytes(1) = &HBB And Bytes(2) = &HBF Then
            IsUTF8 = True
            Exit Function
        End If
        Do While i <= Length - 1
            If Bytes(i) < 128 Then
                i = i + 1
                AscN = AscN + 1
            ElseIf (Bytes(i) And &HE0) = &HC0 And (Bytes(i + 1) And &HC0) = &H80 Then
                i = i + 2
            ElseIf i + 2 < Length Then
                If (Bytes(i) And &HF0) = &HE0 And (Bytes(i + 1) And &HC0) = &H80 And (Bytes(i + 2) And &HC0) = &H80 Then
                     i = i + 3
                Else
                    IsUTF8 = False
                    Exit Function
                End If
            Else
                IsUTF8 = False
                Exit Function
            End If
        Loop
        If AscN = Length Then
            IsUTF8 = False
        Else
            IsUTF8 = True
        End If
End Function
