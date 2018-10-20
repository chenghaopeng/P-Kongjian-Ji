VERSION 5.00
Begin VB.UserControl PWeather 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin P控件集.PNet PN 
      Left            =   2040
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "PWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 480
End Sub

'http://wthrcdn.etouch.cn/weather_mini?city=北京
'http://wthrcdn.etouch.cn/WeatherApi?city=北京

Private Function GetCode1(Optional ByVal strCityName As String = "") As String
    If strCityName = "" Then
        strCityName = PN.GetCurrentIPLoaction
        strCityName = Right(strCityName, Len(strCityName) - InStr(strCityName, "省"))
        If Right(strCityName, 1) = "市" Then strCityName = Left(strCityName, Len(strCityName) - 1)
    End If
    Dim strCode As String
    strCode = PN.GetHtmlCodeByXMLHTTP("http://wthrcdn.etouch.cn/weather_mini?city=" & Encode(strCityName))
    If strCode = "{" & Chr(34) & "desc" & Chr(34) & ":" & Chr(34) & "invilad-citykey" & Chr(34) & "," & Chr(34) & "status" & Chr(34) & ":1002}" Then
        GetCode1 = ""
    Else
        GetCode1 = strCode & "|" & strCityName
    End If
End Function

Public Function GetWethInfo_Today(Optional ByVal strCityName As String = "") As String
    Dim strCode As String
    strCode = GetCode1(strCityName)
    If strCode = "" Then
        GetWethInfo_Today = "ERROR"
        Exit Function
    End If
    Dim s() As String
    s = Split(strCode, Chr(34))
    GetWethInfo_Today = Mid(strCode, InStr(strCode, "|") + 1, Len(strCode) - InStr(strCode, "|")) & "|" & s(33) & "|" & s(29) & "|" & s(37) & "|" & s(11) & "|" & s(21) & "|" & s(25) & "|" & s(15) & "|" & s(41)
End Function

Public Function GetWethInfo_Pred(Optional ByVal strCityName As String = "") As String
    Dim strCode As String
    strCode = GetCode1(strCityName)
    If strCode = "" Then
        GetWethInfo_Pred = "ERROR"
        Exit Function
    End If
    Dim s() As String
    s = Split(strCode, Chr(34))
    GetWethInfo_Pred = Mid(strCode, InStr(strCode, "|") + 1, Len(strCode) - InStr(strCode, "|")) & "|" & s(155) & "|" & s(151) & "|" & s(159) & "|" & s(147) & "|" & s(143) & "|" & s(163)
End Function

Public Function GetWethInfo_Succ(Optional ByVal Days As Integer = 1, Optional ByVal strCityName As String = "") As String
    If Days > 4 Or Days < 1 Then
        GetWethInfo_Succ = "ERROR"
        Exit Function
    End If
    Dim strCode As String
    strCode = GetCode1(strCityName)
    If strCode = "" Then
        GetWethInfo_Succ = "ERROR"
        Exit Function
    End If
    Dim s() As String
    s = Split(strCode, Chr(34))
    GetWethInfo_Succ = Mid(strCode, InStr(strCode, "|") + 1, Len(strCode) - InStr(strCode, "|")) & "|" & s(33 + Days * 24) & "|" & s(29 + Days * 24) & "|" & s(37 + Days * 24) & "|" & s(21 + Days * 24) & "|" & s(25 + Days * 24) & "|" & s(41 + Days * 24)
End Function

Private Function Encode(ByVal strUrl As String) As String
    Dim wch  As String, uch As String, szRet As String, X As Long, intLen As Long, nAsc As Long, nAsc2 As Long, nAsc3 As Long
    If strUrl = "" Then
        Encode = strUrl
        Exit Function
    End If
    intLen = Len(strUrl)
    For X = 1 To intLen
        wch = Mid(strUrl, X, 1)
        nAsc = AscW(wch)
        If nAsc < 0 Then nAsc = nAsc + 65536
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    Encode = szRet
End Function


