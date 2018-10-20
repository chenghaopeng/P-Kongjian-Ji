VERSION 5.00
Begin VB.UserControl PUpdate 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin P¿Ø¼þ¼¯.PNet PNet1 
      Left            =   360
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "PUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const NowVersion = "7"

Public Function CheckUpdate() As String
    Dim strCode As String
    strCode = PNet1.GetHtmlCodeByXMLHTTP("http://p.longdows.cn/things/ver/pcs.txt")
    Dim s() As String
    s = Split(strCode, "}")
    Dim i As Integer
    Dim s1 As String, s2 As String
    s1 = Left(s(0), InStr(s(0), "{") - 1)
    s2 = Right(s(0), Len(s(0)) - InStr(s(0), "{"))
    If s2 <> NowVersion Then
        CheckUpdate = s2
        For i = 1 To UBound(s) - 1
            s1 = Left(s(i), InStr(s(i), "{") - 1)
            s2 = Right(s(i), Len(s(i)) - InStr(s(i), "{"))
            Select Case s1
                Case "note"
                    CheckUpdate = CheckUpdate & "*****" & Replace(s2, "\n", vbCrLf)
                Case "downloadurl"
                    CheckUpdate = CheckUpdate & "*****" & s2
            End Select
        Next
    End If
End Function

Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 480
End Sub
