VERSION 5.00
Begin VB.UserControl PUIMgrPlus 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin P¿Ø¼þ¼¯.PUIMgr PUI 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "PUIMgrPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event MoveSmlyComplete(Control As Object)
Public Event ColorSmlyComplete(Index As Integer)
Public Event ColorSmlyIng(Index As Integer, nColor As Long)

Private Sub PUI_ColorSmlyComplete(Index As Integer)
    RaiseEvent ColorSmlyComplete(Index)
    Unload PUI(Index)
End Sub

Private Sub PUI_ColorSmlyIng(Index As Integer, nColor As Long)
    RaiseEvent ColorSmlyIng(Index, nColor)
End Sub

Private Sub PUI_MoveSmlyComplete(Index As Integer, Control As Object)
    RaiseEvent MoveSmlyComplete(Control)
    Unload PUI(Index)
End Sub

Private Sub UserControl_Resize()
    Width = 480
    Height = 480
End Sub

Public Sub MoveSmly(ByRef Control As Object, ByVal nLeft As Long, ByVal nTop As Long, ByVal Delay As Integer, ByVal Index As Integer, Optional ByVal Speed As Integer = 10)
    If (Delay > 1000) Or (Delay < 1) Then Exit Sub
    On Error Resume Next
    Load PUI(Index + 1)
    PUI(Index + 1).MoveSmly Control, nLeft, nTop, Delay, Speed
End Sub

Public Sub ColorSmly(ByVal CurrentColor As Long, ByVal GoalColor As Long, ByVal CGSPD As Integer, ByVal Delay As Integer, ByVal Index As Integer)
    If (Delay > 1000) Or (Delay < 1) Then Exit Sub
    On Error Resume Next
    Load PUI(Index + 1)
    PUI(Index + 1).ColorSmly CurrentColor, GoalColor, CGSPD, Delay
End Sub

