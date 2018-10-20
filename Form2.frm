VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   255
      Left            =   4080
      TabIndex        =   30
      Top             =   2160
      Width           =   255
      ExtentX         =   450
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
   Begin RichTextLib.RichTextBox bas 
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   27
      Tag             =   "At司仪菌_滚动.bas"
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":0000
   End
   Begin VB.ListBox List1 
      Height          =   1680
      ItemData        =   "Form2.frx":0D4B
      Left            =   2400
      List            =   "Form2.frx":0D9A
      TabIndex        =   26
      Top             =   480
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":0E9D
   End
   Begin VB.TextBox jss 
      Height          =   270
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":8D0F
      Top             =   0
      Width           =   255
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":F8A8
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":12053
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":15D37
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   5
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":18B21
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":21F66
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":24231
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   7
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":27037
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   8
      Left            =   1080
      TabIndex        =   9
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":2ED71
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   10
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":312A6
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   10
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":32E45
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   11
      Left            =   360
      TabIndex        =   12
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":36541
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   12
      Left            =   720
      TabIndex        =   13
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":39B08
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   13
      Left            =   1080
      TabIndex        =   14
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":3AEC0
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   14
      Left            =   1440
      TabIndex        =   15
      Top             =   1080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":3CE24
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   15
      Left            =   0
      TabIndex        =   16
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":40287
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   16
      Left            =   360
      TabIndex        =   17
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":42BE6
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   17
      Left            =   720
      TabIndex        =   18
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":46E3C
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   18
      Left            =   1080
      TabIndex        =   19
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":4B235
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   19
      Left            =   1440
      TabIndex        =   20
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":4E4DE
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   20
      Left            =   0
      TabIndex        =   21
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":4EE41
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   21
      Left            =   360
      TabIndex        =   22
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":4F590
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   22
      Left            =   720
      TabIndex        =   23
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":53860
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   23
      Left            =   1080
      TabIndex        =   24
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":54C5A
   End
   Begin RichTextLib.RichTextBox Code 
      Height          =   375
      Index           =   24
      Left            =   1440
      TabIndex        =   25
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":5CE02
   End
   Begin RichTextLib.RichTextBox bas 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   28
      Tag             =   "Functions.bas"
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":5F9C6
   End
   Begin RichTextLib.RichTextBox bas 
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   29
      Tag             =   "MouseLeave.bas"
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form2.frx":62681
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
