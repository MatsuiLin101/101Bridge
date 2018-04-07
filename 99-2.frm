VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "歡迎光臨"
   ClientHeight    =   2970
   ClientLeft      =   1335
   ClientTop       =   1485
   ClientWidth     =   4425
   Icon            =   "99-2.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame2 
      Caption         =   "電腦速度"
      Height          =   1095
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
      Begin VB.OptionButton Option1 
         Caption         =   "緩慢"
         Height          =   255
         Index           =   45
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "中等"
         Height          =   255
         Index           =   30
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "最快"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "各家電腦名稱"
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Text            =   "電腦3"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Text            =   "電腦2"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Text            =   "電腦1"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Text            =   "玩家"
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "尊姓大名?"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "歡迎加入"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'Form1.Enabled = True
'Kill "101bridge.nam"
'Open "101bridge.nam" For Output As #1
'For i = 0 To 3
'a = Text(i).Text
'Print #1, a
'Form1.labell(i).Caption = a
'Next i
'Close #1
Private Sub Command1_Click()
Form1.Enabled = True
'Kill "101bridge.nam"
'Open "101bridge.nam" For Output As #1
For i = 0 To 3
a = Text1(i).Text
'Print #1, a
Form1.Label1(i).Caption = a
'Form4.Label1(i).Caption = a
Next i
'Close #1
Unload Form3
'Speed = 1
'g_num = 1
Form1.Show
NewGamePut
comx(2, 1) = 30
comx(3, 1) = 224
comx(4, 1) = 420
comx(2, 2) = 43
comx(3, 2) = 233
comx(4, 2) = 433
comx(2, 3) = 47
comx(3, 3) = 237
comx(4, 3) = 437
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
GameSpeed = 15
Option1(GameSpeed).Value = True
Form1.Enabled = False
CloseEnd Me.hwnd
'Open "101bridge.nam" For Input As #1
'For i = 0 To 3
'Input #1, a
'Text1(i).Text = a
'Next i
'Close #1
End Sub

Private Sub Option1_Click(Index As Integer)

GameSpeed = Index

End Sub

