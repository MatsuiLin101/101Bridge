VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "玩法說明"
   ClientHeight    =   3225
   ClientLeft      =   1275
   ClientTop       =   1770
   ClientWidth     =   5055
   Icon            =   "99-4.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "這是我設計的電腦超級弱智版橋牌"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label9 
      Caption         =   "敬請期待"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "如:花橋、2人PK橋、連線對戰版、電腦必贏版、18禁版...等"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label Label7 
      Caption         =   "以後會陸續推出其他版本"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "點數小於4點可以倒牌(可以幫對家倒，電腦不會倒)"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "而是我們班平常在玩的玩法"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "並不是一般比賽所玩的規則"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "其他有問題請問我"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "規則:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Unload Form5
Form1.Show
End Sub

Private Sub Form_Load()
CloseEnd Me.hwnd
End Sub

