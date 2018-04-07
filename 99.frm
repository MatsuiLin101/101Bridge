VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "101bridge"
   ClientHeight    =   6030
   ClientLeft      =   2685
   ClientTop       =   2880
   ClientWidth     =   9870
   Icon            =   "99.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "99.frx":030A
   ScaleHeight     =   6030
   ScaleWidth      =   9870
   Begin VB.CommandButton Command5 
      Caption         =   "關於我"
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "說明"
      Height          =   375
      Left            =   8280
      TabIndex        =   19
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "選項"
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開新牌局"
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   8040
      ScaleHeight     =   675
      ScaleWidth      =   1755
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "敵方"
      Height          =   2415
      Left            =   8880
      TabIndex        =   9
      Top             =   2040
      Width           =   975
      Begin VB.PictureBox Picture7 
         Appearance      =   0  '平面
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   150
         Picture         =   "99.frx":0858
         ScaleHeight     =   2070
         ScaleWidth      =   645
         TabIndex        =   13
         Top             =   240
         Width           =   670
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "我方"
      Height          =   2415
      Left            =   8040
      TabIndex        =   8
      Top             =   2040
      Width           =   975
      Begin VB.PictureBox Picture6 
         Appearance      =   0  '平面
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   110
         Picture         =   "99.frx":4AF6
         ScaleHeight     =   2070
         ScaleWidth      =   645
         TabIndex        =   12
         Top             =   240
         Width           =   670
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "叫牌"
      Height          =   855
      Index           =   0
      Left            =   8040
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   397
      ScaleMode       =   3  '像素
      ScaleWidth      =   533
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Height          =   255
         Index           =   3
         Left            =   6600
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00008000&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   2
         Top             =   5520
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   0
      Picture         =   "99.frx":8D94
      ScaleHeight     =   358
      ScaleMode       =   3  '像素
      ScaleWidth      =   156
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      Begin VB.PictureBox Picture4 
         Appearance      =   0  '平面
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   0
         Picture         =   "99.frx":16BFE
         ScaleHeight     =   96
         ScaleMode       =   3  '像素
         ScaleWidth      =   71
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   2400
      Picture         =   "99.frx":1BD40
      ScaleHeight     =   3000
      ScaleWidth      =   3150
      TabIndex        =   10
      Top             =   120
      Width           =   3150
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      Picture         =   "99.frx":3AB42
      ScaleHeight     =   375
      ScaleWidth      =   3150
      TabIndex        =   14
      Top             =   3360
      Width           =   3150
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   5400
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2520
      Picture         =   "99.frx":3E93C
      ScaleHeight     =   375
      ScaleWidth      =   1200
      TabIndex        =   15
      Top             =   4080
      Width           =   1230
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2520
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   16
      Top             =   4800
      Width           =   670
   End
   Begin VB.Menu game 
      Caption         =   "牌局(&G)"
      Begin VB.Menu NewGame 
         Caption         =   "新的牌局(&N)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu set 
         Caption         =   "選項(&S)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu abc 
         Caption         =   "-"
      End
      Begin VB.Menu quit 
         Caption         =   "結束(&X)"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "說明(&H)"
      Begin VB.Menu Help1 
         Caption         =   "玩法說明(&H)"
      End
      Begin VB.Menu abc1 
         Caption         =   "-"
      End
      Begin VB.Menu Help2 
         Caption         =   "關於(&A)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()

NewGamePut

End Sub

Private Sub Command2_Click()
Form1.Enabled = False
Load Form2
Form2.Show
End Sub

Private Sub Command3_Click(Index As Integer)

Me.Command3(0).Enabled = False

Form8.Command2.Enabled = False
If PlayerPoint(0) < 4 Then Form8.Command2.Enabled = True
If PlayerPoint(2) < 4 Then Form8.Command2.Enabled = True

Form8.Show

'If Index = 0 Then
'    For i = 0 To 35
'    Me.Command1(i).Enabled = True
'    Next i
    'Me.Command2.Enabled = True
'    Me.Command3(0).Enabled = False
'End If

End Sub

Private Sub Command4_Click()
Form1.Enabled = False
Load Form5
Form5.Show
End Sub

Private Sub Command5_Click()
Form1.Enabled = False
Load Form6
Form6.Show
End Sub

Private Sub Form_Load()
GameKing = 0
Form1.Show
Load Form3
Load Form4
Load Form7
Form3.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Help1_Click()
Form1.Enabled = False
Load Form5
Form5.Show
End Sub

Private Sub Help2_Click()
Form1.Enabled = False
Load Form6
Form6.Show
End Sub

Private Sub NewGame_Click()
Form1.Picture1.Line (0, 0)-(1000, 1000), &H8000&, BF
Form1.Picture5.Line (0, 0)-(500, 500), &H8000&, BF
NewGamePut
GameKing = 0
pass = 0
Me.Command3(0).Enabled = True
'g_num = 1
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'If over = 0 Then
CheckCards
If GameKing = 0 Then GoTo out
If ReturnPlayer = 0 Then GoTo out     '檢查換玩家了沒

If y > 295 And y < 391 Then
   If x > 120 And x < 140 + CardsPlace(1) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 1: Exit Sub
       End If
   End If
   If x > 140 And x < 160 + CardsPlace(2) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 2: Exit Sub
      End If
   End If
   If x > 160 And x < 180 + CardsPlace(3) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 3: Exit Sub
      End If
   End If
   If x > 180 And x < 200 + CardsPlace(4) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 4: Exit Sub
      End If
   End If
   If x > 200 And x < 220 + CardsPlace(5) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 5: Exit Sub
      End If
   End If
   If x > 220 And x < 240 + CardsPlace(6) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 6: Exit Sub
       End If
   End If
   If x > 240 And x < 260 + CardsPlace(7) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 7: Exit Sub
       End If
   End If
   If x > 260 And x < 280 + CardsPlace(8) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 8: Exit Sub
       End If
   End If
   If x > 280 And x < 300 + CardsPlace(9) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 9: Exit Sub
       End If
   End If
   If x > 300 And x < 320 + CardsPlace(10) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 10: Exit Sub
       End If
   End If
   If x > 320 And x < 340 + CardsPlace(11) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 11: Exit Sub
       End If
   End If
   If x > 340 And x < 360 + CardsPlace(12) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 12: Exit Sub
       End If
   End If
   If x > 360 And x < 430 + CardsPlace(13) Then
      If Button = 1 Then
         If sp = 0 Then UserShowCards 13: Exit Sub
       End If
   End If
End If
out:
End Sub

Private Sub quit_Click()
End
End Sub

Private Sub resetplay_Click()
Form1.Enabled = False
Form7.Show
End Sub

Private Sub set_Click()
Form1.Enabled = False
Load Form2
Form2.Show
End Sub

Private Sub Timer1_Timer()

Timer1.Enabled = False

End Sub
