VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "叫牌"
   ClientHeight    =   3225
   ClientLeft      =   8430
   ClientTop       =   1200
   ClientWidth     =   4575
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2225.952
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   4296.162
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "牌太爛!!!倒牌!!!"
      Height          =   975
      Left            =   0
      TabIndex        =   36
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   8
      Left            =   1080
      Picture         =   "99-7.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   35
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   13
      Left            =   1080
      Picture         =   "99-7.frx":054E
      Style           =   1  '圖片外觀
      TabIndex        =   34
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   18
      Left            =   1080
      Picture         =   "99-7.frx":0A9C
      Style           =   1  '圖片外觀
      TabIndex        =   33
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   6
      Left            =   360
      Picture         =   "99-7.frx":0FEA
      Style           =   1  '圖片外觀
      TabIndex        =   32
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   16
      Left            =   360
      Picture         =   "99-7.frx":1538
      Style           =   1  '圖片外觀
      TabIndex        =   31
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   21
      Left            =   360
      Picture         =   "99-7.frx":1A86
      Style           =   1  '圖片外觀
      TabIndex        =   30
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   7
      Left            =   720
      Picture         =   "99-7.frx":1FD4
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   12
      Left            =   720
      Picture         =   "99-7.frx":2522
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   17
      Left            =   720
      Picture         =   "99-7.frx":2A70
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   5
      Left            =   0
      Picture         =   "99-7.frx":2FBE
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   10
      Left            =   0
      Picture         =   "99-7.frx":350C
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   15
      Left            =   0
      Picture         =   "99-7.frx":3A5A
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   20
      Left            =   0
      Picture         =   "99-7.frx":3FA8
      Style           =   1  '圖片外觀
      TabIndex        =   23
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1N"
      Height          =   255
      Index           =   4
      Left            =   1440
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2N"
      Height          =   255
      Index           =   9
      Left            =   1440
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3N"
      Height          =   255
      Index           =   14
      Left            =   1440
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4N"
      Height          =   255
      Index           =   19
      Left            =   1440
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   0
      Left            =   0
      Picture         =   "99-7.frx":44F6
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   1
      Left            =   360
      Picture         =   "99-7.frx":4A44
      Style           =   1  '圖片外觀
      TabIndex        =   17
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   2
      Left            =   720
      Picture         =   "99-7.frx":4F92
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   3
      Left            =   1080
      Picture         =   "99-7.frx":54E0
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   11
      Left            =   360
      Picture         =   "99-7.frx":5A2E
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   22
      Left            =   720
      Picture         =   "99-7.frx":5F7C
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   23
      Left            =   1080
      Picture         =   "99-7.frx":64CA
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5N"
      Height          =   255
      Index           =   24
      Left            =   1440
      TabIndex        =   11
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   25
      Left            =   0
      Picture         =   "99-7.frx":6A18
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   26
      Left            =   360
      Picture         =   "99-7.frx":6F66
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   27
      Left            =   720
      Picture         =   "99-7.frx":74B4
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   28
      Left            =   1080
      Picture         =   "99-7.frx":7A02
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6N"
      Height          =   255
      Index           =   29
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   30
      Left            =   0
      Picture         =   "99-7.frx":7F50
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   31
      Left            =   360
      Picture         =   "99-7.frx":849E
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   32
      Left            =   720
      Picture         =   "99-7.frx":89EC
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   33
      Left            =   1080
      Picture         =   "99-7.frx":8F3A
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7N"
      Height          =   255
      Index           =   34
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pass"
      Height          =   255
      Index           =   35
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3225
      ScaleWidth      =   4545
      TabIndex        =   37
      Top             =   0
      Width           =   4575
      Begin VB.Label Label5 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00008000&
         Caption         =   "各家擁有點數"
         Height          =   255
         Left            =   2040
         TabIndex        =   42
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00008000&
         Caption         =   "Label4"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   41
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00008000&
         Caption         =   "Label3"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   40
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00008000&
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   39
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00008000&
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   38
         Top             =   2160
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

If Index = 0 Then
    WhoIsKing 1
    Me.Command1(0).Enabled = False
End If

If Index = 1 Then
    WhoIsKing 2
    For i = 0 To 1
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 2 Then
    WhoIsKing 3
    For i = 0 To 2
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 3 Then
    WhoIsKing 4
    For i = 0 To 3
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 4 Then
    WhoIsKing 5
    For i = 0 To 4
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 5 Then
    WhoIsKing 6
    For i = 0 To 5
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 6 Then
    WhoIsKing 7
    For i = 0 To 6
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 7 Then
    WhoIsKing 8
    For i = 0 To 7
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 8 Then
    WhoIsKing 9
    For i = 0 To 8
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 9 Then
    WhoIsKing 10
    For i = 0 To 9
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 10 Then
    WhoIsKing 11
    For i = 0 To 10
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 11 Then
    WhoIsKing 12
    For i = 0 To 11
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 12 Then
    WhoIsKing 13
    For i = 0 To 12
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 13 Then
    WhoIsKing 14
    For i = 0 To 13
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 14 Then
    WhoIsKing 15
    For i = 0 To 14
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 15 Then
    WhoIsKing 16
    For i = 0 To 15
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 16 Then
    WhoIsKing 17
    For i = 0 To 16
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 17 Then
    WhoIsKing 18
    For i = 0 To 17
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 18 Then
    WhoIsKing 19
    For i = 0 To 18
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 19 Then
    WhoIsKing 20
    For i = 0 To 19
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 20 Then
    WhoIsKing 21
    For i = 0 To 20
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 21 Then
    WhoIsKing 22
    For i = 0 To 21
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 22 Then
    WhoIsKing 23
    For i = 0 To 22
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 23 Then
    WhoIsKing 24
    For i = 0 To 23
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 24 Then
    WhoIsKing 25
    For i = 0 To 24
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 25 Then
    WhoIsKing 26
    For i = 0 To 25
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 26 Then
    WhoIsKing 27
    For i = 0 To 26
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 27 Then
    WhoIsKing 28
    For i = 0 To 27
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 28 Then
    WhoIsKing 29
    For i = 0 To 28
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 29 Then
    WhoIsKing 30
    For i = 0 To 29
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 30 Then
    WhoIsKing 31
    For i = 0 To 30
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 31 Then
    WhoIsKing 32
    For i = 0 To 31
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 32 Then
    WhoIsKing 33
    For i = 0 To 32
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 33 Then
    WhoIsKing 34
    For i = 0 To 33
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 34 Then
    WhoIsKing 35
    For i = 0 To 34
    Me.Command1(i).Enabled = False
    Next i
End If

If Index = 35 Then
    WhoIsKing 36
End If

End Sub

Private Sub Command2_Click()

Unload Form8
NewGamePut

End Sub

Private Sub Form_Load()

'For i = 0 To 3
'    Form8.Label1(i).Caption = PlayerPoint(i)
'Next i

'If Label1(0) < 4 Then Me.Command2.Enabled = True
'If Label1(2) < 4 Then Me.Command2.Enabled = True

CloseEnd Me.hwnd

End Sub

