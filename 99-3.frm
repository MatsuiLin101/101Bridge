VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  '雙線固定對話方塊
   ClientHeight    =   645
   ClientLeft      =   2190
   ClientTop       =   2055
   ClientWidth     =   2265
   Icon            =   "99-3.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   43
   ScaleMode       =   3  '像素
   ScaleWidth      =   151
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command2 
      Caption         =   "離開遊戲"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開新牌局"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Height          =   225
      Index           =   3
      Left            =   4200
      TabIndex        =   4
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label3 
      Height          =   225
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label3 
      Height          =   225
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   225
   End
   Begin VB.Label Label3 
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   225
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Form1.Enabled = Ture
Form4.Hide
Form1.Show
NewGamePut
'Form1.Enabled = True
'If g_num <> 5 Then
'  g_num = g_num + 1
'Else
'  For i = 0 To 19
'  Label2(i).Caption = ""
'  Next i
'  For i = 0 To 3
'  Label3(i).Caption = ""
'  Next i
'  g_num = 1
'End If
'Form4.Hide
'Form4.Command1.Caption = "下一局"
'Form1.Show
'NewGamePut
End Sub

Private Sub Command2_Click()
End
'fp = 1x
'Form7.Show
End Sub

Private Sub Form_Load()
CloseEnd Me.hwnd
End Sub
