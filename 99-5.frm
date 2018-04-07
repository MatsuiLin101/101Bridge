VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "關於"
   ClientHeight    =   2055
   ClientLeft      =   1770
   ClientTop       =   1770
   ClientWidth     =   4410
   Icon            =   "99-5.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "http://www.wretch.cc/blog/lssh101"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      MousePointer    =   10  '往上指
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "我的網站:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "我的E-mail:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "tw19900703@hotmail.com"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1800
      MousePointer    =   10  '往上指
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "99-5.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Unload Form6
Form1.Show
End Sub

Private Sub Form_Load()
CloseEnd Me.hwnd
End Sub

Private Sub Label4_Click()
xreturn = Shell("start.exe mailto:syc837@ms8.hinet.net", 0)
End Sub

Private Sub Label7_Click()
xreturn = Shell("start.exe http://netcity.hinet.net/syc837", 0)
End Sub
