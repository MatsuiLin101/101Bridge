VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "牌局記錄"
   ClientHeight    =   3105
   ClientLeft      =   2625
   ClientTop       =   1920
   ClientWidth     =   3135
   Icon            =   "99-6.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If fp = 1 Then
   Form4.Enabled = True
   Form4.Show
Else
    Form1.Enabled = True
    Form1.Show
End If
Form7.Hide
End Sub

Private Sub Form_Load()
CloseEnd Me.hwnd
End Sub
