VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "���k����"
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
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command1 
      Caption         =   "�T�w"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "�o�O�ڳ]�p���q���W�Ůz�������P"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label9 
      Caption         =   "�q�д���"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "�p:����B2�HPK���B�s�u��Ԫ��B�q����Ĺ���B18�T��...��"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label Label7 
      Caption         =   "�H��|������X��L����"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "�I�Ƥp��4�I�i�H�˵P(�i�H����a�ˡA�q�����|��)"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "�ӬO�ڭ̯Z���`�b�������k"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "�ä��O�@����ɩҪ����W�h"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "��L�����D�аݧ�"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "�W�h:"
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

