VERSION 5.00
Begin VB.Form ѡ�� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ���ʱ�����ͣ�"
   ClientHeight    =   1575
   ClientLeft      =   8115
   ClientTop       =   5835
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4335
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&E)"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "ȫ������"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton Option2 
         Caption         =   "����ʱ"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "����ʱ"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "ѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Option1 Then Z_JiShi.Show
If Option2 Then D_JiShi.Show
Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub
