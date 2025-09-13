VERSION 5.00
Begin VB.Form Z_JiShi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "正计时"
   ClientHeight    =   2940
   ClientLeft      =   8385
   ClientTop       =   5145
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4695
   Begin VB.CommandButton qx 
      Caption         =   "取消"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton zt 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton ks 
      Caption         =   "开始"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Timer JiShi 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   2520
   End
   Begin VB.Label Shi 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "时："
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Miao 
      Caption         =   "0"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Fen 
      Caption         =   "0"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "秒："
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "分："
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "Z_JiShi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '防止误触按钮
    If JiShi.Enabled = True Then
        If MsgBox("倒计时还未结束，您确定要返回计时器类型选择吗？", vbOKCancel + vbDefaultButton2 + vbQuestion) = vbOK Then
            选择.Show
        Else
            Cancel = True
        End If
    Else
        选择.Show
    End If
End Sub

Private Sub ks_Click()
    JiShi.Enabled = True '启动计时代码
    '重新设置按钮可用性
    ks.Enabled = False
    zt.Enabled = True
    qx.Enabled = True
End Sub

Private Sub zt_Click()
    JiShi.Enabled = False '暂停计时代码
    '重新设置按钮可用性
    ks.Enabled = True
    zt.Enabled = False
End Sub

Private Sub qx_Click()
    JiShi.Enabled = False    '暂停计时代码
    '重新设置按钮可用性
    ks.Enabled = True
    zt.Enabled = False
    qx.Enabled = False
    '清零时间
    Shi = 0
    Fen = 0
    Miao = 0
End Sub

Private Sub JiShi_Timer()
    Miao = Miao + 1
    '实现自动进位
    If Miao = 60 Then
        Fen = Fen + 1
        Miao = 0
    End If
    If Fen = 60 Then
        Shi = Shi + 1
        Fen = 0
    End If
End Sub
