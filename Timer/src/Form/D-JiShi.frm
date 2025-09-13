VERSION 5.00
Begin VB.Form D_JiShi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "倒计时"
   ClientHeight    =   3000
   ClientLeft      =   8430
   ClientTop       =   5415
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4695
   Begin VB.CommandButton qx 
      Caption         =   "取消"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton zt 
      Caption         =   "暂停"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton ks 
      Caption         =   "开始"
      Default         =   -1  'True
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   270
      Left            =   840
      TabIndex        =   0
      Text            =   "1"
      Top             =   840
      Width           =   3615
   End
   Begin VB.Timer JiShi 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   2520
   End
   Begin VB.Label Label1 
      Caption         =   "秒数："
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "D_JiShi"
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
    If Text = "0" Then Exit Sub '防止出现bug
    If Text = "" Then Exit Sub '防止发生错误
    Text.Locked = True
    JiShi.Enabled = True
    ks.Enabled = False
    zt.Enabled = True
    qx.Enabled = True
End Sub

Private Sub zt_Click()
JiShi.Enabled = False
ks.Enabled = True
zt.Enabled = False
End Sub

Private Sub qx_Click()
JiShi.Enabled = False
Text = 1
Text.Locked = False
ks.Enabled = True
zt.Enabled = False
qx.Enabled = False
End Sub

Private Sub JiShi_Timer()
Text = Text - 1
If Text.Text = "0" Then
Done
JiShi.Enabled = False
End If
End Sub

Private Sub Done()
ks.Enabled = True
zt.Enabled = False
qx.Enabled = False
Text = 1
Text.Locked = False
MsgBox "倒计时已结束！", vbInformation + vbSystemModal
End Sub
