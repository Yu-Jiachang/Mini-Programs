VERSION 5.00
Begin VB.Form Z_JiShi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ʱ"
   ClientHeight    =   2940
   ClientLeft      =   8385
   ClientTop       =   5145
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4695
   Begin VB.CommandButton qx 
      Caption         =   "ȡ��"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton zt 
      Caption         =   "��ͣ"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton ks 
      Caption         =   "��ʼ"
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
      Caption         =   "ʱ��"
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
      Caption         =   "�룺"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "�֣�"
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
    '��ֹ�󴥰�ť
    If JiShi.Enabled = True Then
        If MsgBox("����ʱ��δ��������ȷ��Ҫ���ؼ�ʱ������ѡ����", vbOKCancel + vbDefaultButton2 + vbQuestion) = vbOK Then
            ѡ��.Show
        Else
            Cancel = True
        End If
    Else
        ѡ��.Show
    End If
End Sub

Private Sub ks_Click()
    JiShi.Enabled = True '������ʱ����
    '�������ð�ť������
    ks.Enabled = False
    zt.Enabled = True
    qx.Enabled = True
End Sub

Private Sub zt_Click()
    JiShi.Enabled = False '��ͣ��ʱ����
    '�������ð�ť������
    ks.Enabled = True
    zt.Enabled = False
End Sub

Private Sub qx_Click()
    JiShi.Enabled = False    '��ͣ��ʱ����
    '�������ð�ť������
    ks.Enabled = True
    zt.Enabled = False
    qx.Enabled = False
    '����ʱ��
    Shi = 0
    Fen = 0
    Miao = 0
End Sub

Private Sub JiShi_Timer()
    Miao = Miao + 1
    'ʵ���Զ���λ
    If Miao = 60 Then
        Fen = Fen + 1
        Miao = 0
    End If
    If Fen = 60 Then
        Shi = Shi + 1
        Fen = 0
    End If
End Sub
