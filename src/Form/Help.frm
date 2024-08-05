VERSION 5.00
Begin VB.Form Help 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "程序使用方式"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6495
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "参数说明"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6255
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.ListBox ListBox 
      Height          =   420
      ItemData        =   "Help.frx":C4F2
      Left            =   120
      List            =   "Help.frx":C4FC
      TabIndex        =   2
      Top             =   720
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "命令行参数格式"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Label Label1 
         Caption         =   "bin.exe number,[title]"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Me.WindowState = 0
End Sub

Private Sub ListBox_Click()
    Select Case ListBox
        Case "number"
            Label2 = "必需的。指定进度条持续的时间(单位：秒)。"
        Case "title"
            Label2 = "可选的。指定窗口标题。如果没有指定，则相当于长度为零的字符串("""")。"
    End Select
End Sub
