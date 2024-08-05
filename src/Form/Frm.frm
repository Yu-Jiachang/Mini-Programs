VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form Frm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   2085
   ClientTop       =   2850
   ClientWidth     =   5835
   Icon            =   "Frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5835
   StartUpPosition =   1  '所有者中心
   WhatsThisHelp   =   -1  'True
   Begin ComctlLib.ProgressBar JDT 
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   2760
   End
   Begin VB.Label Label2 
      Caption         =   "当前进度："
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label JDTS 
      Caption         =   "0%"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   2400
      Width           =   5715
   End
   Begin VB.Label Label1 
      Caption         =   "正在执行操作......"
      Height          =   195
      Left            =   2100
      TabIndex        =   1
      Top             =   540
      Width           =   1635
   End
End
Attribute VB_Name = "Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MF_BYPOSITION = &H400&
Private Const MF_DISABLED = &H2&
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Function DisableCloseMenu(ByVal dwMenu As Long) As Boolean
    Dim hMenu As Long
    Dim nCount As Long
    DisableCloseMenu = False
    hMenu = GetSystemMenu(dwMenu, False)
    If hMenu <> 0 Then
        nCount = GetMenuItemCount(hMenu)
        If RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION) <> 0 Then
            If DrawMenuBar(dwMenu) <> 0 Then
                DisableCloseMenu = True
            End If
        End If
    End If
End Function

Private Sub Form_Load()
    DisableCloseMenu Me.hwnd
    Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> 1 Then Me.WindowState = 0
End Sub

Private Sub Timer_Timer()
    If JDT = JDT.Max Then End
    JDT = JDT + 1
    JDTS.Caption = JDT / JDT.Max * 100 & "%"
End Sub
