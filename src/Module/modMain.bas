Attribute VB_Name = "modMain"
Sub Main()
    If Command = "" Then GoTo Err:
    Dim cssz
    cssz = Split(Command, ",", 2, vbTextCompare)
    Dim cs1 As Single
    cs1 = cssz(0)
    Dim cs2
    cs2 = cssz(1)
    If Int(cs1) = cs1 Then
        Frm.JDT.Max = cs1 * 2
        Frm.Timer.Enabled = True
        Frm.Caption = cs2
        Frm.Show
    Else
        GoTo Err:
    End If
    
Exit Sub
Err:
MsgBox "命令行参数格式错误！" & vbNewLine & _
"点击确定以查看帮助。", 16
Help.Show
End Sub
