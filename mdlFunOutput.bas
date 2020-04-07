Attribute VB_Name = "mdlFunOutput"
Option Explicit

Public Function FunOutput(ByVal FunName As String)
On Error GoTo Error
    Dim X As String
    X = frm.txtMsg.SelText
    If frm.txtMsg.SelText = "" Then
        MsgBox "您没有选择任何内容！", vbExclamation, "输入数据为空"
    Else
        Select Case FunName
        Case "asc": Output Asc(X)
        Case "chr": Output cHr(X)
        Case "cint": Output CInt(X)
        Case "fix": Output Fix(X)
        Case "hex": Output Hex(X)
        Case "int": Output Int(X)
        Case "oct": Output Oct(X)
        
        Case "abs": Output Abs(X)
        Case "atn": Output Atn(X)
        Case "cos": Output Cos(X)
        Case "exp": Output Exp(X)
        Case "log": Output Log(X)
        Case "sgn": Output Sgn(X)
        Case "sin": Output Sin(X)
        Case "sqr": Output Sqr(X)
        Case "tan": Output Tan(X)
        End Select
    End If
    
    Exit Function
    
Error:
    Dim msg As String
    Select Case Err.Number
    Case 6
        msg = "请不要输入过高或过低的数值。" & vbCrLf & _
              "INT: -32,768 至 32,767" & vbCrLf & _
              "LONG: -2,147,483,648 至 2,147,483,647" & vbCrLf & _
              "详情请上网查询相关资料。"
    Case 13
        msg = "数值类型的函数必须接受数值输入。"
    End Select
    MsgBox "输入数据有误！" & vbCrLf & "* 错误代码：" & Err.Number & vbCrLf & "* 错误描述：" & Err.Description & vbCrLf & msg, vbCritical, "错误"
End Function
