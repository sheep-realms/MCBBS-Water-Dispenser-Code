Attribute VB_Name = "mdlFunOutput"
Option Explicit

Public Function FunOutput(ByVal FunName As String)
On Error GoTo Error
    Dim X As String
    X = frm.txtMsg.SelText
    If frm.txtMsg.SelText = "" Then
        MsgBox "��û��ѡ���κ����ݣ�", vbExclamation, "��������Ϊ��"
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
        msg = "�벻Ҫ������߻���͵���ֵ��" & vbCrLf & _
              "INT: -32,768 �� 32,767" & vbCrLf & _
              "LONG: -2,147,483,648 �� 2,147,483,647" & vbCrLf & _
              "������������ѯ������ϡ�"
    Case 13
        msg = "��ֵ���͵ĺ������������ֵ���롣"
    End Select
    MsgBox "������������" & vbCrLf & "* ������룺" & Err.Number & vbCrLf & "* ����������" & Err.Description & vbCrLf & msg, vbCritical, "����"
End Function
