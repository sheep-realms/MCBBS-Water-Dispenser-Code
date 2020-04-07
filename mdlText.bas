Attribute VB_Name = "mdlText"
Option Explicit

Public Function TextInsert(ByVal Text As String, ByVal Symbol As String) As String
    Dim m As Long, i As Long
    m = Len(Text)
    If m < 2 Then TextInsert = Text: Exit Function
    Dim txt() As String
    ReDim txt(1 To m) As String
    For i = 1 To m
        txt(i) = Mid(Text, i, 1)
    Next i
    TextInsert = Join(txt, Symbol)
End Function

Public Function Fb()
    Dim OnlyText As Boolean
    Dim X As String
    If frm.RButton = False Then
        X = GetFile(FileHdata & "\special\format_brush.txt")
        If X = "__SPR:TEXT__" Then frm.MsgBarOut "您的格式刷中没有格式！", "请左右键同时按下格式刷按钮以设置格式刷。", "red", 8: Exit Function
        If InStr(X, "__SPR:NOFB__") <> 0 Then
            X = Replace(X, "__SPR:NOFB__", "")
            GoTo NOFB
        End If
        If InStr(X, "__SPR:ONLYTEXT__") <> 0 Then
            X = Replace(X, "__SPR:ONLYTEXT__", "")
            OnlyText = True
        End If
        
        X = Replace(X, "__SPR:TEXT__", frm.txtMsg.SelText)
        
        If OnlyText = True Then GoTo NOFB
        
        X = Replace(X, "__SPR:/44__", ",")
        X = Replace(X, "__SPR:/BR__", vbCrLf)
        X = Replace(X, "__SPR:/NOBR__" & vbCrLf, "")
        
        X = Replace(X, "__SPR:STRREVERSE__", StrReverse(frm.txtMsg.SelText))
        X = Replace(X, "__SPR:LCASE__", LCase(frm.txtMsg.SelText))
        X = Replace(X, "__SPR:UCASE__", UCase(frm.txtMsg.SelText))
        
        X = Replace(X, "__SPR:NOW__", Now)
        X = Replace(X, "__SPR:DATE__", Date)
        X = Replace(X, "__SPR:TIME__", Time)
        
NOFB:
        frm.txtMsg.SelText = X
        frm.msgPrint "使用 格式刷"
    Else
        SetPage "Special", "format_brush"
        frm.txtMsg.Text = GetFile(FileHdata & "\special\format_brush.txt")
        frm.RButton = False
        frm.msgPrint "设置 格式刷"
    End If
End Function
