Attribute VB_Name = "Module1"
Option Explicit

Public FileHead As String            '软件本体位置
Public FileHdata As String           '数据文件地址
Public FileDatapack As String        '数据包位置
Public FileDatapackName As String    '数据包名称

'Public Page As String
'Public PageData As String            '页面
'Public PageUp As String              '上一个页面
Public PageSelStart As Long          '文本选择位置
Public PageSelLength As Long         '文本选择长度
Public PageMode As Boolean           '页面模式
Public TopMode As Boolean            '置顶记录

'[不需要，删了]

Public BBInputMode As Boolean
Public BBInputV(2) As String
Public BBInputL(2) As String
Public BBInputC As String

Public Function Output(ByVal Text As String, Optional ByVal All As Boolean)
    Dim X As Long
    X = frm.txtMsg.SelStart
    
    If All = False Then
        frm.txtMsg.SelText = Text
        frm.txtMsg.SetFocus
        frm.txtMsg.SelStart = X
        frm.txtMsg.SelLength = Len(Text)
    Else
        frm.txtMsg.Text = Text
        frm.txtMsg.SetFocus
        frm.txtMsg.SelStart = 0
        'frm.txtMsg.SelLength = Len(frm.txtMsg.SelText)
    End If
End Function

Public Function BarMsgOut(ByVal Text As String)
    frm.labBar.Caption = "[" & Time & "] " & Text
End Function

Public Function NumChange(ByVal NumType As String, Value As Integer) As String
    Select Case NumType
    Case "ch"
        Select Case Value
        Case 0: NumChange = "零"
        Case 1: NumChange = "一"
        Case 2: NumChange = "二"
        Case 3: NumChange = "三"
        Case 4: NumChange = "四"
        Case 5: NumChange = "五"
        Case 6: NumChange = "六"
        Case 7: NumChange = "七"
        Case 8: NumChange = "八"
        Case 9: NumChange = "九"
        Case 10: NumChange = "十"
        Case 11: NumChange = "十一"
        Case 12: NumChange = "十二"
        Case 13: NumChange = "十三"
        Case 14: NumChange = "十四"
        Case 15: NumChange = "十五"
        Case 16: NumChange = "十六"
        Case 17: NumChange = "十七"
        Case 18: NumChange = "十八"
        Case 19: NumChange = "十九"
        Case 20: NumChange = "二十"
        Case Else: NumChange = "[哪有这种等级你自裁罢！]"
        End Select
    Case "num"
        If Value <> 9 Then
            NumChange = Value
        Else
            Randomize
            If (Int(Rnd * (10 - 0 + 1)) + 0) > 2 Then
                NumChange = "⑨"
            Else
                NumChange = Value
            End If
        End If
    End Select
End Function

Public Function CanSet(ByVal Index As Integer, Optional ByVal Text As String)
    frm.optCan(Index).Caption = Text
    frm.optCan(Index + 11).Visible = False
    frm.optCan(Index + 11).Enabled = False
End Function

Public Function CanInSet(ByVal Index As Integer, Optional ByVal Text As String)
    frm.txtCan(Index).Visible = True
    frm.txtCan(Index).Enabled = True
    frm.txtCan(Index).Text = Text
End Function

Public Function CanListSet(ByVal Count As Integer)
    Dim i As Integer
    frm.optCan(0).Value = True
    frm.optCan(i).Visible = True
    frm.optCan(i).Enabled = True
    For i = 1 To frm.optCan.Count - 1
        If i <= Count - 1 Then
            frm.optCan(i).Visible = True
            frm.optCan(i).Enabled = True
        Else
            frm.optCan(i).Visible = False
            frm.optCan(i).Enabled = False
        End If
        frm.optCan(i).Value = False
    Next i
    For i = 0 To frm.txtCan.Count - 1
        frm.txtCan(i).Text = ""
        frm.txtCan(i).Visible = False
        frm.txtCan(i).Enabled = False
    Next i
End Function

'////////////////////代码生成函数////////////////////////////////////////////////////////////

Public Function BBCode(ByVal Code As String, Optional ByVal V1 As String, Optional ByVal V2 As String, Optional ByVal V3 As String, Optional ByVal Mode As Boolean)
    Dim X As Long
    Dim i As Integer
    Dim j As Long
    
    If V1 = "" Then V1 = frm.txtMsg.SelText
    i = Len(Code) + 2
    j = Len(V1)
    X = frm.txtMsg.SelStart
    
    If V2 = "" And V3 = "" And Mode = False Then
        If V1 = "" Then
            frm.txtMsg.SelText = "[" & Code & "][/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
        Else
            frm.txtMsg.SelText = "[" & Code & "]" & V1 & "[/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
            frm.txtMsg.SelLength = j
        End If
    ElseIf (Mode = True) Or (V2 <> "" And V3 <> "") Then
        i = i + 2 + Len(V2) + Len(V3)
        If V1 = "" Then
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "," & V3 & "][/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
        Else
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "," & V3 & "]" & V1 & "[/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
            frm.txtMsg.SelLength = j
        End If
    ElseIf V3 = "" Then
        i = i + 1 + Len(V2)
        If V1 = "" Then
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "][/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
        Else
            frm.txtMsg.SelText = "[" & Code & "=" & V2 & "]" & V1 & "[/" & Code & "]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i
            frm.txtMsg.SelLength = j
        End If
    End If
End Function

Public Function BBCodeList(Optional ByVal V1 As String, Optional ByVal V2 As String)
    Dim X As Long
    Dim i As Integer
    Dim j As Long
    
    If V1 = "" Then V1 = frm.txtMsg.SelText
    i = 6
    j = Len(V1)
    X = frm.txtMsg.SelStart
    
    If Mid(frm.txtMsg.Text, frm.txtMsg.SelStart + 3, 7) = "[/list]" Then
        frm.txtMsg.SelText = vbCrLf & "[*]"
        frm.txtMsg.SetFocus
        frm.txtMsg.SelStart = X + 5
    ElseIf V2 = "" Then
        If V1 = "" Then
            frm.txtMsg.SelText = "[list]" & vbCrLf & "[*]" & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
        Else
            V1 = Replace(V1, vbCrLf, vbCrLf & "[*]")
            X = Len(V1)
            frm.txtMsg.SelText = "[list]" & vbCrLf & "[*]" & V1 & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
            'frm.txtMsg.SelLength = j
        End If
    Else
        i = i + 1 + Len(V2)
        If V1 = "" Then
            frm.txtMsg.SelText = "[list=" & V2 & "]" & vbCrLf & "[*]" & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
        Else
            V1 = Replace(V1, vbCrLf, vbCrLf & "[*]")
            X = Len(V1)
            frm.txtMsg.SelText = "[list=" & V2 & "]" & vbCrLf & "[*]" & V1 & vbCrLf & "[/list]"
            frm.txtMsg.SetFocus
            frm.txtMsg.SelStart = X + i + 5
            'frm.txtMsg.SelLength = j
        End If
    End If

End Function

Public Function GetBBCode(ByVal Code As String, Optional ByVal Title As String, Optional ByVal L1 As String, Optional ByVal L2 As String, Optional ByVal L3 As String, Optional ByVal V1 As String, Optional ByVal V2 As String, Optional ByVal V3 As String)
    BBInputC = Code
    BBInputL(0) = L1
    BBInputL(1) = L2
    BBInputL(2) = L3
    If V1 = "" And frm.txtMsg.SelText <> "" Then V1 = frm.txtMsg.SelText
    BBInputV(0) = V1
    BBInputV(1) = V2
    BBInputV(2) = V3
    frmInput.Show
    If Title = "" Then frmInput.Caption = Code Else frmInput.Caption = Title
End Function

'////////////////////页面切换函数////////////////////////////////////////////////////////////

'[不需要，删了]

'////////////////////重置颜色函数////////////////////////////////////////////////////////////

'[不需要，删了]

'////////////////////其他函数////////////////////////////////////////////////////////////

Public Function UrlBox(ByVal Url As String)
    frmUrl.Show
    frmUrl.txtUrl.Text = Url
End Function
