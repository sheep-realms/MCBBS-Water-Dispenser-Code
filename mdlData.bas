Attribute VB_Name = "mdlData"
Option Explicit

Public TextSave(15) As String
Public TextSaveIndex As Integer
Public TextNoSave As Boolean

Public hiragana_katakana() As String

Public LangItem() As String

Dim DataSave As String
Dim DataListX() As String
Dim DataListY() As String

Public DataItem() As String
Public DataItemLength As Integer

Public Type DataList
    Name As String
    Value As String
    Data As String
End Type

Public ColorData() As DataList

Public Function GetDataPack(ByVal DataPath As String)
    On Error Resume Next
    Dim j As Integer
    
    'ǰ��׼��
    ReDim DataItem(127) As String
    ReDim DataListX(127) As String
    
    '��ȡ�봦������
    If Dir(DataPath) = "" Then ErrorInfo DataPath & "�ļ���ʧ��": Exit Function
    DataSave = GetFile(DataPath)
    DataSave = Replace(DataSave, vbCrLf & vbCrLf, vbCrLf)
    DataSave = Replace(DataSave, " ", "")
    DataSave = Replace(DataSave, "__SPR:\0__", " ")
    
    '�ָ�����
    DataListX = Split(DataSave, vbCrLf)
    
    'Ӧ������
    On Error GoTo GetDataPack_End
    For j = 0 To 127
        DataItem(j) = Replace(DataListX(j), "__SPR:\BR__", vbCrLf)
    Next j
    
    Exit Function
    
GetDataPack_End:
    DataItemLength = j - 1
    
End Function

'Public Function GetDataList(ByVal DataType As String, Optional ByVal DataItem As String)
'    Dim i As Integer, j As Integer
'    Select Case LCase(DataType)
'
''//////////��ɫ���б�////////////////////////////////////////////////////////////
'    Case "colordata"
'        'ǰ��׼��
'        ReDim Preserve ColorData(127) As DataList
'        ReDim DataListX(127) As String
'        ReDim DataListY(127) As String
'
'        '��ȡ�봦������
'        If Dir(FileHdata & "\special\configuration\color\color_name.txt") = "" Then ErrorInfo "\data\special\configuration\color\color_name.txt�ļ���ʧ��": Exit Function
'        DataSave = GetFile(FileHdata & "\special\configuration\color\color_name.txt")
'        DataSave = Replace(DataSave, vbCrLf, "")
'        DataSave = Replace(DataSave, " ", "")
'
'        '�ָ�����
'        DataListX = Split(DataSave, ";")
'        'Exit Function
'        For i = 0 To 127
'            On Error GoTo Colordata_Step1
'            DataListY = Split(DataListX(i), ",")
'            ColorData(i).Name = DataListY(1)
'            ColorData(i).Value = DataListY(0)
'            ColorData(i).Data = DataListY(2)
'            If ColorData(i).Name = "" Then ColorData(i).Name = ColorData(i).Value
'        Next i
'
'Colordata_Step1:    'Ӧ������
'        For j = 0 To 39
'            frmColor.Color(j).BackColor = ColorData(j).Data
'            frmColor.Color(j).ToolTipText = ColorData(j).Name
'        Next j
'
''//////////�������ձ�////////////////////////////////////////////////////////////
'    Case "hiragana_katakana"
'        'ǰ��׼��
'        ReDim hiragana_katakana(82, 1) As String
'        ReDim DataListX(82) As String
'        ReDim DataListY(1) As String
'
'        '��ȡ�봦������
'        If Dir(FileHdata & "\datapack\hiragana_katakana.txt") = "" Then ErrorInfo "hiragana_katakana.txt�ļ���ʧ��": Exit Function
'        DataSave = GetFile(FileHdata & "\datapack\hiragana_katakana.txt")
'        DataSave = Replace(DataSave, vbCrLf, "")
'        DataSave = Replace(DataSave, " ", "")
'
'        '�ָ�����
'        DataListX = Split(DataSave, ";")
'        For i = 0 To 82
'            On Error GoTo Hiragana_katakana_Step1
'            DataListY = Split(DataListX(i), ",")
'            hiragana_katakana(i, 0) = DataListY(0)
'            hiragana_katakana(i, 1) = DataListY(1)
'        Next i
'
'Hiragana_katakana_Step1:
'
''//////////���԰���Ŀ////////////////////////////////////////////////////////////
'    Case "langitem"
'        'ǰ��׼��
'        ReDim LangItem(32767, 1) As String
'        ReDim DataListX(32767) As String
'        ReDim DataListY(1) As String
'
'        '��ȡ�봦������
'        If Dir(FileHdata & "\lang\" & DataItem) = "" Then ErrorInfo DataItem & "�ļ���ʧ��": Exit Function
'        DataSave = GetFile(FileHdata & "\lang\" & DataItem)
'        DataSave = Replace(DataSave, vbCrLf & vbCrLf, vbCrLf)
'        DataSave = Replace(DataSave, vbCrLf & vbCrLf, vbCrLf)
'
'        '�ָ�����
'        DataListX = Split(DataSave, vbCrLf)
'        For i = 0 To 32767
'            On Error GoTo LangList_Step1
'            DataListY = Split(DataListX(i), "==")
'            DataListY(1) = Replace(DataListY(1), "\n", vbCrLf)
'            LangItem(i, 0) = DataListY(0)
'            LangItem(i, 1) = DataListY(1)
'        Next i
'
'LangList_Step1:
'
''//////////���԰�Ԥ��////////////////////////////////////////////////////////////
'    Case "langitem_"
'        'ǰ��׼��
'        ReDim LangItem(5, 1) As String
'        ReDim DataListX(5) As String
'        ReDim DataListY(1) As String
'
'        '��ȡ�봦������
'        If Dir(FileHdata & "\lang\" & DataItem) = "" Then ErrorInfo DataItem & "�ļ���ʧ��": Exit Function
'        DataSave = GetFile(FileHdata & "\lang\" & DataItem)
'        DataSave = Replace(DataSave, vbCrLf & vbCrLf, vbCrLf)
'        DataSave = Replace(DataSave, vbCrLf & vbCrLf, vbCrLf)
'
'        '�ָ�����
'        DataListX = Split(DataSave, vbCrLf)
'        For i = 0 To 5
'            On Error GoTo LangList_Step1
'            DataListY = Split(DataListX(i), "==")
'            DataListY(1) = Replace(DataListY(1), "\n", vbCrLf)
'            LangItem(i, 0) = DataListY(0)
'            LangItem(i, 1) = DataListY(1)
'        Next i
'
'LangList__Step1:
'
'    End Select
'End Function
'
'Public Function ClearDataList(ByVal DataType As String)
'On Error Resume Next
'    Dim i As Integer
'    Select Case LCase(DataType)
'    Case "colordata"
'        ReDim ColorData(127) As DataList
'    End Select
'End Function
'
'Public Function TextSaves(Optional ByVal Text As String)
'    Dim i As Integer
'    If Text = "" Then Text = frm.txtMsg.Text
'    If TextSaveIndex = 0 Then
'        For i = 14 To 0 Step -1
'            TextSave(i + 1) = TextSave(i)
'        Next i
'    ElseIf TextSaveIndex <> 1 Then
'        For i = 1 To (15 - TextSaveIndex + 1)
'            TextSave(i) = TextSave(TextSaveIndex + i - 1)
'        Next i
'    End If
'    TextSave(0) = frm.txtMsg.Text
'    TextSaveIndex = 0
'End Function
'
'Public Function TextRevoke()
'    TextSaveIndex = TextSaveIndex + 1
'    If TextSaveIndex = 16 Then
'        TextSaveIndex = 15
'        frm.MsgBarOut "�����ٳ����ˣ�", "�ǲ�ס��ô������"
'        Exit Function
'    End If
'    TextNoSave = True
'    frm.txtMsg.Text = TextSave(TextSaveIndex)
'    TextNoSave = False
'    frm.txtMsg.SelStart = Len(frm.txtMsg.Text)
'End Function
