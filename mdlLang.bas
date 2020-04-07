Attribute VB_Name = "mdlLang"
Option Explicit

'I18N

Public Type LangText
    'AppName As String
    
    'Page As String
    'Code As String
    'Special As String
    'Template As String
    
    'SeniorMode As String
    
    'Menu_file As String
    '    Menu_file_save As String
    '    Menu_file_quit As String
    
    'Menu_edit As String
    '    Menu_edit_allSel As String
    '    Menu_edit_clear As String
    '    Menu_edit_now As String
    
    'Menu_view As String
    '    Menu_view_setSize As String
    
    'Menu_code As String
    '    Menu_code_hide As String
    
    'Menu_template As String
    
    'Menu_tools As String
    '    Menu_tools_fb As String
    '        Menu_tools_fb_cls As String
    '        Menu_tools_fb_text As String
    '        Menu_tools_fb_lcase As String
    '        Menu_tools_fb_ucase As String
    
    'Menu_help As String
    '    Menu_help_github As String
    '    Menu_help_help As String
    '    Menu_help_helpDoc As String
    '    Menu_help_UpdataLog As String
    '    Menu_help_BUGLog As String
    '    Menu_help_about As String
    
    'Menu_debug As String
    
    'frm_cmdClear As String
    'frm_cmdCopy As String
    
    YesA As String
    NoA As String
    CancelA As String
    UnloadA As String
    CopyA As String
    ApplicationA As String
    
End Type

'Public en_US As LangText
'Public zh_CN As LangText
Public Lang As LangText

Public LangList() As String

Public Function LangGet()
On Error Resume Next
    'Copy自网络
    Dim MyDirectory As String
    ReDim LangList(100) As String
    LangList(0) = ""
    Dim MyPath As String
    Dim MyDirectoryName As String
    Dim i As Integer
    i = 0
    MyPath = FileHdata & "\lang\*.txt" ' 指定路径。
    MyDirectoryName = Dir(MyPath, vbDirectory) ' 找寻第一项。
    Do While MyDirectoryName <> "" ' 开始循环。
    ' 跳过当前的目录及上层目录。
    If MyDirectoryName <> "." And MyDirectoryName <> ".." Then
    ' 使用位比较来确定 MyName 代表一目录。
    If (GetAttr(MyPath & MyDirectoryName) And vbDirectory) = vbDirectory Then
    LangList(i) = MyDirectoryName
    i = i + 1
    'MyDirectory = MyDirectory & MyDirectoryName & vbCrLf
    End If
    End If
    MyDirectoryName = Dir ' 查找下一个目录。
    Loop
End Function

Public Function LangItemGet(ByVal key As String) As String
On Error GoTo LangItemGetEnd
    
    Dim i As Integer
    
    For i = 0 To 32767
        If key = LangItem(i, 0) Then LangItemGet = LangItem(i, 1)
    Next i
    
LangItemGetEnd:
End Function

Public Function LangSet(ByVal Item As String)
    GetDataList "langitem", Item
    
    
    Lang.YesA = LangItemGet("tc.yes")
    Lang.NoA = LangItemGet("tc.no")
    Lang.CancelA = LangItemGet("tc.cancel")
    Lang.UnloadA = LangItemGet("tc.unload")
    Lang.CopyA = LangItemGet("tc.copy")
    Lang.ApplicationA = LangItemGet("tc.application")
    
    
    frmLang.cmdUnload.Caption = LangItemGet("tc.unload")
    frmLang.cmdOK.Caption = LangItemGet("tc.application")
    
    frm.cColor.Caption = LangItemGet("c.color")
    frm.cColor.ToolTipText = LangItemGet("c.color.tip")
    frm.cImg.Caption = LangItemGet("c.img")
    frm.cImg.ToolTipText = LangItemGet("c.img.tip")
    frm.cUrl.Caption = LangItemGet("c.url")
    frm.cUrl.ToolTipText = LangItemGet("c.url.tip")
    frm.cList.Caption = LangItemGet("c.list")
    frm.cList.ToolTipText = LangItemGet("c.list.tip")
    
    frm.m0.Caption = LangItemGet("m.file")
    frm.m0_Save.Caption = LangItemGet("m.file.save")
    frm.m0_Exit.Caption = LangItemGet("m.file.exit")
    
    frm.m1.Caption = LangItemGet("m.edit")
    frm.m1_Revoke.Caption = LangItemGet("m.edit.revoke")
    frm.m1_All.Caption = LangItemGet("m.edit.allsel")
    frm.m1_Del.Caption = LangItemGet("m.edit.del")
    frm.m1_Find.Caption = LangItemGet("m.edit.instr")
    frm.m1_fb.Caption = LangItemGet("m.edit.format_brush")
    frm.m1_f1.Caption = LangItemGet("m.edit.fun_convert")
    frm.m1_f2.Caption = LangItemGet("m.edit.fun_mun")
    frm.m1_f3.Caption = LangItemGet("m.edit.fun_str")
    frm.m1_DateTime.Caption = LangItemGet("m.edit.now")
    
    frm.m3.Caption = LangItemGet("m.view")
    
    frm.m5.Caption = LangItemGet("m.code")
    
    frm.m10.Caption = LangItemGet("m.template")
    
    frm.m20.Caption = LangItemGet("m.tool")
    
    frm.m80.Caption = LangItemGet("m.settings")
    
    frm.m90.Caption = LangItemGet("m.help")
    
    frm.m99.Caption = LangItemGet("m.debug")
    
End Function
