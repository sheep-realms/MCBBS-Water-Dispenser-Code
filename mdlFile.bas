Attribute VB_Name = "mdlFile"
Option Explicit

Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function GetFile(ByVal FileName As String) As String    '获取文件内容
On Error Resume Next
    '该函数来自网络
    Dim intFile As Integer
    Dim strData As String
    
    intFile = FreeFile
    Open FileName For Input As intFile
    strData = StrConv(InputB(FileLen(FileName), intFile), vbUnicode)

    Close intFile
    
    If Right(strData, 2) = vbCrLf Then strData = Left(strData, Len(strData) - 2)
    
    GetFile = strData
End Function

Public Function PrintFlie(ByVal FileName As String, Text As String)    '打印文件
On Error Resume Next
    Open FileName For Append As #1
    Print #1, Text
    Close #1
End Function

Public Function SetFiles(ByVal FileName As String)    '创建文件夹
On Error Resume Next
    If Dir(FileName) = "" Then MkDir FileName
End Function

Public Function DirForce(ByVal FileName As String, Text As String)    '强制覆盖文件
On Error Resume Next
    DeleteFile FileName
    Open FileName For Append As #1
    Print #1, Text
    Close #1
End Function
