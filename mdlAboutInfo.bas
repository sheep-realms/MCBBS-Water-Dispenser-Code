Attribute VB_Name = "mdlAboutInfo"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Private Type OSVERSIONINFOEX
    dwOSVersionExInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Function GetOSVersionEx() As String
    Dim retLng As Long, OSVersionEx As OSVERSIONINFOEX
    OSVersionEx.dwOSVersionExInfoSize = Len(OSVersionEx)
    retLng = GetVersionEx(OSVersionEx)
    If retLng = 0 Then
        GetOSVersionEx = "未知"
        Exit Function
    End If
    With OSVersionEx
        GetOSVersionEx = .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber
    End With
End Function

Public Function AboutInfo(Optional ByVal doData As Boolean) As String
    Dim X As String
    If doData = True Then GoTo AboutInfo1
    
    X = X & "SHEEP REALMS - 版权所有" & vbCrLf
    X = X & "作者：我是绵羊Yang_g" & vbCrLf
    X = X & "========================================" & vbCrLf
    X = X & "作者信息" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "联系邮箱：sheep_realms@qq.com" & vbCrLf
    X = X & "" & vbCrLf
    
AboutInfo1:
    X = X & "========================================" & vbCrLf
    X = X & "本体信息" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "版本：" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    X = X & "" & vbCrLf
    X = X & "========================================" & vbCrLf
    X = X & "技术性信息" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "==环境==" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "操作系统：" & Environ("OS") & vbCrLf
    X = X & "内核版本：" & GetOSVersionEx & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "日期：" & Date & vbCrLf
    X = X & "时间：" & Time & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "==通用模块==" & vbCrLf
    X = X & ""
    X = X & "----------------------------------------" & vbCrLf
    X = X & "==主窗体==" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & ""
    X = X & "========================================" & vbCrLf
    X = X & "如果您遇到了什么技术性问题，您可以复制本页面信息提交给作者，并描述详细情况~"

    AboutInfo = X
End Function
