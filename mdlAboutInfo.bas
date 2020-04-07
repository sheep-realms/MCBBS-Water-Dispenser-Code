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
        GetOSVersionEx = "δ֪"
        Exit Function
    End If
    With OSVersionEx
        GetOSVersionEx = .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber
    End With
End Function

Public Function AboutInfo(Optional ByVal doData As Boolean) As String
    Dim X As String
    If doData = True Then GoTo AboutInfo1
    
    X = X & "SHEEP REALMS - ��Ȩ����" & vbCrLf
    X = X & "���ߣ���������Yang_g" & vbCrLf
    X = X & "========================================" & vbCrLf
    X = X & "������Ϣ" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "��ϵ���䣺sheep_realms@qq.com" & vbCrLf
    X = X & "" & vbCrLf
    
AboutInfo1:
    X = X & "========================================" & vbCrLf
    X = X & "������Ϣ" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "�汾��" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
    X = X & "" & vbCrLf
    X = X & "========================================" & vbCrLf
    X = X & "��������Ϣ" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "==����==" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "����ϵͳ��" & Environ("OS") & vbCrLf
    X = X & "�ں˰汾��" & GetOSVersionEx & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "���ڣ�" & Date & vbCrLf
    X = X & "ʱ�䣺" & Time & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & "==ͨ��ģ��==" & vbCrLf
    X = X & ""
    X = X & "----------------------------------------" & vbCrLf
    X = X & "==������==" & vbCrLf
    X = X & "----------------------------------------" & vbCrLf
    X = X & ""
    X = X & "========================================" & vbCrLf
    X = X & "�����������ʲô���������⣬�����Ը��Ʊ�ҳ����Ϣ�ύ�����ߣ���������ϸ���~"

    AboutInfo = X
End Function
