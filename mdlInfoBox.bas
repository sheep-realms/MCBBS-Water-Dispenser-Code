Attribute VB_Name = "mdlInfoBox"
Option Explicit

Public Function InfoBox(ByVal Info As String, Optional ByVal Title As String, Optional ByVal Caption As String, Optional ByVal Copy As Boolean)
    frmInfo.txtInfo.Text = Info
    If Title <> "" Then frmInfo.labTitle.Caption = Title
    If Caption <> "" Then frmInfo.Caption = Caption
    frmInfo.cmdCopy.Visible = Copy
    frmInfo.Show 1
    Title = ""
    Caption = ""
End Function

Public Function ErrorInfo(ByVal Info As String, Optional ByVal Title As String, Optional ByVal Caption As String)
On Error Resume Next
    frmInfo.txtInfo.Text = Info
    If Title <> "" Then frmInfo.labTitle.Caption = Title Else frmInfo.labTitle.Caption = "����"
    If Caption <> "" Then frmInfo.Caption = Caption Else frmInfo.Caption = "����"
    frmInfo.Show 1
    Title = ""
    Caption = ""
End Function

Public Function ErrorInfos(ByVal Number As String, ByVal Description As String, Optional ByVal Locate As String, Optional ByVal Message As String)
    ErrorInfo "SPR-BBCodeTools - �ڲ�����" & vbCrLf & _
              "========================================" & vbCrLf & _
              "������룺" & Number & vbCrLf & _
              "����������" & Description & vbCrLf & _
              "����λ�ã�" & Locate & _
              vbCrLf & _
              vbCrLf & Message & vbCrLf & _
              vbCrLf & _
              "�������������ύ���󱨸棬������ɲ�����б�Ǹ�������������Ϣ��" & vbCrLf & _
              AboutInfo(True) _
              , "�ڲ�����"
End Function
