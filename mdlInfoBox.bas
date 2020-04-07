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
    If Title <> "" Then frmInfo.labTitle.Caption = Title Else frmInfo.labTitle.Caption = "错误"
    If Caption <> "" Then frmInfo.Caption = Caption Else frmInfo.Caption = "错误"
    frmInfo.Show 1
    Title = ""
    Caption = ""
End Function

Public Function ErrorInfos(ByVal Number As String, ByVal Description As String, Optional ByVal Locate As String, Optional ByVal Message As String)
    ErrorInfo "SPR-BBCodeTools - 内部错误" & vbCrLf & _
              "========================================" & vbCrLf & _
              "错误代码：" & Number & vbCrLf & _
              "错误描述：" & Description & vbCrLf & _
              "错误位置：" & Locate & _
              vbCrLf & _
              vbCrLf & Message & vbCrLf & _
              vbCrLf & _
              "您可以向作者提交错误报告，给您造成不便深感抱歉！以下是相关信息。" & vbCrLf & _
              AboutInfo(True) _
              , "内部错误"
End Function
