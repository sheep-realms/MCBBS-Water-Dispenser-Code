Attribute VB_Name = "mdlUrl"
Option Explicit

Public Function GotoUrl(ByVal Link As String, Optional ByVal Mode As Boolean)
    frmUrl.Show
    frmUrl.txtUrl.Text = Link
    If Mode = True Then frmUrl.labMsg = "一个外部程序的命令需要本软件打开网页，这很诡异，请小心。"
End Function
