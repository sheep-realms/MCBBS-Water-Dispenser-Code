Attribute VB_Name = "mdlUrl"
Option Explicit

Public Function GotoUrl(ByVal Link As String, Optional ByVal Mode As Boolean)
    frmUrl.Show
    frmUrl.txtUrl.Text = Link
    If Mode = True Then frmUrl.labMsg = "һ���ⲿ�����������Ҫ���������ҳ����ܹ��죬��С�ġ�"
End Function
