Attribute VB_Name = "mdlSettings"
Option Explicit

Public Function toSettings(ByVal V1 As String, Optional ByVal V2 As String)
    'frmSettings.Show
    frmSettings.lst1.ListIndex = 1
    frmSettings.lst1_Click
    frmSettings.lst2.ListIndex = 0
    frmSettings.lst2_Click
    frmSettings.Show 1
End Function
