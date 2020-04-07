VERSION 5.00
Begin VB.Form frmInfo 
   Caption         =   "Info"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6495
   ScaleWidth      =   6375
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label labTitle 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
    Dim X As String
    X = txtInfo.Text
    Clipboard.Clear
    Clipboard.SetText X
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Paint()
On Error Resume Next
    If (WindowState = 0) Then
        If (Me.Width < 6615) Then     '限制最小宽度
            Me.Enabled = False
            Me.Width = 6615
            Me.Enabled = True
        End If
        If (Me.Height < 7065) Then    '限制最小高度
            Me.Enabled = False
            Me.Height = 7065
            Me.Enabled = True
        End If
    End If
    
    txtInfo.Width = Me.Width - txtInfo.Left - 480
    labTitle.Width = Me.Width - labTitle.Left - 480
    
    cmdOK.Left = Me.Width - cmdOK.Width - 480
    cmdCopy.Left = cmdOK.Left - cmdCopy.Width - 110
    
    txtInfo.Height = Me.Height - txtInfo.Top - 1170
    cmdOK.Top = txtInfo.Top + txtInfo.Height + 110
    cmdCopy.Top = txtInfo.Top + txtInfo.Height + 110
End Sub

Private Sub Form_Resize()
    Form_Paint
End Sub
