VERSION 5.00
Begin VB.Form frmUrl 
   Caption         =   "跳转提示"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUrl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   7215
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "否(&N)"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制(&C)"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "是(&Y)"
      Default         =   -1  'True
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtUrl 
      Height          =   390
      Left            =   240
      TabIndex        =   3
      Text            =   "http://"
      Top             =   1080
      Width           =   6735
   End
   Begin VB.Label labCopy 
      Alignment       =   2  'Center
      Caption         =   "  复制成功！"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labMsg 
      Caption         =   "您的某个操作需要打开网页 ，是否访问？"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   6495
   End
   Begin VB.Label labTitle 
      Caption         =   "您即将访问以下网站"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmUrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtUrl.Text
    labCopy.Visible = True
End Sub

Private Sub cmdNo_Click()
    Unload Me
End Sub

Private Sub cmdYes_Click()
On Error Resume Next
    ShellExecute 0, "open", txtUrl.Text, 0, 0, 1
    Unload Me
End Sub

'Private Sub Form_Load()
'    cmdNo.Caption = Lang.NoA
'    cmdCopy.Caption = Lang.CopyA
'    cmdYes.Caption = Lang.YesA
'End Sub
