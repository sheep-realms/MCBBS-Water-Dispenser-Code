VERSION 5.00
Begin VB.Form frm 
   Caption         =   "MCBBS饮水机"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10665
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制(&C)"
      Height          =   615
      Left            =   8400
      TabIndex        =   47
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton cUrl 
      Caption         =   "链"
      Height          =   375
      Left            =   2160
      TabIndex        =   46
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cImg 
      Caption         =   "图"
      Height          =   375
      Left            =   1800
      TabIndex        =   45
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame fraCan 
      Caption         =   "参数"
      Height          =   4455
      Left            =   2280
      TabIndex        =   9
      Top             =   2160
      Width           =   6015
      Begin VB.TextBox txtCan 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   44
         Top             =   600
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtCan 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   32
         Left            =   4080
         TabIndex        =   42
         Top             =   3960
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   31
         Left            =   4080
         TabIndex        =   41
         Top             =   3600
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   30
         Left            =   4080
         TabIndex        =   40
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   29
         Left            =   4080
         TabIndex        =   39
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   28
         Left            =   4080
         TabIndex        =   38
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   27
         Left            =   4080
         TabIndex        =   37
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   26
         Left            =   4080
         TabIndex        =   36
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   25
         Left            =   4080
         TabIndex        =   35
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   24
         Left            =   4080
         TabIndex        =   34
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   23
         Left            =   4080
         TabIndex        =   33
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   22
         Left            =   4080
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   21
         Left            =   2160
         TabIndex        =   31
         Top             =   3960
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   20
         Left            =   2160
         TabIndex        =   30
         Top             =   3600
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   19
         Left            =   2160
         TabIndex        =   29
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   18
         Left            =   2160
         TabIndex        =   28
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   17
         Left            =   2160
         TabIndex        =   27
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   16
         Left            =   2160
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   15
         Left            =   2160
         TabIndex        =   25
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   14
         Left            =   2160
         TabIndex        =   24
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   13
         Left            =   2160
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   12
         Left            =   2160
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   11
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   20
         Top             =   3960
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   19
         Top             =   3600
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   18
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optCan 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdBar 
      Caption         =   "回复可见 >>"
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton cmdBar 
      Caption         =   "欢迎新人 >>"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cmdBar 
      Caption         =   "领取金锭  >>"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cS 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cU 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cI 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtMsg 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   10215
   End
   Begin VB.Label labBar 
      Caption         =   "欢迎使用！"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ShiftClick As Boolean

Dim MsgType As String

Private Sub cB_Click()
    BBCode "b"
End Sub

Private Sub cI_Click()
    BBCode "i"
End Sub

Private Sub cImg_Click()
    If ShiftClick = False Then
        BBCode "img"
    ElseIf ShiftClick = True Then
        GetBBCode "img", "网络图片", "网络图片地址", "宽(可选)", "高(可选)"
        ShiftClick = False
        frmInput.Move Me.Left + cImg.Left + 54, Me.Top + cImg.Top + cImg.Height + 400
    End If
End Sub

Private Sub cImg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 Then ShiftClick = True Else ShiftClick = False
End Sub

Private Sub cImg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift = 1 Then ShiftClick = True Else ShiftClick = False
End Sub

Private Sub cmdBar_Click(Index As Integer)
    Unload frmInput
    Select Case Index
    Case 0
        MsgType = "gold_ingot"
        CanListSet 1
        CanSet 0, "庆祝升级"
        CanInSet 0, "5"
    Case 1
        MsgType = "new_user"
        CanListSet 2
        CanSet 0, "真萌新"
        CanSet 1, "假萌新"
        optCan_Click 0
    Case 2
        MsgType = "reply_visible"
        CanListSet 1
        CanSet 0, "一般作品"
        'CanSet 1, "教程资料"
        optCan_Click 0
    Case Else
        
    End Select
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtMsg.Text
    BarMsgOut "复制成功！按 Ctrl + V 粘贴"
End Sub

Private Sub cS_Click()
    BBCode "s"
End Sub

Private Sub cU_Click()
    BBCode "u"
End Sub

Private Sub cUrl_Click()
    GetBBCode "url", "超链接", "链接文本", "URL地址"
    frmInput.Move Me.Left + cUrl.Left + 54, Me.Top + cUrl.Top + cUrl.Height + 400
End Sub

Private Sub Form_Click()
    Unload frmInput
End Sub

Private Sub Form_Load()
    FileDatapackName = "main"

    FileHead = App.Path
    FileHdata = App.Path & "\data"
    
    'FileHead = "D:\Users\Administrator\Desktop\MW"
    'FileHdata = "D:\Users\Administrator\Desktop\MW\data"
    
    FileDatapack = FileHdata & "\datapack\" & FileDatapackName
    
    Dim i As Integer
    For i = 0 To optCan.Count - 1
        optCan(i).Visible = False
        optCan(i).Enabled = False
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub fraCan_Click()
    Unload frmInput
End Sub

Private Sub Label1_Click()
    Unload frmInput
End Sub

Private Sub optCan_Click(Index As Integer)
    Randomize
    Dim str As String
    Dim i As Integer
    For i = 0 To optCan.Count - 1
        If i = Index Then
        Else
            optCan(i).Value = False
        End If
    Next i
    
    Select Case MsgType
    Case "gold_ingot"
        Select Case Index
        Case 0
            If IsNumeric(txtCan(Index).Text) = True Then
                GetDataPack FileDatapack & "\gold_ingot\levelUp.txt"
                str = DataItem(Int(Rnd * (DataItemLength - 3 + 1)) + 3)
                Dim strLv As String, strLv1 As String
                Randomize
                If (Int(Rnd * (10 - 0 + 1)) + 0) > 5 Then
                    strLv = NumChange("ch", txtCan(0).Text)
                    strLv1 = NumChange("ch", txtCan(0).Text + 1)
                Else
                    strLv = NumChange("num", txtCan(0).Text)
                    strLv1 = NumChange("num", txtCan(0).Text + 1)
                End If
                
                str = Replace(str, "__SPR:VALUE__", strLv)
                str = Replace(str, "__SPR:VALUE+1__", strLv1)
                txtMsg.Text = str
            End If
        Case Else
            
        End Select
        
    Case "new_user"
        Select Case Index
        Case 0
            GetDataPack FileDatapack & "\new_user\true.txt"
            str = DataItem(Int(Rnd * (DataItemLength - 3 + 1)) + 3)
            txtMsg.Text = str
        Case 1
            GetDataPack FileDatapack & "\new_user\false.txt"
            str = DataItem(Int(Rnd * (DataItemLength - 3 + 1)) + 3)
            txtMsg.Text = str
        End Select
        
    Case "reply_visible"
        Select Case Index
        Case 0
            GetDataPack FileDatapack & "\reply_visible\default.txt"
            str = DataItem(Int(Rnd * (DataItemLength - 3 + 1)) + 3)
            txtMsg.Text = str
        Case 1
            
        End Select
    
    Case Else
    
    End Select
End Sub

Private Sub optCan_DblClick(Index As Integer)
    optCan_Click Index
End Sub

Private Sub txtCan_Change(Index As Integer)
    If IsNumeric(txtCan(Index).Text) = True Then optCan_Click Index
End Sub

Private Sub txtMsg_Click()
    Unload frmInput
End Sub
