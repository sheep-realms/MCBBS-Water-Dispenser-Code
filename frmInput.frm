VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdMore 
      Caption         =   "编辑"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Text            =   "txt"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Text            =   "txt"
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txt 
      Height          =   270
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmInput.frx":577A
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lab 
      Caption         =   "lab"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lab 
      Caption         =   "lab"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lab 
      Caption         =   "lab"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cOk As Boolean

Private Sub cmdMore_Click()
    cmdMore.Caption = "咕"
End Sub

Private Sub cmdOK_Click()
    cOk = True
    Unload Me
End Sub

Public Function frmInputLoad()
    Form_Load
End Function


Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 2
        If BBInputL(i) = "" Then lab(i).Visible = False: txt(i).Visible = False Else lab(i).Visible = True: txt(i).Visible = True
        txt(i).Text = BBInputV(i)
        lab(i).Caption = BBInputL(i)
    Next i
    If txt(0).Text <> "" Then txt(0).SelStart = 0: txt(0).SelLength = Len(txt(0).Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cOk = True Then
        BBCode BBInputC, txt(0).Text, txt(1).Text, txt(2).Text, BBInputMode
        BBInputC = ""
        BBInputMode = False
        Dim i As Integer
        For i = 0 To 2
            BBInputV(i) = ""
            BBInputL(i) = ""
        Next i
    End If
End Sub
