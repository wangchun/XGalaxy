VERSION 5.00
Begin VB.Form FormInterval 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "时间间隔"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   330
      Left            =   1935
      TabIndex        =   4
      Top             =   660
      Width           =   1110
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   330
      Left            =   435
      TabIndex        =   3
      Top             =   660
      Width           =   1110
   End
   Begin VB.TextBox TextInterval 
      Height          =   300
      Left            =   2070
      TabIndex        =   1
      Top             =   180
      Width           =   900
   End
   Begin VB.Label LabelUnit 
      AutoSize        =   -1  'True
      Caption         =   "秒"
      Height          =   180
      Left            =   3090
      TabIndex        =   2
      Top             =   240
      Width           =   180
   End
   Begin VB.Label LabelInterval 
      AutoSize        =   -1  'True
      Caption         =   "两次计算的时间间隔:"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1710
   End
End
Attribute VB_Name = "FormInterval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    TextInterval.Text = Documents(MainForm.ActiveForm.DocumentIndex).Interval
End Sub

Private Sub CommandOK_Click()
    On Error GoTo ErrorHandle
    Documents(MainForm.ActiveForm.DocumentIndex).Interval = TextInterval.Text
    Unload Me
    Exit Sub
ErrorHandle:
    MsgBox "输入数据无效。", vbOKOnly Or vbExclamation
End Sub

Private Sub CommandCancel_Click()
    Unload Me
End Sub
