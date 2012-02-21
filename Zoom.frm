VERSION 5.00
Begin VB.Form FormZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "显示比例"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   Icon            =   "Zoom.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3510
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
   Begin VB.TextBox TextZoom 
      Height          =   300
      Left            =   1170
      TabIndex        =   1
      Top             =   180
      Width           =   1260
   End
   Begin VB.Label LabelUnit 
      AutoSize        =   -1  'True
      Caption         =   "米每像素"
      Height          =   180
      Left            =   2550
      TabIndex        =   2
      Top             =   240
      Width           =   720
   End
   Begin VB.Label LabelZoom 
      AutoSize        =   -1  'True
      Caption         =   "显示比例:"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "FormZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************
' *                                                                  *
' *           X 星系   版权所有(C)   王纯   2000年05月18日           *
' *                                                                  *
' *    EMAIL: wcwcwwc@263.net   HOMEPAGE: http://wcwcwwc.yeah.net    *
' *                                                                  *
' ********************************************************************

Option Explicit

Private Sub Form_Load()
    TextZoom.Text = 15 / Documents(MainForm.ActiveForm.DocumentIndex).Zoom
End Sub

Private Sub CommandOK_Click()
    On Error GoTo ErrorHandle
    Documents(MainForm.ActiveForm.DocumentIndex).Zoom = 15 / TextZoom.Text
    MainForm.ActiveForm.RefreshWindow
    Unload Me
    Exit Sub
ErrorHandle:
    MsgBox "输入数据无效。", vbOKOnly Or vbExclamation
End Sub

Private Sub CommandCancel_Click()
    Unload Me
End Sub
