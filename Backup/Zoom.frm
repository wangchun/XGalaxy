VERSION 5.00
Begin VB.Form FormZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ʾ����"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   1935
      TabIndex        =   4
      Top             =   660
      Width           =   1110
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "ȷ��"
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
      Caption         =   "��ÿ����"
      Height          =   180
      Left            =   2550
      TabIndex        =   2
      Top             =   240
      Width           =   720
   End
   Begin VB.Label LabelZoom 
      AutoSize        =   -1  'True
      Caption         =   "��ʾ����:"
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
' ******************************************************************
' *                                                                *
' *          X ��ϵ   ��Ȩ����(C)   ����   2000��05��16��          *
' *                                                                *
' *    EMAIL:wcwcwwc@263.net   HOMEPAGE:http://wcwcwwc.yeah.net    *
' *                                                                *
' ******************************************************************

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
    MsgBox "����������Ч��", vbOKOnly Or vbExclamation
End Sub

Private Sub CommandCancel_Click()
    Unload Me
End Sub
