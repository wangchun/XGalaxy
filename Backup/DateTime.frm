VERSION 5.00
Begin VB.Form FormDateTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ں�ʱ��"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "DateTime.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   1815
      TabIndex        =   2
      Top             =   660
      Width           =   1230
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   435
      TabIndex        =   1
      Top             =   660
      Width           =   1110
   End
   Begin VB.TextBox TextDateTime 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   3000
   End
End
Attribute VB_Name = "FormDateTime"
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
    TextDateTime.Text = Documents(MainForm.ActiveForm.DocumentIndex).DateTime
End Sub

Private Sub CommandOK_Click()
    On Error GoTo ErrorHandle
    Documents(MainForm.ActiveForm.DocumentIndex).DateTime = TextDateTime.Text
    Unload Me
    Exit Sub
ErrorHandle:
    MsgBox "����������Ч��", vbOKOnly Or vbExclamation
End Sub

Private Sub CommandCancel_Click()
    Unload Me
End Sub
