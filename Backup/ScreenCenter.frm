VERSION 5.00
Begin VB.Form FormScreenCenter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ļ����"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "ScreenCenter.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   1935
      TabIndex        =   7
      Top             =   960
      Width           =   1110
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   435
      TabIndex        =   6
      Top             =   960
      Width           =   1110
   End
   Begin VB.TextBox TextY 
      Height          =   300
      Left            =   1380
      TabIndex        =   4
      Top             =   540
      Width           =   1560
   End
   Begin VB.TextBox TextX 
      Height          =   300
      Left            =   1380
      TabIndex        =   1
      Top             =   180
      Width           =   1560
   End
   Begin VB.Label LabelUnitY 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   3060
      TabIndex        =   5
      Top             =   600
      Width           =   180
   End
   Begin VB.Label LabelY 
      AutoSize        =   -1  'True
      Caption         =   "��ֱλ��(&Y):"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label LabelUnitX 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   3060
      TabIndex        =   2
      Top             =   240
      Width           =   180
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      Caption         =   "ˮƽλ��(&X):"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "FormScreenCenter"
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
    TextX.Text = Documents(MainForm.ActiveForm.DocumentIndex).ScrollX
    TextY.Text = Documents(MainForm.ActiveForm.DocumentIndex).ScrollY
End Sub

Private Sub CommandOK_Click()
    On Error GoTo ErrorHandle
    Documents(MainForm.ActiveForm.DocumentIndex).ScrollX = TextX.Text
    Documents(MainForm.ActiveForm.DocumentIndex).ScrollY = TextY.Text
    MainForm.ActiveForm.RefreshWindow
    Unload Me
    Exit Sub
ErrorHandle:
    MsgBox "����������Ч��", vbOKOnly Or vbExclamation
End Sub

Private Sub CommandCancel_Click()
    Unload Me
End Sub
