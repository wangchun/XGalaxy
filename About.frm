VERSION 5.00
Begin VB.Form FormAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���� X ��ϵ"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CommandOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Top             =   2760
      Width           =   1110
   End
   Begin VB.Label LabelDescription 
      AutoSize        =   -1  'True
      Caption         =   "���棺���������������Ȩ���ı�������������и��Ʋ�ʹ��������������֪ͨ���ߣ��������߱������ɸ��������һ��Ȩ����"
      Height          =   540
      Left            =   240
      TabIndex        =   5
      Top             =   1980
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label LabelDate 
      AutoSize        =   -1  'True
      Caption         =   "2000��01��15�ա�2000��05��18��"
      Height          =   180
      Left            =   1440
      TabIndex        =   3
      Top             =   1140
      Width           =   2700
   End
   Begin VB.Label LabelCopyright 
      AutoSize        =   -1  'True
      Caption         =   "��Ȩ����(C)  ����  ��������Ȩ��"
      Height          =   180
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   2790
   End
   Begin VB.Label LabelVersion 
      AutoSize        =   -1  'True
      Caption         =   "�汾 1.0"
      Height          =   180
      Left            =   1440
      TabIndex        =   1
      Top             =   540
      Width           =   720
   End
   Begin VB.Label LabelTitle 
      AutoSize        =   -1  'True
      Caption         =   "X ��ϵ"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Image Image 
      Height          =   960
      Left            =   240
      Picture         =   "About.frx":030A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   960
   End
   Begin VB.Label LabelHomepage 
      AutoSize        =   -1  'True
      Caption         =   "http://wcwcwwc.yeah.net"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   2070
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   240
      X2              =   4425
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   4440
      Y1              =   1815
      Y2              =   1815
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************************************
' *                                                                  *
' *           X ��ϵ   ��Ȩ����(C)   ����   2000��05��18��           *
' *                                                                  *
' *    EMAIL: wcwcwwc@263.net   HOMEPAGE: http://wcwcwwc.yeah.net    *
' *                                                                  *
' ********************************************************************

Option Explicit

Private Sub CommandOK_Click()
    Unload Me
End Sub
