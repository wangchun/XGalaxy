VERSION 5.00
Begin VB.Form FormLock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "Lock.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   660
      Width           =   1110
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1110
   End
   Begin VB.ListBox List 
      Height          =   2760
      ItemData        =   "Lock.frx":030A
      Left            =   240
      List            =   "Lock.frx":030C
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FormLock"
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

Private Sub Form_Load()
    Dim I As Long
    Dim J As Long
    With Documents(MainForm.ActiveForm.DocumentIndex)
        List.Clear
        List.AddItem "(��)"
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                List.AddItem .Objects(I).Caption
                List.Selected(List.ListCount - 1) = .Objects(I).Selected
                If .Lock = I Then J = List.ListCount - 1
            End If
        Next
        If .Lock = -1 Then List.ListIndex = 0 Else List.ListIndex = J
    End With
End Sub

Private Sub CommandOK_Click()
    Dim I As Long
    With Documents(MainForm.ActiveForm.DocumentIndex)
        If List.ListIndex = 0 Then .Lock = -1
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull And .Objects(I).Caption = List.List(List.ListIndex) Then .Lock = I
        Next
    End With
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    Unload Me
End Sub

Private Sub List_DblClick()
    Dim I As Long
    With Documents(MainForm.ActiveForm.DocumentIndex)
        If List.ListIndex = 0 Then .Lock = -1
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull And .Objects(I).Caption = List.List(List.ListIndex) Then .Lock = I
        Next
    End With
    Unload Me
End Sub
