VERSION 5.00
Begin VB.Form FormSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选定"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "Select.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   660
      Width           =   1110
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1110
   End
   Begin VB.ListBox List 
      Height          =   2760
      ItemData        =   "Select.frx":030A
      Left            =   240
      List            =   "Select.frx":030C
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FormSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******************************************************************
' *                                                                *
' *          X 星系   版权所有(C)   王纯   2000年05月16日          *
' *                                                                *
' *    EMAIL:wcwcwwc@263.net   HOMEPAGE:http://wcwcwwc.yeah.net    *
' *                                                                *
' ******************************************************************

Option Explicit

Private Sub Form_Load()
    Dim I As Long
    With Documents(MainForm.ActiveForm.DocumentIndex)
        List.Clear
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                List.AddItem .Objects(I).Caption
                List.Selected(List.ListCount - 1) = .Objects(I).Selected
            End If
        Next
    End With
End Sub

Private Sub CommandOK_Click()
    Dim I As Long
    Dim J As Long
    With Documents(MainForm.ActiveForm.DocumentIndex)
        For I = 0 To UBound(.Objects) - 1
            .Objects(I).Selected = False
        Next I
        For I = 0 To List.ListCount - 1
            If List.Selected(I) Then
                For J = 0 To UBound(.Objects) - 1
                    If .Objects(I).Style <> osNull Then
                        If List.List(I) = .Objects(I).Caption Then
                            .Objects(I).Selected = True
                            Exit For
                        End If
                    End If
                Next
            End If
        Next
    End With
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    Unload Me
End Sub

Private Sub List_DblClick()
    Dim I As Long
    Dim J As Long
    With Documents(MainForm.ActiveForm.DocumentIndex)
        For I = 0 To UBound(.Objects) - 1
            .Objects(I).Selected = False
        Next I
        For I = 0 To List.ListCount - 1
            If List.Selected(I) Then
                For J = 0 To UBound(.Objects) - 1
                    If .Objects(I).Style <> osNull Then
                        If List.List(I) = .Objects(I).Caption Then
                            .Objects(I).Selected = True
                            Exit For
                        End If
                    End If
                Next
            End If
        Next
    End With
    Unload Me
End Sub
