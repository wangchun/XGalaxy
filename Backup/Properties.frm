VERSION 5.00
Begin VB.Form FormProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Properties.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CommandApply 
      Caption         =   "Ӧ��(&A)"
      Height          =   330
      Left            =   3360
      TabIndex        =   4
      Top             =   1080
      Width           =   1110
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   330
      Left            =   3360
      TabIndex        =   3
      Top             =   660
      Width           =   1110
   End
   Begin VB.CommandButton CommandOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   330
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1110
   End
   Begin VB.ComboBox Combo 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.ListBox List 
      Height          =   2400
      ItemData        =   "Properties.frx":030A
      Left            =   240
      List            =   "Properties.frx":030C
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "FormProperties"
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

Public LastListIndex As Long
Private AObject As TObject
Private ADocument As TDocument
Private Difference() As Boolean

Private Sub LoadListSetting(Index As Long)
    On Error Resume Next
    Combo.Clear
    If Difference(Index) Then
        Combo.Text = ""
    Else
        Select Case List.List(Index)
            Case "����"
                Combo.Text = AObject.Caption
            Case "����"
                Combo.Text = AObject.Mass
            Case "�뾶"
                Combo.Text = AObject.Radius
            Case "ˮƽλ��"
                Combo.Text = AObject.X
            Case "��ֱλ��"
                Combo.Text = AObject.Y
            Case "ˮƽ�ٶ�"
                Combo.Text = AObject.VX
            Case "��ֱ�ٶ�"
                Combo.Text = AObject.VY
            Case "�����ɫ"
                Combo.Text = Hex(AObject.FillColor)
                If Len(Combo.Text) < 6 Then Combo.Text = String(6 - Len(Combo.Text), "0") & Combo.Text
        End Select
    End If
End Sub

Private Function SaveListSetting(Index As Long) As Boolean
    Dim I As Long
    Dim J As Long
    SaveListSetting = False
    On Error GoTo ErrorHandle
    If Combo.Text = "" Then
        Difference(Index) = True
    Else
        Select Case List.List(Index)
            Case "����"
                J = 0
                For I = 0 To UBound(Documents(MainForm.ActiveForm.DocumentIndex).Objects) - 1
                    If Documents(MainForm.ActiveForm.DocumentIndex).Objects(I).Style <> osNull And Documents(MainForm.ActiveForm.DocumentIndex).Objects(I).Caption = Trim(Combo.Text) Then J = J + 1
                Next
                If J >= 1 And (J > 1 Or AObject.Caption <> Trim(Combo.Text)) Then Exit Function
                AObject.Caption = Trim(Combo.Text)
            Case "����"
                AObject.Mass = Combo.Text
            Case "�뾶"
                AObject.Radius = Combo.Text
            Case "ˮƽλ��"
                AObject.X = Combo.Text
            Case "��ֱλ��"
                AObject.Y = Combo.Text
            Case "ˮƽ�ٶ�"
                AObject.VX = Combo.Text
            Case "��ֱ�ٶ�"
                AObject.VY = Combo.Text
            Case "�����ɫ"
                AObject.FillColor = "&H" + Combo.Text
        End Select
        Difference(Index) = False
    End If
    SaveListSetting = True
ErrorHandle:
End Function

Private Sub Form_Load()
    Dim Flag As Boolean
    Dim I As Long
    Dim J As Long
    List.Clear
    J = 0
    For I = 0 To UBound(Documents(MainForm.ActiveForm.DocumentIndex).Objects) - 1
        If Documents(MainForm.ActiveForm.DocumentIndex).Objects(I).Style <> osNull And Documents(MainForm.ActiveForm.DocumentIndex).Objects(I).Selected Then J = J + 1
    Next
    If J < 2 Then List.AddItem "����"
    List.AddItem "����"
    List.AddItem "�뾶"
    List.AddItem "ˮƽλ��"
    List.AddItem "��ֱλ��"
    List.AddItem "ˮƽ�ٶ�"
    List.AddItem "��ֱ�ٶ�"
    List.AddItem "�����ɫ"
    ReDim Difference(0 To List.ListCount - 1) As Boolean
    ADocument = Documents(MainForm.ActiveForm.DocumentIndex)
    With Documents(MainForm.ActiveForm.DocumentIndex)
        Flag = True
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull And .Objects(I).Selected Then
                If Flag Then
                    AObject = .Objects(I)
                    Flag = False
                Else
                    For J = 0 To UBound(Difference)
                        If Not Difference(J) Then
                            Select Case List.List(J)
                                Case "����"
                                    If .Objects(I).Caption <> AObject.Caption Then Difference(J) = True
                                Case "����"
                                    If .Objects(I).Mass <> AObject.Mass Then Difference(J) = True
                                Case "�뾶"
                                    If .Objects(I).Radius <> AObject.Radius Then Difference(J) = True
                                Case "ˮƽλ��"
                                    If .Objects(I).X <> AObject.X Then Difference(J) = True
                                Case "��ֱλ��"
                                    If .Objects(I).Y <> AObject.Y Then Difference(J) = True
                                Case "ˮƽ�ٶ�"
                                    If .Objects(I).VX <> AObject.VX Then Difference(J) = True
                                Case "��ֱ�ٶ�"
                                    If .Objects(I).VY <> AObject.VY Then Difference(J) = True
                                Case "�����ɫ"
                                    If .Objects(I).FillColor <> AObject.FillColor Then Difference(J) = True
                            End Select
                        End If
                    Next
                End If
            End If
        Next
    End With
    LastListIndex = -1
    List.ListIndex = 0
End Sub

Private Sub CommandOK_Click()
    Dim I As Long
    Dim J As Long
    If List.ListIndex <> -1 Then
        If Not SaveListSetting(List.ListIndex) Then
            MsgBox "����ֵ��Ч��", vbOKOnly Or vbExclamation
            LoadListSetting List.ListIndex
            Exit Sub
        End If
    End If
    With Documents(MainForm.ActiveForm.DocumentIndex)
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull And .Objects(I).Selected Then
                For J = 0 To UBound(Difference)
                    If Not Difference(J) Then
                        Select Case List.List(J)
                            Case "����"
                                .Objects(I).Caption = AObject.Caption
                            Case "����"
                                .Objects(I).Mass = AObject.Mass
                            Case "�뾶"
                                .Objects(I).Radius = AObject.Radius
                            Case "ˮƽλ��"
                                .Objects(I).X = AObject.X
                            Case "��ֱλ��"
                                .Objects(I).Y = AObject.Y
                            Case "ˮƽ�ٶ�"
                                .Objects(I).VX = AObject.VX
                            Case "��ֱ�ٶ�"
                                .Objects(I).VY = AObject.VY
                            Case "�����ɫ"
                                .Objects(I).FillColor = AObject.FillColor
                        End Select
                    End If
                Next
            End If
        Next
    End With
    MainForm.ActiveForm.RefreshWindow
    Unload Me
End Sub

Private Sub CommandCancel_Click()
    Documents(MainForm.ActiveForm.DocumentIndex) = ADocument
    MainForm.ActiveForm.RefreshWindow
    Unload Me
End Sub

Private Sub CommandApply_Click()
    Dim I As Long
    Dim J As Long
    If List.ListIndex <> -1 Then
        If Not SaveListSetting(List.ListIndex) Then
            MsgBox "����ֵ��Ч��", vbOKOnly Or vbExclamation
            LoadListSetting List.ListIndex
            Exit Sub
        End If
    End If
    With Documents(MainForm.ActiveForm.DocumentIndex)
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull And .Objects(I).Selected Then
                For J = 0 To UBound(Difference)
                    If Not Difference(J) Then
                        Select Case List.List(J)
                            Case "����"
                                .Objects(I).Caption = AObject.Caption
                            Case "����"
                                .Objects(I).Mass = AObject.Mass
                            Case "�뾶"
                                .Objects(I).Radius = AObject.Radius
                            Case "ˮƽλ��"
                                .Objects(I).X = AObject.X
                            Case "��ֱλ��"
                                .Objects(I).Y = AObject.Y
                            Case "ˮƽ�ٶ�"
                                .Objects(I).VX = AObject.VX
                            Case "��ֱ�ٶ�"
                                .Objects(I).VY = AObject.VY
                            Case "�����ɫ"
                                .Objects(I).FillColor = AObject.FillColor
                        End Select
                    End If
                Next
            End If
        Next
    End With
    MainForm.ActiveForm.RefreshWindow
End Sub

Private Sub List_Click()
    If LastListIndex <> List.ListIndex And LastListIndex <> -1 Then
        If Not SaveListSetting(LastListIndex) Then
            MsgBox "����ֵ��Ч��", vbOKOnly Or vbExclamation
            List.ListIndex = LastListIndex
        End If
    End If
    LoadListSetting List.ListIndex
    LastListIndex = List.ListIndex
End Sub
