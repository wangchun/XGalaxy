VERSION 5.00
Begin VB.Form FormDocument 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9480
   ClipControls    =   0   'False
   Icon            =   "Document.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu MenuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu MenuFileNew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu MenuFileOpen 
         Caption         =   "打开(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu MenuFileClose 
         Caption         =   "关闭(&C)"
      End
      Begin VB.Menu MenuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuFileSaveAs 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu MenuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFilePageSetup 
         Caption         =   "页面设置(&U)..."
      End
      Begin VB.Menu MenuFilePrintPreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu MenuFilePrint 
         Caption         =   "打印(&P)..."
         Shortcut        =   ^P
      End
      Begin VB.Menu MenuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu MenuEditCut 
         Caption         =   "剪切(&T)"
         Shortcut        =   ^X
      End
      Begin VB.Menu MenuEditCopy 
         Caption         =   "复制(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu MenuEditPaste 
         Caption         =   "粘贴(&P)"
         Shortcut        =   ^V
      End
      Begin VB.Menu MenuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEditSelectAll 
         Caption         =   "全选(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu MenuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEditSelect 
         Caption         =   "选定(&S)..."
      End
      Begin VB.Menu MenuEditFind 
         Caption         =   "查找(&F)..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu MenuView 
      Caption         =   "视图(&V)"
      Begin VB.Menu MenuViewToolbar 
         Caption         =   "工具栏(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MenuViewStatusBar 
         Caption         =   "状态栏(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu MenuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenuViewLock 
         Caption         =   "锁定(&L)..."
      End
      Begin VB.Menu MenuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuViewScreenCenter 
         Caption         =   "屏幕中心(&C)..."
      End
      Begin VB.Menu MenuViewZoom 
         Caption         =   "显示比例(&Z)..."
      End
      Begin VB.Menu MenuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuViewCustom 
         Caption         =   "自定义(&C)..."
      End
   End
   Begin VB.Menu MenuGalaxy 
      Caption         =   "星系(&G)"
      Begin VB.Menu MenuGalaxyInsertObject 
         Caption         =   "添加对象(&O)"
      End
      Begin VB.Menu MenuGalaxyBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenuGalaxyDateTime 
         Caption         =   "日期和时间(&D)..."
      End
      Begin VB.Menu MenuGalaxyInterval 
         Caption         =   "时间间隔(&I)..."
      End
      Begin VB.Menu MenuGalaxyBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuGalaxyProperties 
         Caption         =   "对象属性(&P)..."
      End
   End
   Begin VB.Menu MenuRun 
      Caption         =   "运行(&R)"
      Begin VB.Menu MenuRunStart 
         Caption         =   "启动(&S)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuRunBreak 
         Caption         =   "中断(&K)"
      End
      Begin VB.Menu MenuRunEnd 
         Caption         =   "结束(&E)"
      End
      Begin VB.Menu MenuRunRestart 
         Caption         =   "重新启动(&R)"
         Shortcut        =   +{F5}
      End
   End
   Begin VB.Menu MenuWindow 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu MenuWindowCascade 
         Caption         =   "层叠(&C)"
      End
      Begin VB.Menu MenuWindowTileHorizontal 
         Caption         =   "横向平铺(&H)"
      End
      Begin VB.Menu MenuWindowTileVertical 
         Caption         =   "纵向平铺(&V)"
      End
      Begin VB.Menu MenuWindowArrangeIcons 
         Caption         =   "排列图标(&A)"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "关于 X 星系(&A)..."
      End
   End
   Begin VB.Menu MenuPopupObject 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MenuPopupObjectCut 
         Caption         =   "剪切(&T)"
      End
      Begin VB.Menu MenuPopupObjectCopy 
         Caption         =   "复制(&C)"
      End
      Begin VB.Menu MenuPopupObjectPaste 
         Caption         =   "粘贴(&P)"
      End
      Begin VB.Menu MenuPopupObjectDelete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu MenuPopupObjectBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPopupObjectScreenCenter 
         Caption         =   "设为屏幕中心(&E)"
      End
      Begin VB.Menu MenuPopupObjectBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuPopupObjectProperties 
         Caption         =   "对象属性(&P)..."
      End
   End
End
Attribute VB_Name = "FormDocument"
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

Private DragLastX As Double
Private DragLastY As Double
Private MouseX As Double
Private MouseY As Double
Public DocumentIndex As Long

Public Sub RefreshWindow()
    Dim I As Long
    Dim X As Double
    Dim Y As Double
    Dim Radius As Double
    On Error Resume Next
    With Documents(DocumentIndex)
        Cls
        DrawMode = 13
        DrawStyle = 5
        FillStyle = 0
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                FillColor = .Objects(I).FillColor
                X = (.Objects(I).X - .ScrollX) * .Zoom + ScaleWidth / 2
                Y = (.Objects(I).Y - .ScrollY) * .Zoom + ScaleHeight / 2
                Radius = .Objects(I).Radius * .Zoom
                If Radius < 15 Then Radius = 15
                Circle (X, Y), Radius
            End If
        Next
        If Not .Run Or .Break Then
            For I = 0 To UBound(.Objects) - 1
                If .Objects(I).Style <> osNull Then
                    X = (.Objects(I).X - .ScrollX) * .Zoom + ScaleWidth / 2
                    Y = (.Objects(I).Y - .ScrollY) * .Zoom + ScaleHeight / 2
                    Radius = .Objects(I).Radius * .Zoom
                    If Radius < 15 Then Radius = 15
                    CurrentX = -TextWidth(.Objects(I).Caption)
                    CurrentY = -TextHeight(.Objects(I).Caption)
                    CurrentX = X - TextWidth(.Objects(I).Caption) / 2
                    If Sqr(TextWidth(.Objects(I).Caption) ^ 2 + TextHeight(.Objects(I).Caption) ^ 2) >= (Radius - 45) * 2 Then
                        CurrentY = Y + Radius + 45
                        ForeColor = Not BackColor And &HFFFFFF
                    Else
                        CurrentY = Y - TextHeight(.Objects(I).Caption) / 2
                        ForeColor = Not .Objects(I).FillColor And &HFFFFFF
                    End If
                    If CurrentX <> -TextWidth(.Objects(I).Caption) And CurrentY <> -TextHeight(.Objects(I).Caption) Then Print .Objects(I).Caption
                End If
            Next
            DrawMode = 7
            DrawStyle = 0
            For I = 0 To UBound(.Objects) - 1
                If .Objects(I).Style <> osNull And .Objects(I).Selected Then
                    X = (.Objects(I).X - .ScrollX) * .Zoom + ScaleWidth / 2
                    Y = (.Objects(I).Y - .ScrollY) * .Zoom + ScaleHeight / 2
                    Radius = .Objects(I).Radius * .Zoom
                    If Radius < 15 Then Radius = 15
                    Line (X - Radius - 45, Y - Radius - 45)-Step(30, 30), RGB(255, 255, 255), BF
                    Line (X + Radius + 15, Y - Radius - 45)-Step(30, 30), RGB(255, 255, 255), BF
                    Line (X - Radius - 45, Y + Radius + 15)-Step(30, 30), RGB(255, 255, 255), BF
                    Line (X + Radius + 15, Y + Radius + 15)-Step(30, 30), RGB(255, 255, 255), BF
                End If
            Next
        End If
    End With
End Sub

Private Sub Form_Load()
    Caption = Documents(DocumentIndex).Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Documents(DocumentIndex).Title = ""
    ReDim Documents(DocumentIndex).Objects(0 To 0)
    UpdateEnabled
    MainForm.StatusBar.SimpleText = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Documents(DocumentIndex).Dirty Then
        Select Case MsgBox("是否保存对 " & Documents(DocumentIndex).Title & " 的修改？", vbYesNoCancel Or vbExclamation)
            Case vbYes
                If Documents(DocumentIndex).FileName = "" Then
                    With MainForm.CommonDialog
                        .CancelError = False
                        .FileName = ""
                        .Filter = "星系文件 (*.gal)|*.gal|所有文件 (*.*)|*.*"
                        .FilterIndex = 0
                        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
                        .ShowSave
                        If Len(.FileName) = 0 Then Exit Sub
                        Documents(DocumentIndex).FileName = .FileName
                    End With
                End If
                FileSave DocumentIndex
            Case vbCancel
                Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Paint()
    RefreshWindow
End Sub

Private Sub Form_Activate()
    UpdateEnabled
    MainForm.StatusBar.SimpleText = Documents(DocumentIndex).DateTime
End Sub

Private Sub Form_Resize()
    RefreshWindow
End Sub

Private Sub Form_DblClick()
    Dim I As Long
    Dim J As Long
    Dim Min As Double
    Dim PropertiesBox As FormProperties
    If Documents(DocumentIndex).Run And Not Documents(DocumentIndex).Break Then Exit Sub
    With Documents(DocumentIndex)
        J = -1
        Min = 1.79769313486231E+308
        For I = 0 To UBound(.Objects) - 1
            If Sqr((MouseX - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (MouseY - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2) <= Min And Sqr((MouseX - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (MouseY - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2) <= .Objects(I).Radius * .Zoom + 45 Then
                J = I
                Min = Sqr((MouseX - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (MouseY - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2)
            End If
        Next
        If J >= 0 Then
            If Not .Objects(J).Selected Then
                For I = 0 To UBound(.Objects) - 1
                    .Objects(I).Selected = False
                Next
                .Objects(J).Selected = True
            End If
            Set PropertiesBox = New FormProperties
            Load PropertiesBox
            PropertiesBox.Show vbModal
        End If
    End With
    UpdateEnabled
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    Dim J As Long
    Dim Min As Double
    Dim DeltaX As Double
    Dim DeltaY As Double
    If Documents(DocumentIndex).Run And Not Documents(DocumentIndex).Break Then Exit Sub
    MouseX = X
    MouseY = Y
    If Button = 1 Then
        With Documents(DocumentIndex)
            DeltaX = X / .Zoom + .ScrollX - DragLastX
            DeltaY = Y / .Zoom + .ScrollY - DragLastY
            For I = 0 To UBound(.Objects) - 1
                If .Objects(I).Selected Then
                    .Objects(I).X = .Objects(I).X + DeltaX
                    .Objects(I).Y = .Objects(I).Y + DeltaY
                End If
            Next
        End With
        RefreshWindow
    End If
    If Button = 2 Then
        With Documents(DocumentIndex)
            J = -1
            Min = 1.79769313486231E+308
            For I = 0 To UBound(.Objects) - 1
                If Sqr((X - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (Y - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2) <= Min And Sqr((X - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (Y - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2) <= .Objects(I).Radius * .Zoom + 45 Then
                    J = I
                    Min = Sqr((X - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (Y - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2)
                End If
            Next
            If J >= 0 Then
                If Not .Objects(J).Selected Then
                    For I = 0 To UBound(.Objects) - 1
                        .Objects(I).Selected = False
                    Next
                    .Objects(J).Selected = True
                End If
                PopupMenu MenuPopupObject, , , , MenuPopupObjectProperties
            End If
        End With
    End If
    UpdateEnabled
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    Dim J As Long
    Dim Min As Double
    If Documents(DocumentIndex).Run And Not Documents(DocumentIndex).Break Then Exit Sub
    If Button = 1 Then
        With Documents(DocumentIndex)
            DragLastX = X / .Zoom + .ScrollX
            DragLastY = Y / .Zoom + .ScrollY
            J = -1
            Min = 1.79769313486231E+308
            For I = 0 To UBound(.Objects) - 1
                If Sqr((X - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (Y - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2) <= Min And Sqr((X - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (Y - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2) <= .Objects(I).Radius * .Zoom + 45 Then
                    J = I
                    Min = Sqr((X - (.Objects(I).X - .ScrollX) * .Zoom - ScaleWidth / 2) ^ 2 + (Y - (.Objects(I).Y - .ScrollY) * .Zoom - ScaleHeight / 2) ^ 2)
                End If
            Next
            If Shift <> 1 And Shift <> 2 Then
                If J >= 0 Then
                    If Not .Objects(J).Selected Then
                        For I = 0 To UBound(.Objects) - 1
                            .Objects(I).Selected = False
                        Next
                    End If
                Else
                    For I = 0 To UBound(.Objects) - 1
                        .Objects(I).Selected = False
                    Next
                End If
            End If
            If J >= 0 Then .Objects(J).Selected = True
        End With
        RefreshWindow
    End If
    UpdateEnabled
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Long
    Dim DeltaX As Double
    Dim DeltaY As Double
    If Documents(DocumentIndex).Run And Not Documents(DocumentIndex).Break Then Exit Sub
    If Button = 1 Then
        With Documents(DocumentIndex)
            DeltaX = X / .Zoom + .ScrollX - DragLastX
            DeltaY = Y / .Zoom + .ScrollY - DragLastY
            For I = 0 To UBound(.Objects) - 1
                If .Objects(I).Selected Then
                    .Objects(I).X = .Objects(I).X + DeltaX
                    .Objects(I).Y = .Objects(I).Y + DeltaY
                    .Dirty = True
                End If
            Next
            DragLastX = X / .Zoom + .ScrollX
            DragLastY = Y / .Zoom + .ScrollY
        End With
        RefreshWindow
    End If
End Sub

Private Sub Timer_Timer()
    Dim I As Long
    Dim J As Long
    Dim X As Double
    Dim Y As Double
    Dim F As Double
    Dim FX As Double
    Dim FY As Double
    Dim D As Double
    Dim DX As Double
    Dim DY As Double
    Dim VX As Double
    Dim VY As Double
    Dim Radius As Double
    On Error Resume Next
    With Documents(DocumentIndex)
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                For J = I + 1 To UBound(.Objects) - 1
                    If .Objects(J).Style <> osNull Then
                        DX = .Objects(I).X - .Objects(J).X
                        DY = .Objects(I).Y - .Objects(J).Y
                        D = Sqr(DX ^ 2 + DY ^ 2)
                        F = G * .Interval * .Objects(I).Mass * .Objects(J).Mass / D ^ 2
                        FX = -(F * DX / D)
                        FY = -(F * DY / D)
                        VX = FX / .Objects(I).Mass
                        VY = FY / .Objects(I).Mass
                        .Objects(I).VX = .Objects(I).VX + VX
                        .Objects(I).VY = .Objects(I).VY + VY
                        VX = -FX / .Objects(J).Mass
                        VY = -FY / .Objects(J).Mass
                        .Objects(J).VX = .Objects(J).VX + VX
                        .Objects(J).VY = .Objects(J).VY + VY
                    End If
                Next
            End If
        Next
        DrawMode = 13
        DrawStyle = 5
        FillStyle = 0
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                FillColor = BackColor
                X = (.Objects(I).X - .ScrollX) * .Zoom + ScaleWidth / 2
                Y = (.Objects(I).Y - .ScrollY) * .Zoom + ScaleHeight / 2
                Radius = .Objects(I).Radius * .Zoom
                If Radius < 15 Then Radius = 15
                Circle (X, Y), Radius
            End If
        Next
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                .Objects(I).X = .Objects(I).X + .Objects(I).VX * .Interval
                .Objects(I).Y = .Objects(I).Y + .Objects(I).VY * .Interval
                If I = .Lock Then
                    .ScrollX = .Objects(I).X
                    .ScrollY = .Objects(I).Y
                End If
            End If
        Next
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                FillColor = .Objects(I).FillColor
                X = (.Objects(I).X - .ScrollX) * .Zoom + ScaleWidth / 2
                Y = (.Objects(I).Y - .ScrollY) * .Zoom + ScaleHeight / 2
                Radius = .Objects(I).Radius * .Zoom
                If Radius < 15 Then Radius = 15
                Circle (X, Y), Radius
            End If
        Next
        .DateTime = DateAdd("s", .Interval, .DateTime)
        MainForm.StatusBar.SimpleText = .DateTime
    End With
End Sub

Private Sub MenuFileNew_Click()
    FileNew
End Sub

Private Sub MenuFileOpen_Click()
    With MainForm.CommonDialog
        .CancelError = False
        .FileName = ""
        .Filter = "星系文件 (*.gal)|*.gal|所有文件 (*.*)|*.*"
        .FilterIndex = 0
        .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        FileOpen .FileName
    End With
End Sub

Private Sub MenuFileClose_Click()
    Unload Me
End Sub

Private Sub MenuFileSave_Click()
    If Documents(DocumentIndex).FileName = "" Then
        With MainForm.CommonDialog
            .CancelError = False
            .FileName = ""
            .Filter = "星系文件 (*.gal)|*.gal|所有文件 (*.*)|*.*"
            .FilterIndex = 0
            .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
            .ShowSave
            If Len(.FileName) = 0 Then Exit Sub
            Documents(DocumentIndex).FileName = .FileName
        End With
    End If
    FileSave DocumentIndex
End Sub

Private Sub MenuFileSaveAs_Click()
    With MainForm.CommonDialog
        .CancelError = False
        .FileName = ""
        .Filter = "星系文件 (*.gal)|*.gal|所有文件 (*.*)|*.*"
        .FilterIndex = 0
        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        Documents(DocumentIndex).FileName = .FileName
    End With
    FileSave DocumentIndex
End Sub

Private Sub MenuFileExit_Click()
    Unload MainForm
End Sub

Private Sub MenuEditCut_Click()
    Dim I As Long
    Dim S As String
    With Documents(DocumentIndex)
        Clipboard.Clear
        S = ""
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Selected Then
                S = S & .Objects(I).Caption & vbCrLf
                S = S & .Objects(I).Style & vbCrLf
                S = S & .Objects(I).Selected & vbCrLf
                S = S & .Objects(I).X & vbCrLf
                S = S & .Objects(I).Y & vbCrLf
                S = S & .Objects(I).VX & vbCrLf
                S = S & .Objects(I).VY & vbCrLf
                S = S & .Objects(I).Mass & vbCrLf
                S = S & .Objects(I).Radius & vbCrLf
                S = S & .Objects(I).FillColor & vbCrLf
            End If
        Next
        Clipboard.SetText S
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Selected Then DeleteObject DocumentIndex, I
        Next
        RefreshWindow
        UpdateEnabled
    End With
End Sub

Private Sub MenuEditCopy_Click()
    Dim I As Long
    Dim S As String
    With Documents(DocumentIndex)
        Clipboard.Clear
        S = ""
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Selected Then
                S = S & .Objects(I).Caption & vbCrLf
                S = S & .Objects(I).Style & vbCrLf
                S = S & .Objects(I).Selected & vbCrLf
                S = S & .Objects(I).X & vbCrLf
                S = S & .Objects(I).Y & vbCrLf
                S = S & .Objects(I).VX & vbCrLf
                S = S & .Objects(I).VY & vbCrLf
                S = S & .Objects(I).Mass & vbCrLf
                S = S & .Objects(I).Radius & vbCrLf
                S = S & .Objects(I).FillColor & vbCrLf
            End If
        Next
        Clipboard.SetText S
    End With
End Sub

Private Sub MenuEditPaste_Click()
    Dim Flag As Boolean
    Dim I As Long
    Dim J As Long
    Dim S As String
    Dim T As String
    Dim AObject As TObject
    On Error GoTo ErrorHandle
    With Documents(DocumentIndex)
        S = Clipboard.GetText
        For I = 0 To UBound(.Objects)
            .Objects(I).Selected = False
        Next
        Do
            If InStr(S, vbCrLf) = 0 Then Exit Do
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.Caption = T
            Flag = False
            For I = 0 To UBound(.Objects) - 1
                If .Objects(I).Style <> osNull Then
                    If .Objects(I).Caption = AObject.Caption Then
                        Flag = True
                        Exit For
                    End If
                End If
            Next
            If Flag Then
                I = 1
                Do
                    For J = 0 To UBound(.Objects)
                        If .Objects(J).Style <> osNull And .Objects(J).Caption = "对象" & I Then Exit For
                        If J = UBound(.Objects) Then
                            AObject.Caption = "对象" & I
                            Exit Do
                        End If
                    Next
                    I = I + 1
                Loop
            End If
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.Style = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.Selected = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.X = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.Y = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.VX = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.VY = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.Mass = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.Radius = T
            If InStr(S, vbCrLf) = 0 Then GoTo ErrorHandle
            T = Left(S, InStr(S, vbCrLf) - 1)
            S = Right(S, Len(S) - Len(T) - 2)
            AObject.FillColor = T
            NewObject DocumentIndex, AObject
        Loop
    End With
    RefreshWindow
    UpdateEnabled
    Exit Sub
ErrorHandle:
    MsgBox "剪贴板数据无效。", vbOKOnly Or vbExclamation
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuEditDelete_Click()
    Dim I As Long
    For I = 0 To UBound(Documents(DocumentIndex).Objects) - 1
        If Documents(DocumentIndex).Objects(I).Selected Then DeleteObject DocumentIndex, I
    Next
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuEditSelectAll_Click()
    Dim I As Long
    For I = 0 To UBound(Documents(DocumentIndex).Objects) - 1
        Documents(DocumentIndex).Objects(I).Selected = True
    Next
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuEditSelect_Click()
    Dim SelectBox As FormSelect
    Set SelectBox = New FormSelect
    Load SelectBox
    SelectBox.Show vbModal
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuEditFind_Click()
    Dim FindBox As FormFind
    Set FindBox = New FormFind
    Load FindBox
    FindBox.Show vbModal
    RefreshWindow
End Sub

Private Sub MenuViewToolbar_Click()
    MainForm.Toolbar.Visible = Not MainForm.Toolbar.Visible
    MainForm.MenuViewToolbar.Checked = MainForm.Toolbar.Visible
    MenuViewToolbar.Checked = MainForm.Toolbar.Visible
End Sub

Private Sub MenuViewStatusBar_Click()
    MainForm.StatusBar.Visible = Not MainForm.StatusBar.Visible
    MainForm.MenuViewStatusBar.Checked = MainForm.StatusBar.Visible
    MenuViewStatusBar.Checked = MainForm.StatusBar.Visible
End Sub

Private Sub MenuViewLock_Click()
    Dim LockBox As FormLock
    Set LockBox = New FormLock
    Load LockBox
    LockBox.Show vbModal
    UpdateEnabled
End Sub

Private Sub MenuViewScreenCenter_Click()
    Dim ScreenCenterBox As FormScreenCenter
    Set ScreenCenterBox = New FormScreenCenter
    Load ScreenCenterBox
    ScreenCenterBox.Show vbModal
End Sub

Private Sub MenuViewZoom_Click()
    Dim ZoomBox As FormZoom
    Set ZoomBox = New FormZoom
    Load ZoomBox
    ZoomBox.Show vbModal
End Sub

Private Sub MenuViewCustom_Click()
    MainForm.Toolbar.Customize
End Sub

Private Sub MenuGalaxyInsertObject_Click()
    Dim I As Long
    Dim J As Long
    Dim AObject As TObject
    With AObject
        I = 1
        Do
            For J = 0 To UBound(Documents(DocumentIndex).Objects)
                If Documents(DocumentIndex).Objects(J).Style <> osNull And Documents(DocumentIndex).Objects(J).Caption = "对象" & I Then Exit For
                If J = UBound(Documents(DocumentIndex).Objects) Then
                    .Caption = "对象" & I
                    Exit Do
                End If
            Next
            I = I + 1
        Loop
        .Style = osObject
        .Selected = True
        .X = Documents(DocumentIndex).ScrollX
        .Y = Documents(DocumentIndex).ScrollY
        .VX = 0
        .VY = 0
        .Mass = 1
        .Radius = 1
        .FillColor = RGB(255, 255, 255)
    End With
    For I = 0 To UBound(Documents(DocumentIndex).Objects)
        Documents(DocumentIndex).Objects(I).Selected = False
    Next
    NewObject DocumentIndex, AObject
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuGalaxyDateTime_Click()
    Dim DateTimeBox As FormDateTime
    Set DateTimeBox = New FormDateTime
    Load DateTimeBox
    DateTimeBox.Show vbModal
    MainForm.StatusBar.SimpleText = Documents(DocumentIndex).DateTime
End Sub

Private Sub MenuGalaxyInterval_Click()
    Dim IntervalBox As FormInterval
    Set IntervalBox = New FormInterval
    Load IntervalBox
    IntervalBox.Show vbModal
End Sub

Private Sub MenuGalaxyProperties_Click()
    Dim I As Long
    Dim PropertiesBox As FormProperties
    For I = 0 To UBound(Documents(DocumentIndex).Objects)
        If I = UBound(Documents(DocumentIndex).Objects) Then Exit Sub
        If Documents(DocumentIndex).Objects(I).Selected Then Exit For
    Next
    Set PropertiesBox = New FormProperties
    Load PropertiesBox
    PropertiesBox.Show vbModal
End Sub

Private Sub MenuRunStart_Click()
    If Documents(DocumentIndex).Run And Not Documents(DocumentIndex).Break Then Exit Sub
    If Not Documents(DocumentIndex).Break Then
        Documents(DocumentIndex).Run = True
        Documents(DocumentIndex).InitDateTime = Documents(DocumentIndex).DateTime
        Documents(DocumentIndex).InitInterval = Documents(DocumentIndex).Interval
        Documents(DocumentIndex).InitObjects = Documents(DocumentIndex).Objects
    End If
    Documents(DocumentIndex).Break = False
    Timer.Enabled = True
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuRunBreak_Click()
    If Not Documents(DocumentIndex).Run Or Documents(DocumentIndex).Break Then Exit Sub
    Documents(DocumentIndex).Break = True
    Timer.Enabled = False
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuRunEnd_Click()
    If Not Documents(DocumentIndex).Run Then Exit Sub
    Documents(DocumentIndex).Run = False
    Documents(DocumentIndex).Break = False
    Documents(DocumentIndex).DateTime = Documents(DocumentIndex).InitDateTime
    Documents(DocumentIndex).Interval = Documents(DocumentIndex).InitInterval
    Documents(DocumentIndex).Objects = Documents(DocumentIndex).InitObjects
    Timer.Enabled = False
    RefreshWindow
    UpdateEnabled
    MainForm.StatusBar.SimpleText = Documents(DocumentIndex).DateTime
End Sub

Private Sub MenuRunRestart_Click()
    If Not Documents(DocumentIndex).Run Or Not Documents(DocumentIndex).Break Then Exit Sub
    Documents(DocumentIndex).Break = False
    Documents(DocumentIndex).DateTime = Documents(DocumentIndex).InitDateTime
    Documents(DocumentIndex).Interval = Documents(DocumentIndex).InitInterval
    Documents(DocumentIndex).Objects = Documents(DocumentIndex).InitObjects
    Timer.Enabled = True
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuWindowCascade_Click()
    MainForm.Arrange vbCascade
End Sub

Private Sub MenuWindowTileHorizontal_Click()
    MainForm.Arrange vbTileHorizontal
End Sub

Private Sub MenuWindowTileVertical_Click()
    MainForm.Arrange vbTileVertical
End Sub

Private Sub MenuWindowArrangeIcons_Click()
    MainForm.Arrange vbArrangeIcons
End Sub

Private Sub MenuHelpAbout_Click()
    Dim AboutBox As FormAbout
    Set AboutBox = New FormAbout
    Load AboutBox
    AboutBox.Show vbModal
End Sub

Private Sub MenuPopupObjectDelete_Click()
    Dim I As Long
    For I = 0 To UBound(Documents(DocumentIndex).Objects) - 1
        If Documents(DocumentIndex).Objects(I).Selected Then DeleteObject DocumentIndex, I
    Next
    RefreshWindow
    UpdateEnabled
End Sub

Private Sub MenuPopupObjectScreenCenter_Click()
    With Documents(DocumentIndex)
        .ScrollX = (MouseX - ScaleWidth / 2) / .Zoom + .ScrollX
        .ScrollY = (MouseY - ScaleHeight / 2) / .Zoom + .ScrollY
    End With
    RefreshWindow
End Sub

Private Sub MenuPopupObjectProperties_Click()
    Dim I As Long
    Dim PropertiesBox As FormProperties
    For I = 0 To UBound(Documents(DocumentIndex).Objects)
        If I = UBound(Documents(DocumentIndex).Objects) Then Exit Sub
        If Documents(DocumentIndex).Objects(I).Selected Then Exit For
    Next
    Set PropertiesBox = New FormProperties
    Load PropertiesBox
    PropertiesBox.Show vbModal
End Sub
