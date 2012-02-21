VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm FormMain 
   BackColor       =   &H8000000C&
   Caption         =   "X 星系"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9480
   Icon            =   "Main.frx":0000
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      ButtonWidth     =   609
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageListToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "新建"
            Description     =   "新建"
            Object.ToolTipText     =   "新建"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "打开"
            Description     =   "打开"
            Object.ToolTipText     =   "打开"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "保存"
            Description     =   "保存"
            Object.ToolTipText     =   "保存"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "剪切"
            Description     =   "剪切"
            Object.ToolTipText     =   "剪切"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "复制"
            Description     =   "复制"
            Object.ToolTipText     =   "复制"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "粘贴"
            Description     =   "粘贴"
            Object.ToolTipText     =   "粘贴"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "删除"
            Description     =   "删除"
            Object.ToolTipText     =   "删除"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "查找"
            Description     =   "查找"
            Object.ToolTipText     =   "查找"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "对象属性"
            Description     =   "对象属性"
            Object.ToolTipText     =   "对象属性"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "启动"
            Description     =   "启动"
            Object.ToolTipText     =   "启动"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "中断"
            Description     =   "中断"
            Object.ToolTipText     =   "中断"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "结束"
            Description     =   "结束"
            Object.ToolTipText     =   "结束"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "放大"
            Description     =   "放大"
            Object.ToolTipText     =   "放大"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "缩小"
            Description     =   "缩小"
            Object.ToolTipText     =   "缩小"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "向上"
            Description     =   "向上"
            Object.ToolTipText     =   "向上"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "向下"
            Description     =   "向下"
            Object.ToolTipText     =   "向下"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "向左"
            Description     =   "向左"
            Object.ToolTipText     =   "向左"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "向右"
            Description     =   "向右"
            Object.ToolTipText     =   "向右"
            ImageIndex      =   18
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListToolbarIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":030A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":041E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0532
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0646
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":075A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":086E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0982
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0A96
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0FEA
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":10FE
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1652
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1BA6
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":20FA
            Key             =   "ZoomIn"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":264E
            Key             =   "ZoomOut"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":2BA2
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":30F6
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":364A
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":3B9E
            Key             =   "Right"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Begin VB.Menu MenuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileRecent 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileRecent 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileRecent 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileRecent 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileBar1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "退出(&X)"
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
      Begin VB.Menu MenuViewCustom 
         Caption         =   "自定义(&C)..."
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "关于 X 星系(&A)..."
      End
   End
End
Attribute VB_Name = "FormMain"
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

Private Sub MDIForm_Unload(Cancel As Integer)
    SaveSetting "X 星系", "RecentFiles", "RecentFile1", RecentFiles(0)
    SaveSetting "X 星系", "RecentFiles", "RecentFile2", RecentFiles(1)
    SaveSetting "X 星系", "RecentFiles", "RecentFile3", RecentFiles(2)
    SaveSetting "X 星系", "RecentFiles", "RecentFile4", RecentFiles(3)
End Sub

Private Sub MenuFileNew_Click()
    FileNew
End Sub

Private Sub MenuFileOpen_Click()
    With CommonDialog
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

Private Sub MenuFileRecent_Click(Index As Integer)
    FileOpen RecentFiles(Index)
End Sub

Private Sub MenuFileExit_Click()
    Unload Me
End Sub

Private Sub MenuViewToolbar_Click()
    Toolbar.Visible = Not Toolbar.Visible
    MenuViewToolbar.Checked = Toolbar.Visible
End Sub

Private Sub MenuViewStatusBar_Click()
    StatusBar.Visible = Not StatusBar.Visible
    MenuViewStatusBar.Checked = StatusBar.Visible
End Sub

Private Sub MenuViewCustom_Click()
    Toolbar.Customize
End Sub

Private Sub MenuHelpAbout_Click()
    Dim AboutBox As FormAbout
    Set AboutBox = New FormAbout
    Load AboutBox
    AboutBox.Show vbModal
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim Flag As Boolean
    Dim I As Long
    Dim J As Long
    Dim S As String
    Dim T As String
    Dim AObject As TObject
    Dim FindBox As FormFind
    Dim PropertiesBox As FormProperties
    Select Case Button.Key
        Case "新建"
            FileNew
        Case "打开"
            With CommonDialog
                .CancelError = False
                .FileName = ""
                .Filter = "星系文件 (*.gal)|*.gal|所有文件 (*.*)|*.*"
                .FilterIndex = 0
                .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
                .ShowOpen
                If Len(.FileName) = 0 Then Exit Sub
                FileOpen .FileName
            End With
        Case "保存"
            If Not ActiveForm Is Nothing Then
                If Documents(ActiveForm.DocumentIndex).FileName = "" Then
                    With MainForm.CommonDialog
                        .CancelError = False
                        .FileName = ""
                        .Filter = "星系文件 (*.gal)|*.gal|所有文件 (*.*)|*.*"
                        .FilterIndex = 0
                        .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
                        .ShowSave
                        If Len(.FileName) = 0 Then Exit Sub
                        Documents(ActiveForm.DocumentIndex).FileName = .FileName
                    End With
                End If
                FileSave ActiveForm.DocumentIndex
                ActiveForm.Caption = Documents(ActiveForm.DocumentIndex).Title
            End If
        Case "剪切"
            With Documents(ActiveForm.DocumentIndex)
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
                    If .Objects(I).Selected Then DeleteObject ActiveForm.DocumentIndex, I
                Next
                ActiveForm.RefreshWindow
                UpdateEnabled
            End With
        Case "复制"
            With Documents(ActiveForm.DocumentIndex)
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
        Case "粘贴"
            On Error GoTo ErrorHandle
            With Documents(ActiveForm.DocumentIndex)
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
                    NewObject ActiveForm.DocumentIndex, AObject
                Loop
            End With
            ActiveForm.RefreshWindow
        Case "删除"
            If Not ActiveForm Is Nothing Then
                For I = 0 To UBound(Documents(ActiveForm.DocumentIndex).Objects) - 1
                    If Documents(ActiveForm.DocumentIndex).Objects(I).Selected Then DeleteObject ActiveForm.DocumentIndex, I
                Next
                ActiveForm.RefreshWindow
            End If
        Case "查找"
            If Not ActiveForm Is Nothing Then
                Set FindBox = New FormFind
                Load FindBox
                FindBox.Show vbModal
                ActiveForm.RefreshWindow
            End If
        Case "对象属性"
            If Not ActiveForm Is Nothing Then
                For I = 0 To UBound(Documents(ActiveForm.DocumentIndex).Objects)
                    If I = UBound(Documents(ActiveForm.DocumentIndex).Objects) Then Exit Sub
                    If Documents(ActiveForm.DocumentIndex).Objects(I).Selected Then Exit For
                Next
                Set PropertiesBox = New FormProperties
                Load PropertiesBox
                PropertiesBox.Show vbModal
            End If
        Case "放大"
            If Not ActiveForm Is Nothing Then
                Documents(ActiveForm.DocumentIndex).Zoom = Documents(ActiveForm.DocumentIndex).Zoom * 1.1
                ActiveForm.RefreshWindow
            End If
        Case "缩小"
            If Not ActiveForm Is Nothing Then
                Documents(ActiveForm.DocumentIndex).Zoom = Documents(ActiveForm.DocumentIndex).Zoom / 1.1
                ActiveForm.RefreshWindow
            End If
        Case "向上"
            If Not ActiveForm Is Nothing Then
                Documents(ActiveForm.DocumentIndex).OffsetY = Documents(ActiveForm.DocumentIndex).OffsetY - ActiveForm.ScaleHeight / Documents(ActiveForm.DocumentIndex).Zoom * 0.1
                ActiveForm.RefreshWindow
            End If
        Case "向下"
            If Not ActiveForm Is Nothing Then
                Documents(ActiveForm.DocumentIndex).OffsetY = Documents(ActiveForm.DocumentIndex).OffsetY + ActiveForm.ScaleHeight / Documents(ActiveForm.DocumentIndex).Zoom * 0.1
                ActiveForm.RefreshWindow
            End If
        Case "向左"
            If Not ActiveForm Is Nothing Then
                Documents(ActiveForm.DocumentIndex).OffsetX = Documents(ActiveForm.DocumentIndex).OffsetX - ActiveForm.ScaleWidth / Documents(ActiveForm.DocumentIndex).Zoom * 0.1
                ActiveForm.RefreshWindow
            End If
        Case "向右"
            If Not ActiveForm Is Nothing Then
                Documents(ActiveForm.DocumentIndex).OffsetX = Documents(ActiveForm.DocumentIndex).OffsetX + ActiveForm.ScaleWidth / Documents(ActiveForm.DocumentIndex).Zoom * 0.1
                ActiveForm.RefreshWindow
            End If
        Case "启动"
            If Documents(ActiveForm.DocumentIndex).Run And Not Documents(ActiveForm.DocumentIndex).Break Then Exit Sub
            If Not Documents(ActiveForm.DocumentIndex).Break Then
                Documents(ActiveForm.DocumentIndex).Run = True
                Documents(ActiveForm.DocumentIndex).InitDateTime = Documents(ActiveForm.DocumentIndex).DateTime
                Documents(ActiveForm.DocumentIndex).InitInterval = Documents(ActiveForm.DocumentIndex).Interval
                Documents(ActiveForm.DocumentIndex).InitLock = Documents(ActiveForm.DocumentIndex).Lock
                Documents(ActiveForm.DocumentIndex).InitZoom = Documents(ActiveForm.DocumentIndex).Zoom
                Documents(ActiveForm.DocumentIndex).InitOffsetX = Documents(ActiveForm.DocumentIndex).OffsetX
                Documents(ActiveForm.DocumentIndex).InitOffsetY = Documents(ActiveForm.DocumentIndex).OffsetY
                Documents(ActiveForm.DocumentIndex).InitObjects = Documents(ActiveForm.DocumentIndex).Objects
            End If
            Documents(ActiveForm.DocumentIndex).Break = False
            ActiveForm.Timer.Enabled = True
            ActiveForm.RefreshWindow
            UpdateEnabled
        Case "中断"
            If Not Documents(ActiveForm.DocumentIndex).Run Or Documents(ActiveForm.DocumentIndex).Break Then Exit Sub
            Documents(ActiveForm.DocumentIndex).Break = True
            ActiveForm.Timer.Enabled = False
            ActiveForm.RefreshWindow
            UpdateEnabled
        Case "结束"
            If Not Documents(ActiveForm.DocumentIndex).Run Then Exit Sub
            Documents(ActiveForm.DocumentIndex).Run = False
            Documents(ActiveForm.DocumentIndex).Break = False
            Documents(ActiveForm.DocumentIndex).DateTime = Documents(ActiveForm.DocumentIndex).InitDateTime
            Documents(ActiveForm.DocumentIndex).Interval = Documents(ActiveForm.DocumentIndex).InitInterval
            Documents(ActiveForm.DocumentIndex).Lock = Documents(ActiveForm.DocumentIndex).InitLock
            Documents(ActiveForm.DocumentIndex).Zoom = Documents(ActiveForm.DocumentIndex).InitZoom
            Documents(ActiveForm.DocumentIndex).OffsetX = Documents(ActiveForm.DocumentIndex).InitOffsetX
            Documents(ActiveForm.DocumentIndex).OffsetY = Documents(ActiveForm.DocumentIndex).InitOffsetY
            Documents(ActiveForm.DocumentIndex).Objects = Documents(ActiveForm.DocumentIndex).InitObjects
            ActiveForm.Timer.Enabled = False
            ActiveForm.RefreshWindow
            UpdateEnabled
            StatusBar.SimpleText = Documents(ActiveForm.DocumentIndex).DateTime
    End Select
    UpdateEnabled
    Exit Sub
ErrorHandle:
    MsgBox "剪贴板数据无效。", vbOKOnly Or vbExclamation
End Sub
