Attribute VB_Name = "ModuleMain"
' ********************************************************************
' *                                                                  *
' *           X 星系   版权所有(C)   王纯   2000年05月18日           *
' *                                                                  *
' *    EMAIL: wcwcwwc@263.net   HOMEPAGE: http://wcwcwwc.yeah.net    *
' *                                                                  *
' ********************************************************************
Option Explicit

Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Public Type TObject
    Caption As String
    Style As Long
    Selected As Boolean
    X As Double
    Y As Double
    VX As Double
    VY As Double
    Mass As Double
    Radius As Double
    FillColor As Long
End Type

Public Type TDocument
    Title As String
    FileName As String
    Dirty As Boolean
    Run As Boolean
    Break As Boolean
    DateTime As Date
    Interval As Double
    Lock As Long
    Zoom As Double
    OffsetX As Double
    OffsetY As Double
    Objects() As TObject
    InitDateTime As Date
    InitInterval As Double
    InitLock As Long
    InitZoom As Double
    InitOffsetX As Double
    InitOffsetY As Double
    InitObjects() As TObject
End Type

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const osNull = 0
Public Const osObject = 1

Public Const G = 0.00000000006672

Public DocumentCount As Long
Public SplashWnd As FormSplash
Public MainForm As FormMain
Public Documents() As TDocument
Public RecentFiles(0 To 3) As String

Public Sub Main()
    Dim I As Long
    Dim StartTime As Date
    ReDim Documents(0 To 0)
    DocumentCount = 0
    Documents(0).Title = ""
    I = 0
    On Error Resume Next
    RecentFiles(I) = GetSetting("X 星系", "RecentFiles", "RecentFile1", "")
    If RecentFiles(I) <> "" Then I = I + 1
    RecentFiles(I) = GetSetting("X 星系", "RecentFiles", "RecentFile2", "")
    If RecentFiles(I) <> "" Then I = I + 1
    RecentFiles(I) = GetSetting("X 星系", "RecentFiles", "RecentFile3", "")
    If RecentFiles(I) <> "" Then I = I + 1
    RecentFiles(I) = GetSetting("X 星系", "RecentFiles", "RecentFile4", "")
    If RecentFiles(I) <> "" Then I = I + 1
    On Error GoTo 0
    Set SplashWnd = New FormSplash
    Load SplashWnd
    SetWindowPos SplashWnd.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SplashWnd.Show
    SplashWnd.Refresh
    StartTime = Now
    Do
    Loop Until (Now - StartTime) * 86400 >= 2
    Set MainForm = New FormMain
    Load MainForm
    MainForm.Show
    FileNew
    Unload SplashWnd
End Sub

Private Function LRTrim(S As String)
    Dim I As Integer
    If S = "" Then
        LRTrim = ""
        Exit Function
    End If
    I = 1
    Do
        If Mid(S, I, 1) <> " " And Mid(S, I, 1) <> Chr(9) Then Exit Do
        I = I + 1
    Loop
    LRTrim = Right(S, Len(S) - I + 1)
    I = Len(LRTrim)
    Do
        If Mid(LRTrim, I, 1) <> " " And Mid(LRTrim, I, 1) <> Chr(9) Then Exit Do
        I = I - 1
    Loop
    LRTrim = Left(LRTrim, I)
End Function

Public Sub UpdateEnabled()
    Dim Flag As Boolean
    Dim I As Long
    Dim Count As Long
    If MainForm.ActiveForm Is Nothing Then
        With MainForm.Toolbar
            .Buttons("新建").Enabled = True
            .Buttons("打开").Enabled = True
            .Buttons("保存").Enabled = False
            .Buttons("剪切").Enabled = False
            .Buttons("复制").Enabled = False
            .Buttons("粘贴").Enabled = False
            .Buttons("删除").Enabled = False
            .Buttons("查找").Enabled = False
            .Buttons("对象属性").Enabled = False
            .Buttons("启动").Enabled = False
            .Buttons("中断").Enabled = False
            .Buttons("结束").Enabled = False
            .Buttons("放大").Enabled = False
            .Buttons("缩小").Enabled = False
            .Buttons("向上").Enabled = False
            .Buttons("向下").Enabled = False
            .Buttons("向左").Enabled = False
            .Buttons("向右").Enabled = False
        End With
    Else
        With Documents(MainForm.ActiveForm.DocumentIndex)
            Flag = False
            Count = 0
            For I = 0 To UBound(.Objects) - 1
                If .Objects(I).Style <> osNull Then
                    If .Objects(I).Selected Then Flag = True
                    Count = Count + 1
                End If
            Next
            MainForm.Toolbar.Buttons("新建").Enabled = True
            MainForm.Toolbar.Buttons("打开").Enabled = True
            MainForm.Toolbar.Buttons("保存").Enabled = Not .Run
            MainForm.Toolbar.Buttons("剪切").Enabled = Flag And (Not .Run Or .Break)
            MainForm.Toolbar.Buttons("复制").Enabled = Flag And (Not .Run Or .Break)
            MainForm.Toolbar.Buttons("粘贴").Enabled = Not .Run Or .Break
            MainForm.Toolbar.Buttons("删除").Enabled = Flag And (Not .Run Or .Break)
            MainForm.Toolbar.Buttons("查找").Enabled = Count > 0
            MainForm.Toolbar.Buttons("对象属性").Enabled = Flag And (Not .Run Or .Break)
            MainForm.Toolbar.Buttons("启动").Enabled = Not .Run Or .Break
            MainForm.Toolbar.Buttons("中断").Enabled = .Run And Not .Break
            MainForm.Toolbar.Buttons("结束").Enabled = .Run
            MainForm.Toolbar.Buttons("放大").Enabled = True
            MainForm.Toolbar.Buttons("缩小").Enabled = True
            MainForm.Toolbar.Buttons("向上").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.Toolbar.Buttons("向下").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.Toolbar.Buttons("向左").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.Toolbar.Buttons("向右").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.ActiveForm.MenuFileNew.Enabled = True
            MainForm.ActiveForm.MenuFileOpen.Enabled = True
            MainForm.ActiveForm.MenuFileClose.Enabled = True
            MainForm.ActiveForm.MenuFileSave.Enabled = Not .Run
            MainForm.ActiveForm.MenuFileSaveAs.Enabled = Not .Run
            MainForm.ActiveForm.MenuFileExit.Enabled = True
            MainForm.ActiveForm.MenuEditCut.Enabled = Flag And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuEditCopy.Enabled = Flag And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuEditPaste.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuEditDelete.Enabled = Flag And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuEditBringFront.Enabled = Flag And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuEditBringBack.Enabled = Flag And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuEditSelectAll.Enabled = Count > 0 And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuEditSelect.Enabled = Count > 0 And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuEditFind.Enabled = Count > 0
            MainForm.ActiveForm.MenuViewToolbar.Enabled = True
            MainForm.ActiveForm.MenuViewStatusBar.Enabled = True
            MainForm.ActiveForm.MenuViewLock.Enabled = Count > 0
            MainForm.ActiveForm.MenuViewScreenCenter.Enabled = True
            MainForm.ActiveForm.MenuViewZoom.Enabled = True
            MainForm.ActiveForm.MenuViewCustom.Enabled = True
            MainForm.ActiveForm.MenuGalaxyInsertObject.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuGalaxyDateTime.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuGalaxyInterval.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuGalaxyProperties.Enabled = Flag And (Not .Run Or .Break)
            MainForm.ActiveForm.MenuRunStart.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuRunBreak.Enabled = .Run And Not .Break
            MainForm.ActiveForm.MenuRunEnd.Enabled = .Run
            MainForm.ActiveForm.MenuRunRestart.Enabled = .Run And .Break
            MainForm.ActiveForm.MenuWindowCascade.Enabled = True
            MainForm.ActiveForm.MenuWindowTileHorizontal.Enabled = True
            MainForm.ActiveForm.MenuWindowTileVertical.Enabled = True
            MainForm.ActiveForm.MenuWindowArrangeIcons.Enabled = True
            MainForm.ActiveForm.MenuHelpAbout.Enabled = True
        End With
    End If
    Flag = False
    For I = 0 To 3
        If Len(RecentFiles(I)) <> 0 Then
            MainForm.MenuFileRecent(I).Visible = Len(RecentFiles(I)) <> 0
            MainForm.ActiveForm.MenuFileRecent(I).Visible = Len(RecentFiles(I)) <> 0
            Flag = True
        End If
        MainForm.MenuFileRecent(I).Caption = "&" & I + 1 & " " & RecentFiles(I)
        MainForm.ActiveForm.MenuFileRecent(I).Caption = "&" & I + 1 & " " & RecentFiles(I)
    Next
    MainForm.MenuFileBar1.Visible = Flag
    MainForm.ActiveForm.MenuFileBar2.Visible = Flag
End Sub

Public Function NewDocument(ADocument As TDocument) As Long
    Dim I As Long
    Dim DocumentWnd As FormDocument
    Set DocumentWnd = New FormDocument
    For I = 0 To UBound(Documents)
        If Documents(I).Title = "" Then
            NewDocument = I
            Documents(I) = ADocument
            DocumentWnd.DocumentIndex = I
            If I = UBound(Documents) Then
                ReDim Preserve Documents(0 To I + 1)
                Documents(I + 1).Title = ""
            End If
            Exit For
        End If
    Next
    Load DocumentWnd
    DocumentWnd.MenuViewToolbar.Checked = MainForm.Toolbar.Visible
    DocumentWnd.MenuViewStatusBar.Checked = MainForm.StatusBar.Visible
    DocumentWnd.Show
End Function

Public Function NewObject(ADocumentIndex As Long, AObject As TObject) As Long
    Dim I As Long
    Documents(ADocumentIndex).Dirty = True
    For I = 0 To UBound(Documents(ADocumentIndex).Objects)
        If Documents(ADocumentIndex).Objects(I).Style = osNull Then
            NewObject = I
            Documents(ADocumentIndex).Objects(I) = AObject
            If I = UBound(Documents(ADocumentIndex).Objects) Then
                ReDim Preserve Documents(ADocumentIndex).Objects(0 To I + 1)
                Documents(ADocumentIndex).Objects(I + 1).Style = osNull
            End If
            Exit For
        End If
    Next
    UpdateEnabled
End Function

Public Sub DeleteObject(ADocumentIndex As Long, AObjectIndex As Long)
    Documents(ADocumentIndex).Dirty = True
    Documents(ADocumentIndex).Objects(AObjectIndex).Style = osNull
    Documents(ADocumentIndex).Objects(AObjectIndex).Selected = False
    If Documents(ADocumentIndex).Lock = AObjectIndex Then Documents(ADocumentIndex).Lock = -1
    UpdateEnabled
End Sub

Public Function FileNew() As Long
    Dim ADocument As TDocument
    DocumentCount = DocumentCount + 1
    With ADocument
        .Title = "星系" & DocumentCount
        .FileName = ""
        .Dirty = False
        .Run = False
        .Break = False
        .DateTime = Now
        .Interval = 60
        .Lock = -1
        .Zoom = 15
        .OffsetX = 0
        .OffsetY = 0
        ReDim .Objects(0 To 0)
    End With
    FileNew = NewDocument(ADocument)
    UpdateEnabled
    MainForm.StatusBar.SimpleText = ADocument.DateTime
End Function

Public Function FileOpen(FileName As String) As Long
    Dim I As Long
    Dim S As String
    Dim Key As String
    Dim Value As String
    Dim ADocument As TDocument
    Dim AObject As TObject
    FileOpen = -1
    On Error GoTo ErrorHandle
    With ADocument
        ReDim .Objects(0 To 0)
        Open FileName For Input As #1
        .Title = FileName
        .FileName = FileName
        .Dirty = False
        .Run = False
        .Break = False
        .DateTime = Now
        .Interval = 60
        .Lock = -1
        .Zoom = 15
        .OffsetX = 0
        .OffsetY = 0
        Do
            Line Input #1, S
            S = LRTrim(S)
        Loop Until S <> "" And Left(S, 1) <> ";"
        If UCase(Left(S, InStr(S, " ") - 1)) <> "XGALAXY" Then GoTo ErrorHandle
        If UCase(LRTrim(Right(S, Len(S) - InStr(S, " ")))) <> "1.0" Then
            If MsgBox("打开一个较高版本的 X 星系文件会出现不可预料的结果，是否坚持打开？", vbYesNo Or vbExclamation) = vbNo Then Exit Function
        End If
        Do
            Line Input #1, S
            S = LRTrim(S)
        Loop Until S <> "" And Left(S, 1) <> ";"
        If UCase(S) <> "BEGIN" Then GoTo ErrorHandle
        Do
            Do
                Line Input #1, S
                S = LRTrim(S)
            Loop Until S <> "" And Left(S, 1) <> ";"
            If UCase(S) = "END" Then Exit Do
            Key = UCase(LRTrim(Left(S, InStr(S, "=") - 1)))
            Value = LRTrim(Right(S, Len(S) - InStr(S, "=")))
            Select Case Key
                Case "RUN"
                    .Run = Value
                Case "BREAK"
                    .Break = Value
                Case "DATETIME"
                    .DateTime = CDate(Value)
                Case "INTERVAL"
                    .Interval = Value
                Case "LOCK"
                    .Lock = Value
                Case "ZOOM"
                    .Zoom = Value
                Case "OFFSETX"
                    .OffsetX = Value
                Case "OFFSETY"
                    .OffsetY = Value
                Case "OBJECT"
                    AObject.Caption = Value
                    Do
                        Line Input #1, S
                        S = LRTrim(S)
                    Loop Until S <> "" And Left(S, 1) <> ";"
                    If UCase(S) <> "BEGIN" Then GoTo ErrorHandle
                    AObject.Style = osObject
                    AObject.Selected = False
                    AObject.X = .OffsetX
                    AObject.Y = .OffsetY
                    AObject.VX = 0
                    AObject.VY = 0
                    AObject.Mass = 1
                    AObject.Radius = 1
                    AObject.FillColor = RGB(255, 255, 255)
                    Do
                        Do
                            Line Input #1, S
                            S = LRTrim(S)
                        Loop Until S <> "" And Left(S, 1) <> ";"
                        If UCase(S) = "END" Then Exit Do
                        Key = UCase(LRTrim(Left(S, InStr(S, "=") - 1)))
                        Value = LRTrim(Right(S, Len(S) - InStr(S, "=")))
                        Select Case Key
                            Case "X"
                                AObject.X = Value
                            Case "Y"
                                AObject.Y = Value
                            Case "VX"
                                AObject.VX = Value
                            Case "VY"
                                AObject.VY = Value
                            Case "MASS"
                                AObject.Mass = Value
                            Case "RADIUS"
                                AObject.Radius = Value
                            Case "FILLCOLOR"
                                AObject.FillColor = "&H" & Value
                        End Select
                    Loop
                    ADocument.Objects(UBound(ADocument.Objects)) = AObject
                    ReDim Preserve ADocument.Objects(0 To UBound(ADocument.Objects) + 1)
            End Select
        Loop Until EOF(1)
        Close #1
        NewDocument ADocument
        UpdateEnabled
        MainForm.StatusBar.SimpleText = .DateTime
    End With
    Exit Function
ErrorHandle:
    Close #1
    MsgBox "打开文件时发生错误。", vbOKOnly Or vbExclamation
End Function

Public Sub FileSave(ADocumentIndex As Long)
    Dim I As Long
    On Error GoTo ErrorHandle
    With Documents(ADocumentIndex)
        .Title = .FileName
        Open .FileName For Output As #1
        Print #1, "; ********************************************************************"
        Print #1, "; *                                                                  *"
        Print #1, "; *           X 星系   版权所有(C)   王纯   2000年05月16日           *"
        Print #1, "; *                                                                  *"
        Print #1, "; *    EMAIL: wcwcwwc@263.net   HOMEPAGE: http://wcwcwwc.yeah.net    *"
        Print #1, "; *                                                                  *"
        Print #1, "; ********************************************************************"
        Print #1,
        Print #1, "XGalaxy 1.0"
        Print #1, "Begin"
        Print #1, Chr(9) & "Run = " & .Run
        Print #1, Chr(9) & "Break = " & .Break
        Print #1, Chr(9) & "DateTime = " & CDbl(.DateTime)
        Print #1, Chr(9) & "Interval = " & .Interval
        Print #1, Chr(9) & "Lock = " & .Lock
        Print #1, Chr(9) & "Zoom = " & .Zoom
        Print #1, Chr(9) & "OffsetX = " & .OffsetX
        Print #1, Chr(9) & "OffsetY = " & .OffsetY
        For I = 0 To UBound(.Objects) - 1
            If .Objects(I).Style <> osNull Then
                Print #1, Chr(9) & "Object = " & .Objects(I).Caption
                Print #1, Chr(9) & "Begin"
                Print #1, Chr(9) & Chr(9) & "X = " & .Objects(I).X
                Print #1, Chr(9) & Chr(9) & "Y = " & .Objects(I).Y
                Print #1, Chr(9) & Chr(9) & "VX = " & .Objects(I).VX
                Print #1, Chr(9) & Chr(9) & "VY = " & .Objects(I).VY
                Print #1, Chr(9) & Chr(9) & "Mass = " & .Objects(I).Mass
                Print #1, Chr(9) & Chr(9) & "Radius = " & .Objects(I).Radius
                If Len(Hex(.Objects(I).FillColor)) < 6 Then
                    Print #1, Chr(9) & Chr(9) & "FillColor = " & String(6 - Len(Hex(.Objects(I).FillColor)), "0") & Hex(.Objects(I).FillColor)
                Else
                    Print #1, Chr(9) & Chr(9) & "FillColor = " & Hex(.Objects(I).FillColor)
                End If
                Print #1, Chr(9) & "End"
            End If
        Next
        Print #1, "End"
        Close #1
        .Dirty = False
        UpdateEnabled
    End With
    Exit Sub
ErrorHandle:
    Close #1
    MsgBox "保存文件时发生错误。", vbOKOnly Or vbExclamation
End Sub
