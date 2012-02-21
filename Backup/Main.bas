Attribute VB_Name = "ModuleMain"
' ******************************************************************
' *                                                                *
' *          X ��ϵ   ��Ȩ����(C)   ����   2000��05��16��          *
' *                                                                *
' *    EMAIL:wcwcwwc@263.net   HOMEPAGE:http://wcwcwwc.yeah.net    *
' *                                                                *
' ******************************************************************

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
    ScrollX As Double
    ScrollY As Double
    Objects() As TObject
    InitDateTime As Date
    InitInterval As Double
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

Public Sub Main()
    Dim StartTime As Date
    ReDim Documents(0 To 0)
    DocumentCount = 0
    Documents(0).Title = ""
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
    MsgBox "��ֻ��һ��δ��ɵ��ڲ����԰棬�������Բ�����õĽ�������(������󱨸�)�߽��ɵõ����������ʽ���ע���롣�ļ���ȡ�ʹ�ӡ������δ��ɡ���ӭ����������ҳhttp://wcwcwwc.yeah.net��"
End Sub

Public Sub UpdateEnabled()
    Dim Flag As Boolean
    Dim I As Long
    Dim Count As Long
    If MainForm.ActiveForm Is Nothing Then
        With MainForm.Toolbar
            .Buttons("�½�").Enabled = True
            .Buttons("��").Enabled = True
            .Buttons("����").Enabled = False
            .Buttons("��ӡ").Enabled = False
            .Buttons("��ӡԤ��").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("ճ��").Enabled = False
            .Buttons("ɾ��").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("��������").Enabled = False
            .Buttons("�Ŵ�").Enabled = False
            .Buttons("��С").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("����").Enabled = False
            .Buttons("�ж�").Enabled = False
            .Buttons("����").Enabled = False
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
            MainForm.Toolbar.Buttons("�½�").Enabled = True
            MainForm.Toolbar.Buttons("��").Enabled = True
            MainForm.Toolbar.Buttons("����").Enabled = .Dirty And Not .Run
            MainForm.Toolbar.Buttons("��ӡ").Enabled = Not .Run Or .Break
            MainForm.Toolbar.Buttons("��ӡԤ��").Enabled = Not .Run Or .Break
            MainForm.Toolbar.Buttons("����").Enabled = Flag And Not .Run
            MainForm.Toolbar.Buttons("����").Enabled = Flag And Not .Run
            MainForm.Toolbar.Buttons("ճ��").Enabled = Not .Run Or .Break
            MainForm.Toolbar.Buttons("ɾ��").Enabled = Flag And Not .Run
            MainForm.Toolbar.Buttons("����").Enabled = Count > 0
            MainForm.Toolbar.Buttons("��������").Enabled = Flag And (Not .Run Or .Break)
            MainForm.Toolbar.Buttons("�Ŵ�").Enabled = True
            MainForm.Toolbar.Buttons("��С").Enabled = True
            MainForm.Toolbar.Buttons("����").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.Toolbar.Buttons("����").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.Toolbar.Buttons("����").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.Toolbar.Buttons("����").Enabled = Not .Run Or .Break Or .Lock = -1
            MainForm.Toolbar.Buttons("����").Enabled = Not .Run Or .Break
            MainForm.Toolbar.Buttons("�ж�").Enabled = .Run And Not .Break
            MainForm.Toolbar.Buttons("����").Enabled = .Run
            MainForm.ActiveForm.MenuFileNew.Enabled = True
            MainForm.ActiveForm.MenuFileOpen.Enabled = True
            MainForm.ActiveForm.MenuFileClose.Enabled = True
            MainForm.ActiveForm.MenuFileSave.Enabled = .Dirty And Not .Run
            MainForm.ActiveForm.MenuFileSaveAs.Enabled = Not .Run
            MainForm.ActiveForm.MenuFilePageSetup.Enabled = True
            MainForm.ActiveForm.MenuFilePrintPreview.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuFilePrint.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuFileExit.Enabled = True
            MainForm.ActiveForm.MenuEditCut.Enabled = Flag And Not .Run
            MainForm.ActiveForm.MenuEditCopy.Enabled = Flag And Not .Run
            MainForm.ActiveForm.MenuEditPaste.Enabled = Not .Run Or .Break
            MainForm.ActiveForm.MenuEditDelete.Enabled = Flag And Not .Run
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
        .Title = "��ϵ" & DocumentCount
        .FileName = ""
        .Dirty = False
        .Run = False
        .DateTime = Now
        .Interval = 60
        .Lock = -1
        .Zoom = 15
        .ScrollX = 0
        .ScrollY = 0
        ReDim .Objects(0 To 0)
    End With
    FileNew = NewDocument(ADocument)
    UpdateEnabled
    MainForm.StatusBar.SimpleText = ADocument.DateTime
End Function

Public Function FileOpen(FileName As String) As Long
    Dim ADocument As TDocument
    ReDim ADocument.Objects(0 To 0)
    ADocument.Title = "" '***
    ADocument.FileName = FileName
    ADocument.Dirty = False
    ReDim ADocument.Objects(0 To 0) '***
    FileOpen = NewDocument(ADocument)
    UpdateEnabled
    MainForm.StatusBar.SimpleText = ADocument.DateTime
End Function

Public Sub FileSave(ADocumentIndex As Long)
    Documents(ADocumentIndex).Dirty = False
    UpdateEnabled
End Sub
