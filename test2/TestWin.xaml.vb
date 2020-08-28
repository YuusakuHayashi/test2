Imports System.ComponentModel
Public Class TestWin

    Private Const DEFAULTSETTING As String = "defaultSetting.json"

    Private _mnavi As NavigationService
    Private _enavi As NavigationService

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。


        Dim loadproxy As ModelLoadProxy1(Of Model)
        Dim saveproxy As ModelSaveProxy1(Of Model)

        ' Models
        Dim m As Model
        Dim pm As New ProjectModel

        ' ViewModels
        Dim ivm As InitViewModel
        Dim dbevm As DBExplorerViewModel
        'Dim mvm As MenuViewModel

        ' Views
        Dim iv As InitView
        Dim dbev As DBExplorerView
        'Dim mv As MenuView


        Me._mnavi = Me.MainFlame.NavigationService
        Me._enavi = Me.ExplorerFlame.NavigationService

        loadproxy = AddressOf pm.ModelLoad(Of Model)
        saveproxy = AddressOf pm.ModelSave(Of Model)


        ' モデルのロード
        m = loadproxy(DEFAULTSETTING)


        ' なければＪＳＯＮをとりあえず作っておく
        If m Is Nothing Then
            m = New Model
            Call saveproxy(DEFAULTSETTING, m)
        End If



        ' メイン画面ビューモデル
        ivm = New InitViewModel(m)

        ' 初期メイン画面
        If ivm.NextFlag Then
        Else
            iv = New InitView(ivm)
            Me._mnavi.Navigate(iv)
        End If


        ' エクスプローラー画面ビューモデル
        dbevm = New DBExplorerViewModel(m)




        ' 初期エクスプローラー画面
        dbev = New DBExplorerView(dbevm)
        'Me._enavi.Navigate(dbev)

        AddHandler m.PropertyChanged, AddressOf Me._PageNavigation
        'AddHandler mvm.PropertyChanged, AddressOf Me._PageNavigation
    End Sub

    Private Sub _PageNavigation(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        Dim m As Model
        Dim pm As New ProjectModel
        Dim saveproxy As ModelSaveProxy1(Of Model)

        saveproxy = AddressOf pm.ModelSave(Of Model)

        If e.PropertyName = "ChangePageStrings" Then
            m = CType(sender, Model)
            saveproxy(DEFAULTSETTING, m)
            MsgBox("hoge")
        End If

        'Select Case sender.GetType.Name
        '    Case "InitViewModel"
        '        If e.PropertyName = "InitFlag" Then
        '            ivm = CType(sender, InitViewModel)
        '            If ivm.InitFlag Then
        '                mvm = New MenuViewModel
        '                mv = New MenuView(mvm)
        '                Me._mnavi.Navigate(mv)
        '            End If
        '        End If
        '        'Case "MenuViewModel"
        '        '    If e.PropertyName = "MenuFlag" Then
        '        '        mvm = CType(sender, MenuViewModel)
        '        '        If mvm.MenuFlag Then
        '        '            bmv = New BatchMenuView
        '        '            Me._mnavi.Navigate(bmv)
        '        '        End If
        '        '    End If
        '    Case Else
        '        Exit Sub
        'End Select
    End Sub
End Class
