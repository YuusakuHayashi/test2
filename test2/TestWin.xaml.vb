Imports System.ComponentModel
Public Class TestWin

    Private Const DEFAULTSETTING As String = "defaultSetting.json"

    Private _mnavi As NavigationService
    Private _enavi As NavigationService

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。


        Dim loadproxy As ModelLoadProxy(Of Model)

        ' Models
        Dim m As Model
        Dim pm As New ProjectModel

        ' ViewModels
        Dim ivm As InitViewModel
        'Dim mvm As MenuViewModel

        ' Views
        Dim iv As InitView
        'Dim mv As MenuView




        Me._mnavi = Me.MainFlame.NavigationService
        Me._enavi = Me.ExplorerFlame.NavigationService


        loadproxy = AddressOf pm.ModelLoad(Of Model)

        ' モデルのロード
        m = loadproxy(DEFAULTSETTING, "All")


        If m Is Nothing Then
            m = New Model
        End If


        ' ビューモデルのロード
        ' インスタンス化時に設定があれば、それをロードする
        ivm = New InitViewModel(m)
        'mvm = New MenuViewModel(m)
        'bmvm = New BatchMenuViewModel(m)

        ' 起動時画面遷移
        If ivm.NextViewFlag Then
            'If mvm.NextViewFlag Then
            '    bmv = New BatchMenuView(bmvm)
            '    Me._mnavi.Navigate(bmv)
            'Else
            '    mv = New MenuView(mvm)
            '    Me._mnavi.Navigate(mv)
            'End If
        Else
            iv = New InitView(ivm)
            Me._mnavi.Navigate(iv)
        End If


        'エクスプローラー画面
        'dbev = New DBExplorerView(dbevm)
        'Me._enavi.Navigate(dbev)

        'AddHandler ivm.PropertyChanged, AddressOf Me._PageNavigation
        'AddHandler mvm.PropertyChanged, AddressOf Me._PageNavigation
    End Sub

    Private Sub _PageNavigation(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        '    Dim ivm As InitViewModel
        '    Dim mvm As MenuViewModel

        '    Dim mv As MenuView
        '    Dim bmv As BatchMenuView

        '    Select Case sender.GetType.Name
        '        Case "InitViewModel"
        '            If e.PropertyName = "InitFlag" Then
        '                ivm = CType(sender, InitViewModel)
        '                If ivm.InitFlag Then
        '                    mvm = New MenuViewModel
        '                    mv = New MenuView(mvm)
        '                    Me._mnavi.Navigate(mv)
        '                End If
        '            End If
        '        Case "MenuViewModel"
        '            If e.PropertyName = "MenuFlag" Then
        '                mvm = CType(sender, MenuViewModel)
        '                If mvm.MenuFlag Then
        '                    bmv = New BatchMenuView
        '                    Me._mnavi.Navigate(bmv)
        '                End If
        '            End If
        '        Case Else
        '            Exit Sub
        '    End Select
    End Sub
End Class
