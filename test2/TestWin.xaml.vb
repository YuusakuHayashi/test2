Imports System.ComponentModel
Public Class TestWin

    Private _mnavi As NavigationService
    Private _enavi As NavigationService

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        Dim pm As New ProjectModel
        'これら各ビューモデルは、
        'インスタンス化時に設定があれば、それをロードする
        Dim ivm As New InitViewModel
        Dim mvm As New MenuViewModel
        Dim dbevm As New DBExplorerViewModel


        Dim iv As InitView
        Dim mv As MenuView
        Dim bmv As BatchMenuView
        Dim dbev As DBExplorerView


        Me._mnavi = Me.MainFlame.NavigationService
        Me._enavi = Me.ExplorerFlame.NavigationService


        'メイン画面
        If Not ivm.InitFlag Then
            iv = New InitView(ivm)
            Me._mnavi.Navigate(iv)
        Else
            If Not mvm.MenuFlag Then
                mv = New MenuView(mvm)
                Me._mnavi.Navigate(mv)
            Else
                bmv = New BatchMenuView
                Me._mnavi.Navigate(bmv)
            End If
        End If


        'エクスプローラー画面
        dbev = New DBExplorerView(dbevm)
        Me._enavi.Navigate(dbev)


        'Dim tvm As New TreeViewModel
        'Me.treeView.DataContext = tvm

        AddHandler ivm.PropertyChanged, AddressOf Me._PageNavigation
        AddHandler mvm.PropertyChanged, AddressOf Me._PageNavigation
    End Sub

    Private Sub _PageNavigation(ByVal sender As Object, ByVal e As PropertyChangedEventArgs)
        Dim ivm As InitViewModel
        Dim mvm As MenuViewModel

        Dim mv As MenuView
        Dim bmv As BatchMenuView

        Select Case sender.GetType.Name
            Case "InitViewModel"
                If e.PropertyName = "InitFlag" Then
                    ivm = CType(sender, InitViewModel)
                    If ivm.InitFlag Then
                        mvm = New MenuViewModel
                        mv = New MenuView(mvm)
                        Me._mnavi.Navigate(mv)
                    End If
                End If
            Case "MenuViewModel"
                If e.PropertyName = "MenuFlag" Then
                    mvm = CType(sender, MenuViewModel)
                    If mvm.MenuFlag Then
                        bmv = New BatchMenuView
                        Me._mnavi.Navigate(bmv)
                    End If
                End If
            Case Else
                Exit Sub
        End Select
    End Sub
End Class
