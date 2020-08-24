Imports System.ComponentModel
Public Class TestWin

    Private _Navi As NavigationService

    Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        Dim pm As New ProjectModel
        'これら各ビューモデルは、
        'インスタンス化時に設定があれば、それをロードする
        Dim ivm As New InitViewModel
        Dim mvm As New MenuViewModel


        Dim iv As InitView
        Dim mv As MenuView
        Dim bmv As BatchMenuView


        Me._Navi = Me.MainFlame.NavigationService


        '初期時の画面
        If Not ivm.InitFlag Then
            iv = New InitView(ivm)
            Me._Navi.Navigate(iv)
        Else
            If Not mvm.MenuFlag Then
                mv = New MenuView(mvm)
                Me._Navi.Navigate(mv)
            Else
                bmv = New BatchMenuView
                Me._Navi.Navigate(bmv)
            End If
        End If


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
                        Me._Navi.Navigate(mv)
                    End If
                End If
            Case "MenuViewModel"
                If e.PropertyName = "MenuFlag" Then
                    mvm = CType(sender, MenuViewModel)
                    If mvm.MenuFlag Then
                        bmv = New BatchMenuView
                        Me._Navi.Navigate(bmv)
                    End If
                End If
            Case Else
                Exit Sub
        End Select
    End Sub
End Class
