Public Class ViewController : Inherits BaseViewModel2

    Public Overrides Sub Initialize(ByRef adm As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        Me.[AddHandler] = [Delegate].Combine(
            New Action(AddressOf _TabCloseAddHandler),
            New Action(AddressOf _OpenViewAddHandler),
            New Action(AddressOf _MultiViewSizeChangedAddHandler),
            New Action(AddressOf _OpenProjectAddHandler)
        )
        Call BaseInitialize(adm, vm)
    End Sub

    '--- ビュー削除 ------------------------------------------------------------------------------'
    ' 廃止
    ' ビューに紐づいたモデルは個々のモデルなどで削除する
    'Private Sub _DeleteViewAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.DeleteViewRequested,
    '        AddressOf Me._DeleteViewRequestedReview
    'End Sub

    'Private Sub _DeleteViewRequestedReview(ByVal [view] As ViewItemModel, ByVal e As System.EventArgs)
    '    Call _DeleteViewRequestAccept([view])
    'End Sub

    'Private Sub _DeleteViewRequestAccept(ByVal [view] As ViewItemModel)
    '    If ViewModel.Views.Contains([view]) Then
    '        ViewModel.Views.Remove([view])
    '    End If
    'End Sub
    '---------------------------------------------------------------------------------------------'

    '--- タブを開く関連 --------------------------------------------------------------------------'
    Private Sub _OpenProjectAddHandler()
        AddHandler _
            DelegateEventListener.Instance.OpenProjectRequested,
            AddressOf Me._OpenProjectRequestedReview
    End Sub

    Private Sub _OpenProjectRequestedReview(ByVal p As ProjectInfoModel, ByVal e As System.EventArgs)
        Call _OpenProjectRequestAccept(p)
    End Sub

    Private Sub _OpenProjectRequestAccept(ByVal p As ProjectInfoModel)
        If AppInfo.ProjectInfo.DirectoryName <> p.DirectoryName Then
            Call InitializeViewContent()
            Call AllLoad(p)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------'

    '--- ＳＰＬＩＴＴＥＲ変更 --------------------------------------------------------------------'
    Private Sub _MultiViewSizeChangedReview(ByVal sender As Object, ByVal e As System.EventArgs)
        Call _MultiViewSizeChangeAccept(sender)
    End Sub

    Private Sub _MultiViewSizeChangeAccept(ByVal sender As Object)
        Select Case sender.Name
            Case "MainView"
                ViewModel.MultiView.MainGridHeight = New GridLength(sender.ActualHeight)
                ViewModel.MultiView.RightGridWidth = New GridLength(sender.ActualWidth)
            Case "ExplorerView"
                ViewModel.MultiView.ExplorerGridHeight = New GridLength(sender.ActualHeight)
                ViewModel.MultiView.LeftGridWidth = New GridLength(sender.ActualWidth)
            Case "HistoryView"
                ViewModel.MultiView.HistoryGridHeight = New GridLength(sender.ActualHeight)
                ViewModel.MultiView.RightGridWidth = New GridLength(sender.ActualWidth)
            Case Else
                Throw New Exception("Error View Name")
        End Select
    End Sub

    Private Sub _MultiViewSizeChangedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.MultiViewSizeChanged,
            AddressOf Me._MultiViewSizeChangedReview
    End Sub
    '---------------------------------------------------------------------------------------------'

    '--- タブを閉じる関連 ------------------------------------------------------------------------'
    Private Sub _TabCloseRequestedReview(ByVal t As TabItemModel, ByVal e As System.EventArgs)
        Call _TabCloseRequestAccept(t)
    End Sub

    Private Sub _TabCloseRequestAccept(ByVal [tab] As TabItemModel)
        Call RemoveTabItem([tab])
    End Sub

    Private Sub _TabCloseAddHandler()
        AddHandler _
            DelegateEventListener.Instance.TabCloseRequested,
            AddressOf Me._TabCloseRequestedReview
    End Sub
    '---------------------------------------------------------------------------------------------'


    '--- タブを開く関連 --------------------------------------------------------------------------'
    Private Sub _OpenViewAddHandler()
        AddHandler _
            DelegateEventListener.Instance.OpenViewRequested,
            AddressOf Me._OpenViewRequestedReview
    End Sub

    Private Sub _OpenViewRequestedReview(ByVal v As ViewItemModel, ByVal e As System.EventArgs)
        Call _OpenViewRequestAccept(v)
    End Sub

    Private Sub _OpenViewRequestAccept(ByVal v As ViewItemModel)
        'Dim idx = ViewModel.Views.IndexOf(v)
        'ViewModel.Views(idx).OpenState = True
        'Call ViewLoad(ViewModel.Views(idx))
    End Sub
    '---------------------------------------------------------------------------------------------'
End Class

