Public Class ViewController : Inherits BaseViewModel2

    Public Overrides Sub Initialize(ByRef adm As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        Me.[AddHandler] = [Delegate].Combine(
            New Action(AddressOf _TabCloseAddHandler),
            New Action(AddressOf _OpenViewAddHandler)
        )
        Call BaseInitialize(adm, vm)
    End Sub


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
        Dim obj As Object
        Dim [define] = ViewDefineHandler
        Dim idx = ViewModel.Views.IndexOf(v)
        ViewModel.Views(idx).OpenState = True
        obj = [define](ViewModel.Views(idx))
        Call obj.Initialize(AppInfo, ViewModel)
        ViewModel.Views(idx).Content = obj
        Call AddView(ViewModel.Views(idx))
    End Sub
    '---------------------------------------------------------------------------------------------'
End Class

