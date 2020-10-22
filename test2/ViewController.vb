Public Class ViewController : Inherits BaseViewModel2

    Public Overrides ReadOnly Property FrameType As String
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides Sub Initialize(ByRef m As Model, ByRef vm As ViewModel, ByRef adm As AppDirectoryModel, ByRef pim As ProjectInfoModel)
        Me.[AddHandler] = [Delegate].Combine(
            New Action(AddressOf _TabCloseAddHandler),
            New Action(AddressOf _OpenViewAddHandler)
        )
        Call BaseInitialize(m, vm, adm, pim)
    End Sub

    '---------------------------------------------------------------------------------------------'
    ' このオブジェクト自体を上書きしてしまうと、ＭＶＶＭが適用されなくなる
    ' (Ｖｉｅｗへの反映がされない)ため、ロードが必要なメンバをここで別個にセットする
    'Public Overloads Sub Initialize(ByVal vm As ViewModel)
    '    '--- you henkou ----------------------------------'
    '    ViewModel.Views = vm.Views
    '    '-------------------------------------------------'
    'End Sub

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
        Dim idx = ViewModel.Views.IndexOf(v)
        ViewModel.Views(idx).OpenState = True
        If ViewModel.Views(idx).Content Is Nothing Then
            obj = GetViewOfName(ViewModel.Views(idx).Name)
            Call obj.Initialize(Model, ViewModel, AppInfo, ProjectInfo)
            ViewModel.Views(idx).Content = obj
        End If
        Call Me.AddView(ViewModel.Views(idx))
    End Sub
    '---------------------------------------------------------------------------------------------'
End Class

