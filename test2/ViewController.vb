Public Class ViewController : Inherits BaseViewModel2

    Public Overrides Sub Initialize(ByRef adm As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        Me.[AddHandler] = [Delegate].Combine(
            New Action(AddressOf _OpenProjectAddHandler),
            New Action(AddressOf _TabViewClosedAddHandler),
            New Action(AddressOf _ReloadViewsRequestedAddHandler),
            New Action(AddressOf _ViewResizedAddHandler),
            New Action(AddressOf _OpenViewRequestedAddHandler),
            New Action(AddressOf _CloseViewRequestedAddHandler),
            New Action(AddressOf _ResetViewRequestedAddHandler)
        )
        Call BaseInitialize(adm, vm)
    End Sub

    ' ＯｐｅｎＶｉｅｗ
    '---------------------------------------------------------------------------------------------'
    Private Sub _OpenViewRequestedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.OpenViewRequested,
            AddressOf Me._OpenViewRequestedReview
    End Sub

    Private Sub _OpenViewRequestedReview(ByVal vim As ViewItemModel, ByVal e As System.EventArgs)
        Call _OpenViewRequestedAccept(vim)
    End Sub

    Private Sub _OpenViewRequestedAccept(ByRef src As ViewItemModel)
        Dim vim As ViewItemModel

        If ViewModel.Content.MenuViewContent IsNot Nothing Then
            vim = _OpenViewDataSet(src, ViewModel.Content.MenuViewContent)
            If vim IsNot Nothing Then
                ViewModel.Content.MenuViewContent = vim
            End If
        End If
        If ViewModel.Content.MainViewContent IsNot Nothing Then
            vim = _OpenViewDataSet(src, ViewModel.Content.MainViewContent)
            If vim IsNot Nothing Then
                ViewModel.Content.MainViewContent = vim
            End If
        End If
        If ViewModel.Content.LeftExplorerViewContent IsNot Nothing Then
            vim = _OpenViewDataSet(src, ViewModel.Content.LeftExplorerViewContent)
            If vim IsNot Nothing Then
                ViewModel.Content.LeftExplorerViewContent = vim
            End If
        End If
        If ViewModel.Content.RightExplorerViewContent IsNot Nothing Then
            vim = _OpenViewDataSet(src, ViewModel.Content.RightExplorerViewContent)
            If vim IsNot Nothing Then
                ViewModel.Content.RightExplorerViewContent = vim
            End If
        End If
        If ViewModel.Content.HistoryViewContent IsNot Nothing Then
            vim = _OpenViewDataSet(src, ViewModel.Content.HistoryViewContent)
            If vim IsNot Nothing Then
                ViewModel.Content.HistoryViewContent = vim
            End If
        End If

        Call ViewModel.Content.OpenView(vim)
        ViewModel.ReloadViews()
    End Sub

    Private Function _OpenViewDataSet(ByRef src As ViewItemModel, ByRef dst As ViewItemModel) As ViewItemModel
        _OpenViewDataSet = Nothing
        If dst IsNot Nothing Then
            If dst.Name = src.Name Then
                dst.IsVisible = True
                If dst.Content Is Nothing Then
                    dst.Content = dst.RestoreContent
                End If
                If dst.ModelName = "TabViewModel" Then
                    Call _OpenTabViewDataSet(src, dst)
                End If
            Else
                If dst.ModelName = "TabViewModel" Then
                    If src.Parent IsNot Nothing Then
                        If dst.Name = src.Parent.Name Then
                            Call _OpenTabViewDataSetFromChild(src, dst)
                        End If
                    End If
                End If
            End If
            _OpenViewDataSet = dst
        End If
    End Function

    ' タブ子要素をすべて開く
    Private Sub _OpenTabViewDataSet(ByRef src As ViewItemModel,
                                    ByRef dst As ViewItemModel)
        For Each child In dst.Children
            child.IsVisible = True
            If child.Content Is Nothing Then
                child.Content = child.RestoreContent
            End If
            Call dst.Content.AddTab(child)
        Next
    End Sub

    ' タブ子要素を（から）開く
    Private Sub _OpenTabViewDataSetFromChild(ByRef src As ViewItemModel,
                                             ByRef dst As ViewItemModel)
        Dim tvm As TabViewModel
        dst.IsVisible = True
        If dst.Content Is Nothing Then
            dst.Content = dst.RestoreContent
        End If
        For Each child In dst.Children
            If src.Name = child.Name Then
                child.IsVisible = True
                If child.Content Is Nothing Then
                    child.Content = child.RestoreContent
                End If
                Call dst.Content.AddTab(child)
                Exit For
            End If
        Next
    End Sub
    '---------------------------------------------------------------------------------------------'

    ' ビューを閉じる
    '---------------------------------------------------------------------------------------------'
    Private Sub _CloseViewRequestedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.CloseViewRequested,
            AddressOf Me._CloseViewRequestedReview
    End Sub

    Private Sub _CloseViewRequestedReview(ByVal vim As ViewItemModel, ByVal e As System.EventArgs)
        Call _CloseViewRequestedAccept(vim)
    End Sub

    Private Overloads Sub _CloseViewRequestedAccept(ByVal vim As ViewItemModel)
        ViewModel.Content.CloseView(vim)
        ViewModel.ReloadViews()
    End Sub
    '---------------------------------------------------------------------------------------------'

    ' ビューリセット
    '---------------------------------------------------------------------------------------------'
    Private Sub _ResetViewRequestedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.ResetViewRequested,
            AddressOf Me._ResetViewRequestedReview
    End Sub

    Private Sub _ResetViewRequestedReview(ByVal fvm As FrameViewModel, ByVal e As System.EventArgs)
        If fvm Is Nothing Then
            Call _ResetViewRequestedAccept()
        Else
            Call _ResetViewRequestedAccept(fvm)
        End If
    End Sub

    Private Overloads Sub _ResetViewRequestedAccept()
        Call ViewModel.ResetView()
    End Sub

    Private Overloads Sub _ResetViewRequestedAccept(ByVal fvm As FrameViewModel)
        ViewModel.Content = fvm
    End Sub
    '---------------------------------------------------------------------------------------------'

    ' FrameView最適化
    '---------------------------------------------------------------------------------------------'
    'Private Sub _FrameOptimizeRequestedAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.FrameOptimizeRequested,
    '        AddressOf Me._FrameOptimizeRequestedReview
    'End Sub

    'Private Sub _FrameOptimizeRequestedReview(ByVal order As String, ByVal e As System.EventArgs)
    '    Call _FrameOptimizeRequestedAccept()
    'End Sub

    'Private Overloads Sub _FrameOptimizeRequestedAccept()
    '    ViewModel.Content.OptimizeView()
    'End Sub
    '---------------------------------------------------------------------------------------------'


    '--- ビューサイズ変更 -------------------------------------------------------------------------'
    Private Sub _ViewResizedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.ViewResized,
            AddressOf Me._ViewResizedReview
    End Sub

    Private Sub _ViewResizedReview(ByVal sender As Object, ByVal e As SizeChangedEventArgs)
        Call _ViewResizedAccept(sender)
    End Sub

    Private Sub _ViewResizedAccept(ByRef sender As Object)
        Select Case sender.Name
            Case "MenuContent"
                ViewModel.Content.MenuViewHeight = sender.ActualHeight
            Case "LeftExplorerContent"
                ViewModel.Content.LeftExplorerViewWidth = sender.ActualWidth
            Case "MainContent"
                ViewModel.Content.MainViewHeight = sender.ActualHeight
                ViewModel.Content.MainViewWidth = sender.ActualWidth
            Case "RightExplorerContent"
                ViewModel.Content.RightExplorerViewWidth = sender.ActualWidth
            Case "HistoryContent"
                ViewModel.Content.HistoryViewHeight = sender.ActualHeight
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------'

    '--- タブを開く関連 --------------------------------------------------------------------------'
    Private Sub _ReloadViewsRequestedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.ReloadViewsRequested,
            AddressOf Me._ReloadViewsRequestedReview
    End Sub

    Private Sub _ReloadViewsRequestedReview(ByVal sender As Object, ByVal e As System.EventArgs)
        Call _ReloadViewsRequestedAccept()
    End Sub

    Private Sub _ReloadViewsRequestedAccept()
        Call ViewModel.ReloadViews()
    End Sub
    '---------------------------------------------------------------------------------------------'

    '--- タブを閉じる関連 --------------------------------------------------------------------------'
    Private Sub _TabViewClosedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.TabViewClosed,
            AddressOf Me._TabViewClosedReview
    End Sub

    Private Sub _TabViewClosedReview(ByVal sender As Object, ByVal e As System.EventArgs)
        'Call _TabViewClosedAccept(ViewModel.Content)
        'Call DelegateEventListener.Instance.RaiseFlexViewOptimizationRequested()
    End Sub

    Private Sub _TabViewClosedAccept(ByRef fvm As FlexibleViewModel)
        'Dim fnc As Func(Of Object, String, Boolean)
        'fnc = Function(ByVal obj As Object, [name] As String)
        '          Dim b = False
        '          Dim fvm2 As FlexibleViewModel
        '          Select Case [name]
        '              Case "FlexibleViewModel"
        '                  fvm2 = CType(obj, FlexibleViewModel)
        '                  Call _TabViewClosedAccept(fvm2)
        '              Case "TabViewModel"
        '                  If obj.Tabs.Count = 0 Then
        '                      b = True
        '                  End If
        '              Case Else
        '          End Select
        '          Return b
        '      End Function

        'If fvm.MainContent IsNot Nothing Then
        '    If fnc(fvm.MainContent, fvm.MainViewContent.ModelName) Then
        '        fvm.MainContent = Nothing
        '    End If
        'End If
        'If fvm.RightContent IsNot Nothing Then
        '    If fnc(fvm.RightContent, fvm.RightViewContent.ModelName) Then
        '        fvm.RightContent = Nothing
        '    End If
        'End If
        'If fvm.BottomContent IsNot Nothing Then
        '    If fnc(fvm.BottomContent, fvm.BottomViewContent.ModelName) Then
        '        fvm.BottomContent = Nothing
        '    End If
        'End If
    End Sub
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


    ' ビューエクスプローラーからタブ開く
    ' 正直ロジック厳しすぎる
    ''---------------------------------------------------------------------------------------------'
    'Private Sub _OpenViewAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.OpenViewRequested,
    '        AddressOf Me._OpenViewRequestedReview
    'End Sub

    'Private Sub _OpenViewRequestedReview(ByVal v As ViewItemModel, ByVal e As System.EventArgs)
    '    'Dim vims As New List(Of ViewItemModel)
    '    'vims = _CheckViewItemList(vims, v, ViewModel.Content)
    '    'Call _OpenViewRequestAccept(ViewModel.Content, vims)
    'End Sub

    'Private Function _PopData(Of T As {New, IList})(ByVal [old] As T) As T
    '    Dim [new] As New T
    '    For Each o In [old]
    '        If Not [old].IndexOf(o) = 0 Then
    '            [new].Add(o)
    '        End If
    '    Next
    '    _PopData = [new]
    'End Function

    'Private Overloads Sub _OpenViewSet(ByRef obj As Object, ByVal vim As ViewItemModel, ByVal vims As List(Of ViewItemModel))
    '    'Select Case vim.ModelName
    '    '    Case "FlexibleViewModel"
    '    '        Call _OpenFlexViewSet(obj, vim, vims)
    '    '    Case "TabViewModel"
    '    '        Call _OpenTabViewSet(obj, vim, vims)
    '    '    Case Else
    '    '        obj = vim.Content
    '    'End Select
    'End Sub

    'Private Overloads Sub _OpenFlexViewSet(ByRef obj As Object, ByVal vim As ViewItemModel, ByVal vims As List(Of ViewItemModel))
    '    'Dim fvm As FlexibleViewModel
    '    'obj = New FlexibleViewModel With {
    '    '    .Name = vim.Content.Name,
    '    '    .ContentViewWidth = IIf(vim.Content.ContentViewWidth > 0.0, vim.Content.ContentViewWidth, 0.0),
    '    '    .ContentViewHeight = IIf(vim.Content.ContentViewHeight > 0.0, vim.Content.ContentViewHeight, 0.0)
    '    '}
    '    'fvm = CType(obj, FlexibleViewModel)
    '    'Call _OpenViewRequestAccept(fvm, vims)
    'End Sub

    'Private Overloads Sub _OpenTabViewSet(ByRef obj As Object, ByVal vim As ViewItemModel, ByVal vims As List(Of ViewItemModel))
    '    'Dim tvm As TabViewModel
    '    'obj = New TabViewModel
    '    'tvm = CType(obj, TabViewModel)
    '    'Call _OpenViewRequestAccept(tvm, vim, vims)
    'End Sub

    'Private Overloads Sub _OpenViewRequestAccept(ByRef fvm As FlexibleViewModel,
    '                                             ByVal vims As List(Of ViewItemModel))
    '    'If fvm.MainViewContent IsNot Nothing Then
    '    '    If fvm.MainViewContent.Name = vims(0).Name Then
    '    '        If fvm.MainContent Is Nothing Then
    '    '            Call _OpenViewSet(fvm.MainContent, fvm.MainViewContent, vims)
    '    '        Else
    '    '            vims = IIf(fvm.MainViewContent.ModelName = "TabViewModel", vims, _PopData(vims))
    '    '            Call _OpenViewRequestAccept(fvm.MainContent, fvm.MainViewContent, vims)
    '    '        End If
    '    '    End If
    '    'End If
    '    'If fvm.RightViewContent IsNot Nothing Then
    '    '    If fvm.RightViewContent.Name = vims(0).Name Then
    '    '        If fvm.RightContent Is Nothing Then
    '    '            Call _OpenViewSet(fvm.RightContent, fvm.RightViewContent, vims)
    '    '        Else
    '    '            vims = IIf(fvm.RightViewContent.ModelName = "TabViewModel", vims, _PopData(vims))
    '    '            Call _OpenViewRequestAccept(fvm.RightContent, fvm.RightViewContent, vims)
    '    '        End If
    '    '    End If
    '    'End If
    '    'If fvm.BottomViewContent IsNot Nothing Then
    '    '    If fvm.BottomViewContent.Name = vims(0).Name Then
    '    '        If fvm.BottomContent Is Nothing Then
    '    '            Call _OpenViewSet(fvm.BottomContent, fvm.BottomViewContent, vims)
    '    '        Else
    '    '            vims = IIf(fvm.BottomViewContent.ModelName = "TabViewModel", vims, _PopData(vims))
    '    '            Call _OpenViewRequestAccept(fvm.BottomContent, fvm.BottomViewContent, vims)
    '    '        End If
    '    '    End If
    '    'End If
    'End Sub

    'Private Overloads Sub _OpenViewRequestAccept(ByRef tvm As TabViewModel,
    '                                             ByVal vim As ViewItemModel,
    '                                             ByVal vims As List(Of ViewItemModel))
    '    ' a. TabViewModelが存在、TabItemModelがない場合
    '    ' b. TabViewModelが存在しない
    '    ' この２パターンの違いで内部的な動作が異なる
    '    '
    '    ' a.のケースの場合、
    '    ' 存在するTabViewModelで、
    '    ' _OpenViewRequestAccept(TabViewModel, ViewItemModel, List(Of ViewItemModel))
    '    ' を実行する
    '    '
    '    ' b.のケースの場合、
    '    ' _OpenViewSetで新たにTabViewModelを新規作成し、そのTabViewModelで、
    '    ' _OpenViewRequestAccept(TabViewModel, ViewItemModel, List(Of ViewItemModel))
    '    ' を実行する
    '    'If vim.Name = vims(0).Name Then
    '    '    vims = _PopData(vims)
    '    '    For Each vt In vim.Content.ViewContentTabs
    '    '        If vt.Name = vims(0).Name Then
    '    '            tvm.AddTab(vt)
    '    '            Call _OpenViewRequestAccept(tvm.Tabs.Last, vt, vims)
    '    '            Exit For
    '    '        End If
    '    '    Next
    '    'End If
    'End Sub


    'Private Overloads Sub _OpenViewRequestAccept(ByRef obj As Object,
    '                                             ByVal vim As ViewItemModel,
    '                                             ByVal [old] As List(Of ViewItemModel))
    '    'Dim fvm As FlexibleViewModel
    '    'Dim tvm As TabViewModel
    '    ''Dim [new] = New List(Of ViewItemModel)

    '    ''[new] = _PopData([old])

    '    'Select Case vim.ModelName
    '    '    Case "FlexibleViewModel"
    '    '        fvm = CType(obj, FlexibleViewModel)
    '    '        Call _OpenViewRequestAccept(fvm, [old])
    '    '    Case "TabViewModel"
    '    '        tvm = CType(obj, TabViewModel)
    '    '        Call _OpenViewRequestAccept(tvm, vim, [old])
    '    '    Case Else
    '    '        '' UnExpected Case
    '    '        'Throw New Exception("ViewController._OpenViewRequestAccept Error!!!")
    '    'End Select
    'End Sub


    'Private Overloads Function _CheckViewItemList(ByVal [new] As List(Of ViewItemModel),
    '                                              ByVal v As ViewItemModel,
    '                                              ByVal fvm As FlexibleViewModel) As List(Of ViewItemModel)
    '    'If fvm.MainViewContent IsNot Nothing Then
    '    '    [new] = _CheckViewItemList([new], v, fvm.MainViewContent)
    '    'End If
    '    'If fvm.RightViewContent IsNot Nothing Then
    '    '    [new] = _CheckViewItemList([new], v, fvm.RightViewContent)
    '    'End If
    '    'If fvm.BottomViewContent IsNot Nothing Then
    '    '    [new] = _CheckViewItemList([new], v, fvm.BottomViewContent)
    '    'End If
    '    '_CheckViewItemList = [new]
    'End Function

    'Private Overloads Function _CheckViewItemList(ByVal [new] As List(Of ViewItemModel),
    '                                              ByVal v As ViewItemModel,
    '                                              ByVal vim As ViewItemModel) As List(Of ViewItemModel)

    '    'Dim fvm As FlexibleViewModel
    '    'Dim tvm As TabViewModel

    '    'Dim [old] As New List(Of ViewItemModel)
    '    'For Each n In [new]
    '    '    [old].Add(n)
    '    'Next
    '    '[new].Add(vim)

    '    'Dim [old2] As New List(Of ViewItemModel)
    '    'For Each n In [new]
    '    '    [old2].Add(n)
    '    'Next

    '    'Select Case vim.ModelName
    '    '    Case "FlexibleViewModel"
    '    '        fvm = CType(vim.Content, FlexibleViewModel)
    '    '        [new] = _CheckViewItemList([new], v, fvm)
    '    '    Case "TabViewModel"
    '    '        tvm = CType(vim.Content, TabViewModel)
    '    '        [new] = _CheckViewItemList([new], v, tvm)
    '    '        If [new].Count = [old2].Count Then
    '    '            [new] = [old]
    '    '        End If
    '    '    Case v.ModelName
    '    '        If Not v.Name = vim.Name Then
    '    '            [new] = [old]
    '    '        End If
    '    '    Case Else
    '    '        [new] = [old]
    '    'End Select
    '    '_CheckViewItemList = [new]
    'End Function

    'Private Overloads Function _CheckViewItemList(ByVal [new] As List(Of ViewItemModel),
    '                                              ByVal v As ViewItemModel,
    '                                              ByVal tvm As TabViewModel) As List(Of ViewItemModel)
    '    'For Each vt In tvm.ViewContentTabs
    '    '    [new] = _CheckViewItemList([new], v, vt)
    '    'Next
    '    '_CheckViewItemList = [new]
    'End Function

    ''Private Overloads Sub _OpenViewRequestAccept(ByVal vim As ViewItemModel, ByRef fvm As FlexibleViewModel)
    ''    Dim fnc As Func(Of ViewItemModel, ViewItemModel, Boolean)
    ''    fnc = Function(ByVal vim2 As ViewItemModel, vim3 As ViewItemModel) As Boolean
    ''              Dim b = False
    ''              Dim fvm2 As FlexibleViewModel
    ''              Dim tvm As TabViewModel        'Call _OpenViewRequestAccept(obj, vim, vims)

    ''              Select Case vim3.ModelName
    ''                  Case "FlexibleViewModel"
    ''                      fvm2 = CType(vim3.Content, FlexibleViewModel)
    ''                      Call _OpenViewRequestAccept(vim2, fvm2)
    ''                  Case "TabViewModel"
    ''                      tvm = CType(vim3.Content, TabViewModel)
    ''                      Call _OpenTabViewRequestAccept(vim2, tvm)
    ''                  Case vim2.ModelName
    ''                      If vim2.Name = vim3.Name Then
    ''                          b = True
    ''                      End If
    ''              End Select
    ''              Return b
    ''          End Function

    ''    If fvm.MainViewContent IsNot Nothing Then
    ''        If fnc(vim, fvm.MainViewContent) Then
    ''            fvm.MainViewContent.IsVisible = True
    ''            fvm.MainContent = fvm.MainViewContent.Content
    ''        End If
    ''    End If
    ''    If fvm.RightViewContent IsNot Nothing Then
    ''        If fnc(vim, fvm.RightViewContent) Then
    ''            fvm.RightViewContent.IsVisible = True
    ''            fvm.RightContent = fvm.RightViewContent.Content
    ''        End If
    ''    End If
    ''    If fvm.BottomViewContent IsNot Nothing Then
    ''        If fnc(vim, fvm.BottomViewContent) Then
    ''            fvm.BottomViewContent.IsVisible = True
    ''            fvm.BottomContent = fvm.BottomViewContent.Content
    ''        End If
    ''    End If
    ''End Sub

    ''Private Overloads Sub _OpenTabViewRequestAccept(ByVal v As ViewItemModel, ByRef tvm As TabViewModel)
    ''    Dim fvm As FlexibleViewModel
    ''    Dim tvm2 As TabViewModel
    ''    Dim obj As Object
    ''    For Each vt In tvm.ViewContentTabs
    ''        Select Case vt.ModelName
    ''            Case "FlexibleViewModel"
    ''                fvm = CType(vt.Content, FlexibleViewModel)
    ''                Call _OpenViewRequestAccept(v, fvm)
    ''            Case "TabViewModel"
    ''                tvm2 = CType(vt.Content, TabViewModel)
    ''                Call _OpenTabViewRequestAccept(v, tvm2)
    ''            Case v.ModelName
    ''                If v.Name = vt.Name Then
    ''                    If vt.Content Is Nothing Then
    ''                        Throw New Exception("ViewController._OpenTabViewRequestAccept Error!")
    ''                    Else
    ''                        vt.IsVisible = True
    ''                        tvm.AddTab(vt)
    ''                    End If
    ''                    Exit For
    ''                End If
    ''        End Select
    ''    Next
    ''End Sub
    ''---------------------------------------------------------------------------------------------'
End Class

