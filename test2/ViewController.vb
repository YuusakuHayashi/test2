Public Class ViewController : Inherits BaseViewModel2

    Public Overrides Sub Initialize(ByRef adm As AppDirectoryModel,
                                    ByRef vm As ViewModel)
        Me.[AddHandler] = [Delegate].Combine(
            New Action(AddressOf _OpenViewAddHandler),
            New Action(AddressOf _MultiViewSizeChangedAddHandler),
            New Action(AddressOf _OpenProjectAddHandler),
            New Action(AddressOf _TabViewClosedAddHandler)
        )
        Call BaseInitialize(adm, vm)
    End Sub

    '--- ＳＰＬＩＴＴＥＲ変更 --------------------------------------------------------------------'
    'Private Sub _ViewResizedReview(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call _ViewResizedAccept(ViewModel.Content)
    'End Sub

    'Private Sub _ViewResizedAccept(ByRef fvm As FlexibleViewModel)
    '    Dim fnc As Func(Of Object, String, Boolean)
    '    fnc = Function(ByRef obj As Object, ByVal [name] As String)
    '              Dim b = False
    '              Dim fvm2 As FlexibleViewModel
    '              Dim tvm As TabViewModel
    '              Select Case [name]
    '                  Case "FlexibleViewModel"
    '                      fvm2 = CType(obj, FlexibleViewModel)
    '                      Call _ViewResizedAccept(fvm2)
    '                  Case "TabViewModel"
    '                      tvm = CType(obj, TabViewModel) 
    '                      Call _TabViewResizedAccept(tvm)
    '                  Case Else
    '              End Select
    '          End Function
    '   
    '    fvm.ContentViewHeight = fvm.ContentGridHeight.Value
    '    fvm.ContentViewHeight = fvm.ContentGridHeight.Value

    '    If fvm.MainContent IsNot Nothing Then
    '        If fnc(fvm.MainContent, fvm.MainViewContent.ModelName) Then
    '            fvm.MainContent = Nothing
    '        End If
    '    End If
    '    If fvm.RightContent IsNot Nothing Then
    '        If fnc(fvm.RightContent, fvm.RightViewContent.ModelName) Then
    '            fvm.RightContent = Nothing
    '        End If
    '    End If
    '    If fvm.BottomContent IsNot Nothing Then
    '        If fnc(fvm.BottomContent, fvm.BottomViewContent.ModelName) Then
    '            fvm.BottomContent = Nothing
    '        End If
    '    End If
    'End Sub

    'Private Sub _ViewResizedAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.ViewResized,
    '        AddressOf Me._ViewResizedReview
    'End Sub
    '---------------------------------------------------------------------------------------------'

    '--- タブを開く関連 --------------------------------------------------------------------------'
    Private Sub _TabViewClosedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.TabViewClosed,
            AddressOf Me._TabViewClosedReview
    End Sub

    Private Sub _TabViewClosedReview(ByVal sender As Object, ByVal e As System.EventArgs)
        Call _TabViewClosedAccept(ViewModel.Content)
    End Sub

    Private Sub _TabViewClosedAccept(ByRef fvm As FlexibleViewModel)
        Dim fnc As Func(Of Object, String, Boolean)
        fnc = Function(ByVal obj As Object, [name] As String)
                  Dim b = False
                  Dim fvm2 As FlexibleViewModel
                  Select Case [name]
                      Case "FlexibleViewModel"
                          fvm2 = CType(obj, FlexibleViewModel)
                          Call _TabViewClosedAccept(fvm2)
                      Case "TabViewModel"
                          If obj.Tabs.Count = 0 Then
                              b = True
                          End If
                      Case Else
                  End Select
                  Return b
              End Function

        If fvm.MainContent IsNot Nothing Then
            If fnc(fvm.MainContent, fvm.MainViewContent.ModelName) Then
                fvm.MainContent = Nothing
            End If
        End If
        If fvm.RightContent IsNot Nothing Then
            If fnc(fvm.RightContent, fvm.RightViewContent.ModelName) Then
                fvm.RightContent = Nothing
            End If
        End If
        If fvm.BottomContent IsNot Nothing Then
            If fnc(fvm.BottomContent, fvm.BottomViewContent.ModelName) Then
                fvm.BottomContent = Nothing
            End If
        End If
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

    '--- ＳＰＬＩＴＴＥＲ変更 --------------------------------------------------------------------'
    Private Sub _MultiViewSizeChangedReview(ByVal sender As Object, ByVal e As System.EventArgs)
        Call _MultiViewSizeChangeAccept(sender)
    End Sub

    Private Sub _MultiViewSizeChangeAccept(ByVal sender As Object)
        'Select Case sender.Name
        '    Case "MainView"
        '        ViewModel.MultiView.MainGridHeight = New GridLength(sender.ActualHeight)
        '        ViewModel.MultiView.RightGridWidth = New GridLength(sender.ActualWidth)
        '    Case "ExplorerView"
        '        ViewModel.MultiView.ExplorerGridHeight = New GridLength(sender.ActualHeight)
        '        ViewModel.MultiView.LeftGridWidth = New GridLength(sender.ActualWidth)
        '    Case "HistoryView"
        '        ViewModel.MultiView.HistoryGridHeight = New GridLength(sender.ActualHeight)
        '        ViewModel.MultiView.RightGridWidth = New GridLength(sender.ActualWidth)
        '    Case Else
        '        Throw New Exception("Error View Name")
        'End Select
    End Sub

    Private Sub _MultiViewSizeChangedAddHandler()
        'AddHandler _
        '    DelegateEventListener.Instance.MultiViewSizeChanged,
        '    AddressOf Me._MultiViewSizeChangedReview
    End Sub
    '---------------------------------------------------------------------------------------------'


    ' ビューエクスプローラーからタブ開く
    ' 正直ロジック厳しすぎる
    '---------------------------------------------------------------------------------------------'
    Private Sub _OpenViewAddHandler()
        AddHandler _
            DelegateEventListener.Instance.OpenViewRequested,
            AddressOf Me._OpenViewRequestedReview
    End Sub

    Private Sub _OpenViewRequestedReview(ByVal v As ViewItemModel, ByVal e As System.EventArgs)
        Call _OpenViewRequestAccept(v, ViewModel.Content)
    End Sub

    Private Overloads Sub _OpenViewRequestAccept(ByVal vim As ViewItemModel, ByRef fvm As FlexibleViewModel)
        Dim fnc As Func(Of ViewItemModel, ViewItemModel, Boolean)
        fnc = Function(ByVal vim2 As ViewItemModel, vim3 As ViewItemModel) As Boolean
                  Dim b = False
                  Dim fvm2 As FlexibleViewModel
                  Dim tvm As TabViewModel
                  Select Case vim3.ModelName
                      Case "FlexibleViewModel"
                          fvm2 = CType(vim3.Content, FlexibleViewModel)
                          Call _OpenViewRequestAccept(vim2, fvm2)
                      Case "TabViewModel"
                          tvm = CType(vim3.Content, TabViewModel)
                          Call _OpenTabViewRequestAccept(vim2, tvm)
                      Case vim2.ModelName
                          If vim2.Name = vim3.Name Then
                              b = True
                          End If
                  End Select
                  Return b
              End Function

        If fvm.MainViewContent IsNot Nothing Then
            If fnc(vim, fvm.MainViewContent) Then
                fvm.MainContent = fvm.MainViewContent.Content
            End If
        End If
        If fvm.RightViewContent IsNot Nothing Then
            If fnc(vim, fvm.RightViewContent) Then
                fvm.RightContent = fvm.RightViewContent.Content
            End If
        End If
        If fvm.BottomViewContent IsNot Nothing Then
            If fnc(vim, fvm.BottomViewContent) Then
                fvm.BottomContent = fvm.BottomViewContent.Content
            End If
        End If
    End Sub

    Private Overloads Sub _OpenTabViewRequestAccept(ByVal v As ViewItemModel, ByRef tvm As TabViewModel)
        Dim fvm As FlexibleViewModel
        Dim tvm2 As TabViewModel
        Dim obj As Object
        Dim [define] As Func(Of String, Object) _
            = AddressOf AppInfo.ProjectInfo.Model.Data.ViewDefineExecute
        For Each vt In tvm.ViewContentTabs
            Select Case vt.ModelName
                Case "FlexibleViewModel"
                    fvm = CType(vt.Content, FlexibleViewModel)
                    Call _OpenViewRequestAccept(v, fvm)
                Case "TabViewModel"
                    tvm2 = CType(vt.Content, TabViewModel)
                    Call _OpenTabViewRequestAccept(v, tvm2)
                Case v.ModelName
                    If v.Name = vt.Name Then
                        If vt.Content Is Nothing Then
                            Throw New Exception("ViewController._OpenTabViewRequestAccept Error!")
                        Else
                            tvm.AddTab(vt)
                        End If
                        Exit For
                    End If
            End Select
        Next
    End Sub
    '---------------------------------------------------------------------------------------------'
End Class

