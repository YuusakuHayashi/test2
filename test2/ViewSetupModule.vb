Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Module ViewSetupModule
    Public Sub DBTESTViewSetupExecute(ByRef app As AppDirectoryModel,
                                      ByRef vm As ViewModel)
        Dim v1, v2
        Dim cvm = New ConnectionViewModel
        Dim dbtvm = New DBTestViewModel
        Dim dbevm = New DBExplorerViewModel
        Dim vevm = New ViewExplorerViewModel
        'Dim hvm = New HistoryViewModel

        Call cvm.Initialize(app, vm)
        'Call hvm.Initialize(app, vm)

        ' ビューへの追加
        v1 = cvm.AddViewItem(cvm,
                             ViewModel.MULTI_VIEW,
                             MultiViewModel.MAIN_FRAME,
                             MultiViewModel.TAB_VIEW)
        'v2 = hvm.AddViewItem(hvm,
        '                     ViewModel.MULTI_VIEW,
        '                     MultiViewModel.HISTORY_FRAME,
        '                     MultiViewModel.TAB_VIEW)
        Call cvm.AddView(v1)
        'Call hvm.AddView(v2)
    End Sub

    'Public Function DBTESTViewDefineExecute(ByVal [name] As String) As Object
    Public Function DBTESTViewDefineExecute(ByVal v As ViewItemModel) As Object
        Dim obj As Object
        Select Case v.Name
            Case (New ConnectionViewModel).GetType.Name
                obj = New ConnectionViewModel
            Case (New DBTestViewModel).GetType.Name
                obj = New DBTestViewModel
            Case (New DBExplorerViewModel).GetType.Name
                obj = New DBExplorerViewModel
            Case (New ViewExplorerViewModel).GetType.Name
                obj = New ViewExplorerViewModel
            Case Else
                obj = Nothing
        End Select
        DBTESTViewDefineExecute = obj
    End Function



    Public Sub RpaProjectViewSetupExecute(ByRef app As AppDirectoryModel,
                                          ByRef vm As ViewModel)
        Dim rpapvm = New RpaProjectViewModel
        Dim rpapmvm = New RpaProjectMenuViewModel
        Dim v1, v2
        Call rpapvm.Initialize(app, vm)
        Call rpapmvm.Initialize(app, vm)
        v1 = rpapvm.AddViewItem(rpapvm,
                                ViewModel.MULTI_VIEW,
                                MultiViewModel.MAIN_FRAME,
                                MultiViewModel.TAB_VIEW)
        v2 = rpapvm.AddViewItem(rpapmvm,
                                ViewModel.MULTI_VIEW,
                                MultiViewModel.PROJECT_MENU_FRAME,
                                MultiViewModel.NORMAL_VIEW)
        Call rpapvm.AddView(v1)
        Call rpapmvm.AddView(v2)
    End Sub

    Public Sub RpaProjectViewLoadExecute(ByRef app As AppDirectoryModel,
                                         ByRef vm As ViewModel)
        Dim obj As Object
        For Each v In vm.Views
            If v.OpenState Then
                obj = RpaProjectViewDefineExecute(v)
                obj.Initialize(app, vm)
                v.Content = obj
                Call obj.AddView(v)
            End If
        Next
    End Sub

    Public Function RpaProjectViewDefineExecute(ByVal v As ViewItemModel) As Object
        Dim obj As Object
        Dim ck1, ck2, ck3

        ck1 = TryCast(
            _CheckLoadedModel(Of RpaProjectMenuViewModel)(v.Content),
            RpaProjectMenuViewModel
        )
        ck2 = TryCast(
            _CheckLoadedModel(Of RpaProjectViewModel)(v.Content),
            RpaProjectViewModel
        )
        ck3 = TryCast(
            _CheckLoadedModel(Of RpaViewModel)(v.Content),
            RpaViewModel
        )

        If ck1 IsNot Nothing Then
            obj = New RpaProjectMenuViewModel
        End If
        If ck2 IsNot Nothing Then
            obj = New RpaProjectViewModel
        End If
        If ck3 IsNot Nothing Then
            obj = New RpaViewModel
        End If

        RpaProjectViewDefineExecute = obj
    End Function

    Private Function _CheckLoadedModel(Of T As {New})(ByVal old As Object) As Object
        Dim [new] As Object

        Select Case old.GetType
            Case (New Object).GetType
                [new] = CType(old, T)
            Case (New JObject).GetType
                ' Ｊｓｏｎからロードした場合は、JObject型になっている
                [new] = old.ToObject(Of T)
            Case (New T).GetType
                [new] = old
        End Select

        _CheckLoadedModel = [new]
    End Function

End Module
