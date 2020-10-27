Public Module ViewSetupModule
    Public Sub DBTESTViewSetupExecute(ByRef app As AppDirectoryModel,
                                      ByRef vm As ViewModel)
        Dim v1, v2, v3
        Dim cvm = New ConnectionViewModel
        Dim dbtvm = New DBTestViewModel
        Dim dbevm = New DBExplorerViewModel
        Dim vevm = New ViewExplorerViewModel
        Dim hvm = New HistoryViewModel
        Dim mvm = New MenuViewModel

        Call cvm.Initialize(app, vm)
        Call hvm.Initialize(app, vm)
        Call mvm.Initialize(app, vm)

        ' ビューへの追加
        v1 = cvm.AddViewItem(cvm,
                             ViewModel.MULTI_VIEW,
                             MultiViewModel.MAIN_FRAME,
                             MultiViewModel.TAB_VIEW)
        v2 = hvm.AddViewItem(hvm,
                             ViewModel.MULTI_VIEW,
                             MultiViewModel.HISTORY_FRAME,
                             MultiViewModel.TAB_VIEW)
        v3 = mvm.AddViewItem(mvm,
                             ViewModel.MULTI_VIEW,
                             MultiViewModel.MENU_FRAME,
                             MultiViewModel.NORMAL_VIEW)
        Call cvm.AddView(v1)
        Call hvm.AddView(v2)
        Call mvm.AddView(v3)
    End Sub

    Public Function DBTESTViewDefineExecute(ByVal [name] As String) As Object
        Dim obj As Object
        Select Case [name]
            Case (New ConnectionViewModel).GetType.Name
                obj = New ConnectionViewModel
            Case (New DBTestViewModel).GetType.Name
                obj = New DBTestViewModel
            Case (New DBExplorerViewModel).GetType.Name
                obj = New DBExplorerViewModel
            Case (New ViewExplorerViewModel).GetType.Name
                obj = New ViewExplorerViewModel
            Case (New HistoryViewModel).GetType.Name
                obj = New HistoryViewModel
            Case (New MenuViewModel).GetType.Name
                obj = New MenuViewModel
            Case Else
                obj = Nothing
        End Select
        DBTESTViewDefineExecute = obj
    End Function



    Public Sub RpaProjectViewSetupExecute(ByRef app As AppDirectoryModel,
                                          ByRef vm As ViewModel)
        Dim rpapvm = New RpaProjectViewModel
        Dim v As ViewItemModel
        Call rpapvm.Initialize(app, vm)
        v = rpapvm.AddViewItem(rpapvm,
                               ViewModel.MULTI_VIEW,
                               MultiViewModel.MAIN_FRAME,
                               MultiViewModel.TAB_VIEW)
        Call rpapvm.AddView(v)
    End Sub

    Public Function RpaProjectViewDefineExecute(ByVal [name] As String) As Object
        Dim obj As Object
        Select Case [name]
            Case (New RpaProjectViewModel).GetType.Name
                obj = New RpaProjectViewModel
            Case (New RpaViewModel).GetType.Name
                obj = New RpaViewModel
            Case Else
                obj = Nothing
        End Select
        RpaProjectViewDefineExecute = obj
    End Function
End Module
