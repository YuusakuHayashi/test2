'Imports Newtonsoft.Json
'Imports Newtonsoft.Json.Linq

'Public Module ViewSetupModule
'    Public Sub DBTESTViewSetupExecute(ByRef app As AppDirectoryModel,
'                                      ByRef vm As ViewModel)
'        Dim v1, v2
'        Dim cvm = New ConnectionViewModel
'        Dim dbtvm = New DBTestViewModel
'        Dim dbevm = New DBExplorerViewModel
'        Dim vevm = New ViewExplorerViewModel

'        Call cvm.Initialize(app, vm)

'        ' ビューへの追加
'        v1 = New ViewItemModel With {
'            .Content = cvm,
'            .Name = cvm.GetType.Name,
'            .FrameType = MultiViewModel.MAIN_FRAME,
'            .LayoutType = ViewModel.MULTI_VIEW,
'            .ViewType = MultiViewModel.TAB_VIEW,
'            .OpenState = True
'        }
'        Call cvm.AddView(v1)
'    End Sub

'    'Public Function DBTESTViewDefineExecute(ByVal [name] As String) As Object
'    Public Function DBTESTViewDefineExecute(ByVal v As ViewItemModel) As Object
'        Dim obj As Object
'        Select Case v.Name
'            Case (New ConnectionViewModel).GetType.Name
'                obj = New ConnectionViewModel
'            Case (New DBTestViewModel).GetType.Name
'                obj = New DBTestViewModel
'            Case (New DBExplorerViewModel).GetType.Name
'                obj = New DBExplorerViewModel
'            Case (New ViewExplorerViewModel).GetType.Name
'                obj = New ViewExplorerViewModel
'            Case Else
'                obj = Nothing
'        End Select
'        DBTESTViewDefineExecute = obj
'    End Function



'    Public Sub RpaProjectViewSetupExecute(ByRef app As AppDirectoryModel,
'                                          ByRef vm As ViewModel)
'        Dim rpapvm = New RpaProjectViewModel
'        Dim rpapmvm = New RpaProjectMenuViewModel
'        Dim v1, v2
'        Call rpapvm.Initialize(app, vm)
'        Call rpapmvm.Initialize(app, vm)
'        v1 = New ViewItemModel With {
'            .Content = rpapvm,
'            .Name = rpapvm.GetType.Name,
'            .FrameType = MultiViewModel.MAIN_FRAME,
'            .LayoutType = ViewModel.MULTI_VIEW,
'            .ViewType = MultiViewModel.TAB_VIEW,
'            .OpenState = True,
'            .ModelName = rpapvm.GetType.Name
'        }
'        v2 = New ViewItemModel With {
'            .Content = rpapmvm,
'            .Name = rpapmvm.GetType.Name,
'            .FrameType = MultiViewModel.PROJECT_MENU_FRAME,
'            .LayoutType = ViewModel.MULTI_VIEW,
'            .ViewType = MultiViewModel.NORMAL_VIEW,
'            .OpenState = True,
'            .ModelName = rpapmvm.GetType.Name
'        }
'        Call rpapvm.AddViewItem(v1)
'        Call rpapvm.AddView(v1)
'        Call rpapmvm.AddViewItem(v2)
'        Call rpapmvm.AddView(v2)
'    End Sub

'    Public Function RpaProjectViewDefineExecute(ByVal v As ViewItemModel) As Object
'        Dim obj As Object

'        Select Case v.ModelName
'            Case (New AppDirectoryModel).GetType.Name
'                obj = New AppDirectoryModel
'            Case (New ViewModel).GetType.Name
'                obj = New ViewModel
'            Case (New HistoryViewModel).GetType.Name
'                obj = New HistoryViewModel
'            Case (New MenuViewModel).GetType.Name
'                obj = New MenuViewModel
'            Case (New RpaProjectMenuViewModel).GetType.Name
'                obj = New RpaProjectMenuViewModel
'            Case (New RpaProjectViewModel).GetType.Name
'                obj = New RpaProjectViewModel
'            Case (New RpaViewModel).GetType.Name
'                obj = New RpaViewModel
'            Case Else
'                obj = Nothing
'        End Select

'        RpaProjectViewDefineExecute = obj
'    End Function

'    Private Function _CheckLoadedModel(Of T As {New})(ByVal old As Object) As Object
'        Dim [new] As Object

'        Select Case old.GetType
'            Case (New Object).GetType
'                [new] = CType(old, T)
'            Case (New JObject).GetType
'                ' Ｊｓｏｎからロードした場合は、JObject型になっている
'                [new] = old.ToObject(Of T)
'            Case (New T).GetType
'                [new] = old
'        End Select

'        _CheckLoadedModel = [new]
'    End Function

'End Module
