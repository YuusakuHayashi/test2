Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Collections.ObjectModel

'Public Class BaseViewModel2(Of T As {New})
Public MustInherit Class BaseViewModel2
    Implements INotifyPropertyChanged, BaseViewModelInterface

    '--- INortify -------------------------------------------------------------------------------------'
    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub
    '--------------------------------------------------------------------------------------------------'

    Private _AppInfo As AppDirectoryModel
    <JsonIgnore>
    Public Property AppInfo As AppDirectoryModel
        Get
            Return _AppInfo
        End Get
        Set(value As AppDirectoryModel)
            _AppInfo = value
        End Set
    End Property

    Private _ViewModel As ViewModel
    <JsonIgnore>
    Public Property ViewModel As ViewModel
        Get
            If Me._ViewModel Is Nothing Then
                Me._ViewModel = New ViewModel
            End If
            Return _ViewModel
        End Get
        Set(value As ViewModel)
            _ViewModel = value
        End Set
    End Property

    Protected Sub AppLoad()
        Dim jh As New JsonHandler(Of AppDirectoryModel)
        AppInfo = jh.ModelLoad(AppDirectoryModel.ModelFileName)
    End Sub

    Public Sub AppSave()
        Call AppInfo.ModelSave(AppDirectoryModel.ModelFileName, AppInfo)
    End Sub

    '---------------------------------------------------------------------------------------------'
    ' プロジェクトのロード
    Public Overloads Sub ProjectLoad()
        Call ProjectLoad(AppInfo.ProjectInfo)
    End Sub
    Public Overloads Sub ProjectLoad(ByVal project As ProjectInfoModel)
        Dim jh = New JsonHandler(Of ProjectInfoModel)
        AppInfo.ProjectInfo = jh.ModelLoad(project.ProjectInfoFileName)
    End Sub
    '---------------------------------------------------------------------------------------------'

    '---------------------------------------------------------------------------------------------'
    ' プロジェクトのセーブ
    Public Overloads Sub ProjectSave()
        Call ProjectSave(AppInfo.ProjectInfo)
    End Sub
    Public Overloads Sub ProjectSave(ByVal project As ProjectInfoModel)
        Call project.ModelSave(project.ProjectInfoFileName, project)
    End Sub
    '---------------------------------------------------------------------------------------------'

    Public Overloads Sub ProjectModelSave(ByVal project As ProjectInfoModel)
        Call project.Model.ModelSave(project.ModelFileName, project.Model)
    End Sub
    Public Overloads Sub ProjectModelSave()
        Call AppInfo.ProjectInfo.Model.ModelSave(AppInfo.ProjectInfo.ModelFileName, AppInfo.ProjectInfo.Model)
    End Sub

    Public Overloads Sub ProjectModelLoad()
        Dim jh = New JsonHandler(Of Model)
        AppInfo.ProjectInfo.Model = jh.ModelLoad(AppInfo.ProjectInfo.ModelFileName)
    End Sub

    Public Overloads Sub ProjectViewModelSave(ByVal project As ProjectInfoModel)
        Call ViewModel.ModelSave(project.ViewModelFileName, ViewModel)
    End Sub
    Public Overloads Sub ProjectViewModelSave()
        Call ViewModel.ModelSave(AppInfo.ProjectInfo.ViewModelFileName, ViewModel)
    End Sub

    ' ロード時にはロードしたプロジェクトをカレントに追加しない
    ' (セーブ時に追加させる)
    Public Overloads Sub AllLoad(project)
        'AppInfo.~のルートであるAppInfoをロードして直接更新すると、参照渡ししている各ビューモデルでおかしくなる
        'Call AppLoad()
        Call ProjectLoad(project)
        Call ProjectSetup()
        Call PushProject(AppInfo.ProjectInfo)
        Call ProjectModelLoad()
        Call ModelSetup()
        Call ProjectViewModelLoad()
        Call ViewModelSetup()
    End Sub

    Public Overloads Sub AllSave()
        Call ProjectModelSave()
        Call ProjectViewModelSave()
        Call ProjectSave()
        Call AppSave()
    End Sub

    ' Resave時
    Public Overloads Sub AllSave(ByVal project As ProjectInfoModel)
        Call ProjectModelSave(project)
        Call ProjectViewModelSave(project)
        Call ProjectSave(project)
        Call AppSave()
    End Sub

    Protected Overloads Sub PushProject()
        Call PushProject(AppInfo.ProjectInfo)
    End Sub

    Protected Overloads Sub PushProject(ByVal project As ProjectInfoModel)
        Dim [new] As New ObservableCollection(Of ProjectInfoModel)

        [new].Add(project)
        If project.[Index] = 0 Then
            Call _AssignProjectIndex(project)
        End If

        For Each p In AppInfo.CurrentProjects
            If project.[Index] <> p.[Index] Then
                [new].Add(p)
            End If
            If [new].Count >= 5 Then
                Exit For
            End If
        Next

        AppInfo.CurrentProjects = [new]
    End Sub

    Private Sub _AssignProjectIndex(ByRef project As ProjectInfoModel)
        Dim idx = 1
        Dim b = False

        Do Until True = False
            b = False
            For Each p In AppInfo.CurrentProjects
                If p.[Index] = idx Then
                    idx += 1
                    b = True
                    Exit For
                End If
            Next
            If Not b Then
                Exit Do
            End If
        Loop
        project.[Index] = idx
    End Sub

    ' プロジェクト起動時にステータスを変更
    Protected Overloads Sub ProjectSetup()
        For Each p In AppInfo.CurrentProjects
            p.ActiveStatus = vbNullString
        Next
        AppInfo.ProjectInfo.ActiveStatus = "(Active)"
    End Sub

    ' ＶｉｅｗＭｏｄｅｌはロードするとＭＶＶＭが機能しなくなるので、
    ' 部分的に更新する
    Public Overloads Sub ProjectViewModelLoad()
        Dim jh = New JsonHandler(Of ViewModel)
        Dim vm = jh.ModelLoad(AppInfo.ProjectInfo.ViewModelFileName)

        ' ロードしたいメンバーをここに追加していく
        ViewModel.Views = vm.Views
        ViewModel.MultiView.MainGridHeight _
            = New GridLength(vm.MultiView.MainViewHeight)
        ViewModel.MultiView.RightGridWidth _
            = New GridLength(vm.MultiView.RightViewWidth)
    End Sub


    Private Sub _DataInitialize(Of T As {New})()
        Dim old As Object
        Dim [new] As Object
        If AppInfo.ProjectInfo.Model.Data IsNot Nothing Then
            old = AppInfo.ProjectInfo.Model.Data
            Select Case old.GetType
                Case (New Object).GetType
                    [new] = CType(old, T)
                Case (New JObject).GetType
                    ' Ｊｓｏｎからロードした場合は、JObject型になっている
                    [new] = old.ToObject(Of T)
                Case (New T).GetType
                    [new] = old
            End Select
        Else
            [new] = New T
        End If
        AppInfo.ProjectInfo.Model.Data = [new]
    End Sub


    ' << Add Case >>
    Public Sub ModelSetup()
        '-- you henkou --------------------------------'
        Select Case AppInfo.ProjectInfo.Kind
            Case AppDirectoryModel.DBTEST
                Call Me._DataInitialize(Of DBTestModel)()
            Case AppDirectoryModel.RpaProject
                Call Me._DataInitialize(Of RpaProjectModel)()
            Case Else
        End Select
        '----------------------------------------------'
    End Sub


    <JsonIgnore>
    Protected Property InitializeHandler As Action
    <JsonIgnore>
    Protected Property CheckCommandEnabledHandler As Action
    <JsonIgnore>
    Protected Property [AddHandler] As Action


    ' Model型のDataObjectはJsonからデシリアライズされたものはJObject型になっている
    Protected Sub BaseInitialize(ByVal app As AppDirectoryModel,
                                 ByRef vm As ViewModel)
        Dim ih = Me.InitializeHandler
        Dim cceh = Me.CheckCommandEnabledHandler
        Dim ah = Me.[AddHandler]

        AppInfo = app
        ViewModel = vm

        ' 自身(ビューモデル）の初期化設定を行います
        If ih <> Nothing Then
            Call ih()
        End If
        ' コマンド実行可否の設定
        If cceh <> Nothing Then
            Call cceh()
        End If
        ' イベントハンドラの登録
        If ah <> Nothing Then
            Call ah()
        End If
    End Sub

    Public MustOverride Sub Initialize(ByRef app As AppDirectoryModel,
                                       ByRef vm As ViewModel) Implements BaseViewModelInterface.Initialize

    Public Sub InitializeViewContent()
        'ViewModel.SingleView = Nothing
        'ViewModel.MultiView = Nothing
        ViewModel.Content = Nothing
    End Sub

    Private Sub _SetContent(ByVal obj As Object)
        ViewModel.Content = obj
    End Sub

    ' ＳＩＮＧＬＥＶＩＥＷ -------------------------------------------------------------------------------------------'
    Private Sub _SetItemToSingleViewContent(ByVal v As ViewItemModel)
        ViewModel.SingleView.Content = v.Content
        Call _SetContent(ViewModel.SingleView)
    End Sub

    '---------------------------------------------------------------------------------------------'
    ' ＤＹＮＡＭＩＣＶＩＥＷ
    Private Sub _InitializeDynamicView()
        'ViewModel.DynamicView = Nothing
    End Sub

    ' 必ず、ShowDynamicView()を通じてViewModel.Contentにセットする

    'Private Overloads Sub _CreateViews(ByVal dvm As DynamicViewModel)
    '    Dim act As Action(Of Object, String)
    '    act = Sub(ByVal obj As Object, ByVal [name] As String)
    '              Dim dvm2 As DynamicViewModel
    '              Dim tvm As TabViewModel
    '              Select Case obj.GetType.Name
    '                  Case "DynamicViewModel"
    '                      dvm2 = CType(obj, DynamicViewModel)
    '                      Call _CreateViews(dvm2)
    '                  Case "TabViewModel"
    '                      tvm = CType(obj, TabViewModel)
    '                      Call _CreateViews(tvm)
    '                  Case Else
    '                      ViewModel.Views.Add(
    '                         New ViewItemModel With {
    '                             .Name = [name]
    '                         }
    '                     )
    '              End Select
    '          End Sub
    '    Call act(dvm.MainContent, dvm.MainContentName)
    '    If dvm.RightContent IsNot Nothing Then
    '        Call act(dvm.RightContent, dvm.RightContentName)
    '    End If
    '    If dvm.BottomContent IsNot Nothing Then
    '        Call act(dvm.BottomContent, dvm.BottomContentName)
    '    End If
    'End Sub

    'Private Overloads Sub _CreateViews(ByVal tvm As TabViewModel)
    '    Dim dvm As DynamicViewModel
    '    Dim tvm2 As TabViewModel
    '    For Each t In tvm.Tabs
    '        Select Case t.Content.GetType.Name
    '            Case "DynamicViewModel"
    '                dvm = CType(t.Content, DynamicViewModel)
    '                Call _CreateViews(dvm)
    '            Case "TabViewModel"
    '                tvm2 = CType(t.Content, TabViewModel)
    '                Call _CreateViews(tvm2)
    '            Case Else
    '                ViewModel.Views.Add(
    '                    New ViewItemModel With {
    '                        .Name = t.Name
    '                    }
    '                )
    '        End Select
    '    Next
    'End Sub
    '---------------------------------------------------------------------------------------------'

    ' ＭＵＬＴＩＶＩＥＷ ---------------------------------------------------------------------------------------------'
    ' コンテントディクショナリが存在しない場合、新規作成し
    ' ビューディクショナリが存在しない場合、追加します
    Private Sub _RegisterViewToDictionary(ByVal v As ViewItemModel)
        Dim frame = v.FrameType
        If ViewModel.MultiView.ContentDictionary Is Nothing Then
            ViewModel.MultiView.ContentDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
        End If
        If Not ViewModel.MultiView.ContentDictionary.ContainsKey(frame) Then
            ViewModel.MultiView.ContentDictionary.Add(frame, New Dictionary(Of String, Object))
        End If
    End Sub

    ' ビューのDataContentに実際にセットします
    Private Sub _SetItemToMultiViewContent(ByVal v As ViewItemModel)
        Dim o As Object
        o = ViewModel.MultiView.ContentDictionary(v.FrameType)(v.Name)

        Select Case v.FrameType
            Case MultiViewModel.MAIN_FRAME
                ViewModel.MultiView.MainViewContent = o
            Case MultiViewModel.EXPLORER_FRAME
                ViewModel.MultiView.ExplorerViewContent = o
            Case MultiViewModel.HISTORY_FRAME
                ViewModel.MultiView.HistoryViewContent = o
            Case MultiViewModel.MENU_FRAME
                ViewModel.MultiView.MenuViewContent = o
            Case MultiViewModel.PROJECT_MENU_FRAME
                ViewModel.MultiView.ProjectMenuViewContent = o
            Case Else
                ' Nothing To Do
        End Select
        Call _SetContent(ViewModel.MultiView)
    End Sub

    Private Sub _AddItem(ByVal v As ViewItemModel)
        Dim obj = v.Content
        Dim frame = v.FrameType
        Dim [name] = obj.GetType.Name
        If ViewModel.MultiView.ContentDictionary(frame).ContainsKey([name]) Then
            ViewModel.MultiView.ContentDictionary(frame)([name]) = obj
        Else
            ViewModel.MultiView.ContentDictionary(frame).Add([name], obj)
        End If
    End Sub
    ' --------------------------------------------------------------------------------------------'

    ' タブコレクションディクショナリが存在しない場合、新規作成し
    ' タブコレクションが存在しない場合、追加します
    Private Sub _RegisterTabViewToDictionary(ByVal v As ViewItemModel)
        Dim frame = v.FrameType
        Dim tvm As TabViewModel
        If Not ViewModel.MultiView.TabsDictionary.ContainsKey(frame) Then
            tvm = New TabViewModel
            ViewModel.MultiView.TabsDictionary.Add(frame, tvm)
        End If
    End Sub


    Private Sub _SetTabItemToMultiViewContent(ByVal v As ViewItemModel)
        Dim o As TabViewModel
        o = ViewModel.MultiView.TabsDictionary(v.FrameType)

        Select Case v.FrameType
            Case MultiViewModel.MAIN_FRAME
                ViewModel.MultiView.MainViewContent = o
            Case MultiViewModel.EXPLORER_FRAME
                ViewModel.MultiView.ExplorerViewContent = o
            Case MultiViewModel.HISTORY_FRAME
                ViewModel.MultiView.HistoryViewContent = o
            Case Else
                ' Nothing To Do
        End Select
        Call _SetContent(ViewModel.MultiView)
    End Sub

    Public Overloads Sub RemoveTabItem(ByVal [tab] As TabItemModel)
        For Each pair In ViewModel.MultiView.TabsDictionary
            If pair.Value.Tabs.Contains([tab]) Then
                pair.Value.Tabs.Remove([tab])
                Call _ResizeMultiViewOfRemove(pair)
                Exit For
            End If
        Next

        ' ビューからの開閉状態記録
        For Each v In ViewModel.Views
            If v.Name = [tab].Name Then
                v.OpenState = False
            End If
        Next
    End Sub

    Private Sub _ResizeMultiViewOfRemove(ByVal pair As KeyValuePair(Of String, TabViewModel))
        Dim tvm = pair.Value

        If tvm.Tabs.Count = 0 Then
            Select Case pair.Key
                Case MultiViewModel.MAIN_FRAME
                Case MultiViewModel.EXPLORER_FRAME
                    ViewModel.MultiView.LeftViewPreservedWidth _
                        = ViewModel.MultiView.LeftViewWidth
                    ViewModel.MultiView.RightGridWidth _
                        = New GridLength(ViewModel.MultiView.RightViewWidth + ViewModel.MultiView.LeftViewWidth)
                Case MultiViewModel.HISTORY_FRAME
                    ViewModel.MultiView.HistoryViewPreservedHeight _
                        = ViewModel.MultiView.HistoryViewHeight
                    ViewModel.MultiView.MainGridHeight _
                        = New GridLength(ViewModel.MultiView.MainViewHeight + ViewModel.MultiView.HistoryViewHeight)
            End Select
        End If
    End Sub

    ' コレクションにタブがなければ追加、あれば更新
    'Private Overloads Sub _AddTabItem(ByVal [tab] As TabItemModel)
    'Private Sub _AddTabItem(ByVal v As ViewItemModel)
    '    Dim [tab] = New TabItemModel With {
    '        .Name = v.Name,
    '        .[Alias] = v.[Alias],
    '        .FrameType = v.FrameType,
    '        .Content = v.Content
    '    }

    '    Dim idx = -1
    '    Dim frame = [tab].FrameType
    '    For Each t In ViewModel.MultiView.TabsDictionary(frame).Tabs
    '        If [tab].Name = t.Name Then
    '            idx = ViewModel.MultiView.TabsDictionary(frame).Tabs.IndexOf(t)
    '        End If
    '    Next
    '    If idx = -1 Then
    '        ' 閉じるコマンドのセット
    '        ViewModel.MultiView.TabsDictionary(frame).Tabs.Add([tab])
    '        idx = ViewModel.MultiView.TabsDictionary(frame).Tabs.IndexOf([tab])
    '        Call _ResizeMultiViewOfAdd(frame)
    '    Else
    '        ViewModel.MultiView.TabsDictionary(frame).Tabs(idx).Content = [tab].Content
    '    End If
    '    ViewModel.MultiView.TabsDictionary(frame).SelectedIndex = idx
    'End Sub

    Private Sub _ResizeMultiViewOfAdd(ByVal frame As String)
        If ViewModel.MultiView.TabsDictionary(frame).Tabs.Count > 0 Then
            Select Case frame
                Case MultiViewModel.MAIN_FRAME
                Case MultiViewModel.EXPLORER_FRAME
                    If ViewModel.MultiView.LeftViewWidth = 0.0 Then
                        If ViewModel.MultiView.LeftViewPreservedWidth > 0.0 Then
                            ViewModel.MultiView.RightGridWidth _
                                = New GridLength(ViewModel.MultiView.RightViewWidth - ViewModel.MultiView.LeftViewPreservedWidth)
                        Else
                            ' 面倒なので、後で実装（実装の必要性ある？）
                            'Throw New Exception("LeftGridWidth Error")
                        End If
                    End If
                Case MultiViewModel.HISTORY_FRAME
                    If ViewModel.MultiView.HistoryViewHeight = 0.0 Then
                        If ViewModel.MultiView.HistoryViewPreservedHeight > 0.0 Then
                            ViewModel.MultiView.MainGridHeight _
                                = New GridLength(ViewModel.MultiView.MainViewHeight - ViewModel.MultiView.HistoryViewPreservedHeight)
                        Else
                            ' 面倒なので、後で実装（実装の必要性ある？）
                            'Throw New Exception("HistoryGridHeight Error")
                        End If
                    End If
            End Select
        End If
    End Sub

    Public Sub AddView(ByVal v As ViewItemModel)
        '    Select Case v.LayoutType
        '        Case ViewModel.SINGLE_VIEW
        '            Call _SetItemToSingleViewContent(v)
        '        Case ViewModel.MULTI_VIEW
        '            Select Case v.ViewType
        '                Case MultiViewModel.NORMAL_VIEW
        '                    Call _RegisterViewToDictionary(v)
        '                    Call _AddItem(v)
        '                    Call _SetItemToMultiViewContent(v)
        '                Case MultiViewModel.TAB_VIEW
        '                    Call _RegisterTabViewToDictionary(v)
        '                    Call _AssignIconOfView(v)
        '                    Call _AddTabItem(v)
        '                    Call _SetTabItemToMultiViewContent(v)
        '                Case Else
        '                    MsgBox("ViewModel.AddView")
        '                    Exit Sub
        '            End Select
        '    End Select
    End Sub

    Private Sub _AssignIconOfView(ByRef v As ViewItemModel)
        Dim iconf As String
        Dim f As Func(Of String, BitmapImage)
        Dim bi As BitmapImage

        Select Case v.FrameType
            Case MultiViewModel.MAIN_FRAME
                iconf = AppDirectoryModel.AppImageDirectory & "\mainframe.png"
            Case MultiViewModel.EXPLORER_FRAME
                iconf = AppDirectoryModel.AppImageDirectory & "\explorerframe.png"
            Case MultiViewModel.MENU_FRAME
                iconf = AppDirectoryModel.AppImageDirectory & "\menuframe.png"
            Case MultiViewModel.PROJECT_MENU_FRAME
                iconf = AppDirectoryModel.AppImageDirectory & "\projectmenuframe.png"
            Case MultiViewModel.HISTORY_FRAME
                iconf = AppDirectoryModel.AppImageDirectory & "\historyframe.png"
            Case Else
                iconf = vbNullString
        End Select

        If Not String.IsNullOrEmpty(iconf) Then
            bi = New BitmapImage
            bi.BeginInit()
            bi.UriSource = New Uri(
                iconf,
                UriKind.Absolute
            )
            bi.EndInit()
            v.Icon = bi
        End If
    End Sub

    Private Delegate Sub ViewSetupDelegater(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
    ''<< Add Case >>
    '<JsonIgnore>
    'Public ReadOnly Property ViewSetupHandler As ViewSetupDelegater
    '    Get
    '        Return AddressOf AppInfo.ProjectInfo.Model.Data.ViewSetupExecute
    '    End Get
    'End Property

    ''<< Add Case >>
    '<JsonIgnore>
    'Public ReadOnly Property ViewDefineHandler As Func(Of ViewItemModel, Object)
    '    Get
    '        Return AddressOf AppInfo.ProjectInfo.Model.Data.ViewDefineExecute
    '    End Get
    'End Property

    'Private Sub _ViewModelLoad(ByRef obj1 As Object, ByRef obj2 As Object)
    '    Dim obj As Object
    '    Dim [define] As Func(Of ViewItemModel, Object) _
    '        = AddressOf AppInfo.ProjectInfo.Model.Data.ViewDefineExecute
    '    For Each v In ViewModel.Views
    '        If v.OpenState Then
    '            obj = [define](v)
    '            obj.Initialize(AppInfo, ViewModel)
    '            v.Content = obj
    '            Call AddView(v)
    '        End If
    '    Next
    'End Sub

    Public Sub AddViewItem(ByVal [view] As ViewItemModel)
        Dim idx = -1
        For Each v In ViewModel.Views
            If [view].Name = v.Name Then
                idx = ViewModel.Views.IndexOf(v)
                Exit For
            End If
        Next
        If idx = -1 Then
            ' 閉じるコマンドのセット
            ViewModel.Views.Add([view])
        Else
            ViewModel.Views(idx) = [view]
        End If
    End Sub

    Protected Sub ShowViewExplorer()
        '    Dim v = New ViewItemModel With {
        '        .Content = ViewModel,
        '        .Name = ViewModel.GetType.Name,
        '        .ModelName = ViewModel.GetType.Name,
        '        .FrameType = MultiViewModel.EXPLORER_FRAME,
        '        .LayoutType = ViewModel.MULTI_VIEW,
        '        .ViewType = MultiViewModel.TAB_VIEW,
        '        .OpenState = True
        '    }
        '    Call AddViewItem(v)
        '    Call AddView(v)
    End Sub

    Protected Sub ShowProjectExplorer()
        '    Dim v = New ViewItemModel With {
        '        .Content = AppInfo,
        '        .Name = AppInfo.GetType.Name,
        '        .ModelName = AppInfo.GetType.Name,
        '        .FrameType = MultiViewModel.EXPLORER_FRAME,
        '        .LayoutType = ViewModel.MULTI_VIEW,
        '        .ViewType = MultiViewModel.TAB_VIEW,
        '        .OpenState = True
        '    }
        '    Call AddViewItem(v)
        '    Call AddView(v)
    End Sub

    Protected Sub ShowMenu()
        '    Dim mvm = New MenuViewModel
        '    mvm.Initialize(AppInfo, ViewModel)
        '    Dim v = New ViewItemModel With {
        '        .Content = mvm,
        '        .Name = mvm.GetType.Name,
        '        .ModelName = mvm.GetType.Name,
        '        .FrameType = MultiViewModel.MENU_FRAME,
        '        .LayoutType = ViewModel.MULTI_VIEW,
        '        .ViewType = MultiViewModel.NORMAL_VIEW,
        '        .OpenState = True
        '    }
        '    Call AddViewItem(v)
        '    Call AddView(v)
    End Sub

    Protected Sub ShowHistory()
        '    Dim hvm = New HistoryViewModel
        '    hvm.Initialize(AppInfo, ViewModel)
        '    Dim v = New ViewItemModel With {
        '        .Content = hvm,
        '        .Name = hvm.GetType.Name,
        '        .ModelName = hvm.GetType.Name,
        '        .FrameType = MultiViewModel.HISTORY_FRAME,
        '        .LayoutType = ViewModel.MULTI_VIEW,
        '        .ViewType = MultiViewModel.TAB_VIEW,
        '        .OpenState = True
        '    }
        '    Call AddViewItem(v)
        '    Call AddView(v)
    End Sub

    Public Overloads Sub ViewModelSetup()
        Dim v2, v3
        Dim mvm, hvm
        Dim obj
        Dim [setup] As ViewSetupDelegater _
            = AddressOf AppInfo.ProjectInfo.Model.Data.ViewSetupExecute
        Dim [define] As Func(Of ViewItemModel, Object) _
            = AddressOf AppInfo.ProjectInfo.Model.Data.ViewDefineExecute

        'If ViewModel.Views.Count < 1 Then
        '    Call [setup](AppInfo, ViewModel)

        '    ' 特殊なビューのセット(ViewModel)
        '    ' ViewDefineHandlerに登録したメソッドが定義してなければ、ビューは登録されない
        '    If [define]((New ViewItemModel With {.ModelName = ViewModel.GetType.Name})).GetType.Name = ViewModel.GetType.Name Then
        '        Call ShowViewExplorer()
        '    End If
        '    If [define]((New ViewItemModel With {.ModelName = AppInfo.GetType.Name})).GetType.Name = AppInfo.GetType.Name Then
        '        Call ShowProjectExplorer()
        '    End If
        '    If [define]((New ViewItemModel With {.ModelName = (New MenuViewModel).GetType.Name})).GetType.Name = (New MenuViewModel).GetType.Name Then
        '        Call ShowMenu()
        '    End If
        '    If [define]((New ViewItemModel With {.ModelName = (New HistoryViewModel).GetType.Name})).GetType.Name = (New HistoryViewModel).GetType.Name Then
        '        Call ShowHistory()
        '    End If
        'Else
        '    For Each v In ViewModel.Views
        '        If v.OpenState Then
        '            Call ViewLoad(v)
        '        End If
        '    Next
        'End If
        Call InitializeViewContent()

        If ViewModel.Content Is Nothing Then
            Call [setup](AppInfo, ViewModel)
        Else
            For Each v In ViewModel.Views
                If v.OpenState Then
                    Call ViewLoad(v)
                End If
            Next
        End If
    End Sub

    Protected Sub ViewLoad(ByVal v As ViewItemModel)
        'Dim [define] As Func(Of ViewItemModel, Object) _
        '    = AddressOf AppInfo.ProjectInfo.Model.Data.ViewDefineExecute

        'Dim obj = [define](v)
        'If obj IsNot Nothing Then
        '    Select Case v.ModelName
        '        Case (New ViewModel).GetType.Name
        '            v.Content = ViewModel
        '            Call AddView(v)
        '        Case (New AppDirectoryModel).GetType.Name
        '            v.Content = AppInfo
        '            Call AddView(v)
        '        Case Else
        '            obj.Initialize(AppInfo, ViewModel)
        '            v.Content = obj
        '            Call AddView(v)
        '    End Select
        'End If
    End Sub

End Class
