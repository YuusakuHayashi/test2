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

    Public Sub AppSave()
        Call AppInfo.ModelSave(AppDirectoryModel.ModelFileName, AppInfo)
    End Sub

    Public Overloads Sub ProjectLoad()
        Dim jh = New JsonHandler(Of ProjectInfoModel)
        AppInfo.ProjectInfo = jh.ModelLoad(AppInfo.ProjectInfo.ProjectInfoFileName)
    End Sub
    Public Overloads Sub ProjectLoad(ByVal project As ProjectInfoModel)
        Dim jh = New JsonHandler(Of ProjectInfoModel)
        AppInfo.ProjectInfo = jh.ModelLoad(project.ProjectInfoFileName)
    End Sub

    Public Overloads Sub ProjectSave()
        Call ProjectSave(AppInfo.ProjectInfo)
    End Sub
    Public Overloads Sub ProjectSave(ByVal project As ProjectInfoModel)
        project.LastUpdate = DateTime.Now.ToString("yyyy/MM/dd")
        Call PushProject(project)
        Call project.ModelSave(
            project.ProjectInfoFileName, project
        )
        Call AppSave()
    End Sub

    Public Overloads Sub ProjectModelSave(ByVal project As ProjectInfoModel)
        Call project.Model.ModelSave(
            project.ModelFileName, project.Model
        )
    End Sub
    Public Overloads Sub ProjectModelSave()
        Call AppInfo.ProjectInfo.Model.ModelSave(
            AppInfo.ProjectInfo.ModelFileName, AppInfo.ProjectInfo.Model
        )
    End Sub

    Public Overloads Sub ProjectModelLoad()
        Dim jh = New JsonHandler(Of Model)
        AppInfo.ProjectInfo.Model = jh.ModelLoad(AppInfo.ProjectInfo.ModelFileName)
    End Sub

    Public Overloads Sub ProjectViewModelSave(ByVal project As ProjectInfoModel)
        Call ViewModel.ModelSave(
            project.ViewModelFileName, ViewModel
        )
    End Sub
    Public Overloads Sub ProjectViewModelSave()
        Call ViewModel.ModelSave(
            AppInfo.ProjectInfo.ViewModelFileName, ViewModel
        )
    End Sub

    'Public Overloads Sub AllLoad()
    '    Call ProjectLoad()
    '    Call ProjectModelLoad()
    '    Call ModelSetup()
    '    Call ProjectViewModelLoad()
    '    Call ViewModelSetup()
    'End Sub

    Public Overloads Sub AllLoad(project)
        Call ProjectLoad(project)
        Call ProjectModelLoad()
        Call ModelSetup()
        Call ProjectViewModelLoad()
        Call ViewModelSetup()
    End Sub

    Public Overloads Sub AllSave()
        Call ProjectModelSave()
        Call ProjectViewModelSave()
        Call ProjectSave()
    End Sub

    Public Overloads Sub AllSave(ByVal project As ProjectInfoModel)
        Call ProjectModelSave(project)
        Call ProjectViewModelSave(project)
        Call ProjectSave(project)
    End Sub

    Public Sub PushProject(ByVal project As ProjectInfoModel)
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

    ' ＶｉｅｗＭｏｄｅｌはロードするとＭＶＶＭが機能しなくなるので、
    ' 部分的に更新する
    Public Overloads Sub ProjectViewModelLoad()
        Dim jh = New JsonHandler(Of ViewModel)
        Dim vm = jh.ModelLoad(AppInfo.ProjectInfo.ViewModelFileName)

        ' ロードしたいメンバーをここに追加していく
        ViewModel.Views = vm.Views
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
        ViewModel.SingleView = Nothing
        ViewModel.MultiView = Nothing
    End Sub

    Private Sub _SetContent(ByVal obj As Object)
        ViewModel.Content = obj
    End Sub

    ' ＳＩＮＧＬＥＶＩＥＷ -------------------------------------------------------------------------------------------'
    Private Sub _SetItemToSingleViewContent(ByVal v As ViewItemModel)
        ViewModel.SingleView.Content = v.Content
        Call _SetContent(ViewModel.SingleView)
    End Sub

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
        For Each p In ViewModel.MultiView.TabsDictionary
            If p.Value.Tabs.Contains([tab]) Then
                p.Value.Tabs.Remove([tab])
            End If
        Next

        ' ビューからの開閉状態記録
        For Each v In ViewModel.Views
            If v.Name = [tab].Name Then
                v.OpenState = False
            End If
        Next
    End Sub

    ' コレクションにタブがなければ追加、あれば更新
    'Private Overloads Sub _AddTabItem(ByVal [tab] As TabItemModel)
    Private Sub _AddTabItem(ByVal v As ViewItemModel)
        Dim [tab] = New TabItemModel With {
            .Name = v.Name,
            .[Alias] = v.[Alias],
            .FrameType = v.FrameType,
            .Content = v.Content
        }

        Dim idx = -1
        Dim frame = [tab].FrameType
        For Each t In ViewModel.MultiView.TabsDictionary(frame).Tabs
            If [tab].Name = t.Name Then
                idx = ViewModel.MultiView.TabsDictionary(frame).Tabs.IndexOf(t)
            End If
        Next
        If idx = -1 Then
            ' 閉じるコマンドのセット
            ViewModel.MultiView.TabsDictionary(frame).Tabs.Add([tab])
            idx = ViewModel.MultiView.TabsDictionary(frame).Tabs.IndexOf([tab])
        Else
            ViewModel.MultiView.TabsDictionary(frame).Tabs(idx).Content = [tab].Content
            'idx = idx
        End If
        ViewModel.MultiView.TabsDictionary(frame).SelectedIndex = idx
    End Sub

    Public Sub AddView(ByVal v As ViewItemModel)
        Select Case v.LayoutType
            Case ViewModel.SINGLE_VIEW
                Call _SetItemToSingleViewContent(v)
            Case ViewModel.MULTI_VIEW
                Select Case v.ViewType
                    Case MultiViewModel.NORMAL_VIEW
                        Call _RegisterViewToDictionary(v)
                        Call _AddItem(v)
                        Call _SetItemToMultiViewContent(v)
                    Case MultiViewModel.TAB_VIEW
                        Call _RegisterTabViewToDictionary(v)
                        Call _AddTabItem(v)
                        Call _SetTabItemToMultiViewContent(v)
                    Case Else
                        MsgBox("ViewModel.AddView")
                        Exit Sub
                End Select
        End Select
    End Sub

    'Public Function AddViewItem(ByVal obj As Object,
    '                            ByVal [layout] As String,
    '                            ByVal frame As String,
    '                            ByVal view As String,
    '                            Optional ByVal nm As String = vbNullString) As ViewItemModel
    'Dim [name] = IIf(String.IsNullOrEmpty(nm), obj.GetType.Name, nm)
    'Dim v As New ViewItemModel With {
    '    .Name = [name],
    '    .FrameType = frame,
    '    .ViewType = view,
    '    .LayoutType = [layout],
    '    .OpenState = True,
    '    .Content = obj
    '}

    ' ViewModel, AppDirectoryModel, ProjectInfoModelは登録しない
    'Select Case True
    '    Case obj.Equals(ViewModel)
    '    Case obj.Equals(AppInfo)
    '    Case obj.GetType.Name = (New ProjectInfoModel).GetType.Name
    '    Case obj.GetType.Name = (New ProjectViewModel).GetType.Name
    '    Case obj.GetType.Name = (New MenuViewModel).GetType.Name
    '    Case obj.GetType.Name = (New HistoryViewModel).GetType.Name
    '    Case Else
    '        ViewModel.Views.Add(v)
    'End Select
    'AddViewItem = v
    'End Function

    Public Sub AddViewItem(ByVal v As ViewItemModel)
        Dim obj = v.Content
        Select Case True
            Case obj.Equals(ViewModel)
            Case obj.Equals(AppInfo)
            Case obj.GetType.Name = (New ProjectInfoModel).GetType.Name
            Case obj.GetType.Name = (New ProjectViewModel).GetType.Name
            Case obj.GetType.Name = (New MenuViewModel).GetType.Name
            Case obj.GetType.Name = (New HistoryViewModel).GetType.Name
            Case Else
                ViewModel.Views.Add(v)
        End Select
    End Sub

    Public Delegate Sub ViewSetupDelegater(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
    '<< Add Case >>
    <JsonIgnore>
    Public ReadOnly Property ViewSetupHandler As ViewSetupDelegater
        Get
            Select Case AppInfo.ProjectInfo.Kind
                Case AppDirectoryModel.DBTEST
                    Return AddressOf ViewSetupModule.DBTESTViewSetupExecute
                Case AppDirectoryModel.RpaProject
                    Return AddressOf ViewSetupModule.RpaProjectViewSetupExecute
                Case Else
                    Throw New Exception("No ViewSetupHandler")
            End Select
        End Get
    End Property

    'Public Delegate Sub ViewLoadDelegater(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
    ''<< Add Case >>
    '<JsonIgnore>
    'Public ReadOnly Property ViewLoadHandler As ViewLoadDelegater
    '    Get
    '        Select Case AppInfo.ProjectInfo.Kind
    '            Case AppDirectoryModel.DBTEST
    '                Return AddressOf _ViewModelLoad
    '            Case AppDirectoryModel.RpaProject
    '                Return AddressOf ViewSetupModule.RpaProjectViewLoadExecute
    '            Case Else
    '                Throw New Exception("No ViewLoadHandler")
    '        End Select
    '    End Get
    'End Property

    '<< Add Case >>
    <JsonIgnore>
    Public ReadOnly Property ViewDefineHandler As Func(Of ViewItemModel, Object)
        'Public ReadOnly Property ViewDefineHandler As Func(Of String, Object)
        Get
            Select Case AppInfo.ProjectInfo.Kind
                Case AppDirectoryModel.DBTEST
                    Return AddressOf ViewSetupModule.DBTESTViewDefineExecute
                Case AppDirectoryModel.RpaProject
                    Return AddressOf ViewSetupModule.RpaProjectViewDefineExecute
                Case Else
                    Throw New Exception("No ViewDefineHandler")
            End Select
        End Get
    End Property

    Private Sub _ViewModelLoad(ByRef obj1 As Object, ByRef obj2 As Object)
        Dim obj As Object
        Dim [define] = Me.ViewDefineHandler
        For Each v In ViewModel.Views
            If v.OpenState Then
                obj = [define](v)
                obj.Initialize(AppInfo, ViewModel)
                v.Content = obj
                Call AddView(v)
            End If
        Next
    End Sub

    Public Overloads Sub ViewModelSetup()
        Dim v0, v1, v2, v3
        Dim mvm, hvm
        Dim obj
        Dim [setup] = ViewSetupHandler
        'Dim [load] = ViewLoadHandler
        Dim [define] = ViewDefineHandler

        If ViewModel.Views.Count < 1 Then
            Call [setup](AppInfo, ViewModel)
        Else
            'Call [load](AppInfo, ViewModel)
            For Each v In ViewModel.Views
                If v.OpenState Then
                    obj = [define](v)
                    obj.Initialize(AppInfo, ViewModel)
                    v.Content = obj
                    Call AddView(v)
                End If
            Next
            'For Each v In ViewModel.Views
            '    If v.OpenState Then
            '        obj = [define](v.Name)
            '        obj.Initialize(AppInfo, ViewModel)
            '        v.Content = obj
            '        Call AddView(v)
            '    End If
            'Next
        End If

        ' 特殊なビューのセット(ViewModel)
        v0 = New ViewItemModel With {
            .Content = ViewModel,
            .Name = ViewModel.GetType.Name,
            .FrameType = MultiViewModel.EXPLORER_FRAME,
            .LayoutType = ViewModel.MULTI_VIEW,
            .ViewType = MultiViewModel.TAB_VIEW,
            .OpenState = True
        }
        v1 = New ViewItemModel With {
            .Content = AppInfo,
            .Name = AppInfo.GetType.Name,
            .FrameType = MultiViewModel.EXPLORER_FRAME,
            .LayoutType = ViewModel.MULTI_VIEW,
            .ViewType = MultiViewModel.TAB_VIEW,
            .OpenState = True
        }
        mvm = New MenuViewModel
        mvm.Initialize(AppInfo, ViewModel)
        v2 = New ViewItemModel With {
            .Content = mvm,
            .Name = mvm.GetType.Name,
            .FrameType = MultiViewModel.MENU_FRAME,
            .LayoutType = ViewModel.MULTI_VIEW,
            .ViewType = MultiViewModel.NORMAL_VIEW,
            .OpenState = True
        }
        hvm = New HistoryViewModel
        hvm.Initialize(AppInfo, ViewModel)
        v3 = New ViewItemModel With {
            .Content = hvm,
            .Name = hvm.GetType.Name,
            .FrameType = MultiViewModel.HISTORY_FRAME,
            .LayoutType = ViewModel.MULTI_VIEW,
            .ViewType = MultiViewModel.TAB_VIEW,
            .OpenState = True
        }
        Call AddViewItem(v0)
        Call AddViewItem(v1)
        Call AddViewItem(v2)
        Call AddViewItem(v3)
        Call AddView(v0)
        Call AddView(v1)
        Call AddView(v2)
        Call AddView(v3)
    End Sub
End Class
