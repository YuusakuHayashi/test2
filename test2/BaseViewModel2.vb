Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Collections.ObjectModel
Imports System.IO

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
        Call ProjectModelLoad()
        Call ModelSetup()
        Call ProjectSetup()
        Call PushProject(AppInfo.ProjectInfo)
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
        ' ViewModel.Views = vm.Views
        ViewModel.SaveContent = vm.Content
        'ViewModel.MultiView.MainGridHeight _
        '    = New GridLength(vm.MultiView.MainViewHeight)
        'ViewModel.MultiView.RightGridWidth _
        '    = New GridLength(vm.MultiView.RightViewWidth)
    End Sub


    ' ConvertJObjToObj に統合予定
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

        ' IConFileNameはプロジェクトのモデルが確定した時点で決定する
        If File.Exists(AppInfo.ProjectInfo.Model.Data.IconFileName) Then
            AppInfo.ProjectInfo.IconFileName _
                = AppInfo.ProjectInfo.Model.Data.IconFileName
        End If
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

    Private Delegate Function ViewSetupDelegater(ByRef app As AppDirectoryModel, ByRef vm As ViewModel) As ViewItemModel

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

    Public Overloads Sub ViewModelSetup()
        Dim [setup] As ViewSetupDelegater _
            = AddressOf AppInfo.ProjectInfo.Model.Data.ViewSetupExecute
        Dim [define] As Func(Of String, Object) _
            = AddressOf AppInfo.ProjectInfo.Model.Data.ViewDefineExecute

        Dim sv As FlexibleViewModel
        Dim obj As Object
        Dim ld As FlexibleViewModel
        Call InitializeViewContent()

        If ViewModel.SaveContent Is Nothing Then
            Call _NewViewSetup()
        Else
            sv = ViewModel.SaveContent
            ld = FlexibleViewLoad(sv)
            Call ViewModel.VisualizeView(ld)
        End If
    End Sub

    Private Sub _NewViewSetup()
        Dim [setup] As ViewSetupDelegater _
            = AddressOf AppInfo.ProjectInfo.Model.Data.ViewSetupExecute
        Dim fvm As FlexibleViewModel
        Dim mvm = New MenuViewModel
        Dim pevm = New ProjectExplorerViewModel
        Dim vevm = New ViewExplorerViewModel
        Dim hvm = New HistoryViewModel

        Dim etvm = New TabViewModel
        Dim htvm = New TabViewModel

        Call hvm.Initialize(AppInfo, ViewModel)
        Call mvm.Initialize(AppInfo, ViewModel)
        Call pevm.Initialize(AppInfo, ViewModel)
        Call vevm.Initialize(AppInfo, ViewModel)

        Call etvm.AddTab(New ViewItemModel With {
            .Name = "ViewExp",
            .Content = vevm
        })
        Call etvm.AddTab(New ViewItemModel With {
            .Name = "ProjExp",
            .Content = pevm
        })

        Call htvm.AddTab(New ViewItemModel With {
            .Name = "History",
            .Content = hvm
        })

        fvm = New FlexibleViewModel With {
            .ContentViewHeight = 25.0,
            .MainViewContent = New ViewItemModel With {
                .Name = "Menu",
                .Content = mvm
            },
            .BottomViewContent = New ViewItemModel With {
                .Name = "LeftView",
                .Content = New FlexibleViewModel With {
                    .ContentViewWidth = 200.0,
                    .MainViewContent = New ViewItemModel With {
                        .Name = "ExpTabs",
                        .Content = etvm
                    },
                    .RightViewContent = New ViewItemModel With {
                        .Name = "RightView",
                        .Content = New FlexibleViewModel With {
                            .MainViewContent = [setup](AppInfo, ViewModel),
                            .BottomViewHeight = 200.0,
                            .BottomViewContent = New ViewItemModel With {
                                .Name = "HistTabs",
                                .Content = htvm
                            }
                        }
                    }
                }
            }
        }
        Call _ViewItemSetup(fvm)
        Call ViewModel.VisualizeView(fvm)
    End Sub

    '---------------------------------------------------------------------------------------------'
    ' 初回時(NewViewSetup時のみ実行され、各ViewItemに必要情報を付加する)

    Private Overloads Sub _ViewItemSetup(ByRef fvm As FlexibleViewModel)
        If fvm.MainViewContent IsNot Nothing Then
            Call _ViewItemSetup(fvm.MainViewContent)
        End If
        If fvm.RightViewContent IsNot Nothing Then
            Call _ViewItemSetup(fvm.RightViewContent)
        End If
        If fvm.BottomViewContent IsNot Nothing Then
            Call _ViewItemSetup(fvm.BottomViewContent)
        End If
    End Sub

    Private Overloads Sub _ViewItemSetup(ByRef vim As ViewItemModel)
        Dim fvm As FlexibleViewModel
        Dim tvm As TabViewModel

        ' ここに初回時にセットさせたい項目を記述
        '---------------------------------------'
        vim.IsVisible = True
        '---------------------------------------'

        Select Case vim.Content.GetType.Name
            Case "FlexibleViewModel"
                fvm = CType(vim.Content, FlexibleViewModel)
                Call _ViewItemSetup(fvm)
            Case "TabViewModel"
                tvm = CType(vim.Content, TabViewModel)
                Call _ViewItemSetup(tvm)
            Case Else
        End Select
    End Sub

    Private Overloads Sub _ViewItemSetup(ByRef tvm As TabViewModel)
        For Each vt In tvm.ViewContentTabs
            Call _ViewItemSetup(vt)
        Next
    End Sub
    '---------------------------------------------------------------------------------------------'


    Private Function FlexibleViewLoad(ByVal [save] As FlexibleViewModel) As FlexibleViewModel
        If [save].MainViewContent IsNot Nothing Then
            [save].MainViewContent = _ViewItemLoad([save].MainViewContent)
        End If
        If [save].RightViewContent IsNot Nothing Then
            [save].RightViewContent = _ViewItemLoad([save].RightViewContent)
        End If
        If [save].BottomViewContent IsNot Nothing Then
            [save].BottomViewContent = _ViewItemLoad([save].BottomViewContent)
        End If
        FlexibleViewLoad = [save]
    End Function

    Private Function _ViewItemLoad(ByVal [save] As ViewItemModel)
        Dim obj As Object
        Dim [define] As Func(Of String, Object) _
            = AddressOf AppInfo.ProjectInfo.Model.Data.ViewDefineExecute
        Select Case [save].ModelName
            Case "FlexibleViewModel"
                obj = _ConvertJObjToObj(Of FlexibleViewModel)([save].Content)
                [save].Content = FlexibleViewLoad(obj)
            Case "TabViewModel"
                obj = _ConvertJObjToObj(Of TabViewModel)([save].Content)
                obj = _TabViewLoad(obj)
                obj = _ViewContentTabViewLoad(obj)
                [save].Content = obj
            Case Else
                obj = [define]([save].ModelName)
                obj.Initialize(AppInfo, ViewModel)
                [save].Content = obj
        End Select
        _ViewItemLoad = [save]
    End Function

    Private Function _TabViewLoad(ByRef [save] As TabViewModel) As TabViewModel
        Dim vim As ViewItemModel
        For Each t In [save].Tabs
            vim = _ViewItemLoad(t.ViewContent)
            t.ViewContent = vim
        Next
        _TabViewLoad = [save]
    End Function

    Private Function _ViewContentTabViewLoad(ByRef [save] As TabViewModel) As TabViewModel
        For Each vt In [save].ViewContentTabs
            vt = _ViewItemLoad(vt)
        Next
        _ViewContentTabViewLoad = [save]
    End Function

    Private Function _ConvertJObjToObj(Of T As {New})(ByVal jobj As Object) As T
        Dim obj As Object
        Select Case jobj.GetType
            Case (New Object).GetType
                obj = CType(jobj, T)
            Case (New JObject).GetType
                ' Ｊｓｏｎからロードした場合は、JObject型になっている
                obj = jobj.ToObject(Of T)
            Case (New T).GetType
                obj = jobj
            Case Else
                Throw New Exception("BaseViewModel2.ConvertJObjToObj Error!!!")
        End Select
        _ConvertJObjToObj = obj
    End Function

End Class
