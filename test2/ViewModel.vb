Imports System.Collections.ObjectModel

Public Class ViewModel
    Inherits BaseModel(Of ViewModel)

    Public Const MAIN_VIEW As String = "MainView"
    Public Const EXPLORER_VIEW As String = "ExplorerView"
    Public Const HISTORY_VIEW As String = "HistoryView"

    Private _MainViewContent As Object
    Public Property MainViewContent As Object
        Get
            Return Me._MainViewContent
        End Get
        Set(value As Object)
            Me._MainViewContent = value
            RaisePropertyChanged("MainViewContent")
        End Set
    End Property

    Private _ExplorerViewContent As Object
    Public Property ExplorerViewContent As Object
        Get
            Return Me._ExplorerViewContent
        End Get
        Set(value As Object)
            Me._ExplorerViewContent = value
            RaisePropertyChanged("ExplorerViewContent")
        End Set
    End Property

    Private _HistoryViewContent As Object
    Public Property HistoryViewContent As Object
        Get
            Return Me._HistoryViewContent
        End Get
        Set(value As Object)
            Me._HistoryViewContent = value
            RaisePropertyChanged("HistoryViewContent")
        End Set
    End Property

    'Private _MainTabViewContents As ObservableCollection(Of TabItemModel)
    'Public Property MainTabViewContents As ObservableCollection(Of TabItemModel)
    '    Get
    '        Return Me._MainTabViewContents
    '    End Get
    '    Set(value As ObservableCollection(Of TabItemModel))
    '        Me._MainTabViewContents = value
    '        RaisePropertyChanged("MainTabViewContents")
    '    End Set
    'End Property

    'Private _MainViewHeight As Integer
    'Public Property MainViewHeight As Integer
    '    Get
    '        Return Me._MainViewHeight
    '    End Get
    '    Set(value As Integer)
    '        Me._MainViewHeight = value
    '        RaisePropertyChanged("MainViewHeight")
    '    End Set
    'End Property

    'Private _MainViewWidth As Integer
    'Public Property MainViewWidth As Integer
    '    Get
    '        Return Me._MainViewWidth
    '    End Get
    '    Set(value As Integer)
    '        Me._MainViewWidth = value
    '        RaisePropertyChanged("MainViewWidth")
    '    End Set
    'End Property


    'Private _ExplorerViewHeight As Integer
    'Public Property ExplorerViewHeight As Integer
    '    Get
    '        Return Me._ExplorerViewHeight
    '    End Get
    '    Set(value As Integer)
    '        Me._ExplorerViewHeight = value
    '        RaisePropertyChanged("ExplorerViewHeight")
    '    End Set
    'End Property

    'Private _ExplorerViewWidth As Integer
    'Public Property ExplorerViewWidth As Integer
    '    Get
    '        Return Me._ExplorerViewWidth
    '    End Get
    '    Set(value As Integer)
    '        Me._ExplorerViewWidth = value
    '        RaisePropertyChanged("ExplorerViewWidth")
    '    End Set
    'End Property

    'Private _HistoryViewHeight As Integer
    'Public Property HistoryViewHeight As Integer
    '    Get
    '        Return Me._HistoryViewHeight
    '    End Get
    '    Set(value As Integer)
    '        Me._HistoryViewHeight = value
    '        RaisePropertyChanged("HistoryViewHeight")
    '    End Set
    'End Property

    'Private _HistoryViewWidth As Integer
    'Public Property HistoryViewWidth As Integer
    '    Get
    '        Return Me._HistoryViewWidth
    '    End Get
    '    Set(value As Integer)
    '        Me._HistoryViewWidth = value
    '        RaisePropertyChanged("HistoryViewWidth")
    '    End Set
    'End Property

    ' コンテントディクショナリ関連 ---------------------------------------------------------------'
    ' コンテントディクショナリ
    Private _ContentDictionary As Dictionary(Of String, Dictionary(Of String, Object))
    Public Property ContentDictionary As Dictionary(Of String, Dictionary(Of String, Object))
        Get
            Return Me._ContentDictionary
        End Get
        Set(value As Dictionary(Of String, Dictionary(Of String, Object)))
            Me._ContentDictionary = value
        End Set
    End Property

    ' コンテントディクショナリが存在しない場合、新規作成し
    ' ビューディクショナリが存在しない場合、追加します
    Private Sub _RegisterViewToDictionary(ByVal view As String)
        If Me.ContentDictionary Is Nothing Then
            Me.ContentDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
        End If
        If Not Me.ContentDictionary.ContainsKey(view) Then
            Me.ContentDictionary.Add(view, New Dictionary(Of String, Object))
        End If
    End Sub

    ' ビューのDataContentに実際にセットします
    Private Sub _SetContentObject(ByVal view As String, ByVal nm As String)
        Dim o As Object
        o = Me.ContentDictionary(view)(nm)

        Select Case view
            Case MAIN_VIEW
                Me.MainViewContent = o
            Case EXPLORER_VIEW
                Me.ExplorerViewContent = o
            Case HISTORY_VIEW
                Me.HistoryViewContent = o
            Case Else
                ' Nothing To Do
        End Select
    End Sub

    ' コンテントをディクショナリにセットします
    Private Sub _AddContentToDictionary(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._RegisterViewToDictionary(viewName)
        If Not Me.ContentDictionary(viewName).ContainsKey(modelName) Then
            Me.ContentDictionary(viewName).Add(modelName, context)
        End If
    End Sub

    ' コンテントをディクショナリにセット＆更新します
    Private Sub _UpdateContentToDictionary(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._AddContentToDictionary(viewName, modelName, context)
        Me.ContentDictionary(viewName)(modelName) = context
    End Sub

    ' コンテントをディクショナリにセット＆ビューの切り替えを行います
    Public Sub ChangeContent(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._AddContentToDictionary(viewName, modelName, context)
        Call Me._SetContentObject(viewName, modelName)
    End Sub

    ' コンテントをディクショナリにセット＆更新＆ビューの切り替えを行います
    Public Sub SetContent(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._UpdateContentToDictionary(viewName, modelName, context)
        Call Me._SetContentObject(viewName, modelName)
    End Sub
    ' --------------------------------------------------------------------------------------------'

    ' --------------------------------------------------------------------------------------------'
    ' タブコレクション関連
    ' タブディクショナリ
    Private _TabsDictionary As Dictionary(Of String, TabViewModel)
    Public Property TabsDictionary As Dictionary(Of String, TabViewModel)
        Get
            Return Me._TabsDictionary
        End Get
        Set(value As Dictionary(Of String, TabViewModel))
            Me._TabsDictionary = value
        End Set
    End Property

    ' タブディクショナリが存在しない場合、新規作成し
    ' ビューディクショナリが存在しない場合、追加します
    Private Sub _RegisterTabViewToDictionary(ByVal view As String)
        If Me.TabsDictionary Is Nothing Then
            Me.TabsDictionary = New Dictionary(Of String, TabViewModel)
        End If
        If Not Me.TabsDictionary.ContainsKey(view) Then
            Me.TabsDictionary.Add(view, New TabViewModel)
        End If
        If Me.TabsDictionary(view).Tabs Is Nothing Then
            Me.TabsDictionary(view).Tabs = New ObservableCollection(Of TabItemModel)
        End If
    End Sub

    ' ビューのDataContentに実際にセットします
    Private Sub _SetTabsObject(ByVal view As String)
        Dim o As TabViewModel
        o = Me.TabsDictionary(view)

        Select Case view
            Case MAIN_VIEW
                Me.MainViewContent = o
            Case EXPLORER_VIEW
                Me.ExplorerViewContent = o
            Case HISTORY_VIEW
                Me.HistoryViewContent = o
            Case Else
                ' Nothing To Do
        End Select
    End Sub

    ' タブをコレクションにセットします
    Private Sub _AddTabsToCollection(ByVal view As String, ByVal [tab] As TabItemModel)
        Call Me._RegisterTabViewToDictionary(view)
        If Not Me.TabsDictionary(view).Tabs.Contains([tab]) Then
            Me.TabsDictionary(view).Tabs.Add([tab])
        End If
    End Sub

    ' タブをコレクションにセット＆更新します
    Private Sub _UpdateTabsToCollection(ByVal view As String, ByVal [tab] As TabItemModel)
        Call Me._AddTabsToCollection(view, [tab])
        Me.TabsDictionary(view).Tabs(Me.TabsDictionary(view).Tabs.IndexOf([tab])) = [tab]
    End Sub

    ' タブをコレクションにセット＆ビューの切り替えを行います
    Public Sub ChangeTabs(ByVal view As String, ByVal [tab] As TabItemModel)
        Call Me._AddTabsToCollection(view, [tab])
        Call Me._SetTabsObject(view)
    End Sub

    ' タブをコレクションにセット＆更新＆ビューの切り替えを行います
    Public Sub SetTabs(ByVal view As String, ByVal [tab] As TabItemModel)
        Call Me._UpdateTabsToCollection(view, [tab])
        Call Me._SetTabsObject(view)
    End Sub
    ' --------------------------------------------------------------------------------------------'

    Public Delegate Sub InitializeProxy(ByVal pk As String, ByVal m As Model, ByVal vm As ViewModel, ByVal adm As AppDirectoryModel, ByVal pim As ProjectInfoModel)

    Public Sub NoInitialize(ByVal pk As String,
                            ByVal m As Model,
                            ByVal vm As ViewModel,
                            ByVal adm As AppDirectoryModel,
                            ByVal pim As ProjectInfoModel)
        ' Nothing To Do
    End Sub

    ' 初回時に必要なビューモデルコンテントを全てディクショナリにセットします
    Public Overloads Sub InitializeViewModelsOfProject(ByVal pk As String,
                                                       ByVal m As Model,
                                                       ByVal vm As ViewModel,
                                                       ByVal adm As AppDirectoryModel,
                                                       ByVal pim As ProjectInfoModel)
        Dim cvm As ConnectionViewModel
        Dim t_cvm As TabItemModel
        Dim dbtvm As DBTestViewModel
        Dim t_dbtvm As TabItemModel
        Dim dbevm As DBExplorerViewModel
        Dim hvm As HistoryViewModel

        '-- you henkou --------------------------------'
        Select Case pk
            Case AppDirectoryModel.DB_TEST
                cvm = New ConnectionViewModel()
                dbtvm = New DBTestViewModel()
                dbevm = New DBExplorerViewModel()
                hvm = New HistoryViewModel()

                'Call _UpdateContentToDictionary(ViewModel.MAIN_VIEW, cvm.GetType.Name, cvm)
                'Call _UpdateContentToDictionary(ViewModel.MAIN_VIEW, dbtvm.GetType.Name, dbtvm)
                'Call _UpdateContentToDictionary(ViewModel.EXPLORER_VIEW, dbevm.GetType.Name, dbevm)
                'Call _UpdateContentToDictionary(ViewModel.HISTORY_VIEW, hvm.GetType.Name, hvm)
                t_cvm = New TabItemModel With {.Name = cvm.GetType.Name, .Content = cvm}
                t_dbtvm = New TabItemModel With {.Name = dbtvm.GetType.Name, .Content = dbtvm}
                Call _UpdateTabsToCollection(ViewModel.MAIN_VIEW, t_cvm)
                Call _UpdateTabsToCollection(ViewModel.MAIN_VIEW, t_dbtvm)

                Call cvm.MyInitializing(m, vm, adm, pim)

                'Call ChangeContent(ViewModel.MAIN_VIEW, dbtvm.GetType.Name, dbtvm)
                'Call ChangeContent(ViewModel.MAIN_VIEW, cvm.GetType.Name, cvm)
                'Call ChangeContent(ViewModel.EXPLORER_VIEW, dbevm.GetType.Name, dbevm)
                'Call ChangeContent(ViewModel.HISTORY_VIEW, hvm.GetType.Name, hvm)
                Call ChangeTabs(ViewModel.MAIN_VIEW, t_cvm)
                'Call ChangeTabs(ViewModel.EXPLORER_VIEW, dbevm.GetType.Name, dbevm)
                'Call ChangeTabs(ViewModel.HISTORY_VIEW, hvm.GetType.Name, hvm)
            Case Else
        End Select
        '----------------------------------------------'
    End Sub
End Class
