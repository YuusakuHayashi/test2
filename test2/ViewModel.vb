Imports System.Collections.ObjectModel

Public Class ViewModel
    Inherits BaseModel(Of ViewModel)

    Public Const MAIN_VIEW As String = "MainView"
    Public Const EXPLORER_VIEW As String = "ExplorerView"
    Public Const HISTORY_VIEW As String = "HistoryView"
    Public Const MENU_VIEW As String = "MenuView"

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

    Private _MenuViewContent As Object
    Public Property MenuViewContent As Object
        Get
            Return Me._MenuViewContent
        End Get
        Set(value As Object)
            Me._MenuViewContent = value
            RaisePropertyChanged("MenuViewContent")
        End Set
    End Property


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
            Case MENU_VIEW
                Me.MenuViewContent = o
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

    ' タブコレクションディクショナリが存在しない場合、新規作成し
    ' タブコレクションが存在しない場合、追加します
    Private Sub _RegisterTabViewToDictionary(ByVal view As String)
        Dim tvm As TabViewModel
        If Me.TabsDictionary Is Nothing Then
            Me.TabsDictionary = New Dictionary(Of String, TabViewModel)
        End If
        If Not Me.TabsDictionary.ContainsKey(view) Then
            tvm = New TabViewModel
            tvm.Initialize()
            Me.TabsDictionary.Add(view, tvm)
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
    'Private Sub _AddTabsToCollection(ByVal view As String, ByVal [tab] As TabItemModel)
    '    Call Me._RegisterTabViewToDictionary(view)
    '    If Not Me.TabsDictionary(view).Tabs.Contains([tab]) Then
    '        Call [tab].Initialize()
    '        Me.TabsDictionary(view).Tabs.Add([tab])
    '    End If
    'End Sub

    ' タブをコレクションにセット＆更新します
    Private Sub _UpdateTabsToCollection(ByVal view As String, ByVal [tab] As TabItemModel)
        Dim idx = -1
        Call Me._RegisterTabViewToDictionary(view)
        For Each t In Me.TabsDictionary(view).Tabs
            If [tab].Name = t.Name Then
                idx = Me.TabsDictionary(view).Tabs.IndexOf(t)
            End If
        Next

        Call [tab].Initialize()
        If idx = -1 Then
            Me.TabsDictionary(view).Tabs.Add([tab])
        Else
            Me.TabsDictionary(view).Tabs(idx) = [tab]
        End If
        Me.TabsDictionary(view).SelectedIndex _
            = Me.TabsDictionary(view).Tabs.IndexOf([tab])
        'Call Me._AddTabsToCollection(view, [tab])
        'Me.TabsDictionary(view).Tabs(Me.TabsDictionary(view).Tabs.IndexOf([tab])) = [tab]
    End Sub

    ' タブをコレクションにセット＆ビューの切り替えを行います
    'Public Sub ChangeTabs(ByVal view As String, ByVal [tab] As TabItemModel)
    '    Call Me._AddTabsToCollection(view, [tab])
    '    Call Me._SetTabsObject(view)
    'End Sub

    ' タブをコレクションにセット＆更新＆ビューの切り替えを行います
    'Public Sub SetTabs(ByVal view As String, ByVal [tab] As TabItemModel)
    '    Call Me._UpdateTabsToCollection(view, [tab])
    '    Call Me._SetTabsObject(view)
    'End Sub
    Public Sub SetTabs(ByVal [tab] As TabItemModel)
        Dim view = [tab].Content.ViewType
        Call Me._UpdateTabsToCollection(view, [tab])
        Call Me._SetTabsObject(view)
    End Sub
    ' --------------------------------------------------------------------------------------------'

    Public Sub NoInitialize(ByVal pk As String,
                            ByVal m As Model,
                            ByVal vm As ViewModel,
                            ByVal adm As AppDirectoryModel,
                            ByVal pim As ProjectInfoModel)
        ' Nothing To Do
    End Sub

    ' 初回時に必要なビューモデルコンテントを全てディクショナリにセットします
    Public Overloads Sub Setup(ByVal m As Model,
                               ByVal vm As ViewModel,
                               ByVal adm As AppDirectoryModel,
                               ByVal pim As ProjectInfoModel)
        'Dim cvm As ConnectionViewModel
        'Dim dbtvm As DBTestViewModel
        'Dim dbevm As DBExplorerViewModel
        'Dim vevm As ViewExplorerViewModel
        'Dim hvm As HistoryViewModel
        'Dim mvm As MenuViewModel
        Dim cvm As Object
        Dim dbtvm As Object
        Dim dbevm As Object
        Dim vevm As Object
        Dim hvm As Object
        Dim mvm As Object
        Dim t_cvm As TabItemModel
        Dim t_dbtvm As TabItemModel
        Dim t_dbevm As TabItemModel
        Dim t_vevm As TabItemModel
        Dim t_hvm As TabItemModel

        '-- you henkou --------------------------------'
        Select Case pim.Kind
            Case AppDirectoryModel.DB_TEST
                cvm = New ConnectionViewModel
                dbtvm = New DBTestViewModel
                dbevm = New DBExplorerViewModel
                vevm = New ViewExplorerViewModel
                hvm = New HistoryViewModel
                mvm = New MenuViewModel

                Call cvm.Initialize(m, vm, adm, pim)
                Call hvm.Initialize(m, vm, adm, pim)
                Call mvm.Initialize(m, vm, adm, pim)
                Call vevm.Initialize(m, vm, adm, pim)

                t_cvm = New TabItemModel With {
                    .Name = cvm.GetType.Name,
                    .Content = cvm
                }
                t_dbtvm = New TabItemModel With {
                    .Name = dbtvm.GetType.Name,
                    .Content = dbtvm
                }
                t_dbevm = New TabItemModel With {
                    .Name = dbevm.GetType.Name,
                    .Content = dbevm
                }
                t_hvm = New TabItemModel With {
                    .Name = hvm.GetType.Name,
                    .Content = hvm
                }
                t_vevm = New TabItemModel With {
                    .Name = vevm.GetType.Name,
                    .Content = vevm
                }

                Call vevm.RegisterViews(t_cvm, t_dbtvm, t_dbevm, t_hvm, t_vevm)

                Call SetTabs(t_cvm)
                Call SetTabs(t_dbtvm)
                Call SetTabs(t_dbevm)
                Call SetTabs(t_vevm)
                Call SetTabs(t_hvm)
                Call ChangeContent(ViewModel.MENU_VIEW, mvm.GetType.Name, mvm)
            Case Else
        End Select
        '----------------------------------------------'
    End Sub

    Public Function GetViewOfName(ByVal [name] As String) As Object
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
        GetViewOfName = obj
    End Function
End Class
