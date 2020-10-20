Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ViewModel
    Inherits JsonHandler(Of ViewModel)
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler _
        Implements INotifyPropertyChanged.PropertyChanged

    Public Overridable Sub RaisePropertyChanged(ByVal PropertyName As String)
        RaiseEvent PropertyChanged(
            Me, New PropertyChangedEventArgs(PropertyName)
        )
    End Sub

    Public Const MAIN_VIEW As String = "MainView"
    Public Const EXPLORER_VIEW As String = "ExplorerView"
    Public Const HISTORY_VIEW As String = "HistoryView"
    Public Const MENU_VIEW As String = "MenuView"

    Private _MainViewContent As Object
    <JsonIgnore>
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
    <JsonIgnore>
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
    <JsonIgnore>
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
    <JsonIgnore>
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
    <JsonIgnore>
    Public Property ContentDictionary As Dictionary(Of String, Dictionary(Of String, Object))
        Get
            If Me._ContentDictionary Is Nothing Then
                Return New Dictionary(Of String, Dictionary(Of String, Object))
            Else
                Return Me._ContentDictionary
            End If
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
    'Private _OpenTabsDictionary As Dictionary(Of String, List(Of String))
    'Public Property OpenTabsDictionary As Dictionary(Of String, List(Of String))
    '    Get
    '        Return Me._OpenTabsDictionary
    '    End Get
    '    Set(value As Dictionary(Of String, List(Of String)))
    '        Me._OpenTabsDictionary = value
    '    End Set
    'End Property

    Private _OpenTabs As List(Of String)
    Public Property OpenTabs As List(Of String)
        Get
            If Me._OpenTabs Is Nothing Then
                Return New List(Of String)
            Else
                Return Me._OpenTabs
            End If
        End Get
        Set(value As List(Of String))
            Me._OpenTabs = value
        End Set
    End Property

    ' タブコレクションディクショナリが存在しない場合、新規作成し
    ' タブコレクションが存在しない場合、追加します
    Public Sub AddOpenTabs(ByVal tab As TabItemModel)
        If Not Me.OpenTabs.Contains(tab.Name) Then
            Me.OpenTabs.Add(tab.Name)
        End If
    End Sub

    Public Sub RemoveOpenTabs(ByVal tab As TabItemModel)
        If Me.OpenTabs.Contains(tab.Name) Then
            Me.OpenTabs.Remove(tab.Name)
        End If
    End Sub

    Private _TabsDictionary As Dictionary(Of String, TabViewModel)
    <JsonIgnore>
    Public Property TabsDictionary As Dictionary(Of String, TabViewModel)
        Get
            If Me._TabsDictionary Is Nothing Then
                Return New Dictionary(Of String, TabViewModel)
            Else
                Return Me._TabsDictionary
            End If
        End Get
        Set(value As Dictionary(Of String, TabViewModel))
            Me._TabsDictionary = value
        End Set
    End Property

    ' タブコレクションディクショナリが存在しない場合、新規作成し
    ' タブコレクションが存在しない場合、追加します
    Private Sub _RegisterTabViewToDictionary(ByVal view As String)
        Dim tvm As TabViewModel
        'If Me.TabsDictionary Is Nothing Then
        '    Me.TabsDictionary = New Dictionary(Of String, TabViewModel)
        'End If
        If Not Me.TabsDictionary.ContainsKey(view) Then
            tvm = New TabViewModel
            ' 閉じるハンドラーのセット
            tvm.Initialize()
            Me.TabsDictionary.Add(view, tvm)
        End If
        'If Me.TabsDictionary(view).Tabs Is Nothing Then
        '    Me.TabsDictionary(view).Tabs = New ObservableCollection(Of TabItemModel)
        'End If
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
        For Each t In Me.TabsDictionary(view).Tabs
            If [tab].Name = t.Name Then
                idx = Me.TabsDictionary(view).Tabs.IndexOf(t)
            End If
        Next

        ' 閉じるコマンドのセット
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
    'Public Sub ShowTabs(ByVal view As String, ByVal [tab] As TabItemModel)
    '    Call Me._UpdateTabsToCollection(view, [tab])
    '    Call Me._SetTabsObject(view)
    'End Sub

    ' タブを表示する唯一の公開メソッドとすること
    Public Sub ShowTabs(ByVal [tab] As TabItemModel)
        Dim view = [tab].Content.ViewType
        Call Me._RegisterTabViewToDictionary(view)
        Call Me._UpdateTabsToCollection(view, [tab])
        Call Me.AddOpenTabs([tab])
        Call Me._SetTabsObject(view)
    End Sub
    ' --------------------------------------------------------------------------------------------'


    ' 初回時に必要なビューモデルコンテントを全てディクショナリにセットします
    Public Sub InitializeTabs(ParamArray views() As Object)
        Dim t As TabItemModel
        For Each v In views
            t = New TabItemModel With {
                .Name = v.GetType.Name,
                .Content = v
            }
            Call ShowTabs(t)
        Next
    End Sub

    Public Overloads Sub Setup(ByVal m As Model,
                               ByVal vm As ViewModel,
                               ByVal adm As AppDirectoryModel,
                               ByVal pim As ProjectInfoModel)
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
                If Me.OpenTabs Is Nothing Then
                    ' ビューの登録
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

                    Call InitializeTabs(cvm, dbtvm, dbevm, vevm, hvm)
                Else
                    Call 
                End If

                't_cvm = New TabItemModel With {
                '    .Name = cvm.GetType.Name,
                '    .Content = cvm
                '}
                't_dbtvm = New TabItemModel With {
                '    .Name = dbtvm.GetType.Name,
                '    .Content = dbtvm
                '}
                't_dbevm = New TabItemModel With {
                '    .Name = dbevm.GetType.Name,
                '    .Content = dbevm
                '}
                't_hvm = New TabItemModel With {
                '    .Name = hvm.GetType.Name,
                '    .Content = hvm
                '}
                't_vevm = New TabItemModel With {
                '    .Name = vevm.GetType.Name,
                '    .Content = vevm
                '}

                Call vevm.RegisterViews(t_cvm, t_dbtvm, t_dbevm, t_hvm, t_vevm)

                Call ShowTabs(t_cvm)
                Call ShowTabs(t_dbtvm)
                Call ShowTabs(t_dbevm)
                Call ShowTabs(t_vevm)
                Call ShowTabs(t_hvm)
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
