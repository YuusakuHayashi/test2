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

    Public Const MAIN_FRAME As String = "MainView"
    Public Const EXPLORER_FRAME As String = "ExplorerView"
    Public Const HISTORY_FRAME As String = "HistoryView"
    Public Const MENU_FRAME As String = "MenuView"

    Public Const NORMAL_VIEW As String = "Normal"
    Public Const TAB_VIEW As String = "Tab"

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
                Me._ContentDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
            End If
            Return Me._ContentDictionary
        End Get
        Set(value As Dictionary(Of String, Dictionary(Of String, Object)))
            Me._ContentDictionary = value
        End Set
    End Property

    ' コンテントディクショナリが存在しない場合、新規作成し
    ' ビューディクショナリが存在しない場合、追加します
    'Private Sub _RegisterViewToDictionary(ByVal v As ViewItemModel)
    '    Dim frame = v.FrameType
    '    If Me.ContentDictionary Is Nothing Then
    '        Me.ContentDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
    '    End If
    '    If Not Me.ContentDictionary.ContainsKey(frame) Then
    '        Me.ContentDictionary.Add(frame, New Dictionary(Of String, Object))
    '    End If
    'End Sub

    '' ビューのDataContentに実際にセットします
    'Private Sub _SetItemToContent(ByVal v As ViewItemModel)
    '    Dim o As Object
    '    o = Me.ContentDictionary(v.FrameType)(v.Name)

    '    Select Case v.FrameType
    '        Case MAIN_FRAME
    '            Me.MainViewContent = o
    '        Case EXPLORER_FRAME
    '            Me.ExplorerViewContent = o
    '        Case HISTORY_FRAME
    '            Me.HistoryViewContent = o
    '        Case MENU_FRAME
    '            Me.MenuViewContent = o
    '        Case Else
    '            ' Nothing To Do
    '    End Select
    'End Sub

    'Private Sub _AddItem(ByVal v As ViewItemModel)
    '    Dim obj = v.Content
    '    Dim frame = obj.FrameType
    '    Dim [name] = obj.GetType.Name
    '    If Me.ContentDictionary(frame).ContainsKey([name]) Then
    '        Me.ContentDictionary(frame)([name]) = obj
    '    Else
    '        Me.ContentDictionary(frame).Add([name], obj)
    '    End If
    'End Sub
    ' --------------------------------------------------------------------------------------------'

    ' --------------------------------------------------------------------------------------------'
    Private _Views As ObservableCollection(Of ViewItemModel)
    Public Property Views As ObservableCollection(Of ViewItemModel)
        Get
            If Me._Views Is Nothing Then
                Me._Views = New ObservableCollection(Of ViewItemModel)
            End If
            Return Me._Views
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Me._Views = value
        End Set
    End Property
    ' --------------------------------------------------------------------------------------------'


    ' タブコレクション関連
    ' タブディクショナリ
    ' --------------------------------------------------------------------------------------------'
    Private _TabsDictionary As Dictionary(Of String, TabViewModel)
    <JsonIgnore>
    Public Property TabsDictionary As Dictionary(Of String, TabViewModel)
        Get
            If Me._TabsDictionary Is Nothing Then
                Me._TabsDictionary = New Dictionary(Of String, TabViewModel)
            End If
            Return Me._TabsDictionary
        End Get
        Set(value As Dictionary(Of String, TabViewModel))
            Me._TabsDictionary = value
        End Set
    End Property

    '' タブコレクションディクショナリが存在しない場合、新規作成し
    '' タブコレクションが存在しない場合、追加します
    'Private Sub _RegisterTabViewToDictionary(ByVal v As ViewItemModel)
    '    Dim frame = v.FrameType
    '    Dim tvm As TabViewModel
    '    If Not Me.TabsDictionary.ContainsKey(frame) Then
    '        tvm = New TabViewModel
    '        Me.TabsDictionary.Add(frame, tvm)
    '    End If
    'End Sub

    '' ビューのDataContentに実際にセットします
    'Private Sub _SetTabItemToContent(ByVal v As ViewItemModel)
    '    Dim o As TabViewModel
    '    o = Me.TabsDictionary(v.FrameType)

    '    Select Case v.FrameType
    '        Case MAIN_FRAME
    '            Me.MainViewContent = o
    '        Case EXPLORER_FRAME
    '            Me.ExplorerViewContent = o
    '        Case HISTORY_FRAME
    '            Me.HistoryViewContent = o
    '        Case Else
    '            ' Nothing To Do
    '    End Select
    'End Sub

    'Public Overloads Sub RemoveTabItem(ByVal [tab] As TabItemModel)
    '    For Each p In Me.TabsDictionary
    '        If p.Value.Tabs.Contains([tab]) Then
    '            p.Value.Tabs.Remove([tab])
    '        End If
    '    Next

    '    ' ビューからの開閉状態記録
    '    For Each v In Me.Views
    '        If v.Name = [tab].Name Then
    '            v.OpenState = False
    '        End If
    '    Next
    'End Sub

    '' コレクションにタブがなければ追加、あれば更新
    'Private Overloads Sub _AddTabItem(ByVal [tab] As TabItemModel)
    '    Dim idx = -1
    '    Dim frame = [tab].FrameType
    '    For Each t In Me.TabsDictionary(frame).Tabs
    '        If [tab].Name = t.Name Then
    '            idx = Me.TabsDictionary(frame).Tabs.IndexOf(t)
    '        End If
    '    Next
    '    If idx = -1 Then
    '        ' 閉じるコマンドのセット
    '        Me.TabsDictionary(frame).Tabs.Add([tab])
    '        idx = Me.TabsDictionary(frame).Tabs.IndexOf([tab])
    '    Else
    '        Me.TabsDictionary(frame).Tabs(idx).Content = [tab].Content
    '        'idx = idx
    '    End If
    '    Me.TabsDictionary(frame).SelectedIndex = idx
    'End Sub

    '' コレクションにタブがなければ追加、あれば更新
    'Private Overloads Sub _AddTabItem(ByVal v As ViewItemModel)
    '    Dim t = New TabItemModel With {
    '        .Name = v.Name,
    '        .FrameType = v.FrameType,
    '        .Content = v.Content
    '    }
    '    Call _AddTabItem(t)
    'End Sub

    'Public Sub AddView(ByVal v As ViewItemModel)
    '    ' ビュータイプから、セットする辞書を判別する
    '    Select Case v.ViewType
    '        Case ViewModel.NORMAL_VIEW
    '            Call _RegisterViewToDictionary(v)
    '            Call _AddItem(v)
    '            Call _SetItemToContent(v)
    '        Case ViewModel.TAB_VIEW
    '            Call _RegisterTabViewToDictionary(v)
    '            Call _AddTabItem(v)
    '            Call _SetTabItemToContent(v)
    '        Case Else
    '            MsgBox("ViewModel.AddView")
    '            Exit Sub
    '    End Select
    'End Sub

    'Public Function AddViewItem(ByVal obj As Object,
    '                            ByVal frame As String,
    '                            ByVal view As String) As ViewItemModel
    '    Dim v As New ViewItemModel With {
    '        .Name = obj.GetType.Name,
    '        .FrameType = frame,
    '        .ViewType = view,
    '        .OpenState = True,
    '        .Content = obj
    '    }

    '    ' 自身の表示はしない
    '    If Not obj.Equals(Me) Then
    '        Me.Views.Add(v)
    '    End If
    '    AddViewItem = v
    'End Function
    '---------------------------------------------------------------------------------------------'

    'Private Sub _LoadSetup(ByVal m As Model,
    '                       ByVal vm As ViewModel,
    '                       ByVal adm As AppDirectoryModel,
    '                       ByVal pim As ProjectInfoModel)
    '    Dim obj As Object
    '    Dim t As TabItemModel

    '    For Each v In Views
    '        If v.OpenState Then
    '            obj = GetViewOfName(v.Name)
    '            obj.Initialize(m, vm, adm, pim)
    '            v.Content = obj
    '            Call AddView(v)
    '        End If
    '    Next
    'End Sub

    'Public Overloads Sub Setup(ByVal m As Model,
    '                           ByVal vm As ViewModel,
    '                           ByVal adm As AppDirectoryModel,
    '                           ByVal pim As ProjectInfoModel)
    '    Dim cvm, dbtvm, dbevm, vevm, hvm, mvm
    '    Dim v0, v1, v2, v3, v4, v5, v6, v7, v8, v9

    '    Select Case pim.Kind
    '        '-- you henkou --------------------------------'
    '        Case AppDirectoryModel.DB_TEST
    '            If Me.Views.Count < 1 Then
    '                cvm = New ConnectionViewModel
    '                dbtvm = New DBTestViewModel
    '                dbevm = New DBExplorerViewModel
    '                vevm = New ViewExplorerViewModel
    '                hvm = New HistoryViewModel
    '                mvm = New MenuViewModel

    '                Call cvm.Initialize(m, vm, adm, pim)
    '                Call hvm.Initialize(m, vm, adm, pim)
    '                Call mvm.Initialize(m, vm, adm, pim)

    '                ' ビューへの追加
    '                v1 = AddViewItem(cvm, MAIN_FRAME, TAB_VIEW)
    '                v2 = AddViewItem(hvm, HISTORY_FRAME, TAB_VIEW)
    '                v3 = AddViewItem(mvm, MENU_FRAME, NORMAL_VIEW)

    '                Call AddView(v1)
    '                Call AddView(v2)
    '                Call AddView(v3)
    '            Else
    '                Call _LoadSetup(m, vm, adm, pim)
    '            End If
    '            ' 特殊なビューのセット(ViewModel)
    '            v0 = AddViewItem(Me, EXPLORER_FRAME, TAB_VIEW)
    '            Call AddView(v0)
    '            'Case Else
    '        Case Else
    '            '----------------------------------------------'
    '    End Select
    'End Sub

    'Public Function GetViewOfName(ByVal [name] As String) As Object
    '    Dim obj As Object
    '    Select Case [name]
    '        Case (New ConnectionViewModel).GetType.Name
    '            obj = New ConnectionViewModel
    '        Case (New DBTestViewModel).GetType.Name
    '            obj = New DBTestViewModel
    '        Case (New DBExplorerViewModel).GetType.Name
    '            obj = New DBExplorerViewModel
    '        Case (New ViewExplorerViewModel).GetType.Name
    '            obj = New ViewExplorerViewModel
    '        Case (New HistoryViewModel).GetType.Name
    '            obj = New HistoryViewModel
    '        Case (New MenuViewModel).GetType.Name
    '            obj = New MenuViewModel
    '        Case Else
    '            obj = Nothing
    '    End Select
    '    GetViewOfName = obj
    'End Function

    ''---------------------------------------------------------------------------------------------'
    '' このオブジェクト自体を上書きしてしまうと、ＭＶＶＭが適用されなくなる
    '' (Ｖｉｅｗへの反映がされない)ため、ロードが必要なメンバをここで別個にセットする
    'Public Overloads Sub Initialize(ByVal vm As ViewModel)
    '    '--- you henkou ----------------------------------'
    '    Me.Views = vm.Views
    '    '-------------------------------------------------'
    '    Call Me.Initialize()
    'End Sub

    'Public Overloads Sub Initialize()
    '    Call _TabCloseAddHandler()
    '    Call _OpenViewAddHandler()
    'End Sub
    ''---------------------------------------------------------------------------------------------'


    ''--- タブを閉じる関連 ------------------------------------------------------------------------'
    'Private Sub _TabCloseRequestedReview(ByVal t As TabItemModel, ByVal e As System.EventArgs)
    '    Call _TabCloseRequestAccept(t)
    'End Sub

    'Private Sub _TabCloseRequestAccept(ByVal [tab] As TabItemModel)
    '    Call RemoveTabItem([tab])
    'End Sub

    'Private Sub _TabCloseAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.TabCloseRequested,
    '        AddressOf Me._TabCloseRequestedReview
    'End Sub
    ''---------------------------------------------------------------------------------------------'


    ''--- タブを開く関連 --------------------------------------------------------------------------'
    'Private Sub _OpenViewAddHandler()
    '    AddHandler _
    '        DelegateEventListener.Instance.OpenViewRequested,
    '        AddressOf Me._OpenViewRequestedReview
    'End Sub

    'Private Sub _OpenViewRequestedReview(ByVal v As ViewItemModel, ByVal e As System.EventArgs)
    '    Call _OpenViewRequestAccept(v)
    'End Sub

    'Private Sub _OpenViewRequestAccept(ByVal v As ViewItemModel)
    '    Dim idx = Me.Views.IndexOf(v)
    '    Me.Views(idx).OpenState = True
    '    Call Me.AddView(Me.Views(idx))
    'End Sub
    ''---------------------------------------------------------------------------------------------'
End Class
