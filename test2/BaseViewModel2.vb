Imports System.ComponentModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

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
    Private _Model As Model
    Public Property Model As Model
        Get
            Return _Model
        End Get
        Set(value As Model)
            _Model = value
        End Set
    End Property

    Private _ViewModel As ViewModel
    Public Property ViewModel As ViewModel
        Get
            Return _ViewModel
        End Get
        Set(value As ViewModel)
            _ViewModel = value
        End Set
    End Property

    Private _AppInfo As AppDirectoryModel
    Public Property AppInfo As AppDirectoryModel
        Get
            Return _AppInfo
        End Get
        Set(value As AppDirectoryModel)
            _AppInfo = value
        End Set
    End Property

    Private _ProjectInfo As ProjectInfoModel
    Public Property ProjectInfo As ProjectInfoModel
        Get
            Return _ProjectInfo
        End Get
        Set(value As ProjectInfoModel)
            _ProjectInfo = value
        End Set
    End Property

    Protected Property InitializeHandler As Action
    Protected Property CheckCommandEnabledHandler As Action
    Protected Property [AddHandler] As Action

    Public MustOverride ReadOnly Property FrameType As String Implements BaseViewModelInterface.FrameType

    Protected Overridable Sub ViewInitializing()
        'Nothing To Do
    End Sub

    Protected Overridable Sub ContextModelCheck()
        'Nothing To Do
    End Sub


    ' Model型のDataObjectはJsonからデシリアライズされたものはJObject型になっている
    Protected Sub BaseInitialize(ByVal m As Model,
                                 ByVal vm As ViewModel,
                                 ByVal adm As AppDirectoryModel,
                                 ByVal pim As ProjectInfoModel)
        Dim obj As Object
        Dim ih = Me.InitializeHandler
        Dim cceh = Me.CheckCommandEnabledHandler
        Dim ah = Me.[AddHandler]

        Me.Model = m
        Me.ViewModel = vm
        Me.AppInfo = adm
        Me.ProjectInfo = pim

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

    Public MustOverride Sub Initialize(ByRef m As Model, ByRef vm As ViewModel, ByRef adm As AppDirectoryModel, ByRef pim As ProjectInfoModel) Implements BaseViewModelInterface.Initialize

    ' コンテントディクショナリが存在しない場合、新規作成し
    ' ビューディクショナリが存在しない場合、追加します
    Private Sub _RegisterViewToDictionary(ByVal v As ViewItemModel)
        Dim frame = v.FrameType
        If ViewModel.ContentDictionary Is Nothing Then
            ViewModel.ContentDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
        End If
        If Not ViewModel.ContentDictionary.ContainsKey(frame) Then
            ViewModel.ContentDictionary.Add(frame, New Dictionary(Of String, Object))
        End If
    End Sub

    ' ビューのDataContentに実際にセットします
    Private Sub _SetItemToContent(ByVal v As ViewItemModel)
        Dim o As Object
        o = ViewModel.ContentDictionary(v.FrameType)(v.Name)

        Select Case v.FrameType
            Case ViewModel.MAIN_FRAME
                ViewModel.MainViewContent = o
            Case ViewModel.EXPLORER_FRAME
                ViewModel.ExplorerViewContent = o
            Case ViewModel.HISTORY_FRAME
                ViewModel.HistoryViewContent = o
            Case ViewModel.MENU_FRAME
                ViewModel.MenuViewContent = o
            Case Else
                ' Nothing To Do
        End Select
    End Sub

    Private Sub _AddItem(ByVal v As ViewItemModel)
        Dim obj = v.Content
        Dim frame = obj.FrameType
        Dim [name] = obj.GetType.Name
        If ViewModel.ContentDictionary(frame).ContainsKey([name]) Then
            ViewModel.ContentDictionary(frame)([name]) = obj
        Else
            ViewModel.ContentDictionary(frame).Add([name], obj)
        End If
    End Sub
    ' --------------------------------------------------------------------------------------------'

    ' タブコレクションディクショナリが存在しない場合、新規作成し
    ' タブコレクションが存在しない場合、追加します
    Private Sub _RegisterTabViewToDictionary(ByVal v As ViewItemModel)
        Dim frame = v.FrameType
        Dim tvm As TabViewModel
        If Not ViewModel.TabsDictionary.ContainsKey(frame) Then
            tvm = New TabViewModel
            ViewModel.TabsDictionary.Add(frame, tvm)
        End If
    End Sub

    ' ビューのDataContentに実際にセットします
    Private Sub _SetTabItemToContent(ByVal v As ViewItemModel)
        Dim o As TabViewModel
        o = ViewModel.TabsDictionary(v.FrameType)

        Select Case v.FrameType
            Case ViewModel.MAIN_FRAME
                ViewModel.MainViewContent = o
            Case ViewModel.EXPLORER_FRAME
                ViewModel.ExplorerViewContent = o
            Case ViewModel.HISTORY_FRAME
                ViewModel.HistoryViewContent = o
            Case Else
                ' Nothing To Do
        End Select
    End Sub

    Public Overloads Sub RemoveTabItem(ByVal [tab] As TabItemModel)
        For Each p In ViewModel.TabsDictionary
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
    Private Overloads Sub _AddTabItem(ByVal [tab] As TabItemModel)
        Dim idx = -1
        Dim frame = [tab].FrameType
        For Each t In ViewModel.TabsDictionary(frame).Tabs
            If [tab].Name = t.Name Then
                idx = ViewModel.TabsDictionary(frame).Tabs.IndexOf(t)
            End If
        Next
        If idx = -1 Then
            ' 閉じるコマンドのセット
            ViewModel.TabsDictionary(frame).Tabs.Add([tab])
            idx = ViewModel.TabsDictionary(frame).Tabs.IndexOf([tab])
        Else
            ViewModel.TabsDictionary(frame).Tabs(idx).Content = [tab].Content
            'idx = idx
        End If
        ViewModel.TabsDictionary(frame).SelectedIndex = idx
    End Sub

    ' コレクションにタブがなければ追加、あれば更新
    Private Overloads Sub _AddTabItem(ByVal v As ViewItemModel)
        Dim t = New TabItemModel With {
            .Name = v.Name,
            .FrameType = v.FrameType,
            .Content = v.Content
        }
        Call _AddTabItem(t)
    End Sub

    Public Sub AddView(ByVal v As ViewItemModel)
        ' ビュータイプから、セットする辞書を判別する
        Select Case v.ViewType
            Case ViewModel.NORMAL_VIEW
                Call _RegisterViewToDictionary(v)
                Call _AddItem(v)
                Call _SetItemToContent(v)
            Case ViewModel.TAB_VIEW
                Call _RegisterTabViewToDictionary(v)
                Call _AddTabItem(v)
                Call _SetTabItemToContent(v)
            Case Else
                MsgBox("ViewModel.AddView")
                Exit Sub
        End Select
    End Sub

    Public Function AddViewItem(ByVal obj As Object,
                                ByVal frame As String,
                                ByVal view As String) As ViewItemModel
        Dim v As New ViewItemModel With {
            .Name = obj.GetType.Name,
            .FrameType = frame,
            .ViewType = view,
            .OpenState = True,
            .Content = obj
        }

        ' 自身の表示はしない
        If Not obj.Equals(ViewModel) Then
            ViewModel.Views.Add(v)
        End If
        AddViewItem = v
    End Function

    Private Sub _LoadSetup(ByVal m As Model,
                       ByVal vm As ViewModel,
                       ByVal adm As AppDirectoryModel,
                       ByVal pim As ProjectInfoModel)
        Dim obj As Object
        Dim t As TabItemModel

        For Each v In ViewModel.Views
            If v.OpenState Then
                obj = GetViewOfName(v.Name)
                obj.Initialize(m, vm, adm, pim)
                v.Content = obj
                Call AddView(v)
            End If
        Next
    End Sub

    Public Overloads Sub Setup(ByVal m As Model,
                               ByVal vm As ViewModel,
                               ByVal adm As AppDirectoryModel,
                               ByVal pim As ProjectInfoModel)
        Dim cvm, dbtvm, dbevm, vevm, hvm, mvm
        Dim v0, v1, v2, v3, v4, v5, v6, v7, v8, v9

        Select Case pim.Kind
            '-- you henkou --------------------------------'
            Case AppDirectoryModel.DB_TEST
                If ViewModel.Views.Count < 1 Then
                    cvm = New ConnectionViewModel
                    dbtvm = New DBTestViewModel
                    dbevm = New DBExplorerViewModel
                    vevm = New ViewExplorerViewModel
                    hvm = New HistoryViewModel
                    mvm = New MenuViewModel

                    Call cvm.Initialize(Model, ViewModel, AppInfo, ProjectInfo)
                    Call hvm.Initialize(Model, ViewModel, AppInfo, ProjectInfo)
                    Call mvm.Initialize(Model, ViewModel, AppInfo, ProjectInfo)

                    ' ビューへの追加
                    v1 = AddViewItem(cvm, ViewModel.MAIN_FRAME, ViewModel.TAB_VIEW)
                    v2 = AddViewItem(hvm, ViewModel.HISTORY_FRAME, ViewModel.TAB_VIEW)
                    v3 = AddViewItem(mvm, ViewModel.MENU_FRAME, ViewModel.NORMAL_VIEW)

                    Call AddView(v1)
                    Call AddView(v2)
                    Call AddView(v3)
                Else
                    Call _LoadSetup(m, vm, adm, pim)
                End If
                ' 特殊なビューのセット(ViewModel)
                v0 = AddViewItem(ViewModel, ViewModel.EXPLORER_FRAME, ViewModel.TAB_VIEW)
                Call AddView(v0)
                'Case Else
            Case Else
                '----------------------------------------------'
        End Select
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
