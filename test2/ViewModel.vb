Public Class ViewModel
    Inherits BaseModel(Of ViewModel)

    Public Const MAIN_VIEW As String = "MainView"
    Public Const EXPLORER_VIEW As String = "ExplorerView"
    Public Const HISTORY_VIEW As String = "HistoryView"


    'Private _ProjectKind As String
    'Public Property ProjectKind As String
    '    Get
    '        Return Me._ProjectKind
    '    End Get
    '    Set(value As String)
    '        Me._ProjectKind = value
    '        RaisePropertyChanged("ProjectKind")
    '    End Set
    'End Property

    Private _MainViewContext As Object
    Public Property MainViewContext As Object
        Get
            Return Me._MainViewContext
        End Get
        Set(value As Object)
            Me._MainViewContext = value
            RaisePropertyChanged("MainViewContext")
        End Set
    End Property

    Private _MainViewHeight As Integer
    Public Property MainViewHeight As Integer
        Get
            Return Me._MainViewHeight
        End Get
        Set(value As Integer)
            Me._MainViewHeight = value
            RaisePropertyChanged("MainViewHeight")
        End Set
    End Property

    Private _MainViewWidth As Integer
    Public Property MainViewWidth As Integer
        Get
            Return Me._MainViewWidth
        End Get
        Set(value As Integer)
            Me._MainViewWidth = value
            RaisePropertyChanged("MainViewWidth")
        End Set
    End Property

    Private _ExplorerViewContext As Object
    Public Property ExplorerViewContext As Object
        Get
            Return Me._ExplorerViewContext
        End Get
        Set(value As Object)
            Me._ExplorerViewContext = value
            RaisePropertyChanged("ExplorerViewContext")
        End Set
    End Property

    Private _ExplorerViewHeight As Integer
    Public Property ExplorerViewHeight As Integer
        Get
            Return Me._ExplorerViewHeight
        End Get
        Set(value As Integer)
            Me._ExplorerViewHeight = value
            RaisePropertyChanged("ExplorerViewHeight")
        End Set
    End Property

    Private _ExplorerViewWidth As Integer
    Public Property ExplorerViewWidth As Integer
        Get
            Return Me._ExplorerViewWidth
        End Get
        Set(value As Integer)
            Me._ExplorerViewWidth = value
            RaisePropertyChanged("ExplorerViewWidth")
        End Set
    End Property

    Private _HistoryViewHeight As Integer
    Public Property HistoryViewHeight As Integer
        Get
            Return Me._HistoryViewHeight
        End Get
        Set(value As Integer)
            Me._HistoryViewHeight = value
            RaisePropertyChanged("HistoryViewHeight")
        End Set
    End Property

    Private _HistoryViewWidth As Integer
    Public Property HistoryViewWidth As Integer
        Get
            Return Me._HistoryViewWidth
        End Get
        Set(value As Integer)
            Me._HistoryViewWidth = value
            RaisePropertyChanged("HistoryViewWidth")
        End Set
    End Property

    Private _HistoryViewContext As Object
    Public Property HistoryViewContext As Object
        Get
            Return Me._HistoryViewContext
        End Get
        Set(value As Object)
            Me._HistoryViewContext = value
            RaisePropertyChanged("HistoryViewContext")
        End Set
    End Property

    Private _ContextDictionary As Dictionary(Of String, Dictionary(Of String, Object))
    Public Property ContextDictionary As Dictionary(Of String, Dictionary(Of String, Object))
        Get
            Return Me._ContextDictionary
        End Get
        Set(value As Dictionary(Of String, Dictionary(Of String, Object)))
            Me._ContextDictionary = value
        End Set
    End Property

    ' コンテキストディクショナリが存在しない場合、新規作成し
    ' ビューディクショナリが存在しない場合、追加します
    Private Sub _AddViewToDictionary(ByVal view As String)
        If Me.ContextDictionary Is Nothing Then
            Me.ContextDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
        End If
        If Not Me.ContextDictionary.ContainsKey(view) Then
            Me.ContextDictionary.Add(view, New Dictionary(Of String, Object))
        End If
    End Sub

    ' ビューのDataContextに実際にセットします
    Private Sub _SetContextObject(ByVal view As String, ByVal nm As String)
        Dim o As Object
        o = Me.ContextDictionary(view)(nm)

        Select Case view
            Case MAIN_VIEW
                Me.MainViewContext = o
            Case EXPLORER_VIEW
                Me.ExplorerViewContext = o
            Case HISTORY_VIEW
                Me.HistoryViewContext = o
            Case Else
                ' Nothing To Do
        End Select
    End Sub

    ' コンテキストをディクショナリにセットします
    Private Sub _AddContextToDictionary(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._AddViewToDictionary(viewName)
        If Not Me.ContextDictionary(viewName).ContainsKey(modelName) Then
            Me.ContextDictionary(viewName).Add(modelName, context)
        End If
    End Sub

    ' コンテキストをディクショナリにセット＆更新します
    Private Sub _UpdateContextToDictionary(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._AddContextToDictionary(viewName, modelName, context)
        Me.ContextDictionary(viewName)(modelName) = context
    End Sub

    ' コンテキストをディクショナリにセット＆ビューの切り替えを行います
    Public Sub ChangeContext(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._AddContextToDictionary(viewName, modelName, context)
        Call Me._SetContextObject(viewName, modelName)
    End Sub

    ' コンテキストをディクショナリにセット＆更新＆ビューの切り替えを行います
    Public Sub SetContext(ByVal viewName As String, ByVal modelName As String, ByRef context As Object)
        Call Me._UpdateContextToDictionary(viewName, modelName, context)
        Call Me._SetContextObject(viewName, modelName)
    End Sub

    Public Delegate Sub InitializeProxy(ByVal pk As String, ByVal m As Model, ByVal vm As ViewModel, ByVal adm As AppDirectoryModel, ByVal pim As ProjectInfoModel)


    ' 初回時に必要なビューモデルコンテキストを全てディクショナリにセットします
    Public Overloads Sub InitializeContext(ByVal pk As String,
                                           ByVal m As Model,
                                           ByVal vm As ViewModel,
                                           ByVal adm As AppDirectoryModel,
                                           ByVal pim As ProjectInfoModel)
        Dim cvm As ConnectionViewModel
        Dim dbtvm As DBTestViewModel
        Dim dbevm As DBExplorerViewModel
        Dim hvm As HistoryViewModel

        '-- you henkou --------------------------------'
        Select Case pk
            Case AppDirectoryModel.DB_TEST
                cvm = New ConnectionViewModel()
                dbtvm = New DBTestViewModel()
                dbevm = New DBExplorerViewModel()
                hvm = New HistoryViewModel()

                Call _UpdateContextToDictionary(ViewModel.MAIN_VIEW, cvm.GetType.Name, cvm)
                Call _UpdateContextToDictionary(ViewModel.MAIN_VIEW, dbtvm.GetType.Name, dbtvm)
                Call _UpdateContextToDictionary(ViewModel.EXPLORER_VIEW, dbevm.GetType.Name, dbevm)
                Call _UpdateContextToDictionary(ViewModel.HISTORY_VIEW, hvm.GetType.Name, hvm)

                Call cvm.MyInitializing(m, vm, adm, pim)

                'Call ChangeContext(ViewModel.MAIN_VIEW, dbtvm.GetType.Name, dbtvm)
                Call ChangeContext(ViewModel.MAIN_VIEW, cvm.GetType.Name, cvm)
                Call ChangeContext(ViewModel.EXPLORER_VIEW, dbevm.GetType.Name, dbevm)
                Call ChangeContext(ViewModel.HISTORY_VIEW, hvm.GetType.Name, hvm)
            Case Else
        End Select
        '----------------------------------------------'
    End Sub
End Class
