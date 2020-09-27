Public Class ViewModel
    Inherits BaseViewModel

    Public Const MAIN_VIEW As String = "MainView"
    Public Const EXPLORER_VIEW As String = "ExplorerView"
    Public Const HISTORY_VIEW As String = "HistoryView"

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

    Private Sub _DictionaryCheck(ByVal view As String, ByVal nm As String)
        If Me.ContextDictionary Is Nothing Then
            Me.ContextDictionary = New Dictionary(Of String, Dictionary(Of String, Object))
        End If
        If Not Me.ContextDictionary.ContainsKey(view) Then
            Me.ContextDictionary.Add(view, New Dictionary(Of String, Object))
        End If
    End Sub

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

    ' Protectedだとアクセスできない。原因追及後回し
    Public Sub SetContext(ByVal view As String, ByVal nm As String, ByRef m As Object)
        Call Me._DictionaryCheck(view, nm)
        If Not Me.ContextDictionary(view).ContainsKey(nm) Then
            Me.ContextDictionary(view).Add(nm, m)
        Else
            Me.ContextDictionary(view)(nm) = m
        End If
        Call Me._SetContextObject(view, nm)
    End Sub

    Public Sub ChangeContext(ByVal view As String, ByVal nm As String, ByRef m As Object)
        Call Me._DictionaryCheck(view, nm)
        If Not Me.ContextDictionary(view).ContainsKey(nm) Then
            Me.ContextDictionary(view).Add(nm, m)
        End If
        Call Me._SetContextObject(view, nm)
    End Sub

    Private Sub _SetDefaultHeight()
        Me.ExplorerViewWidth = 100
    End Sub

    Sub New()
        Call _SetDefaultHeight()
    End Sub
End Class
