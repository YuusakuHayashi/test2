Imports System.Collections.ObjectModel

Public Class ViewExplorerViewModel
    'Inherits BaseViewModel2

    '''Public Overrides ReadOnly Property FrameType As String
    '''    Get
    '''        Return ViewModel.EXPLORER_FRAME
    '''    End Get
    '''End Property

    ''Private _Views As ObservableCollection(Of ViewItemModel)
    ''Public Property Views As ObservableCollection(Of ViewItemModel)
    ''    Get
    ''        If Me._Views Is Nothing Then
    ''            Me._Views = New ObservableCollection(Of ViewItemModel)
    ''        End If
    ''        Return Me._Views
    ''    End Get
    ''    Set(value As ObservableCollection(Of ViewItemModel))
    ''        Me._Views = value
    ''        RaisePropertyChanged("Views")
    ''    End Set
    ''End Property

    ''Private Sub _OpenViewHandler()
    ''    AddHandler _
    ''        DelegateEventListener.Instance.OpenViewRequested,
    ''        AddressOf Me._OpenViewRequestedReview
    ''End Sub

    ''Private Sub _OpenViewRequestedReview(ByVal v As ViewItemModel, ByVal e As System.EventArgs)
    ''    Call _OpenViewRequestAccept(v)
    ''End Sub

    ''Private Sub _OpenViewRequestAccept(ByVal v As ViewItemModel)
    ''    'Dim [tab] As TabItemModel
    ''    'Dim obj = ViewModel.GetViewOfName(v.Name)
    ''    'Call obj.Initialize(Model, ViewModel, AppInfo, ProjectInfo)
    ''    '[tab] = New TabItemModel With {
    ''    '    .Name = v.Name,
    ''    '    .[Alias] = v.[Alias],
    ''    '    .Content = obj
    ''    '}
    ''    'Call ViewModel.ShowTabs([tab])
    ''End Sub


    ''Private Sub _ShowViews()
    ''    For Each v In ViewModel.Views
    ''        ' 自身の表示はしない
    ''        If Not v.Name = Me.GetType.Name Then
    ''            Me.Views.Add(v)
    ''        End If
    ''    Next
    ''End Sub


    'Public Overrides Sub Initialize(ByRef app As AppDirectoryModel,
    '                                ByRef vm As ViewModel)
    '    InitializeHandler = [Delegate].Combine(
    '        New Action(AddressOf _ShowViews)
    '    )
    '    [AddHandler] = [Delegate].Combine(
    '        New Action(AddressOf _OpenViewHandler)
    '    )
    '    Call BaseInitialize(app, vm)
    'End Sub
End Class
