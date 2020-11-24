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

    Private _VisibleIconFileName As String _
        = AppDirectoryModel.AppImageDirectory & "\visible.png"
    Private _InvisibleIconFileName As String _
        = AppDirectoryModel.AppImageDirectory & "\invisible.png"

    Private _WindowHeight As Double
    Public Property WindowHeight As Double
        Get
            If Me._WindowHeight = 0.0 Then
                Me._WindowHeight = 500.0
            End If
            Return Me._WindowHeight
        End Get
        Set(value As Double)
            Me._WindowHeight = value
            RaisePropertyChanged("WindowHeight")
        End Set
    End Property

    Private _WindowWidth As Double
    Public Property WindowWidth As Double
        Get
            If Me._WindowWidth = 0.0 Then
                Me._WindowWidth = 1000.0
            End If
            Return Me._WindowWidth
        End Get
        Set(value As Double)
            Me._WindowWidth = value
            RaisePropertyChanged("WindowWidth")
        End Set
    End Property

    Private _Content As FrameViewModel
    Public Property Content As FrameViewModel
        Get
            Return Me._Content
        End Get
        Set(value As FrameViewModel)
            Me._Content = value
            RaisePropertyChanged("Content")
        End Set
    End Property

    Private _SaveContent As FrameViewModel
    <JsonIgnore>
    Public Property SaveContent As FrameViewModel
        Get
            Return Me._SaveContent
        End Get
        Set(value As FrameViewModel)
            Me._SaveContent = value
        End Set
    End Property

    'Private _Content As FlexibleViewModel
    'Public Property Content As FlexibleViewModel
    '    Get
    '        Return Me._Content
    '    End Get
    '    Set(value As FlexibleViewModel)
    '        Me._Content = value
    '        RaisePropertyChanged("Content")
    '    End Set
    'End Property

    'Private _SaveContent As FlexibleViewModel
    '<JsonIgnore>
    'Public Property SaveContent As FlexibleViewModel
    '    Get
    '        Return Me._SaveContent
    '    End Get
    '    Set(value As FlexibleViewModel)
    '        Me._SaveContent = value
    '    End Set
    'End Property

    'Public Sub VisualizeView(ByVal fvm As FlexibleViewModel)
    '    Me.Views = Nothing
    '    Me.Views = _CreateViews(fvm, Me.Views)
    '    Me.Content = fvm
    'End Sub
    Private _VisibleIconBase As BitmapImage
    <JsonIgnore>
    Public Property VisibleIconBase As BitmapImage
        Get
            If _VisibleIconBase Is Nothing Then
                Dim bi = New BitmapImage
                bi.BeginInit()
                bi.UriSource = New Uri(
                    Me._VisibleIconFileName,
                    UriKind.Absolute
                )
                bi.EndInit()
                Me._VisibleIconBase = bi
            End If
            Return Me._VisibleIconBase
        End Get
        Set(value As BitmapImage)
            Me._VisibleIconBase = value
        End Set
    End Property

    Private _InvisibleIconBase As BitmapImage
    <JsonIgnore>
    Public Property InvisibleIconBase As BitmapImage
        Get
            If _InvisibleIconBase Is Nothing Then
                Dim bi = New BitmapImage
                bi.BeginInit()
                bi.UriSource = New Uri(
                    Me._InvisibleIconFileName,
                    UriKind.Absolute
                )
                bi.EndInit()
                Me._InvisibleIconBase = bi
            End If
            Return Me._InvisibleIconBase
        End Get
        Set(value As BitmapImage)
            Me._InvisibleIconBase = value
        End Set
    End Property

    Public Sub ReloadViews()
        Me.Views = Nothing
        Call _CreateViews(Me.Content)
    End Sub

    Public Sub ResetView()
        Dim fvm = Me.Content.DeepCopy()
        Me.Content = fvm
    End Sub

    Private Sub _CreateViews(ByVal fvm As FrameViewModel)
        Call _AddView(fvm.MenuViewContent)
        Call _AddView(fvm.MainViewContent)
        Call _AddView(fvm.LeftExplorerViewContent)
        Call _AddView(fvm.RightExplorerViewContent)
        Call _AddView(fvm.HistoryViewContent)
    End Sub

    Private Sub _AddView(ByRef vim As ViewItemModel)
        Dim idx = -1
        If vim IsNot Nothing Then
            vim.VisibleIcon = IIf(vim.IsVisible, Me.VisibleIconBase, Me.InvisibleIconBase)
            For Each child In vim.Children
                child.VisibleIcon = IIf(child.IsVisible, Me.VisibleIconBase, Me.InvisibleIconBase)
            Next
            If Not Me.Views.Contains(vim) Then
                Me.Views.Add(vim)
            Else
                idx = Me.Views.IndexOf(vim)
                Me.Views(idx) = vim
            End If
        End If
    End Sub

    Public Sub VisualizeView(ByVal fvm As FrameViewModel)
        Me.Views = Nothing
        Call _CreateViews(fvm)
        Me.Content = fvm
    End Sub

    '---------------------------------------------------------------------------------------------'
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
            'Call DelegateEventListener.Instance.RaiseViewsChanged()
        End Set
    End Property
    '---------------------------------------------------------------------------------------------'
End Class
