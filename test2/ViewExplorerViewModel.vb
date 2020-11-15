Imports System.Collections.ObjectModel
Imports test2
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ViewExplorerViewModel
    Inherits BaseViewModel2

    Private _VisibleIconFileName As String _
        = AppDirectoryModel.AppImageDirectory & "\visible.png"
    Private _InvisibleIconFileName As String _
        = AppDirectoryModel.AppImageDirectory & "\invisible.png"

    Private Property _VisibleIconBase As BitmapImage
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

    Private Property _InvisibleIconBase As BitmapImage
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

    Private _FontSize As Double
    <JsonIgnore>
    Public Property FontSize As Double
        Get
            If Me._FontSize = 0.0 Then
                Me._FontSize = 11.0
            End If
            Return Me._FontSize
        End Get
        Set(value As Double)
            Me._FontSize = value
            RaisePropertyChanged("FontSize")
        End Set
    End Property

    Private _Views As ObservableCollection(Of ViewItemModel)
    <JsonIgnore>
    Public Property Views As ObservableCollection(Of ViewItemModel)
        Get
            If Me._Views Is Nothing Then
                Me._Views = New ObservableCollection(Of ViewItemModel)
            End If
            Return Me._Views
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Call _AssignVisibleIcon(value)
            Me._Views = value
            RaisePropertyChanged("Views")
        End Set
    End Property

    ' Visibleに関するアイコンの割り当て
    Private Sub _AssignVisibleIcon(ByRef vims As ObservableCollection(Of ViewItemModel))
        For Each vim In vims
            vim.VisibleIcon = IIf(vim.IsVisible, Me.VisibleIconBase, Me.InvisibleIconBase)
        Next
    End Sub

    Private Sub _ViewInitializing()
        Me.Views = ViewModel.Views
    End Sub

    '---------------------------------------------------------------------------------------------'
    Private Sub _ViewsChangedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.ViewsChanged,
            AddressOf Me._ViewsChangedReview
    End Sub

    Private Sub _ViewsChangedReview(ByVal sender As Object, ByVal e As System.EventArgs)
        Call _ViewsChangedAccept()
    End Sub

    Private Sub _ViewsChangedAccept()
        Me.Views = ViewModel.Views
    End Sub
    '---------------------------------------------------------------------------------------------'

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitializing
        [AddHandler] = [Delegate].Combine(
            New Action(AddressOf _ViewsChangedAddHandler)
        )
        Call BaseInitialize(app, vm)
    End Sub

End Class
