Imports System.Collections.ObjectModel
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ProjectExplorerViewModel
    Inherits BaseViewModel2

    Private __BitmapImageDictionary As Dictionary(Of String, BitmapImage)
    Private Property _BitmapImageDictionary As Dictionary(Of String, BitmapImage)
        Get
            If Me.__BitmapImageDictionary Is Nothing Then
                Me.__BitmapImageDictionary = New Dictionary(Of String, BitmapImage)
            End If
            Return Me.__BitmapImageDictionary
        End Get
        Set(value As Dictionary(Of String, BitmapImage))
            Me.__BitmapImageDictionary = value
        End Set
    End Property

    Private _Projects As ObservableCollection(Of ProjectInfoModel)
    <JsonIgnore>
    Public Property Projects As ObservableCollection(Of ProjectInfoModel)
        Get
            If Me._Projects Is Nothing Then
                Me._Projects = New ObservableCollection(Of ProjectInfoModel)
            End If
            Return Me._Projects
        End Get
        Set(value As ObservableCollection(Of ProjectInfoModel))
            Me._Projects = value
            RaisePropertyChanged("Projects")
        End Set
    End Property

    Private Sub _AssignIconOfProject()
        For Each p In Me.Projects
            If p.Icon Is Nothing Then
            End If
        Next
    End Sub

    Private Sub _ViewInitializing()
        Me.Projects = AppInfo.CurrentProjects
    End Sub

    '---------------------------------------------------------------------------------------------'
    Private Sub _ProjectsChangedAddHandler()
        AddHandler _
            DelegateEventListener.Instance.ProjectsChanged,
            AddressOf Me._ProjectsChangedReview
    End Sub

    Private Sub _ProjectsChangedReview(ByVal sender As Object, ByVal e As System.EventArgs)
        Call _ProjectsChangedAccept()
    End Sub

    Private Sub _ProjectsChangedAccept()
        Me.Projects = AppInfo.CurrentProjects
    End Sub
    '---------------------------------------------------------------------------------------------'

    Public Overrides Sub Initialize(ByRef app As AppDirectoryModel, ByRef vm As ViewModel)
        InitializeHandler = AddressOf _ViewInitializing
        [AddHandler] = [Delegate].Combine(
            New Action(AddressOf _ProjectsChangedAddHandler)
        )
        Call BaseInitialize(app, vm)
    End Sub

End Class
