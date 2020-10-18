Imports System.Collections.ObjectModel

Public Class ViewExplorerViewModel
    Inherits BaseViewModel2

    Public Overrides ReadOnly Property ViewType As String
        Get
            Return ViewModel.EXPLORER_VIEW
        End Get
    End Property

    Private _Views As ObservableCollection(Of ViewItemModel)
    Public Property Views As ObservableCollection(Of ViewItemModel)
        Get
            Return Me._Views
        End Get
        Set(value As ObservableCollection(Of ViewItemModel))
            Me._Views = value
            RaisePropertyChanged("Views")
        End Set
    End Property

    Public Sub RegisterViews(ParamArray ByVal tabs() As TabItemModel)
        Dim b = False
        Dim vim As ViewItemModel
        If Me.Views Is Nothing Then
            Me.Views = New ObservableCollection(Of ViewItemModel)
        End If
        For Each t In tabs
            b = False
            For Each v In Views
                If v.Name = t.Name Then
                    b = True
                End If
            Next
            If Not b Then
                vim = New ViewItemModel With {
                    .Name = t.Name,
                    .[Alias] = t.[Alias]
                }
                vim.Initialize()
                Me.Views.Add(vim)
            End If
        Next
    End Sub

    Private Sub _OpenViewHandler()
        AddHandler _
            DelegateEventListener.Instance.OpenViewRequested,
            AddressOf Me._OpenViewRequestedReview
    End Sub

    Private Sub _OpenViewRequestedReview(ByVal v As ViewItemModel, ByVal e As System.EventArgs)
        Call _OpenViewRequestAccept(v)
    End Sub

    Private Sub _OpenViewRequestAccept(ByVal v As ViewItemModel)
        Dim [tab] As TabItemModel
        Dim obj = ViewModel.GetViewOfName(v.Name)
        Call obj.Initialize(Model, ViewModel, AppInfo, ProjectInfo)
        [tab] = New TabItemModel With {
            .Name = v.Name,
            .[Alias] = v.[Alias],
            .Content = obj
        }
        Call ViewModel.SetTabs([tab])
    End Sub

    Public Sub Initialize(ByRef m As Model,
                          ByRef vm As ViewModel,
                          ByRef adm As AppDirectoryModel,
                          ByRef pim As ProjectInfoModel)
        [AddHandler] = [Delegate].Combine(
            New Action(AddressOf _OpenViewHandler)
        )
        Call BaseInitialize(m, vm, adm, pim)
    End Sub
End Class
