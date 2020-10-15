Imports System.Collections.ObjectModel

Public Class TabViewModel : Inherits BaseViewModel
    Private _Tabs As ObservableCollection(Of TabItemModel)
    Public Property Tabs As ObservableCollection(Of TabItemModel)
        Get
            Return Me._Tabs
        End Get
        Set(value As ObservableCollection(Of TabItemModel))
            Me._Tabs = value
        End Set
    End Property

    Private Sub _TabCloseRequestedReview(ByVal t As TabItemModel, ByVal e As System.EventArgs)
        Call _TabCloseRequestAccept(t)
    End Sub

    Private Sub _TabCloseRequestAccept(ByVal t As TabItemModel)
        Tabs.Remove(t)
    End Sub

    Public Sub Initialize()
        AddHandler _
            DelegateEventListener.Instance.TabCloseRequested,
            AddressOf Me._TabCloseRequestedReview
    End Sub
End Class
