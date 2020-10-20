Imports System.Collections.ObjectModel

Public Class TabViewModel : Inherits BaseViewModel
    Private Property _SelectedIndex As Integer
    Public Property SelectedIndex As Integer
        Get
            Return Me._SelectedIndex
        End Get
        Set(value As Integer)
            Me._SelectedIndex = value
            RaisePropertyChanged("SelectedIndex")
            'If value = -1 Then
            '    Me.SelectedIndex = 0
            '    RaisePropertyChanged("SelectedIndex")
            'End If
        End Set
    End Property

    Private Property _SelectedItem As TabItemModel
    Public Property SelectedItem As TabItemModel
        Get
            Return Me._SelectedItem
        End Get
        Set(value As TabItemModel)
            Me._SelectedItem = value
            'If value Is Nothing Then
            '    If Tabs(0) IsNot Nothing Then
            '        Me.SelectedItem = Tabs(0)
            '    End If
            'End If
            'RaisePropertyChanged("SelectedItem")
        End Set
    End Property

    Private _Tabs As ObservableCollection(Of TabItemModel)
    Public Property Tabs As ObservableCollection(Of TabItemModel)
        Get
            If Me._Tabs Is Nothing Then
                Return New ObservableCollection(Of TabItemModel)
            Else
                Return Me._Tabs
            End If
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
