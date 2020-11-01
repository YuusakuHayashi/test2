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
            RaisePropertyChanged("SelectedItem")
            Call _ChangeTabCloseButtonVisibilities(value)
            Call _ChangeTabColors(value)
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
                Me._Tabs = New ObservableCollection(Of TabItemModel)
            End If
            Return Me._Tabs
        End Get
        Set(value As ObservableCollection(Of TabItemModel))
            Me._Tabs = value
        End Set
    End Property

    Private Sub _ChangeTabCloseButtonVisibilities(ByVal [tab] As TabItemModel)
        For Each t In Me.Tabs
            If t.Equals([tab]) Then
                t.TabCloseButtonVisibility = Visibility.Visible
            Else
                t.TabCloseButtonVisibility = Visibility.Collapsed
            End If
        Next
    End Sub

    Private Sub _ChangeTabColors(ByVal [tab] As TabItemModel)
        For Each t In Me.Tabs
            If t.Equals([tab]) Then
                t.Color = Colors.Goldenrod
            Else
                t.Color = Colors.White
            End If
        Next
    End Sub
End Class
