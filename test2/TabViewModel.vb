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
End Class
