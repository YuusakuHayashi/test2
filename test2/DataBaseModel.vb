Public Class DataBaseModel : Inherits BaseModel
    Private _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
            'RaisePropertyChanged("Name")
        End Set
    End Property

    Private _DataTables As List(Of DataTableModel)
    Public Property DataTables As List(Of DataTableModel)
        Get
            Return Me._DataTables
        End Get
        Set(value As List(Of DataTableModel))
            Me._DataTables = value
            'RaisePropertyChanged("Name")
        End Set
    End Property

    Private _IsChecked As Boolean
    Public Property IsChecked As Boolean
        Get
            Return Me._IsChecked
        End Get
        Set(value As Boolean)
            Me._IsChecked = value
            RaisePropertyChanged("_IsChecked")
        End Set
    End Property
End Class
