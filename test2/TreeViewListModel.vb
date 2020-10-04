Public Class TreeViewListModel : Inherits BaseModel(Of TreeViewListModel)
    Private _ModelList As List(Of TreeViewModel)
    Public Property ModelList As List(Of TreeViewModel)
        Get
            Return _ModelList
        End Get
        Set(value As List(Of TreeViewModel))
            _ModelList = value
            RaisePropertyChanged("ModelList")
        End Set
    End Property

    Sub New(ByRef tvml As List(Of TreeViewModel))
        ModelList = tvml
    End Sub
End Class
