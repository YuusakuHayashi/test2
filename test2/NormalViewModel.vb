Public Class NormalViewModel : Inherits BaseViewModel
    Private _Content As Object
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
            RaisePropertyChanged("Content")
        End Set
    End Property
End Class
