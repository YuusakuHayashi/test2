Public Class TabItemViewModel : Inherits ViewModelBase(Of TabItemViewModel)
    Private _Content As Object
    Public Property Content As Object
        Get
            Return Me._Content
        End Get
        Set(value As Object)
            Me._Content = value
        End Set
    End Property

    Private _Header As String
    Public Property Header As String
        Get
            Return Me._Header
        End Get
        Set(value As String)
            Me._Header = value
        End Set
    End Property
End Class
