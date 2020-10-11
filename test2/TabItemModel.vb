Public Class TabItemModel
    Private Property _Name As String
    Public Property Name As String
        Get
            Return Me._Name
        End Get
        Set(value As String)
            Me._Name = value
        End Set
    End Property

    Private Property _Content As Object
    Public Property Content As Object
        Get
            Return Me._Content
       End Get
        Set(value As Object)
            Me._Content = value
        End Set
    End Property
End Class
