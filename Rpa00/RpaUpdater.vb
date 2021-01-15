Public Class RpaUpdater
    Private _ReleaseDate As Date
    Public Property ReleaseDate As Date
        Get
            Return Me._ReleaseDate
        End Get
        Set(value As Date)
            Me._ReleaseDate = value
        End Set
    End Property
End Class
