Public Class RpaUpdater
    Private _ReleaseDate As String
    Public Property ReleaseDate As String
        Get
            Return Me._ReleaseDate
        End Get
        Set(value As String)
            Me._ReleaseDate = value
        End Set
    End Property

    Private _ReleaseTitle As String
    Public Property ReleaseTitle As String
        Get
            Return Me._ReleaseTitle
        End Get
        Set(value As String)
            Me._ReleaseTitle = value
        End Set
    End Property

    Private _ReleaseNote As String
    Public Property ReleaseNote As String
        Get
            Return Me._ReleaseNote
        End Get
        Set(value As String)
            Me._ReleaseNote = value
        End Set
    End Property

    Private _PackageDirectory As String
    Public Property PackageDirectory As String
        Get
            Return Me._PackageDirectory
        End Get
        Set(value As String)
            Me._PackageDirectory = value
        End Set
    End Property

    Private _ReleaseTargets As List(Of String)
    Public Property ReleaseTargets As List(Of String)
        Get
            If Me._ReleaseTargets Is Nothing Then
                Me._ReleaseTargets = New List(Of String)
            End If
            Return Me._ReleaseTargets
        End Get
        Set(value As List(Of String))
            Me._ReleaseTargets = value
        End Set
    End Property
End Class
