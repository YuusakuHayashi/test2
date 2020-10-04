Public Class InitModel
    Inherits BaseModel(Of InitModel)
    Implements ProjectInterface

    Private _InitFlag As Boolean
    Public Property InitFlag As Boolean
        Get
            Return Me._InitFlag
        End Get
        Set(value As Boolean)
            Me._InitFlag = value
        End Set
    End Property

    Private _ServerName As String
    Public Property ServerName As String
        Get
            Return Me._ServerName
        End Get
        Set(value As String)
            Me._ServerName = value
        End Set
    End Property

    Private _DataBaseName As String
    Public Property DataBaseName As String
        Get
            Return Me._DataBaseName
        End Get
        Set(value As String)
            Me._DataBaseName = value
        End Set
    End Property


    Private _ConnectionString As String
    Public Property ConnectionString As String
        Get
            Return Me._ConnectionString
        End Get
        Set(value As String)
            Me._ConnectionString = value
        End Set
    End Property

    Public OtherProperty As String
End Class
