Public Class ConnectionModel 
    Inherits BaseModel(Of ConnectionModel)

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

    'Private _TestTableName As String
    'Public Property TestTableName As String
    '    Get
    '        Return Me._TestTableName
    '    End Get
    '    Set(value As String)
    '        Me._TestTableName = value
    '    End Set
    'End Property

    Private _ConnectionString As String
    Public Property ConnectionString As String
        Get
            Return Me._ConnectionString
        End Get
        Set(value As String)
            Me._ConnectionString = value
        End Set
    End Property

    Private _ConnectionResult As Boolean
    Public Property ConnectionResult As Boolean
        Get
            Return Me._ConnectionResult
        End Get
        Set(value As Boolean)
            Me._ConnectionResult = value
        End Set
    End Property
End Class
