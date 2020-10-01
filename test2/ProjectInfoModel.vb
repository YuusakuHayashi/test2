Public Class ProjectInfoModel : Inherits AppDirectoryModel
    Private _UserDirectoryName As String
    Public Property UserDirectoryName As String
        Get
            Return _UserDirectoryName
        End Get
        Set(value As String)
            _UserDirectoryName = value
        End Set
    End Property

    Private _ProjectName As String
    Public Property ProjectName As String
        Get
            Return _ProjectName
        End Get
        Set(value As String)
            _ProjectName = value
        End Set
    End Property

    Private _ProjectDirectoryName As String
    Public Property ProjectDirectoryName As String
        Get
            Return _ProjectDirectoryName
        End Get
        Set(value As String)
            _ProjectDirectoryName = value
        End Set
    End Property

    Protected ReadOnly Property _ProjectIniFileName As String
        Get
            Return ProjectDirectoryName & "\ProjectInit"
        End Get
    End Property

    Protected ReadOnly Property _ProjectModelFileName As String
        Get
            Return ProjectDirectoryName & "\Model.json"
        End Get
    End Property
End Class
